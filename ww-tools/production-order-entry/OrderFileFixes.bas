Attribute VB_Name = "OrderFileFixes"
Option Explicit

'===============================================================================
' MODULE: OrderFileFixes
' PURPOSE: Reliable ZIP packaging and explicit file-handling diagnostics.
'
' REQUIRED VBA REFERENCES (Tools > References):
'   Microsoft Scripting Runtime              -> Scripting.* objects
'   Windows Script Host Object Model         -> IWshRuntimeLibrary.* objects
'   Microsoft Excel Object Library           -> Excel.Workbook (normally present)
'
' DESIGN:
'   Objects use early binding so IntelliSense and compile-time member checking
'   are available. Explicit New assignments make object lifetime visible.
'   Order locks use PtrSafe Windows API declarations and LongPtr handles.
'   Requires VBA7 (Excel 2010 or later); supports both 32-bit and 64-bit Office.
'   The PowerShell script must be allowed by the existing company policy.
'
' TROUBLESHOOTING:
'   ZIP logs: %LOCALAPPDATA%\OrderEntryLogs
'   Breakpoints: before sh.Exec, inside Failed, and before Name sourcePath.
'   Locks are created with delete-on-close and disappear after handle release.
'   Existing old placeholders are adopted and removed by the next updated run.
'   VBA Reset/End does not close API handles; close Excel if debugging leaked one.
'===============================================================================

' Native Windows file handles are pointer-sized. DWORD flags remain 32-bit Long.
' StrPtr supplies the Unicode path to CreateFileW without ANSI conversion.
' Zero sharing denies competing read/write/delete handles until release.
Private Declare PtrSafe Function OrderLockCreateFileW Lib "kernel32" Alias "CreateFileW" ( _
    ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, _
    ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As LongPtr) As LongPtr

Private Declare PtrSafe Function OrderLockCloseHandle Lib "kernel32" Alias "CloseHandle" ( _
    ByVal hObject As LongPtr) As Long

Private Const ORDER_LOCK_GENERIC_READ As Long = &H80000000
Private Const ORDER_LOCK_GENERIC_WRITE As Long = &H40000000
Private Const ORDER_LOCK_DELETE_ACCESS As Long = &H10000
Private Const ORDER_LOCK_OPEN_ALWAYS As Long = 4
Private Const ORDER_LOCK_ATTRIBUTE_NORMAL As Long = &H80
Private Const ORDER_LOCK_DELETE_ON_CLOSE As Long = &H4000000
Private Const ORDER_LOCK_INVALID_HANDLE As Long = -1

'-------------------------------------------------------------------------------
' CreateZipFile
' Inputs: source folder and destination ZIP, both fully qualified path strings.
' Success: returns only after compression, verification, and publication finish.
' Failure: raises an error to the caller; it does not clear the order worksheet.
' Side effects: launches PowerShell; writes logs and temporary ZIP/lock files.
' The worker publishes the verified ZIP and preserves failed .pending candidates.
'-------------------------------------------------------------------------------
Public Sub CreateZipFile(ByVal folderToZipPath As String, ByVal zippedFileFullName As String)
    Const TIMEOUT_SECONDS As Long = 600
    ' File-system and process-launch objects are separate libraries.
    Dim fs As Scripting.FileSystemObject
    Dim sourceFolder As Scripting.Folder
    Dim sh As IWshRuntimeLibrary.WshShell
    Dim child As IWshRuntimeLibrary.WshExec
    Dim stdoutStream As Scripting.TextStream
    Dim stderrStream As Scripting.TextStream
    Dim scriptPath As String, sourcePath As String, destinationPath As String
    Dim logFolder As String, logPath As String, exePath As String, command As String
    Dim started As Date, stderrText As String, stdoutText As String
    ' StatusBar returns either Boolean False or a String, so Variant is required.
    Dim oldStatus As Variant
    Dim errorNumber As Long, errorText As String
    ' DoEvents allows UI callbacks during the wait. Guard against re-entry into
    ' this procedure before taking ownership of its saved state.
    Static busy As Boolean

    If busy Then Err.Raise vbObjectError + 2400, "CreateZipFile", "ZIP packaging is already running in this Excel instance."
    busy = True
    oldStatus = Application.StatusBar
    On Error GoTo Failed
    Set fs = New Scripting.FileSystemObject
    ' GetFolder validates the source and normalizes its path. The Folder object
    ' itself is not a document lock; the worker locks each document while reading.
    Set sourceFolder = fs.GetFolder(folderToZipPath)
    sourcePath = sourceFolder.Path
    destinationPath = fs.GetAbsolutePathName(zippedFileFullName)
    'scriptPath = fs.BuildPath(ThisWorkbook.Path, "OrderZip.ps1")
    scriptPath = "\\cop.hi.corp\fileserver\CFE\CPS\WWQ-FS01\Orders\Common Files\CPE_Schedule\OrderZip.ps1"
    If Not fs.FileExists(scriptPath) Then
        Err.Raise vbObjectError + 2401, "CreateZipFile", "Missing packaging script: " & scriptPath
    End If
    ' A unique local log avoids collisions between users and packaging attempts.
    ' Startup errors (such as blocked script execution) may occur before logging.
    logFolder = fs.BuildPath(Environ$("LOCALAPPDATA"), "OrderEntryLogs")
    If Not fs.FolderExists(logFolder) Then fs.CreateFolder logFolder
    logPath = fs.BuildPath(logFolder, "Zip-" & fs.GetTempName & ".log")
    exePath = Environ$("SystemRoot") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
    ' Exec launches PowerShell directly, without cmd.exe or shell redirection.
    ' Each path is quoted separately. Execution policy is not overridden.
    command = QuoteArgument(exePath) & " -NoLogo -NoProfile -NonInteractive -File " & _
              QuoteArgument(scriptPath) & " -Source " & QuoteArgument(sourcePath) & _
              " -Destination " & QuoteArgument(destinationPath) & " -Log " & QuoteArgument(logPath)
    Application.StatusBar = "Compressing and verifying order files..."
    Set sh = New IWshRuntimeLibrary.WshShell
    Set child = sh.Exec(command)
    started = Now
    ' Wait for process completion, not for a ZIP folder-item count. Worker output
    ' is intentionally small; per-file progress goes to the log, not the pipes.
    Do While child.Status = IWshRuntimeLibrary.WshRunning
        DoEvents
        If DateDiff("s", started, Now) >= TIMEOUT_SECONDS Then
            child.Terminate
            Err.Raise vbObjectError + 2402, "CreateZipFile", _
                "ZIP packaging exceeded 10 minutes. Check the log and destination before retrying." & vbCrLf & logPath
        End If
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    ' These are typed Scripting.TextStream objects returned by WshExec.
    ' ReadAll is performed after the process exits so it cannot wait on a running
    ' worker that is still expected to produce output.
    Set stdoutStream = child.StdOut
    Set stderrStream = child.StdErr
    stdoutText = stdoutStream.ReadAll
    stderrText = stderrStream.ReadAll
    If child.ExitCode <> 0 Then
        Err.Raise vbObjectError + 2403, "CreateZipFile", _
            "ZIP packaging failed." & vbCrLf & Left$(stderrText & stdoutText, 1800) & _
            vbCrLf & "Diagnostic log (if the script started): " & logPath
    End If
    Application.StatusBar = oldStatus
    busy = False
    Exit Sub
Failed:
    ' Save the originating error before cleanup can replace Err. Resume Next is
    ' restricted to cleanup; packaging/file-operation failures are not ignored.
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    If Not child Is Nothing Then
        If child.Status = IWshRuntimeLibrary.WshRunning Then child.Terminate
    End If
    Application.StatusBar = oldStatus
    busy = False
    On Error GoTo 0
    Err.Raise errorNumber, "CreateZipFile", errorText
End Sub

'-------------------------------------------------------------------------------
' QuoteArgument: quotes one Windows command-line argument and rejects embedded
' quotes/newlines. Paths supplied here are ordinary Windows file/folder paths.
' The source folder is normalized first and has no trailing backslash.
'-------------------------------------------------------------------------------
Private Function QuoteArgument(ByVal value As String) As String
    If InStr(value, Chr$(34)) > 0 Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
        Err.Raise vbObjectError + 2404, "QuoteArgument", "Invalid character in file argument."
    End If
    QuoteArgument = Chr$(34) & value & Chr$(34)
End Function

'-------------------------------------------------------------------------------
' AcquireOrderProcessingLock
' Returns a native LongPtr handle, not a VBA file number. Call the matching
' ReleaseOrderProcessingLock on every exit; do NOT use Close # on this handle.
' The file is removed by Windows when its owning handle closes. OPEN_ALWAYS
' adopts any unlocked placeholder left by an older package without truncation.
' Competing active handles cause an error; they are never forcibly removed.
'-------------------------------------------------------------------------------
Public Function AcquireOrderProcessingLock(ByVal basePath As String) As LongPtr
    AcquireOrderProcessingLock = AcquireDeleteOnCloseLockFile(basePath & ".processing.lock")
End Function

Private Function AcquireDeleteOnCloseLockFile(ByVal lockPath As String) As LongPtr
    Dim handle As LongPtr
    Dim windowsError As Long
    handle = OrderLockCreateFileW(StrPtr(lockPath), _
        ORDER_LOCK_GENERIC_READ Or ORDER_LOCK_GENERIC_WRITE Or ORDER_LOCK_DELETE_ACCESS, _
        0, 0, ORDER_LOCK_OPEN_ALWAYS, _
        ORDER_LOCK_ATTRIBUTE_NORMAL Or ORDER_LOCK_DELETE_ON_CLOSE, 0)
    If handle = ORDER_LOCK_INVALID_HANDLE Then
        ' Capture LastDllError before another API call can replace it.
        windowsError = Err.LastDllError
        Err.Raise vbObjectError + 2411, "AcquireDeleteOnCloseLockFile", _
            "Cannot acquire temporary order lock. Another process may be active, " & _
            "or the share may deny access/delete permission." & vbCrLf & _
            lockPath & vbCrLf & "Windows error: " & CStr(windowsError)
    End If
    AcquireDeleteOnCloseLockFile = handle
End Function

Public Sub ReleaseOrderProcessingLock(ByRef lockHandle As LongPtr)
    Dim windowsError As Long
    If lockHandle = 0 Then Exit Sub
    If OrderLockCloseHandle(lockHandle) = 0 Then
        windowsError = Err.LastDllError
        Err.Raise vbObjectError + 2412, "ReleaseOrderProcessingLock", _
            "Cannot close temporary order lock handle. Windows error: " & CStr(windowsError)
    End If
    ' Prevent cleanup from closing a later handle that reuses the same value.
    ' Delete-on-close removes the lock file without a separate Close-then-Kill race.
    lockHandle = 0
End Sub

'-------------------------------------------------------------------------------
' OpenOrderTemplate
' Returns: the exact workbook opened by this call; caller owns SaveAs/Close.
' Refuses to reuse an already-open workbook in this Excel instance, preventing
' changes to or closure of a template the user has open for another purpose.
' ReadOnly protects the shared source template; SaveAs creates the order copy.
'-------------------------------------------------------------------------------
Public Function OpenOrderTemplate(ByVal templatePath As String) As Excel.Workbook
    Dim existing As Excel.Workbook
    For Each existing In Application.Workbooks
        If StrComp(existing.FullName, templatePath, vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 2405, "OpenOrderTemplate", _
                "Close the template in this Excel instance before processing: " & templatePath
        End If
    Next existing
    Set OpenOrderTemplate = Workbooks.Open(filename:=templatePath, UpdateLinks:=0, _
        ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
End Function

'-------------------------------------------------------------------------------
' MoveSchedulingFile
' Called after the scheduling document is added to the displayed Outlook draft.
' Only the expected Processed subfolder is created; the base must already exist.
' Existing destination files are never replaced. A locked source is an error.
' This does not prove the draft was sent and is not an email transaction.
'-------------------------------------------------------------------------------
Public Sub MoveSchedulingFile(ByVal sourcePath As String, ByVal processedBase As String)
    Dim fs As Scripting.FileSystemObject
    Dim sourceFile As Scripting.File
    Dim filename As String, destinationFolder As String, destination As String
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set fs = New Scripting.FileSystemObject
    If Not fs.FileExists(sourcePath) Then
        Err.Raise vbObjectError + 2406, "MoveSchedulingFile", "Source file does not exist."
    End If
    Set sourceFile = fs.GetFile(sourcePath)
    filename = sourceFile.Name
    If Not fs.FolderExists(processedBase) Then
        Err.Raise vbObjectError + 2407, "MoveSchedulingFile", "Processed base folder is unavailable: " & processedBase
    End If
    If Left$(filename, 4) = "1100" Then
        destinationFolder = fs.BuildPath(processedBase, Left$(filename, 7) & "xxx")
    Else
        destinationFolder = fs.BuildPath(processedBase, "INTERNAL_Orders")
    End If
    If Not fs.FolderExists(destinationFolder) Then fs.CreateFolder destinationFolder
    destination = fs.BuildPath(destinationFolder, filename)
    If fs.FileExists(destination) Or fs.FolderExists(destination) Then
        Err.Raise vbObjectError + 2408, "MoveSchedulingFile", "Destination already exists; no file was overwritten."
    End If
    ' Keep the existing VBA move operation. Failure leaves the source in place
    ' in normal access/collision cases and is reported with both paths below.
    Name sourcePath As destination
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Err.Raise errorNumber, "MoveSchedulingFile", _
        "Scheduling file move failed." & vbCrLf & "Source: " & sourcePath & vbCrLf & _
        "Destination: " & destination & vbCrLf & errorText
End Sub

'-------------------------------------------------------------------------------
' RetryOrderZipOnly
' Operator recovery for packaging failure after documents were already created.
' Does not open the database, create transmittals, or clear the entry form.
' If document generation failed, complete those documents before using this.
' Uses the same source-folder layout and processing lock as NewOrder2NewDB.
'-------------------------------------------------------------------------------
Public Sub RetryOrderZipOnly()
    Const BASE As String = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE\Temporary_Order_Files\"
    ' InputBox(Type:=2) returns a String, or Boolean False for Cancel.
    ' Variant is deliberate here; converting before checking would lose Cancel.
    Dim inputValue As Variant
    Dim orderText As String, basePath As String
    Dim lockHandle As LongPtr
    Dim errorText As String
    On Error GoTo Failed
    inputValue = Application.InputBox("Order number to package again (no database inserts):", _
        "Retry order ZIP only", Type:=2)
    If VarType(inputValue) = vbBoolean Then Exit Sub
    orderText = Trim$(CStr(inputValue))
    If Len(orderText) = 0 Then Exit Sub
    If orderText Like "*[!0-9]*" Then
        Err.Raise vbObjectError + 2409, "RetryOrderZipOnly", "Enter digits only."
    End If
    basePath = BASE & orderText
    lockHandle = AcquireOrderProcessingLock(basePath)
    CreateZipFile basePath, basePath & ".zip"
    ReleaseOrderProcessingLock lockHandle
    ' The worker has exited and the order handle has closed. Both temporary lock
    ' files are removed by delete-on-close before this success notification.
    MsgBox "ZIP created and verified:" & vbCrLf & basePath & ".zip" & vbCrLf & vbCrLf & _
        "Temporary processing locks released and removed.", _
        vbInformation, "ZIP retry complete"
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If lockHandle <> 0 Then ReleaseOrderProcessingLock lockHandle
    On Error GoTo 0
    MsgBox errorText, vbExclamation, "ZIP retry failed"
End Sub

'-------------------------------------------------------------------------------
' CleanupReleasedOrderLocks
' Optional one-time migration for an order's old permanent placeholders.
' Acquires BOTH paths exclusively with delete-on-close; if a path is actively
' owned or inaccessible, reports failure without deleting that active file.
' Does not access the database, source documents, or ZIP archive itself.
'-------------------------------------------------------------------------------
Public Sub CleanupReleasedOrderLocks()
    Const BASE As String = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE\Temporary_Order_Files\"
    Dim inputValue As Variant
    Dim orderText As String, basePath As String, errorText As String
    Dim orderHandle As LongPtr, zipHandle As LongPtr
    On Error GoTo Failed
    inputValue = Application.InputBox("Order number whose old lock files should be removed:", _
        "Clean up released order locks", Type:=2)
    If VarType(inputValue) = vbBoolean Then Exit Sub
    orderText = Trim$(CStr(inputValue))
    If Len(orderText) = 0 Then Exit Sub
    If orderText Like "*[!0-9]*" Then
        Err.Raise vbObjectError + 2409, "CleanupReleasedOrderLocks", "Enter digits only."
    End If
    basePath = BASE & orderText
    orderHandle = AcquireOrderProcessingLock(basePath)
    zipHandle = AcquireDeleteOnCloseLockFile(basePath & ".zip.lock")
    ReleaseOrderProcessingLock zipHandle
    ReleaseOrderProcessingLock orderHandle
    MsgBox "Released lock files cleaned up for order " & orderText & ".", vbInformation
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If zipHandle <> 0 Then ReleaseOrderProcessingLock zipHandle
    If orderHandle <> 0 Then ReleaseOrderProcessingLock orderHandle
    On Error GoTo 0
    MsgBox errorText, vbExclamation, "Lock cleanup stopped"
End Sub

'-------------------------------------------------------------------------------
' TestOrderLockCleanup
' Local smoke check for the native VBA handle: present while held, exclusive
' against a second opener, and absent after release. No production/database work.
'-------------------------------------------------------------------------------
Public Sub TestOrderLockCleanup()
    Dim fs As Scripting.FileSystemObject
    Dim basePath As String, errorText As String
    Dim ownerHandle As LongPtr, competingHandle As LongPtr
    Dim competingError As Long
    On Error GoTo Failed
    Set fs = New Scripting.FileSystemObject
    basePath = fs.BuildPath(fs.GetSpecialFolder(Scripting.TemporaryFolder).Path, fs.GetTempName)
    ownerHandle = AcquireOrderProcessingLock(basePath)
    If Not fs.FileExists(basePath & ".processing.lock") Then
        Err.Raise vbObjectError + 2413, "TestOrderLockCleanup", "Lock file was not present while owned."
    End If
    ' A second handle, including one in this process, must be refused.
    On Error Resume Next
    competingHandle = AcquireOrderProcessingLock(basePath)
    competingError = Err.Number
    Err.Clear
    On Error GoTo Failed
    If competingError = 0 Then
        Err.Raise vbObjectError + 2414, "TestOrderLockCleanup", "Competing opener was not refused."
    End If
    ReleaseOrderProcessingLock ownerHandle
    If fs.FileExists(basePath & ".processing.lock") Then
        Err.Raise vbObjectError + 2415, "TestOrderLockCleanup", "Lock file remained after release."
    End If
    MsgBox "Local order-lock checks passed: exclusivity and automatic removal.", vbInformation
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If competingHandle <> 0 Then ReleaseOrderProcessingLock competingHandle
    If ownerHandle <> 0 Then ReleaseOrderProcessingLock ownerHandle
    On Error GoTo 0
    MsgBox errorText, vbExclamation, "Order-lock check failed"
End Sub
