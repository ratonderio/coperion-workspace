Attribute VB_Name = "Module1"
'Attribute VB_Name = "Module1"
'===============================================================================
' MODULE: ProdEng_Order_Entry (Module1)
'===============================================================================
' PURPOSE:
'   This module handles the Production Engineering order entry pipeline.
'   It is the primary interface between the ProdEng_Order_Entry.xlsm workbook
'   and the PAC1CPE database (Prod_Eng, cpe_scheduling, cpe_schedule tables).
'
' KEY FUNCTIONS:
'   - Insert_SAP_Data:       Imports SAP data exports (Open.XLS, Parts.XLS) into
'                            the workbook for manipulation.
'   - ProdCont_Query:        Queries the cpe_scheduling DB for order info and
'                            populates the New Order worksheet from SAP data sheets.
'   - Print_SAP_Order:       Automates SAP GUI to print/archive a sales order PDF.
'   - NewOrder2NewDB:        Inserts a new order into Prod_Eng DB and sets up the
'                            folder structure + transmittal templates for Windchill upload.
'   - SendNew_Email:         Sends a "New Order" notification email to assigned engineers.
'   - Send_Change_Order:     Sends a "Change Order" notification email to engineers.
'   - WhoHasOrder:           Looks up order assignment and engineering comments.
'   - CreateZipFile:         Zips a folder for Windchill upload.
'   - UpdateAllToCurrentYear: Bulk updates Active_Year on unreleased orders.
'
' KNOWN ISSUES / TECHNICAL DEBT:
'   [SECURITY] Database credentials are hardcoded in this module. These should be
'       moved to a secure credential store, environment variable, or Windows
'       Integrated Authentication (Trusted_Connection).
'   [SQL INJECTION] All SQL queries use string concatenation with cell values.
'       Parameterized queries via ADODB.Command should be used instead.
'   [GLOBALS] Excessive use of Public module-level variables creates tight coupling
'       and makes debugging difficult. Variables like I, J, K are reused across
'       multiple subroutines as loop counters, which is fragile.
'   [ERROR HANDLING] Error handling is inconsistent. Some subs have none; others
'       use On Error Resume Next as flow control rather than true error handling.
'   [SELECT/SELECTION] Several routines use .Select/.Selection anti-patterns
'       that slow execution and make the code harder to follow.
'   [MAGIC NUMBERS] Cell references are hardcoded throughout (e.g., Cells(5,2)
'       for order number). Named ranges or constants would improve readability.
'   [ACTION QUERIES] INSERT and UPDATE statements are executed via Recordset.Open
'       instead of Connection.Execute. This works accidentally but is incorrect usage.
'   [DATE VALIDATION] Dates are validated against the magic number 40000 (roughly
'       6/23/2009 in Excel serial date). This is a proxy for "is this cell not empty/zero"
'       but is fragile and unclear.
'
' REFACTORING HISTORY:
'   2026-03-05 - Added comprehensive code comments throughout module.
'              - Fixed NewOrder2NewDB: Replaced destructive delete-and-recreate
'                directory logic with safe check-and-create-if-missing approach.
'              - Removed J: drive (Path_J_Drive) references; Order Reviews now
'                point to Temporary_Order_Files. Removed Publish2Intranet sub.
'===============================================================================

Option Explicit  ' ADDED: Forces variable declaration; catches typos at compile time.

'===============================================================================
' PUBLIC VARIABLES
'===============================================================================
' NOTE: These module-level Public variables are a significant code smell.
' They create hidden dependencies between subroutines and make it impossible
' to reason about state. Ideally, each sub should declare its own locals and
' pass values via parameters. Refactoring these out is a large effort but would
' dramatically improve maintainability.
'
' COUNTER / INDEX VARIABLES
' Loop counters are declared locally with descriptive names in each sub.
Public Ord As Long, Ignore As Long, ActiveYr As Long, TectCode As Long, CodeIndex As Long
Public TotalLines As Long, TotalOrders As Long, SearchOrd As Long, LastRow As Long, pdfRow As Long
Public PrevOrd As Long, LineNum As Long, PrevLine As Long, PC_Num As Long, PCnum As Long
Public ConfMon As Long, ConfYr As Long

' FLAGS
Public SAB_Appr As Boolean, Flg1 As Boolean

' NUMERIC VALUES
Public LineQuan As Double
Public LineVal As Currency, PerVal As Currency, OrdVal As Currency
Public FieldServ As Currency, PrevVal As Currency

' DATES
Public CExD As Date, PO_Date As Date, CreatedOn As Date, Released As Date
Public PC_Rel As Date, ME_Rel As Date, PrevDate As Date
Public IntoProdEng As Date, PC_A_Out As Date, ME_A_Out As Date
Public PC_A_Back As Date, ME_A_Back As Date
Public StartSpan As Date, EndSpan As Date

' STRINGS
Public CustName As String, MatNum As String, Desc As String, SoldTo As String
Public LineDesc As String, FileNo As String
Public Docs As String, PC_Eng As String, ME_Eng As String
Public PCStatus As String, MEStatus As String, PrjLvl As String, Customer As String
Public ProdCont As String, SU As String, AssyStr As String, OrdStr As String
Public URL_Text As String, URL_Review As String
Public AttachFile As String, AttachPath As String, LookupPath As String
Public AttachPathME As String, AttachPathEE As String
Public FoundDesignSheet As String, DesignSheetPath As String
Public CcList As String, CurrentUser As String

' OBJECTS
Public FSO As Object

'===============================================================================
' CONSTANTS - FILE PATHS
'===============================================================================
' Path_SAP: Network share where SAP data exports (Open.XLS, Parts.XLS) are stored.
' These files are generated by SAP and dropped here for the scheduling tool to consume.
Public Const Path_SAP = "\\USWWQ-P-FS01.cps.local\PAC5-W\Manufacturing\Projects\Data"


'===============================================================================
' CONSTANTS - DATABASE CONNECTION
'===============================================================================
' [SECURITY WARNING] Hardcoded credentials in source code. This is a significant
' security risk as anyone with access to this .xlsm file can see the DB password.
'
' RECOMMENDATION: Switch to Windows Integrated Authentication:
'   "Driver={SQL Server};Server=USLXA-P-SQL01.cps.local;Database=PAC1CPE;Trusted_Connection=yes;"
' This uses the logged-in Windows user's credentials instead of a shared password.
' If a shared service account is required, store credentials in a separate config
' file with restricted permissions, or use Windows Credential Manager.
Public Const serverName As String = "USLXA-P-SQL01.cps.local"
Public Const databaseName As String = "PAC1CPE"
Public Const databaseUserID As String = "dash"
Public Const databasePassword As String = "manage_DB"

Public Const databaseConnectionString As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & databaseName & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"

'===============================================================================
' CONSTANT - TEMPORARY ORDER FILES BASE PATH
'===============================================================================
' This is the network share where order folder structures and transmittal templates
' are created before being zipped and uploaded into Windchill.
Private Const TEMP_ORDER_BASE As String = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE\Temporary_Order_Files\"


'===============================================================================
' SUB: Insert_SAP_Data
'===============================================================================
' PURPOSE: Opens SAP data exports (Open.XLS for open equipment, Parts.XLS for
'          parts) from the W: drive network share and copies them into this
'          workbook's hidden data sheets (Open_EquipWS, Open_PartsWS).
'
' PROCESS:
'   1. Clears existing data on the Open_EquipWS sheet
'   2. Opens Open.XLS as a tab-delimited text file
'   3. Copies all data from Open.XLS into Open_EquipWS
'   4. Closes Open.XLS
'   5. Repeats steps 1-4 for Parts.XLS into Open_PartsWS
'
' BAD PRACTICES:
'   - Hardcoded row limit of 65000 (was old Excel limit; modern limit is 1,048,576)
'     Now uses UsedRange to avoid this, but note UsedRange can include blank rows
'     if the source file has formatting applied below the data.
'   - No error handling if files don't exist or network path is unavailable
'===============================================================================
Sub Insert_SAP_Data()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- LOAD OPEN EQUIPMENT DATA ---
    Open_EquipWS.Range("A1:AB25000").ClearContents

    Workbooks.OpenText Filename:=Path_SAP & "\Open.XLS", Origin:=xlWindows _
                            , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                            ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
                            , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
                            Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
                            Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
                            16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
                            Array(23, 1)), TrailingMinusNumbers:=True

    Workbooks("Open.XLS").Sheets(1).UsedRange.Copy
    Open_EquipWS.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Workbooks("Open.XLS").Close SaveChanges:=False

    ' --- LOAD PARTS DATA ---
    Open_PartsWS.Range("A1:AB25000").ClearContents

    Workbooks.OpenText Filename:=Path_SAP & "\Parts.XLS", Origin:=xlWindows _
                            , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                            ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
                            , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
                            Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
                            Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
                            16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
                            Array(23, 1)), TrailingMinusNumbers:=True

    Workbooks("Parts.XLS").Sheets(1).UsedRange.Copy
    Open_PartsWS.Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Workbooks("Parts.XLS").Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


'===============================================================================
' SUB: ProdCont_Query
'===============================================================================
' PURPOSE: Given an order number entered in New_OrderWS.Cells(5,2), queries the
'          cpe_scheduling database for scheduling info, then searches the SAP
'          data sheets (Open_EquipWS, Open_PartsWS) for matching line items.
'
' PROCESS:
'   1. Queries cpe_scheduling for: me_netact, scheddate, appdate, reldate,
'      proddate, shipdate, file_num, po_num, scheduler
'   2. Formats the file number (inserts a period: "XXXX YYYYY" -> "XXXX.YYYYY")
'   3. Loops through Open_EquipWS looking for matching order/line items
'   4. Loops through Open_PartsWS looking for matching order/line items
'
' SQL INJECTION: The order number is concatenated directly into the query string.
'   While it's a Long (numeric), the pattern should still use parameterized queries
'   for consistency and safety.
'
' PERFORMANCE: The loops through 10,000 rows could be replaced with Range.Find
'   or VLOOKUP for better performance, though they work fine for current data volumes.
'===============================================================================
Sub ProdCont_Query()
    Dim rowIdx As Long, lineIdx As Long

    ' Clear previous query results from the New Order worksheet
    New_OrderWS.Range("C5:AB5").ClearContents
    New_OrderWS.Range("F9:J55").ClearContents
    New_OrderWS.Range("B9:D9").ClearContents

    ' Open database connection
    ' TODO: Consider using a shared function for DB connection setup/teardown
    Dim oCon As ADODB.Connection: Set oCon = New ADODB.Connection
    oCon.ConnectionString = databaseConnectionString
    oCon.Open

    ' Read the order number from the worksheet
    Ord = New_OrderWS.Cells(5, 2).value
    OrdStr = Str(Ord)

    ' Query cpe_scheduling for this order's scheduling data
    ' Fields returned: me_netact, scheddate, appdate, reldate, proddate,
    '                  shipdate, file_num, po_num, scheduler
    Dim oRS As ADODB.Recordset: Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT me_netact, scheddate, appdate, reldate, proddate, " & _
                 "shipdate, file_num, po_num, scheduler " & _
                 "FROM cpe_scheduling WHERE cpe_scheduling.orderno=" & Ord & ";"
    oRS.Open

    ' Dump query results to a hidden "QuerySheet" then copy values to New_OrderWS.
    ' This two-step approach (paste to hidden sheet, then copy values) avoids
    ' formatting issues from CopyFromRecordset.
    QuerySheet.Range("C5").CopyFromRecordset oRS
    QuerySheet.Range("C5:L5").Copy
    New_OrderWS.Range("C5:L5").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                       :=False, Transpose:=False

    ' Format the file number: Remove spaces, then insert a period after the 4th character.
    ' SAP stores file numbers like "1234 56789"; we need "1234.56789"
    FileNo = New_OrderWS.Range("I5").value
    FileNo = Replace(FileNo, " ", "", 1)
    FileNo = Left(FileNo, 4) & "." & Right(FileNo, 5)
    New_OrderWS.Range("I5").value = FileNo

    ' Clean up the staging area
    QuerySheet.Range("C5:L1048576").ClearContents

    ' --- SEARCH OPEN EQUIPMENT DATA ---
    ' Loop through the Open Equipment SAP data to find line items matching
    ' this order number. Populates line details into New_OrderWS starting at row 9.
    '
    ' Column mapping for Open_EquipWS:
    '   Col 3 = Order Number
    '   Col 5 = Line Number
    '   Col 6 = Material Number
    '   Col 7 = Description
    '   Col 10 = Customer Name
    '   Col 15 = Price (when LineNum=0, this is the order-level price)
    lineIdx = 0
    For rowIdx = 8 To 10000
        SearchOrd = Open_EquipWS.Cells(rowIdx, 3).value
        LineNum = Open_EquipWS.Cells(rowIdx, 5).value

        ' Match on order number with a non-zero line number (actual line items)
        If Ord = SearchOrd And LineNum <> 0 Then
            New_OrderWS.Cells(9, 2).value = Open_EquipWS.Cells(rowIdx, 10).value  ' Customer Name
            lineIdx = lineIdx + 1
            New_OrderWS.Cells(lineIdx + 8, 6).value = LineNum                                   ' Line Number
            New_OrderWS.Cells(lineIdx + 8, 7).value = Open_EquipWS.Cells(rowIdx, 7).value  ' Description
            New_OrderWS.Cells(lineIdx + 8, 10).value = Open_EquipWS.Cells(rowIdx, 6).value ' Material Number
        End If

        ' Line 0 entry holds the order-level price
        If Ord = SearchOrd And LineNum = 0 Then
            New_OrderWS.Cells(17, 3).value = Open_EquipWS.Cells(rowIdx, 15).value
        End If

        ' Data is sorted by order number; if we've passed valid orders, stop early
        If SearchOrd < 1 Then Exit For

    Next rowIdx

    ' --- SEARCH PARTS DATA ---
    ' Same logic as above but against the Parts SAP data sheet.
    ' Parts data may contain additional line items not in the Equipment export.
    '
    ' NOTE: This overwrites any data from the Equipment loop above for the same
    ' line numbers. This appears intentional (Parts data takes priority) but
    ' could be a bug if both exports have overlapping line items with different data.
    lineIdx = 0
    For rowIdx = 8 To 10000
        SearchOrd = Open_PartsWS.Cells(rowIdx, 3).value
        LineNum = Open_PartsWS.Cells(rowIdx, 5).value
        If Ord = SearchOrd And LineNum <> 0 Then
            New_OrderWS.Cells(9, 2).value = Open_PartsWS.Cells(rowIdx, 10).value
            lineIdx = lineIdx + 1
            New_OrderWS.Cells(lineIdx + 8, 6).value = LineNum
            New_OrderWS.Cells(lineIdx + 8, 7).value = Open_PartsWS.Cells(rowIdx, 7).value
            New_OrderWS.Cells(lineIdx + 8, 10).value = Open_PartsWS.Cells(rowIdx, 6).value
        End If

        If Ord = SearchOrd And LineNum = 0 Then
            New_OrderWS.Cells(17, 3).value = Open_PartsWS.Cells(rowIdx, 15).value
        End If

        If SearchOrd < 1 Then Exit For

    Next rowIdx

    ' Reset number format on the line item area (CopyFromRecordset may set dates)
    New_OrderWS.Range("F9:J1048576").NumberFormat = "General"

End Sub


'===============================================================================
' SUB: Print_SAP_Order
'===============================================================================
' PURPOSE: Automates the SAP GUI to print/archive a sales order confirmation PDF.
'          Connects to a running SAP session, navigates to transaction VA02
'          (Change Sales Order), sets the print output to "Archive Only", then
'          retrieves the most recently created PDF attachment.
'
' PROCESS:
'   1. Searches for an open "SAP Easy Access" window across up to 7 sessions
'   2. Opens VA02 (Change Sales Order) with the order number
'   3. Changes the print output mode to "Archive Only" (key "2")
'   4. Triggers a print, then filters attachments by today's date
'   5. Opens the most recent PDF attachment
'
' BAD PRACTICES:
'   - Uses GoTo for normal control flow (NoHomeScreen, SkipNoHomeScreen)
'   - Uses On Error Resume Next / GoTo as a means of iterating SAP sessions
'   - "While 1 < 2" infinite loop (line ~284) with error-based exit to count
'     rows in the SAP attachment list. There's no SAP Scripting API method to
'     get the row count, so this hack is somewhat understandable, but it should
'     at least have a safety counter to prevent true infinite loops.
'   - The SAPIndex variable on line ~248 is never declared or set in this sub;
'     it likely works because another sub sets it, which is fragile.
'   - Multiple GoTo-based error handlers for closing SAP popup windows make
'     the control flow very hard to follow.
'===============================================================================
Sub Print_SAP_Order()
    Dim sessionIdx As Long
    On Error GoTo NoHomeScreen

    Dim Session
    ' Search through up to 7 SAP sessions looking for one with "SAP Easy Access" title
    For sessionIdx = 0 To 6
        If Left(GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(sessionIdx).findById("wnd[0]").Text, 15) = "SAP Easy Access" Then
            Set Session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(sessionIdx)
            On Error GoTo 0
            Exit For
        End If
    Next sessionIdx

    GoTo SkipNoHomeScreen:

NoHomeScreen:
    ' Error handler: SAP is either not running or no Easy Access window is open
    If sessionIdx = 0 Then
        MsgBox "Cannot connect to SAP. Open an SAP Easy Access window to run the program.", Title:="Error!"
        Exit Sub
    Else
        MsgBox "An SAP Easy Access window must be open to run the program.", Title:="Error!"
        Exit Sub
    End If
    
SkipNoHomeScreen:

    ' Read order number and navigate to VA02 (Change Sales Order)
    Ord = New_OrderWS.Cells(5, 2).value
    Session.findById("wnd[0]").maximize
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "VA02"
    Session.findById("wnd[0]").sendVKey 0
    Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = Ord
    Session.findById("wnd[0]").sendVKey 0

    ' --- SET PRINT OPTIONS TO ARCHIVE ONLY ---
    ' Navigate: Extras > Output > Header > Display
    On Error GoTo CloseWin3
    Session.findById("wnd[0]/mbar/menu[3]/menu[13]/menu[0]/menu[0]").Select
    On Error GoTo 0

    ' Select the first output row and check its dispatch mode
    Session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(0).Selected = True
    Session.findById("wnd[0]/tbar[1]/btn[2]").press  ' Detail button

    ' If dispatch mode is already "2" (Archive), just back out
    If Session.findById("wnd[0]/usr/cmbNAST-TDARMOD").Key = "2" Then
        Session.findById("wnd[0]/tbar[0]/btn[3]").press 'back
        Session.findById("wnd[0]/tbar[0]/btn[3]").press 'back
        Session.findById("wnd[0]/tbar[0]/btn[3]").press 'back
    Else
        ' Otherwise, change the dispatch mode to "2" (Archive Only) and save
        Session.findById("wnd[0]/tbar[0]/btn[3]").press 'back
        Session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").getAbsoluteRow(0).Selected = True
        Session.findById("wnd[0]/tbar[1]/btn[6]").press   ' Change button
        Session.findById("wnd[0]/tbar[1]/btn[2]").press   ' Detail button
        Session.findById("wnd[0]/usr/cmbNAST-TDARMOD").Key = "2"  ' Set to Archive
        Session.findById("wnd[0]/tbar[0]/btn[3]").press 'back
        On Error GoTo CloseWin1
        Session.findById("wnd[0]/tbar[0]/btn[11]").press 'save
        On Error GoTo 0
    End If

    ' Handle any popup confirmation dialogs
    If Left(GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0).findById("wnd[0]").Text, 34) <> "Change Sales Order: Initial Screen" Then
        On Error Resume Next
        Session.findById("wnd[1]").sendVKey 0
        Session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press
        On Error GoTo 0
    End If

    ' Print the order confirmation
    Session.findById("wnd[0]/mbar/menu[0]/menu[5]").Select

    On Error GoTo CloseWin1
    Session.findById("wnd[1]/usr/tblSAPLVMSGTABCONTROL").getAbsoluteRow(0).Selected = True
    On Error GoTo 0

    Session.findById("wnd[1]/tbar[0]/btn[86]").press
    Session.findById("wnd[0]").sendVKey 0

ResumeHere:
    ' Open the attachment list (GOS Toolbox)
    On Error GoTo CloseWin1
    Session.findById("wnd[0]/titl/shellcont/shell").pressButton "%GOS_TOOLBOX"
    On Error GoTo 0

    On Error GoTo CloseWin2
    Session.findById("wnd[0]/shellcont/shell").pressButton "VIEW_ATTA"
    On Error GoTo 0

    ' Filter the attachment list to only show PDFs created today
    Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").setCurrentCell -1, "CREADATE"
    Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectColumn "CREADATE"
    Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").contextMenu
    Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectContextMenuItem "&FILTER"
    Session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = Date
    Session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = Date
    Session.findById("wnd[2]/tbar[0]/btn[0]").press

    ' Count rows in the filtered attachment list.
    ' BAD PRACTICE: This uses an intentional error (selecting a non-existent row)
    ' to find the last row. There's no RowCount property in the SAP ALV Grid control,
    ' so this is a common workaround, but adding a safety limit (e.g., max 1000)
    ' would prevent a true infinite loop if something goes wrong.
    pdfRow = 0
    On Error GoTo FoundLastEntry
    While 1 < 2
        Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = pdfRow
        pdfRow = pdfRow + 1
    Wend

FoundLastEntry:
    On Error GoTo 0
    ' Select the last (most recent) PDF entry
    Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = (pdfRow - 1)

    ' Open the PDF and close out
    Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
    Session.findById("wnd[1]/tbar[0]/btn[12]").press
    Session.findById("wnd[0]/shellcont").Close
    Session.findById("wnd[0]/tbar[0]/btn[3]").press  ' Back
    Session.findById("wnd[0]/tbar[0]/btn[3]").press  ' Back
    Exit Sub

CloseWin1:
    ' Error handler: dismiss a popup and retry the failed statement
    Session.findById("wnd[0]").sendVKey 0
    Resume

CloseWin2:
    ' Error handler: dismiss a popup and jump to ResumeHere
    Session.findById("wnd[0]").sendVKey 0
    Resume ResumeHere

CloseWin3:
    ' Error handler: dismiss a secondary window popup and retry
    Session.findById("wnd[1]").sendVKey 0
    Resume

End Sub


'===============================================================================
' SUB: NewOrder2NewDB
'===============================================================================
' PURPOSE: Core order entry routine. Inserts a new order (with all line items)
'          into the Prod_Eng database table, creates the folder structure for
'          transmittal documents in Temporary_Order_Files, populates Customer
'          Release Forms and Production Release Forms from templates, and zips
'          the result for Windchill upload.
'
' PROCESS:
'   1. Opens a DB connection and reads order details from New_OrderWS
'   2. For each line item (rows 9+), builds a massive INSERT statement and
'      executes it against Prod_Eng
'   3. Creates folder structure:
'        Temporary_Order_Files\{Ord}\
'        Temporary_Order_Files\{Ord}\{Ord}\
'        Temporary_Order_Files\{Ord}\{Ord}\{Ord}M\  (if ME assigned)
'        Temporary_Order_Files\{Ord}\{Ord}\{Ord}E\  (if PC/EE assigned)
'   4. Copies and populates template workbooks:
'        - XXXXXXM.xlsm -> {Ord}M.xlsm (Customer Release - Mechanical)
'        - XXXXXXE.xlsm -> {Ord}E.xlsm (Customer Release - Electrical)
'        - F7-113...MR.xlsm -> {Ord}MR.xlsm (Production Transmittal - Mech)
'        - F7-113...ER.xlsm -> {Ord}ER.xlsm (Production Transmittal - Elec)
'   5. Moves any incoming design sheets into the ME folder
'   6. Zips the folder for Windchill upload
'   7. Clears the order entry form
'
' BAD PRACTICES:
'   [FIXED] Previously used On Error GoTo DeleteDirectory to handle the case
'       where the top-level folder already existed. The error handler would
'       DELETE the entire folder tree and resume, losing any files that had
'       been placed there by earlier pipeline steps. Now uses EnsureFolder()
'       helper to safely check-and-create.
'   [SQL INJECTION] The INSERT statement is built entirely via string concatenation.
'       Apostrophes in customer names are "handled" by replacing them with carets (^),
'       which corrupts the data. Should use ADODB.Command with parameters.
'   [ACTION QUERY VIA RECORDSET] The INSERT is executed via oRS.Open instead of
'       oCon.Execute. This works but is semantically incorrect.
'   [MAGIC NUMBERS] Every cell reference is a raw row/column number. Named ranges
'       or constants would make this much easier to maintain.
'===============================================================================
Sub NewOrder2NewDB()
    Dim lineIdx As Long, rowIdx As Long
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Open database connection
    Dim oCon As ADODB.Connection: Set oCon = New ADODB.Connection
    oCon.ConnectionString = databaseConnectionString
    oCon.Open

    ' Read the order number from the New Order worksheet
    Ord = New_OrderWS.Cells(5, 2).value

    ' Find the last row with line item data (column F = line numbers)
    LastRow = New_OrderWS.Cells(Rows.Count, 6).End(xlUp).Row

    ' --- INSERT EACH LINE ITEM INTO Prod_Eng ---
    For lineIdx = 9 To LastRow

        ' Stop if we hit an empty line number (end of line items)
        If New_OrderWS.Cells(lineIdx, 6).value = "" Then Exit For

        ' Sanitize the description: replace apostrophes with carets to prevent
        ' SQL syntax errors. NOTE: This is a lossy workaround. The real fix is
        ' to use parameterized queries (ADODB.Command) which handle special
        ' characters properly without corrupting the data.
        LineDesc = New_OrderWS.Cells(lineIdx, 7).value
        LineDesc = Replace(LineDesc, "'", Chr(94))  ' Chr(94) = "^"

        ' Build the VALUES clause for the INSERT statement.
        ' This is a single massive concatenated string mapping worksheet cells
        ' to Prod_Eng columns. The format is positional and must match the
        ' column list in the INSERT INTO clause exactly.
        AssyStr = ""
        AssyStr = New_OrderWS.Cells(5, 2).value & ",'"    ' Order_Num (int, no quotes)
        AssyStr = AssyStr & New_OrderWS.Cells(9, 2).value & "','"    ' Customer_Name
        AssyStr = AssyStr & New_OrderWS.Cells(lineIdx, 6).value & "','"    ' Line_Num
        AssyStr = AssyStr & New_OrderWS.Cells(lineIdx, 10).value & "','"   ' Material
        AssyStr = AssyStr & LineDesc & "','"                          ' Description
        AssyStr = AssyStr & New_OrderWS.Cells(5, 3).value & "','"    ' Network1
        AssyStr = AssyStr & New_OrderWS.Cells(6, 3).value & "','"    ' Network2
        AssyStr = AssyStr & New_OrderWS.Cells(7, 3).value & "','"    ' Network3
        AssyStr = AssyStr & New_OrderWS.Cells(8, 3).value & "','"    ' Network4
        AssyStr = AssyStr & New_OrderWS.Cells(11, 3).value & "','"   ' Industry
        AssyStr = AssyStr & New_OrderWS.Cells(11, 2).value & "','"   ' DocTypes
        AssyStr = AssyStr & New_OrderWS.Cells(5, 9).value & "','"    ' FileNum
        AssyStr = AssyStr & New_OrderWS.Cells(17, 2).value & "','"   ' Region_AE

        ' PC1 (Primary Electrical/Controls Engineer) - default to "N/A" if unassigned
        If New_OrderWS.Cells(13, 2).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(13, 2).value & "','"
        Else
            AssyStr = AssyStr & "N/A" & "','"
        End If

        AssyStr = AssyStr & New_OrderWS.Cells(13, 5).value & "','"   ' PC2 (Secondary EE)

        ' ME1 (Primary Mechanical Engineer) - default to "N/A" if unassigned
        If New_OrderWS.Cells(15, 2).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(15, 2).value & "','"
        Else
            AssyStr = AssyStr & "N/A" & "','"
        End If

        AssyStr = AssyStr & New_OrderWS.Cells(15, 5).value & "','"   ' ME2 (Secondary ME)
        AssyStr = AssyStr & New_OrderWS.Cells(17, 5).value & "','"   ' PM (Project Manager)
        AssyStr = AssyStr & New_OrderWS.Cells(19, 5).value & "','"   ' PE (Process Engineer)
        AssyStr = AssyStr & New_OrderWS.Cells(5, 11).value & "','"   ' Scheduler
        AssyStr = AssyStr & "','',"                                   ' PC_Chkr & ME_Chkr (empty)

        ' Numeric estimates (hours) - the 0# + forces numeric evaluation
        AssyStr = AssyStr & (0# + New_OrderWS.Cells(lineIdx, 11).value) & ", 0.0,"  ' PC_Est, PC_CO_Est
        AssyStr = AssyStr & (0# + New_OrderWS.Cells(lineIdx, 12).value) & ", 0.0,"  ' ME_Est, ME_CO_Est
        AssyStr = AssyStr & (0# + New_OrderWS.Cells(17, 4).value) & ","       ' Supp_Est
        AssyStr = AssyStr & (0# + New_OrderWS.Cells(13, 4).value) & ","       ' Soft_Est (Programming)
        AssyStr = AssyStr & (0# + New_OrderWS.Cells(15, 4).value) & ","       ' PM_Est

        AssyStr = AssyStr & "'WW','"                                  ' Eng_Loc (always "WW")
        AssyStr = AssyStr & New_OrderWS.Cells(19, 3).value & "','"   ' MultiPlnt
        AssyStr = AssyStr & New_OrderWS.Cells(19, 2).value & "','"   ' Pjt_Lvl

        ' Set initial status: "IN QUEUE" if an engineer is assigned, blank otherwise
        If New_OrderWS.Cells(13, 2).value <> "N/A" Then
            AssyStr = AssyStr & "IN QUEUE','"    ' PC_Status
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(15, 2).value <> "N/A" Then
            AssyStr = AssyStr & "IN QUEUE','"    ' ME_Status
        Else
            AssyStr = AssyStr & "','"
        End If

        AssyStr = AssyStr & "','',"              ' PC_Override & ME_Override (empty)
        AssyStr = AssyStr & New_OrderWS.Cells(11, 4).value & ",'"    ' Active_Year

        ' Date fields - only include if non-empty (empty dates cause SQL errors)
        If New_OrderWS.Cells(5, 4).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 4).value & "','"  ' Into_Eng
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(5, 5).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 5).value & "','"  ' Apps_Out_0
        Else
            AssyStr = AssyStr & "','"
        End If

        AssyStr = AssyStr & "','','','','"       ' PC/ME Apps Out & Back (empty initially)

        If New_OrderWS.Cells(21, 3).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(21, 3).value & "','"  ' PC_PreRel
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(19, 4).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(19, 4).value & "','"  ' ME_PreRel
        Else
            AssyStr = AssyStr & "','"
        End If

        AssyStr = AssyStr & "','','"             ' PC_Act_PreRel & ME_Act_PreRel (empty)

        If New_OrderWS.Cells(5, 6).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 6).value & "','"  ' PC_Rel_F
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(5, 6).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 6).value & "','"  ' ME_Rel_F
        Else
            AssyStr = AssyStr & "','"
        End If

        AssyStr = AssyStr & "','','"             ' PC_Act_Rel & ME_Act_Rel (empty)
        AssyStr = AssyStr & "','"                ' ShippedDate (empty)

        If New_OrderWS.Cells(5, 6).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 6).value & "','"  ' PC_Rel_0 (initial)
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(5, 6).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 6).value & "','"  ' ME_Rel_0 (initial)
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(5, 7).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 7).value & "','"  ' Prod_Rel_0 (initial)
        Else
            AssyStr = AssyStr & "','"
        End If

        If New_OrderWS.Cells(5, 8).value <> "" Then
            AssyStr = AssyStr & New_OrderWS.Cells(5, 8).value & "','"  ' Ship_Date_0 (initial)
        Else
            AssyStr = AssyStr & "','"
        End If

        AssyStr = AssyStr & New_OrderWS.Cells(17, 3).value & "'"      ' Price

        Debug.Print AssyStr

        ' Execute the INSERT statement
        Dim insertSQL As String
        insertSQL = "INSERT INTO Prod_Eng(uniqid, Order_Num, Customer_Name, Line_Num, " & _
                    "Material, Description, Network1, Network2, Network3, Network4, " & _
                    "Industry, DocTypes, FileNum, Region_AE, PC1, PC2, ME1, ME2, PM, PE, " & _
                    "Scheduler, PC_Chkr, ME_Chkr, PC_Est, PC_CO_Est, ME_Est, ME_CO_Est, " & _
                    "Supp_Est, Soft_Est, PM_Est, Eng_Loc, MultiPlnt, Pjt_Lvl, PC_Status, " & _
                    "ME_Status, PC_Override, ME_Override, Active_Year, Into_Eng, Apps_Out_0, " & _
                    "PC_Apps_Out, ME_Apps_Out, PC_Apps_Back, ME_Apps_Back, PC_PreRel, " & _
                    "ME_PreRel, PC_Act_PreRel, ME_Act_PreRel, PC_Rel_F, ME_Rel_F, " & _
                    "PC_Act_Rel, ME_Act_Rel, ShippedDate, PC_Rel_0, ME_Rel_0, Prod_Rel_0, " & _
                    "Ship_Date_0, Price) VALUES(newid(), " & _
                    AssyStr & _
                    ")"
        Debug.Print insertSQL
        oCon.Execute insertSQL

    Next lineIdx

    If Not oCon Is Nothing Then Set oCon = Nothing

    ' --- CREATE FOLDER STRUCTURE FOR TRANSMITTAL DOCUMENTS ---

    ' Calculate a text code used for the Confirmation Page folder name.
    ' This sums the digits of the order number (positions 5-10) plus 2.
    ' Example: Ord = 1100102999 -> digits at positions 5-10 = 1,0,2,9,9,9
    '          TextCode = 2 + 1 + 0 + 2 + 9 + 9 + 9 = 32
    
    Dim TextCode: TextCode = 2
    For CodeIndex = 5 To 10
        TextCode = TextCode + Mid(Ord, CodeIndex, 1)
    Next CodeIndex

    '===========================================================================
    ' [FIXED] SAFE DIRECTORY CREATION
    '===========================================================================
    ' PREVIOUS BEHAVIOR (DANGEROUS):
    '   Used "On Error GoTo DeleteDirectory" which would delete the entire folder
    '   tree if the top-level folder already existed, then retry MkDir. This
    '   destroyed any files that had been placed there by earlier pipeline steps
    '   (e.g., files from scheduling, design sheets, etc.).
    '
    ' NEW BEHAVIOR:
    '   Uses the EnsureFolder() helper function to check if each directory exists
    '   before attempting to create it. Existing directories are left untouched,
    '   preserving any files already present.
    '===========================================================================
    Dim basePath As String
    basePath = TEMP_ORDER_BASE & Ord

    ' Create the top-level order folder: Temporary_Order_Files\{Ord}\
    EnsureFolder basePath

    ' Create the inner order folder: Temporary_Order_Files\{Ord}\{Ord}\
    EnsureFolder basePath & "\" & Ord

    ' --- MECHANICAL ENGINEERING (ME) TRANSMITTALS ---
    ' Only create ME folders and templates if an ME engineer is assigned
    If New_OrderWS.Cells(15, 2).value <> "N/A" And New_OrderWS.Cells(15, 2).value <> "" Then

        ' Create ME subfolder: ...\{Ord}\{Ord}\{Ord}M\
        EnsureFolder basePath & "\" & Ord & "\" & Ord & "M"

        ' Open the Customer Release template for ME, populate it, and save as order-specific
        ' Template location: Temporary_Order_Files\XXXXXXM.xlsm
        Dim ShtERS
        ShtERS = TEMP_ORDER_BASE & "XXXXXXM.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "XXXXXXM.xlsm"
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(9, 4).value = New_OrderWS.Cells(9, 2).value   ' Customer Name
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(10, 4).value = New_OrderWS.Cells(5, 2).value  ' Order Number
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(11, 4).value = New_OrderWS.Cells(5, 10).value ' PO Number
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(13, 4).value = New_OrderWS.Cells(15, 13).value ' Designer Name (ME)
        ' Save as: ...\{Ord}\{Ord}\{Ord}M\{Ord}M.xlsm  (e.g., 1100102999M.xlsm)
        ActiveWorkbook.SaveAs Filename:=basePath & "\" & Ord & "\" & Ord & "M\" & Ord & "M.xlsm", _
                              FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
    End If

    ' --- ELECTRICAL ENGINEERING (PC/EE) TRANSMITTALS ---
    ' Only create EE folders and templates if a PC/EE engineer is assigned
    If New_OrderWS.Cells(13, 2).value <> "N/A" And New_OrderWS.Cells(13, 2).value <> "" Then

        ' Create EE subfolder: ...\{Ord}\{Ord}\{Ord}E\
        EnsureFolder basePath & "\" & Ord & "\" & Ord & "E"

        ' Open the Customer Release template for EE, populate it, and save
        ShtERS = TEMP_ORDER_BASE & "XXXXXXE.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "XXXXXXE.xlsm"
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(9, 4).value = New_OrderWS.Cells(9, 2).value   ' Customer Name
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(10, 4).value = New_OrderWS.Cells(5, 2).value  ' Order Number
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(11, 4).value = New_OrderWS.Cells(5, 10).value ' PO Number
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(13, 4).value = New_OrderWS.Cells(13, 13).value ' Designer Name (EE)
        ' Save as: ...\{Ord}\{Ord}\{Ord}E\{Ord}E.xlsm  (e.g., 1100102999E.xlsm)
        ActiveWorkbook.SaveAs Filename:=basePath & "\" & Ord & "\" & Ord & "E\" & Ord & "E.xlsm", _
                              FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
    End If

    ' --- PRODUCTION RELEASE TRANSMITTALS ---
    ' These are the "F7-113" forms that go to Manufacturing Engineering

    ' ME Production Release Transmittal (MR)
    If New_OrderWS.Cells(15, 2).value <> "N/A" And New_OrderWS.Cells(15, 2).value <> "" Then
        ShtERS = TEMP_ORDER_BASE & "F7-113 Eng Production Transmittal MR.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "F7-113 Eng Production Transmittal MR.xlsm"
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(8, 3).value = New_OrderWS.Cells(9, 2).value   ' Customer Name
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(9, 3).value = New_OrderWS.Cells(5, 2).value   ' Order Number
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(11, 3).value = New_OrderWS.Cells(15, 13).value ' Designer (ME)
        ' Save as: ...\{Ord}\{Ord}\{Ord}M\{Ord}MR.xlsm  (e.g., 1100102999MR.xlsm)
        ActiveWorkbook.SaveAs Filename:=basePath & "\" & Ord & "\" & Ord & "M\" & Ord & "MR.xlsm", _
                              FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
    End If

    ' EE Production Release Transmittal (ER)
    If New_OrderWS.Cells(13, 2).value <> "N/A" And New_OrderWS.Cells(13, 2).value <> "" Then
        ShtERS = TEMP_ORDER_BASE & "F7-113 Eng Production Transmittal ER.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "F7-113 Eng Production Transmittal ER.xlsm"
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(8, 3).value = New_OrderWS.Cells(9, 2).value   ' Customer Name
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(9, 3).value = New_OrderWS.Cells(5, 2).value   ' Order Number
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(11, 3).value = New_OrderWS.Cells(13, 13).value ' Designer (EE)
        ' Save as: ...\{Ord}\{Ord}\{Ord}E\{Ord}ER.xlsm  (e.g., 1100102999ER.xlsm)
        ActiveWorkbook.SaveAs Filename:=basePath & "\" & Ord & "\" & Ord & "E\" & Ord & "ER.xlsm", _
                              FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
        ActiveWorkbook.Close
    End If

    ' --- ZIP THE FOLDER FOR WINDCHILL UPLOAD ---
    ' Zips the entire order folder structure into a single .zip file.
    ' The engineers will then upload this zip into Windchill.
    ' TODO: Eventually replace this with direct Windchill API upload.
    Call CreateZipFile((basePath & "\"), (basePath & ".zip"))

    ' NOTE: The following lines were previously used to delete the folder after zipping.
    ' They have been commented out because the folder may still be needed, and deletion
    ' should be a deliberate user action, not automatic.
    'Set FSO = CreateObject("scripting.filesystemobject")
    'FSO.deletefolder basePath

SKIP2:
    ' --- CLEAR THE ORDER ENTRY FORM ---
    ' Reset the New Order worksheet for the next order entry
    New_OrderWS.Range("A5:L5").ClearContents
    New_OrderWS.Range("F9:J26").ClearContents
    New_OrderWS.Range("B9:D9").ClearContents

    New_OrderWS.Cells(11, 2).value = ""     ' DocTypes
    New_OrderWS.Cells(11, 3).value = ""     ' Industry
    New_OrderWS.Cells(13, 2).value = ""     ' PC1
    New_OrderWS.Cells(13, 5).value = ""     ' PC2
    New_OrderWS.Cells(15, 2).value = ""     ' ME1
    New_OrderWS.Cells(15, 5).value = ""     ' ME2
    New_OrderWS.Cells(17, 2).value = ""     ' Region_AE
    New_OrderWS.Cells(17, 5).value = ""     ' PM
    New_OrderWS.Cells(19, 2).value = ""     ' Pjt_Lvl
    New_OrderWS.Cells(19, 3).value = ""     ' MultiPlnt
    New_OrderWS.Cells(19, 4).value = ""     ' ME_PreRel
    New_OrderWS.Cells(19, 5).value = ""     ' PE
    New_OrderWS.Cells(21, 3).value = ""     ' PC_PreRel

    New_OrderWS.Cells(13, 4).value = 0      ' Soft_Est (Programming hours)
    New_OrderWS.Cells(15, 4).value = 0      ' PM_Est
    New_OrderWS.Cells(17, 3).value = 0      ' Price
    New_OrderWS.Cells(17, 4).value = 0      ' Supp_Est

    ' Clear per-line estimate columns
    For rowIdx = 9 To 40
        New_OrderWS.Cells(rowIdx, 11).value = 0  ' PC_Est per line
        New_OrderWS.Cells(rowIdx, 12).value = 0  ' ME_Est per line
    Next rowIdx

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Exit Sub

    ' NOTE: The old DeleteDirectory error handler has been removed entirely.
    ' It previously deleted the folder tree and resumed MkDir, which was destructive.
    ' The EnsureFolder() helper now handles directory creation safely.

End Sub


'===============================================================================
' SUB: EnsureFolder (HELPER)
'===============================================================================
' PURPOSE: Safely creates a directory if it doesn't already exist.
'          Does NOT delete or modify existing directories.
'
' PARAMETERS:
'   folderPath - Full path to the directory to create
'
' NOTE: This replaces the old pattern of:
'     On Error GoTo DeleteDirectory
'     MkDir "..."
'   which would delete the entire folder tree on collision. This new approach
'   preserves any files that may already exist in the directory from earlier
'   pipeline steps (scheduling, design sheet uploads, etc.).
'===============================================================================
Private Sub EnsureFolder(ByVal folderPath As String)
    ' Dir() with vbDirectory returns the folder name if it exists, empty string if not
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    ' If the folder already exists, we simply do nothing and preserve its contents.
End Sub


'===============================================================================
' SUB: SendNew_Email
'===============================================================================
' PURPOSE: Composes and displays a "New Order" notification email in Outlook.
'          The email is sent to the assigned PC (Electrical) and ME (Mechanical)
'          engineers with:
'            - A link to the order text PDF
'            - A link to the order review form
'            - The scheduling PDF as an attachment
'            - Order review checksheets as attachments (from Temporary_Order_Files)
'
' PROCESS:
'   1. Determines the current user (Randy or Brian) for signature purposes
'   2. Reads engineer assignments and email addresses from EmailsWS
'   3. Builds an HTML email body with links and greetings
'   4. Finds and attaches scheduling PDFs from CPE_Schedule folder
'   5. Moves attached scheduling files to a "Processed" subfolder
'   6. Finds and attaches Order Review PDFs from Temporary_Order_Files
'   7. Displays the email in Outlook for manual review before sending
'
' BAD PRACTICES:
'   - Hardcoded usernames for CurrentUser detection
'   - Hardcoded email distribution lists
'   - Uses On Error GoTo for folder creation (CreateFolder handler)
'   - Multiple On Error Resume Next blocks obscure real errors
'===============================================================================
Sub SendNew_Email()

    Dim EBody As String, PC_EMail As String, ME_EMail As String
    Dim OutApp As Object
    Dim OutMail As Object

    ' Determine which supervisor is running the tool (for email signature)
    ' NOTE: This should use a config setting instead of hardcoded usernames.
    If UCase(Environ("Username")) = "R.MIELKE" Then
        CurrentUser = "Randy"
    Else
        CurrentUser = "Brian"
    End If

    ' Default unassigned engineers to "N/A"
    If EmailsWS.Cells(3, 5).value = "" Then
        EmailsWS.Cells(3, 5).value = "N/A"
    End If
    If EmailsWS.Cells(3, 6).value = "" Then
        EmailsWS.Cells(3, 6).value = "N/A"
    End If

    ' Read order details and engineer assignments from the Emails worksheet
    Ord = EmailsWS.Cells(3, 2).value
    CustName = EmailsWS.Cells(3, 3).value
    PC_Eng = EmailsWS.Cells(7, 5).value       ' EE Engineer name
    ME_Eng = EmailsWS.Cells(7, 6).value       ' ME Engineer name
    PC_EMail = EmailsWS.Cells(8, 7).value     ' EE Engineer email
    ME_EMail = EmailsWS.Cells(8, 8).value     ' ME Engineer email
    URL_Text = EmailsWS.Cells(13, 8).value    ' Link to order text PDF
    URL_Review = EmailsWS.Cells(15, 8).value  ' Link to order review form

    ' Create Outlook email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)  ' 0 = olMailItem

    ' Build HTML email body with appropriate greeting based on assignments
    EBody = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"

    If ((PC_Eng <> "0" And PC_Eng <> "N/A") And (ME_Eng <> "0" And ME_Eng <> "N/A")) Then
        ' Both EE and ME assigned
        EBody = EBody & PC_Eng & " and " & ME_Eng & ",<BR><BR>Please process the attached NEW order notification.<BR><BR>"
        CcList = "fpm_PAC_EngineeringClerk@coperion.com; joey.purdon@coperion.com; brian.schmoldt@coperion.com; scowen@coperionktron.com"
    ElseIf (ME_Eng = "0" Or ME_Eng = "N/A") Then
        ' Only EE assigned (no ME)
        EBody = EBody & PC_Eng & ",<BR><BR>Please process the attached NEW order notification.<BR><BR>"
        CcList = "fpm_PAC_EngineeringClerk@coperion.com; joey.purdon@coperion.com; brian.schmoldt@coperion.com; scowen@coperionktron.com"
    ElseIf (PC_Eng = "0" Or PC_Eng = "N/A") Then
        ' Only ME assigned (no EE)
        EBody = EBody & ME_Eng & ",<BR><BR>Please process the attached NEW order notification.<BR><BR>"
        CcList = "fpm_PAC_EngineeringClerk@coperion.com; brian.schmoldt@coperion.com; scowen@coperionktron.com"
    End If

    ' Add link to the latest order text document
    EBody = EBody & "Use the following link to review the latest order text:<BR>"
    EBody = EBody & "<A HREF='" & URL_Text & "'>"
    EBody = EBody & Ord & "_" & CustName & "</A><BR><BR>"
    
    OrdStr = Trim$(Str(Ord))
    AttachPath = TEMP_ORDER_BASE & OrdStr & "\" & OrdStr & "\Order_Review.pdf"
    AttachPathME = TEMP_ORDER_BASE & OrdStr & "\" & OrdStr & "\" & OrdStr & "M\" & _
                   "OrderReview_" & OrdStr & "_ME.pdf"
    AttachPathEE = TEMP_ORDER_BASE & OrdStr & "\" & OrdStr & "\" & OrdStr & "E\" & _
                   "OrderReview_" & OrdStr & "_EE.pdf"

    ' Check if any Order Review form exists and add link or warning
    Dim FoundOrderReview As String
    FoundOrderReview = Dir(AttachPath) & Dir(AttachPathEE) & Dir(AttachPathME)

    If FoundOrderReview = "" Then
        EBody = EBody & "<span style=""color:#FF0000"">-No Order Review Found-</span style=""color:#FF0000""><BR>"
    End If

    ' Set up the email with recipients, CC, and subject
    With OutMail
        .To = ME_EMail & "; " & PC_EMail
        .CC = CcList
        .BCC = ""
        .Subject = "New Order " & Ord & " - " & CustName
        .Display  ' Display for review; change to .Send for auto-send
    End With

    ' Capture the personal Outlook signature that was auto-appended on .Display
    Dim Signature: Signature = OutMail.HTMLBody
    ' Prepend our email body before the signature
    With OutMail
        .HTMLBody = "<html><body>" & EBody & Signature & "</body></html>"
    End With

    ' --- ATTACH SCHEDULING PDFs ---
    ' Find all files matching the order number in the CPE_Schedule folder
    LookupPath = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\" & Ord & "*"
    AttachFile = Dir(LookupPath, vbDirectory)

    While AttachFile <> ""
        AttachPath = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\" & AttachFile
        With OutMail
            .Attachments.Add (AttachPath)
        End With

        ' Move the attached file to a "Processed" subfolder to prevent re-attachment.
        ' Files starting with "1100" go into order-specific subfolders;
        ' others go into INTERNAL_Orders.
        If Left(AttachFile, 4) = "1100" Then
            On Error GoTo CreateFolder
            Name AttachPath As "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\Processed\" & Left(AttachFile, 7) & "xxx\" & AttachFile
            On Error GoTo 0
        Else
            Name AttachPath As "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\Processed\INTERNAL_Orders\" & AttachFile
        End If

        AttachFile = Dir(LookupPath, vbDirectory)
    Wend

    ' --- ATTACH ORDER REVIEW PDFs ---
    ' Paths built above (Order_Review.pdf, OrderReview_xxx_ME.pdf, OrderReview_xxx_EE.pdf)
    On Error Resume Next
    With OutMail
        If (Dir(AttachPath) <> "") Then
            .Attachments.Add (AttachPath)
        End If

        If (Dir(AttachPathME) <> "") Then
            .Attachments.Add (AttachPathME)
        End If

        If (Dir(AttachPathEE) <> "") Then
            .Attachments.Add (AttachPathEE)
        End If
    End With
    On Error GoTo 0

    ' Clean up Outlook objects
    Set OutMail = Nothing
    Set OutApp = Nothing

    ' Clear the email setup area for the next order
    EmailsWS.Range("B3:Q5").ClearContents
    Exit Sub

CreateFolder:
    ' Error handler: If the "Processed\{prefix}xxx" folder doesn't exist, create it
    ' then retry the file move. This is another instance of error-driven folder creation
    ' but is less dangerous since it only creates, never deletes.
    MkDir "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\Processed\" & Left(AttachFile, 7) & "xxx"
    Resume
End Sub


'===============================================================================
' SUB: Send_Change_Order
'===============================================================================
' PURPOSE: Similar to SendNew_Email but for change order notifications.
'          Composes and displays a "Change Order" notification email.
'          Also attaches any incoming design sheets and deletes them from
'          the incoming folder after attachment.
'
' DIFFERENCES FROM SendNew_Email:
'   - Subject line says "Change Order" instead of "New Order"
'   - CC list is shorter (no full clerk distribution)
'   - Attaches design sheets from the Incoming Design Sheets folder
'   - Deletes design sheets after attaching (they need to go into Windchill)
'   - Adds a red note to the email if design sheets are present
'   - Does NOT look up Order Review forms from the J: drive
'
' BAD PRACTICES:
'   - Kills (deletes) design sheet files after attaching. If the email is
'     discarded without sending, the files are permanently lost.
'   - Same error-driven folder creation pattern as SendNew_Email
'   - CreateFolder1 label shadows the CreateFolder label in SendNew_Email
'     (different subs so no actual collision, but confusing naming)
'===============================================================================
Sub Send_Change_Order()
    Dim EBody As String, PC_EMail As String, ME_EMail As String
    Dim OutApp As Object
    Dim OutMail As Object

    ' Determine current user for signature purposes
    If UCase(Environ("Username")) = "SCHMOLDT-BRI" Then
        CurrentUser = "Brian"
    ElseIf UCase(Environ("Username")) = "SCOWEN" Then
        CurrentUser = "Steve"
    Else
        CurrentUser = vbNullString
    End If

    ' Default unassigned engineers to "N/A"
    If EmailsWS.Cells(3, 5).value = "" Then
        EmailsWS.Cells(3, 5).value = "N/A"
    End If
    If EmailsWS.Cells(3, 6).value = "" Then
        EmailsWS.Cells(3, 6).value = "N/A"
    End If

    ' Read order details from the Emails worksheet
    Ord = EmailsWS.Cells(3, 2).value
    CustName = EmailsWS.Cells(3, 3).value
    PC_Eng = EmailsWS.Cells(7, 5).value
    ME_Eng = EmailsWS.Cells(7, 6).value
    PC_EMail = EmailsWS.Cells(8, 7).value
    ME_EMail = EmailsWS.Cells(8, 8).value
    URL_Text = EmailsWS.Cells(13, 8).value
    URL_Review = EmailsWS.Cells(15, 8).value

    ' Create Outlook email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Build HTML email body
    EBody = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"

    If ((PC_Eng <> "0" And PC_Eng <> "N/A") And (ME_Eng <> "0" And ME_Eng <> "N/A")) Then
        EBody = EBody & PC_Eng & " and " & ME_Eng & ",<BR><BR>Please process the attached CHANGE order notification.<BR><BR>"
        CcList = "joey.purdon@coperion.com"
    ElseIf (ME_Eng = "0" Or ME_Eng = "N/A") Then
        EBody = EBody & PC_Eng & ",<BR><BR>Please process the attached CHANGE order notification.<BR><BR>"
        CcList = "joey.purdon@coperion.com"
    ElseIf (PC_Eng = "0" Or PC_Eng = "N/A") Then
        EBody = EBody & ME_Eng & ",<BR><BR>Please process the attached CHANGE order notification.<BR><BR>"
        CcList = ""
    End If

    ' Add link to order text
    EBody = EBody & "Use the following link to review the latest order text:"
    EBody = EBody & "<A HREF='" & URL_Text & "'><BR>"
    EBody = EBody & Ord & "_" & CustName & "</A><BR><BR>"

    ' Set up email recipients and display
    With OutMail
        .To = ME_EMail & "; " & PC_EMail
        .CC = CcList
        .BCC = ""
        .Subject = "Change Order for " & Ord
        .Display
    End With

    ' --- ATTACH SCHEDULING PDFs ---
    ' Same pattern as SendNew_Email: find, attach, move to Processed
    LookupPath = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\" & Ord & "*"
    AttachFile = Dir(LookupPath, vbDirectory)

    While AttachFile <> ""
        AttachPath = "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\" & AttachFile
        OutMail.Attachments.Add (AttachPath)

        If Left(AttachFile, 4) = "1100" Then
            On Error GoTo CreateFolder1
            Name AttachPath As "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\Processed\" & Left(AttachFile, 7) & "xxx\" & AttachFile
            On Error GoTo 0
        Else
            Name AttachPath As "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\Processed\INTERNAL_Orders\" & AttachFile
        End If

        AttachFile = Dir(LookupPath, vbDirectory)
    Wend

    ' --- ATTACH INCOMING DESIGN SHEETS ---
    ' Look for design sheets matching this order number
    LookupPath = "\\USWWQ-P-FS01.cps.local\orders\Common Files\Documents\Incoming Design Sheets\" & Ord & "*"
    FoundDesignSheet = Dir(LookupPath, vbDirectory)

    ' If design sheets exist, add a red note to the email telling ME to upload to Windchill
    If FoundDesignSheet <> "" Then
        EBody = EBody & "<span style=""color:#FF0000"">" & ME_Eng & _
                ",<BR>Please place the attached design sheet(s) into Windchill." & _
                "</span style=""color:#FF0000"">"
    End If

    ' Attach each design sheet and then DELETE it from the incoming folder.
    ' WARNING: If the email is discarded, these files are lost permanently.
    ' Consider copying instead of deleting, or moving to a "Processed" folder.
    While FoundDesignSheet <> ""
        DesignSheetPath = "\\USWWQ-P-FS01.cps.local\orders\Common Files\Documents\Incoming Design Sheets\" & FoundDesignSheet
        With OutMail
            .Attachments.Add (DesignSheetPath)
        End With

        ' Remove readonly attribute (if set) and delete the file
        SetAttr DesignSheetPath, vbNormal
        Kill DesignSheetPath

        FoundDesignSheet = Dir(LookupPath, vbDirectory)
    Wend

    ' Capture signature and prepend our email body
    Dim Signature: Signature = OutMail.HTMLBody
    With OutMail
        .HTMLBody = "<html><body>" & EBody & Signature & "</body></html>"
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

    EmailsWS.Range("B3:Q5").ClearContents

    Exit Sub

CreateFolder1:
    ' Error handler: Move file to INTERNAL_Orders if the order-specific
    ' Processed subfolder can't be created
    Name AttachPath As "\\USWWQ-P-FS01.cps.local\orders\Common Files\CPE_Schedule\Processed\INTERNAL_Orders\" & AttachFile
    Resume Next

End Sub


'===============================================================================
' SUB: WhoHasOrder
'===============================================================================
' PURPOSE: Looks up an order's current engineering assignment, status, and
'          engineering comments. Populates the EmailsWS sheet with:
'            - Order details from Prod_Eng
'            - EE and ME engineer comments from cpe_schedule
'            - Line item descriptions
'            - [J-DRIVE] Hyperlinks to order documents
'
' PROCESS:
'   1. Queries Prod_Eng for order details (assignment, status, dates)
'   2. Copies results from a staging area into the visible Emails sheet
'   3. Queries cpe_schedule for EE engineer comments (date-stamped notes)
'   4. Queries cpe_schedule for ME engineer comments
'   5. Formats comments with date stamps: [MM/DD] comment text
'   6. [J-DRIVE] Looks up order folder to build links to order text and review
'
' BAD PRACTICES:
'   - Massive block of cell-by-cell value copying (lines ~881-896) instead of
'     a single Range.Copy or array assignment
'   - Comments loop uses hardcoded column indices (5 and 6) for month/day
'     formatting, but these columns aren't in the query output. This suggests
'     there's a formula or format on CheckWS that extracts month/day from
'     column 1 (datestamp). If that breaks, the date display breaks silently.
'   - AssyStr is reused here as a comment accumulator, which is confusing
'     since it's named for "Assembly String" in NewOrder2NewDB.
'===============================================================================
Sub WhoHasOrder()
    Dim E_eng As String, M_eng As String, EngComments As String
    Dim commentIdx As Long, lineIdx As Long

    Ord = EmailsWS.Cells(3, 2).value

    ' Open database connection
    Dim oCon As ADODB.Connection: Set oCon = New ADODB.Connection
    oCon.ConnectionString = databaseConnectionString
    oCon.Open

    ' Query Prod_Eng for this order's engineering assignment and status
    ' Returns: Order_Num, Customer_Name, Into_Eng, PC1, ME1, PC_Status,
    '          ME_Status, DocTypes, Apps_Out_0, PC_Apps_Out, ME_Apps_Out,
    '          PC_Apps_Back, ME_Apps_Back, MultiPlnt, PC_Act_Rel, ME_Act_Rel,
    '          Line_Num, Description
    Dim oRS As ADODB.Recordset: Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT Prod_Eng.Order_Num, Customer_Name, Into_Eng, PC1, ME1, " & _
                 "PC_Status, ME_Status, DocTypes, Apps_Out_0, PC_Apps_Out, ME_Apps_Out, " & _
                 "PC_Apps_Back, ME_Apps_Back, MultiPlnt, PC_Act_Rel, ME_Act_Rel, " & _
                 "Line_Num, Description FROM Prod_Eng " & _
                 "WHERE Prod_Eng.Order_Num=" & Ord & " ORDER BY Prod_Eng.Line_Num;"
    oRS.Open

    ' Dump query results to staging columns (AA onward), then copy to visible area
    EmailsWS.Select
    EmailsWS.Range("AA3:AS100").ClearContents
    EmailsWS.Range("AA3").CopyFromRecordset oRS

    CheckWS.Range("A1:B500").ClearContents

    ' Copy first row of results to the visible area (columns C through Q, row 3)
    ' This maps the 15 query columns to their display positions.
    ' Column AA (27) = Order_Num -> Col C (3), AB (28) = Customer_Name -> Col D (4), etc.
    EmailsWS.Cells(3, 3).value = EmailsWS.Cells(3, 28).value     ' Customer_Name
    EmailsWS.Cells(3, 4).value = EmailsWS.Cells(3, 29).value     ' Into_Eng
    EmailsWS.Cells(3, 5).value = EmailsWS.Cells(3, 30).value     ' PC1
    EmailsWS.Cells(3, 6).value = EmailsWS.Cells(3, 31).value     ' ME1
    EmailsWS.Cells(3, 7).value = EmailsWS.Cells(3, 32).value     ' PC_Status
    EmailsWS.Cells(3, 8).value = EmailsWS.Cells(3, 33).value     ' ME_Status
    EmailsWS.Cells(3, 9).value = EmailsWS.Cells(3, 34).value     ' DocTypes
    EmailsWS.Cells(3, 10).value = EmailsWS.Cells(3, 35).value    ' Apps_Out_0
    EmailsWS.Cells(3, 11).value = EmailsWS.Cells(3, 36).value    ' PC_Apps_Out
    EmailsWS.Cells(3, 12).value = EmailsWS.Cells(3, 37).value    ' ME_Apps_Out
    EmailsWS.Cells(3, 13).value = EmailsWS.Cells(3, 38).value    ' PC_Apps_Back
    EmailsWS.Cells(3, 14).value = EmailsWS.Cells(3, 39).value    ' ME_Apps_Back
    EmailsWS.Cells(3, 15).value = EmailsWS.Cells(3, 40).value    ' MultiPlnt
    EmailsWS.Cells(3, 16).value = EmailsWS.Cells(3, 41).value    ' PC_Act_Rel
    EmailsWS.Cells(3, 17).value = EmailsWS.Cells(3, 42).value    ' ME_Act_Rel

    E_eng = EmailsWS.Cells(3, 5).value  ' EE Engineer name
    M_eng = EmailsWS.Cells(3, 6).value  ' ME Engineer name

    ' --- RETRIEVE EE ENGINEER COMMENTS ---
    ' Query cpe_schedule for date-stamped comments from the EE engineer
    If E_eng <> "N/A" Then
        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "SELECT cpe_schedule.datestamp, comments FROM cpe_schedule " & _
                     "WHERE cpe_schedule.engineer='" & E_eng & "' AND comments>' ' " & _
                     "AND orderno=" & Ord & " ORDER BY cpe_schedule.datestamp;"
        oRS.Open

        CheckWS.Range("A1:B202").ClearContents
        CheckWS.Range("A3").CopyFromRecordset oRS
    End If

    ' --- RETRIEVE ME ENGINEER COMMENTS ---
    ' Same query but for the ME engineer, placed in rows 203+ to avoid overlap
    If M_eng <> "N/A" Then
        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "SELECT cpe_Schedule.datestamp, comments FROM cpe_schedule " & _
                     "WHERE cpe_schedule.engineer='" & M_eng & "' AND comments>' ' " & _
                     "AND orderno=" & Ord & " ORDER BY cpe_schedule.datestamp;"
        oRS.Open

        CheckWS.Range("A203:B500").ClearContents
        CheckWS.Range("A203").CopyFromRecordset oRS

    End If

    oRS.Close
    oCon.Close

    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing

    ' --- FORMAT EE COMMENTS ---
    ' Build a newline-delimited string of date-stamped comments.
    ' NOTE: Columns 5 and 6 on CheckWS appear to extract Month and Day from
    ' the datestamp in column 1 via worksheet formulas. If those formulas are
    ' missing or broken, the [/] date prefix will be empty.
    EngComments = ""
    For commentIdx = 3 To 202
        If CheckWS.Cells(commentIdx, 2).value <> "" Then
            EngComments = EngComments & Chr(10) & "[" & CheckWS.Cells(commentIdx, 5).value & "/" & _
                         CheckWS.Cells(commentIdx, 6).value & "] " & CheckWS.Cells(commentIdx, 2).value
        Else
            Exit For
        End If
    Next commentIdx
    EmailsWS.Cells(5, 7).value = EngComments  ' EE Comments

    ' --- FORMAT ME COMMENTS ---
    EngComments = ""
    For commentIdx = 203 To 500
        If CheckWS.Cells(commentIdx, 2).value <> "" Then
            EngComments = EngComments & Chr(10) & "[" & CheckWS.Cells(commentIdx, 5).value & "/" & _
                         CheckWS.Cells(commentIdx, 6).value & "] " & CheckWS.Cells(commentIdx, 2).value
        Else
            Exit For
        End If
    Next commentIdx
    EmailsWS.Cells(5, 8).value = EngComments  ' ME Comments

    ' --- FORMAT LINE ITEM DESCRIPTIONS ---
    ' Build a string of [Line_Num] Description for all lines in the order
    ' NOTE: AssyStr is reused here as a string accumulator. In NewOrder2NewDB it's
    ' used for SQL VALUES strings. This dual-purpose usage is confusing.
    EngComments = ""
    Dim DescStr As String: DescStr = ""
    For lineIdx = 3 To 50
        If EmailsWS.Cells(lineIdx, 43).value <> "" Then
            DescStr = DescStr & Chr(10) & "[" & EmailsWS.Cells(lineIdx, 43).value & "] " & _
                     EmailsWS.Cells(lineIdx, 44).value
        Else
            Exit For
        End If
    Next lineIdx
    EmailsWS.Cells(5, 3).value = DescStr
    DescStr = ""

    ' --- ORDER DOCUMENT LINKS ---
    ' Order Reviews are now stored in Temporary_Order_Files.
    ' Order Text PDFs are managed in Windchill (no direct file-share path available).
    OrdStr = Trim$(Str(Ord))
    CustName = EmailsWS.Cells(3, 3).value

    Dim ordNumW As String
    ordNumW = Trim(OrdStr)

    ' Order Text: now in Windchill - no direct path; leave blank for manual lookup
    Dim windchill_obj As WindchillObject: Set windchill_obj = New WindchillObject
    URL_Text = windchill_obj.GetFolderUrlByOrderNumber(OrdStr)
    EmailsWS.Cells(12, 7).value = ""
    EmailsWS.Cells(13, 7).value = Ord & "_" & CustName
    EmailsWS.Cells(13, 8).value = URL_Text

    ' Order Review: located in Temporary_Order_Files EE subfolder
    URL_Review = TEMP_ORDER_BASE & ordNumW & "\" & ordNumW & "\Order_Review.pdf"

    EmailsWS.Cells(15, 8).value = URL_Review

    ' Default unassigned engineers to "N/A" for downstream email logic
    If EmailsWS.Cells(3, 5).value = "" Then
        EmailsWS.Cells(3, 5).value = "N/A"
    End If
    If EmailsWS.Cells(3, 6).value = "" Then
        EmailsWS.Cells(3, 6).value = "N/A"
    End If

End Sub




'===============================================================================
' SUB: CreateZipFile
'===============================================================================
' PURPOSE: Creates a .zip file from a folder using the Windows Shell.
'          Used to zip the order transmittal folder for Windchill upload.
'
' PARAMETERS:
'   folderToZipPath  - Full path to the folder to compress
'   zippedFileFullName - Full path for the output .zip file
'
' PROCESS:
'   1. Creates an empty .zip file with the correct header bytes
'   2. Uses Shell.Application to copy the folder contents into the zip
'   3. Waits (polling every 1 second) until the zip item count matches the source
'
' BAD PRACTICES:
'   - On Error Resume Next during the wait loop silently eats errors if
'     the zipping process fails.
'   - No timeout on the wait loop; if zipping fails, this loops forever.
'   - The empty zip header is a magic byte sequence. This is a well-known
'     technique but should at least be documented.
'
' TODO: This entire process should eventually be replaced with direct
'       Windchill API upload, eliminating the need for intermediate zip files.
'===============================================================================
Sub CreateZipFile(folderToZipPath As Variant, zippedFileFullName As Variant)

    Dim ShellApp As Object

    ' Create an empty zip file with the minimum valid zip header.
    ' Bytes: PK (50 4B) + version needed (05 06) + 18 null bytes
    ' This is the "end of central directory record" for an empty zip.
    Open zippedFileFullName For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1

    ' Use the Windows Shell to copy files into the zip
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).items

    ' Wait for the zipping to complete by comparing item counts.
    ' The On Error Resume Next is here because the zip namespace may not be
    ' immediately readable while Windows is still writing to it.
    On Error Resume Next
    Do Until ShellApp.Namespace(zippedFileFullName).items.Count = ShellApp.Namespace(folderToZipPath).items.Count
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    On Error GoTo 0

End Sub


'===============================================================================
' SUB: UpdateAllToCurrentYear
'===============================================================================
' PURPOSE: Bulk updates the Active_Year field on all unreleased orders in
'          Prod_Eng to the current calendar year. This is typically run at
'          the start of a new year to carry forward active work.
'
' LOGIC: Updates where either PC or ME has a non-empty, non-RELEASED status.
'        This catches all orders that are still being worked on.
'
' BAD PRACTICES:
'   - No error handling if the DB connection fails.
'   - No confirmation prompt before bulk-updating potentially hundreds of records.
'===============================================================================
Sub UpdateAllToCurrentYear()

    ActiveYr = Year(Date)

    ' Open database connection
    Dim oCon As ADODB.Connection: Set oCon = New ADODB.Connection
    oCon.ConnectionString = databaseConnectionString
    oCon.Open

    oCon.Execute "UPDATE Prod_Eng SET Prod_Eng.Active_Year='" & ActiveYr & "' " & _
                 "WHERE ((Prod_Eng.ME_Status<>'RELEASED' AND ME_Status<>'' AND ME_Status IS NOT NULL) " & _
                 "OR (Prod_Eng.PC_Status<>'RELEASED' AND PC_Status<>'' AND PC_Status IS NOT NULL))"

    If Not oCon Is Nothing Then Set oCon = Nothing

    MsgBox "All Active Orders have been updated to " & ActiveYr, Title:="Update Complete!"

End Sub

