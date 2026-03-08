Attribute VB_Name = "Module1"
Option Explicit
' Scheduling Tool
' Table added to CPE Database Version 1.0 - CPE Manager
' Original - Version 1.0 - Dated 2/7/2012
' Revision 1.5 - Dated 7/23/2012
'Revision 2.0 - Dated 7/10/2015 - Moved databasse to new server

Dim DEBUG_FLAG As Boolean
Const CONNECTION_STRING As String = "Driver={SQL Server};Server=USLXA-P-SQL01.cps.local;Database=PAC1CPE;UID=dash;Pwd=manage_DB;"

'@EntryPoint
Sub Button1_Click()
    
    '===================
    'DEBUG MODE
    '!!!!!!!!!!!!!!!!!!!
    DEBUG_FLAG = False
    '!!!!!!!!!!!!!!!!!!!
    '===================
    
    ' Order Scheduling Tool
    ' Revised 2/7/2018 to add Multi-Plant field
    '
    Dim Dur As Long

    Dim TempStr As String
    Dim BuildHTML As String
    Dim Scheduler As String

    Dim OrdNum
    Dim RevNum
    Dim Dpath As String
    Dim Spath As String
    Dim File_Num As String
    Dim P_O As String

    Dim SchDate As Date
    Dim AssyStr As String
    Dim PreME As Date

    Dim FiltNotes As String
    Dim SchNote As String
    
    Dim oCon As ADODB.Connection: Set oCon = New ADODB.Connection
    Dim oRS As ADODB.Recordset
    
    oCon.ConnectionString = CONNECTION_STRING
    oCon.Open
    Set oRS = oCon.Execute("SELECT WW_Initials from WW_Prod_Eng_Names WHERE WW_Name = '" & Environ$("Username") & "'")

    Dpath = "R:\Common Files\CPE_Schedule\"
    'Spath = "\\uslxa-p-web01.cps.local\Intranet\Database\CPE\"
    Spath = "R:\Common Files\CPE\Temporary_Order_Files\Staging\"
    
    SchDate = SchedulingWS.Cells(3, 5).Value
    OrdNum = SchedulingWS.Cells(5, 5).Value
    RevNum = SchedulingWS.Cells(5, 7).Value
    File_Num = SchedulingWS.Cells(5, 11).Value
    P_O = SchedulingWS.Cells(7, 11).Value
    Dur = SchedulingWS.Cells(15, 11).Value
    PreME = SchedulingWS.Cells(21, 5).Value
    
    ' Error handling for setting Scheduler
    On Error Resume Next
    Scheduler = oRS.Fields().Item(0).Value
    On Error GoTo 0 ' Reset error handling
 
    TempStr = Spath & "Blank_Sched.xls"
    
    Dim blank_wb As Workbook: Set blank_wb = Application.Workbooks.Open(Filename:=TempStr)
    Dim blank_ws As Worksheet: Set blank_ws = blank_wb.Worksheets("Blank")

    
    ' Write the Date, Order Number and Revision Number
    blank_ws.Cells.Item(3, 5).Value = SchDate
    blank_ws.Cells(5, 5).Value = OrdNum
    blank_ws.Cells(5, 7).Value = RevNum
    blank_ws.Cells(5, 11).Value = File_Num
    blank_ws.Cells(7, 11).Value = P_O
    blank_ws.Cells(15, 11).Value = Dur

    blank_ws.Cells(26, 7).Value = SchedulingWS.[E_Kronos]
    blank_ws.Cells(30, 7).Value = SchedulingWS.[M_Kronos]
    
    If PreME > 5 Then
        blank_ws.Cells(21, 5).Value = PreME
    End If
    
    ' Write the Approval settings
    blank_ws.Cells(10, 2).Value = SchedulingWS.Cells(10, 2).Value
    blank_ws.Cells(12, 2).Value = SchedulingWS.Cells(12, 2).Value
    blank_ws.Cells(14, 2).Value = SchedulingWS.Cells(14, 2).Value
    
    ' Write the Certified settings
    blank_ws.Cells(10, 6).Value = SchedulingWS.Cells(10, 6).Value
    blank_ws.Cells(12, 6).Value = SchedulingWS.Cells(12, 6).Value
    blank_ws.Cells(14, 6).Value = SchedulingWS.Cells(14, 6).Value
    
    ' Write the Work Orders and Steps
    blank_ws.Cells(17, 5).Value = SchedulingWS.Cells(17, 5).Value
    blank_ws.Cells(19, 5).Value = SchedulingWS.Cells(19, 5).Value
    
    'blank_ws.Cells(21, 5).Value = SchedulingWS.Cells(21, 5).Value
    blank_ws.Cells(17, 8).Value = SchedulingWS.Cells(17, 8).Value
    blank_ws.Cells(19, 8).Value = SchedulingWS.Cells(19, 8).Value
    'blank_ws.Cells(21, 8).Value = SchedulingWS.Cells(21, 8).Value

    ' Write the Scheduled Dates
    blank_ws.Cells(24, 5).Value = SchedulingWS.Cells(24, 5).Value
    blank_ws.Cells(26, 5).Value = SchedulingWS.Cells(26, 5).Value
    blank_ws.Cells(28, 5).Value = SchedulingWS.Cells(28, 5).Value
    blank_ws.Cells(30, 5).Value = SchedulingWS.Cells(30, 5).Value
    blank_ws.Cells(32, 5).Value = SchedulingWS.Cells(32, 5).Value
    blank_ws.Cells(10, 11).Value = SchedulingWS.Cells(10, 11).Value
    
    ' Write the Comments
    blank_ws.Cells(33, 3).Value = Scheduler
    blank_ws.Cells(35, 2).Value = SchedulingWS.Cells(35, 2).Value

    '    Message = " Button Press Acknowledged!"
    '    MsgBox Message
    '--------------------------------------------------------------------------------
    TempStr = Dpath & OrdNum & "-" & RevNum & ".pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
                                    TempStr, Quality:= _
                                    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
                                    OpenAfterPublish:=False
    '-------------------------------------------------------------------------------
    Dim OutApp: Set OutApp = CreateObject("Outlook.Application")
    Dim OutMail: Set OutMail = OutApp.CreateItem(0)
        
    BuildHTML = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    BuildHTML = BuildHTML & "Order " & OrdNum & " Rev " & RevNum & " has been scheduled and a document has been saved to the R Drive." & "<BR>"
    With OutMail
        .To = "brian.schmoldt@coperion.com"
        .CC = "joey.purdon@coperion.com"
        .BCC = vbNullString
        .Subject = "Order " & OrdNum & " has been Scheduled"
        .HTMLBody = BuildHTML
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

    '-------------------------------------------------------------------------------
    'Save to CPE Database
    
    AssyStr = "'" & Date$ & "', "
    AssyStr = AssyStr & SchedulingWS.Cells(5, 5).Value & ", " ' orderno
    AssyStr = AssyStr & SchedulingWS.Cells(5, 7).Value & ", " 'revno
    
    'Approvals
    AssyStr = AssyStr & IIf(SchedulingWS.Cells(12, 2).Value = vbNullString, "0, ", "1, ")
    'Electrical Approvals
    AssyStr = AssyStr & IIf(SchedulingWS.Cells(14, 2).Value = vbNullString, "0, ", "1, ")
    'Mechanical Approvals
    AssyStr = AssyStr & IIf(SchedulingWS.Cells(10, 6).Value = vbNullString, "0, ", "1, ")
    'Certifieds
    AssyStr = AssyStr & IIf(SchedulingWS.Cells(10, 6).Value = vbNullString, "0, ", "1, ")
    'Electrical Certifieds
    AssyStr = AssyStr & IIf(SchedulingWS.Cells(12, 6).Value = vbNullString, "0, ", "1, ")
    'Mechanical Certifieds
    AssyStr = AssyStr & IIf(SchedulingWS.Cells(14, 6).Value = vbNullString, "0, '", "1, '")
    
    AssyStr = AssyStr & SchedulingWS.Cells(17, 5).Value & "', '" ' EE Network
    AssyStr = AssyStr & SchedulingWS.Cells(17, 8).Value & "', '" ' EE Activity
    AssyStr = AssyStr & SchedulingWS.Cells(19, 5).Value & "', '" ' ME Network
    AssyStr = AssyStr & SchedulingWS.Cells(19, 8).Value & "', '" ' ME Activity
    AssyStr = AssyStr & SchedulingWS.Cells(21, 5).Value & "', '" ' PE Network
    AssyStr = AssyStr & SchedulingWS.Cells(21, 8).Value & "', '" ' PE Activity
    AssyStr = AssyStr & SchedulingWS.Cells(24, 5).Value & "', '" ' Scheduled Date
    AssyStr = AssyStr & SchedulingWS.Cells(26, 5).Value & "', '" ' Approvals Date
    AssyStr = AssyStr & SchedulingWS.Cells(28, 5).Value & "', '" ' Release Date
    AssyStr = AssyStr & SchedulingWS.Cells(30, 5).Value & "', '" ' Production Date
    AssyStr = AssyStr & SchedulingWS.Cells(32, 5).Value & "', '" ' Ship Date
    
    SchNote = SchedulingWS.Cells(35, 2).Value
        
    ' Define characters to replace
    Dim charactersToReplace As Variant
    charactersToReplace = Array("'", """", vbCrLf)
    SchNote = MultiReplace(SchNote, charactersToReplace, "_")
    
    AssyStr = AssyStr & Mid$(FiltNotes, 1, 254) & "', '" ' Notes
    AssyStr = AssyStr & SchedulingWS.Cells(5, 11).Value & "', '" ' File Number
    AssyStr = AssyStr & SchedulingWS.Cells(7, 11).Value & "', " ' Customer PO
    AssyStr = AssyStr & SchedulingWS.Cells(15, 11).Value & ", '" ' Duration
    AssyStr = AssyStr & SchedulingWS.Cells(10, 11).Value & "', '" ' MultiPlant
    AssyStr = AssyStr & Scheduler & "'"
    
    Set oRS = oCon.Execute("INSERT INTO cpe_scheduling(uniqid, datestamp, orderno, revno, approvals, ee_app, me_app, certifieds, ee_certs, " & _
                    "me_certs, ee_netact, ee_step, me_netact, me_step, pe_netact, pe_step, scheddate, appdate, reldate, proddate, shipdate, comments, " & _
                    "file_num, po_num, duration, MultiPlnt, scheduler) VALUES(newid(), " & AssyStr & ")")

    
    If Not DEBUG_FLAG Then clear_worksheet

    '-------------------------------------------------------------------------------
    Workbooks("Blank_Sched.xls").Close SaveChanges:=False

End Sub

Sub CopyRangeValues(sourceSheet As Worksheet, targetSheet As Worksheet, rangeAddress As String)
    targetSheet.Range(rangeAddress).Value = sourceSheet.Range(rangeAddress).Value
End Sub

Sub clear_worksheet()

    'Clear worksheets
    SchedulingWS.Cells(5, 5).Value = vbNullString
    SchedulingWS.Cells(5, 7).Value = vbNullString
    SchedulingWS.Cells(5, 11).Value = vbNullString
    
    SchedulingWS.Cells(7, 11).Value = vbNullString
    SchedulingWS.Cells(10, 2).Value = vbNullString
    SchedulingWS.Cells(12, 2).Value = vbNullString
    SchedulingWS.Cells(14, 2).Value = vbNullString
    
    ThisWorkbook.Worksheets("Scheduling").[E_Kronos] = vbNullString
    ThisWorkbook.Worksheets("Scheduling").[M_Kronos] = vbNullString
    
    ' Write the Certified settings
    SchedulingWS.Cells(10, 6).Value = vbNullString
    SchedulingWS.Cells(12, 6).Value = vbNullString
    SchedulingWS.Cells(14, 6).Value = vbNullString
    
    ' Write the Work Orders and Steps
    SchedulingWS.Cells(17, 5).Value = vbNullString
    SchedulingWS.Cells(19, 5).Value = vbNullString
    SchedulingWS.Cells(21, 5).Value = vbNullString
    SchedulingWS.Cells(17, 8).Value = vbNullString
    SchedulingWS.Cells(19, 8).Value = vbNullString
    SchedulingWS.Cells(21, 8).Value = vbNullString

    ' Write the Scheduled Dates
    SchedulingWS.Cells(24, 5).Value = vbNullString
    SchedulingWS.Cells(26, 5).Value = vbNullString
    SchedulingWS.Cells(28, 5).Value = vbNullString
    SchedulingWS.Cells(30, 5).Value = vbNullString
    SchedulingWS.Cells(32, 5).Value = vbNullString
    
    ' Write the Comments
    SchedulingWS.Cells(35, 2).Value = vbNullString
    
    ' Clear Multi Plant
    SchedulingWS.Cells(10, 11).Value = vbNullString

End Sub


Function MultiReplace(originalText As String, charactersToReplace As Variant, replacementCharacter As String) As String
    Dim characterIndex As Integer
    
    For characterIndex = LBound(charactersToReplace) To UBound(charactersToReplace)
        originalText = Replace(originalText, charactersToReplace(characterIndex), replacementCharacter)
    Next characterIndex
    
    MultiReplace = originalText
End Function


