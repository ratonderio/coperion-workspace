Attribute VB_Name = "Emails"
Option Explicit
'@IgnoreModule SetAssignmentWithIncompatibleObjectType, ProcedureNotUsed, MemberNotOnInterface
'@Folder "Modules"
Dim HTMLString As String
Dim SelectedOption As String
Dim EngineeringType As String
Dim SplitArray() As String
Dim Debg As Boolean
Dim ProjResponsible As String
Dim AEIndex As Long

Sub SendApprovalTest()
    Dim hasReminderText As Boolean: hasReminderText = False
    Dim lastUpdate As Date: lastUpdate = Date
    Dim orderNumber As Long: orderNumber = 1100102179
        
    SendApprovalEmail hasReminderText, orderNumber:=orderNumber
End Sub


Sub SendApprovalEmail(Optional ByVal hasReminderText As Boolean = False, Optional ByVal lastUpdate As Date = 0, Optional ByVal orderNumber As Long = 0)
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'If order number not 0
    If orderNumber <> 0 Then Ord = orderNumber
    
    'Get PO Number
    Set oCon = New ADODB.connection: oCon.ConnectionString = databaseConnectionString: oCon.Open
    
    ' Retrieve records
    Set oRs = oCon.Execute("SELECT DISTINCT po_num FROM cpe_scheduling WHERE cpe_scheduling.orderno=" & Ord)
    
    'Copy recordset to Schedule Worksheet AM1
    ScheduleWS.Range("AM1").CopyFromRecordset oRs
    
    'Close recordset
    oRs.Close
    
    If hasReminderText Then
        If ScheduleWS.Range("AJ2").Value = "PC" Then
            oCon.Execute ("UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Out='" & Date & "'WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.PC1 = '" & ProdEng & "'")
        Else
            oCon.Execute ("UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Out='" & Date & "'WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.ME1 = '" & ProdEng & "'")
        End If
    End If
    
    PONum = ScheduleWS.Range("AM1").Value
    
    ScheduleWS.Range("AM1").Value = vbNullString
    ScheduleWS.Range("AM1:AM10000").ClearContents

    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    
    'Eng Type
    If UCase$(ScheduleWS.Range("AJ2").Value) = "PC" Then
        EngType = "Electrical"
        OthEngType = "Mechanical"
    Else
        EngType = "Mechanical"
        OthEngType = "Electrical"
    End If
    
    'Other engineer's email, either PC or ME
    Dim OtherEngEmail As String

    
    'Create new connection and open
    Set oCon = New ADODB.connection: oCon.ConnectionString = databaseConnectionString: oCon.Open
    
    'Set recordset to execution of query using either PC1 or ME1
    Set oRs = IIf(OthEngType = "Electrical", _
                   oCon.Execute("SELECT WW_Email FROM WW_Prod_Eng_Names WHERE WW_Initials IN (SELECT DISTINCT PC1 FROM Prod_Eng WHERE Order_Num = " & Ord & ")"), _
                   oCon.Execute("SELECT WW_Email FROM WW_Prod_Eng_Names WHERE WW_Initials IN (SELECT DISTINCT ME1 FROM Prod_Eng WHERE Order_Num = " & Ord & ")"))
    
    'Copy recordset results to BD4 on ScheduleWS
    ScheduleWS.Range("BD4").CopyFromRecordset oRs
    oRs.Close
    
    'Get last row with a value in column "BD" on schedulews
    Dim last_email_row As Long: last_email_row = get_last_row_in_column("BD", ScheduleWS)
    
    'Iterate through the starting row "4" to the last discovered row
    Dim email_iter As Long
    For email_iter = 4 To last_email_row
        'If it is the first iteration, assign the email directly, else append
        OtherEngEmail = IIf(email_iter = 4, ScheduleWS.Range("BD" & email_iter), OtherEngEmail & ";" & ScheduleWS.Range("BD" & email_iter))
    Next email_iter
    
    'Clear out recordset results
    ScheduleWS.Range("BC4:BD100").ClearContents
    
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    
    'Greetings
    BuildHTML = "<strong><span style='color: red;'>ATTACH CUSTOMER TRANSMITTAL, DRAWINGS PDF, &amp; DRAWING FILES AS A ZIP FILE AND DELETE THIS LINE.</span></strong><br>"
    BuildHTML = BuildHTML & "<b>K-Tron Equipment Order:</b> " & Ord & "<BR>"
    BuildHTML = BuildHTML & "<b>Ship To:</b> " & CustName & "<BR>"
    BuildHTML = BuildHTML & "<b>Current Status:</b> Out for Approval Hold<BR><BR>"

    'Greeting
    BuildHTML = BuildHTML & "Hello,<BR><BR>"

    If hasReminderText Then
        BuildHTML = BuildHTML & "I am following up on the drawing approvals I last sent on " & lastUpdate & ".  " & _
                                    "Please let me know if there are any additional questions or concerns.<br><br><hr><br>"
    End If

    'Message Text
    BuildHTML = BuildHTML & "The " & LCase$(EngType) & " approval documents and drawings for your equipment order are ready and attached to this email.<BR>"
    BuildHTML = BuildHTML & "Please review them and reply with any required changes. (See markup guidelines below.)<BR>"
    BuildHTML = BuildHTML & "Please note that this email only covers your " & LCase$(EngType) & " drawings. You will receive a separate email for your " & _
                                LCase$(OthEngType) & " approval documents.<BR><BR>"
    BuildHTML = BuildHTML & "<b>Project Status &amp; Next Steps</b>"
    BuildHTML = BuildHTML & "<ul>"
    BuildHTML = BuildHTML & "<li><b>Current Hold:</b> This order is currently on an <b>Out for Approval Hold</b> and is not yet scheduled for shipment. " & _
                                "K-Tron will assign an official ship date once the approval process is complete.</li>"
    BuildHTML = BuildHTML & "<li><b>Change Policy:</b> Any modifications to equipment design or configuration during this review may impact your final " & _
                                "ship date and may result in change order fees.</li>"
    BuildHTML = BuildHTML & "<li><b>How to Submit Changes:</b> We prefer mark-ups on a PDF or scanned file. However, we also accept direct electronic edits " & _
                                "as long as every change is clearly highlighted using revision clouds, distinct markings, and notes.</li>"
    BuildHTML = BuildHTML & "</ul>"
    BuildHTML = BuildHTML & "If you have questions, comments, or concerns; please do not hesitate to contact me."
    'Display email
    OutMail.Display
    
    'Grab the signature from the default email generation
    Dim Signature As String: Signature = OutMail.HTMLBody
    
    'Fill out email and fields
    On Error Resume Next
    With OutMail
        .To = "Look up who to send it to via SAP"
        .CC = ENGINEERING_CLERK_EMAIL & "; brian.schmoldt@ktron.com; " & OtherEngEmail
        .BCC = vbNullString
        .Subject = CustName & " PO # " & PONum & " / Coperion Order # " & Ord & " : " & EngType & " Approval Documents"
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body>" & BuildHTML & "<br>" & Signature & "</body></html>"
        .Attachments.Add ("\\uswwq-p-fs01\orders\Common Files\CPE_Schedule\Coperion Food and Performance Materials (FPM) is migrating to Coperion.pdf")
    End With
    
    On Error GoTo 0
    
    'Find person to email through SAP
    FindEmailToPerson
    
    'Clear objects, probably doesn't need to be done
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    'If this is a reminder email
    If hasReminderText Then Order_Query
    
End Sub

Sub SendCertifiedTest()

    SendCertifiedEmail "1100102179"

End Sub


Sub SendCertifiedEmail(ByVal order_number As String)
    Dim Debg As Boolean: Debg = False
    Ord = order_number
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'Get PO Number
    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString
    oCon.Open
    
    ' Database Query
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    oRs.Source = "SELECT DISTINCT po_num FROM cpe_scheduling WHERE cpe_scheduling.orderno=" & order_number & ";"
    
    ' Retrieve records
    oRs.Open
    'Stop
    ScheduleWS.Range("AM1").CopyFromRecordset oRs
    PONum = ScheduleWS.Range("AM1").Value

    ScheduleWS.Range("AM1:AM50000").ClearContents

    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    
    'Eng Type
    EngType = IIf(UCase$(ScheduleWS.Range("AJ2").Value) = "PC", "Electrical", "Mechanical")
    
    'Other engineer's email, either PC or ME
    Dim OtherEngEmail As String
    
    'Create new connection and open
    Set oCon = New ADODB.connection: oCon.ConnectionString = databaseConnectionString: oCon.Open
    
    'Set recordset to execution of query using either PC1 or ME1
    Set oRs = IIf(OthEngType = "Electrical", _
                   oCon.Execute("SELECT WW_Email FROM WW_Prod_Eng_Names WHERE WW_Initials IN (SELECT DISTINCT PC1 FROM Prod_Eng WHERE Order_Num = " & Ord & ")"), _
                   oCon.Execute("SELECT WW_Email FROM WW_Prod_Eng_Names WHERE WW_Initials IN (SELECT DISTINCT ME1 FROM Prod_Eng WHERE Order_Num = " & Ord & ")"))
    'Copy recordset results to BD4 on ScheduleWS
    ScheduleWS.Range("BD4").CopyFromRecordset oRs
    oRs.Close
    
    'Get last row with a value in column "BD" on schedulews
    Dim last_email_row As Long: last_email_row = get_last_row_in_column("BD", ScheduleWS)
    
    'Iterate through the starting row "4" to the last discovered row
    Dim email_iter As Long
    For email_iter = 4 To last_email_row
        'If it is the first iteration, assign the email directly, else append
        OtherEngEmail = IIf(email_iter = 4, ScheduleWS.Range("BD" & email_iter), OtherEngEmail & ";" & ScheduleWS.Range("BD" & email_iter))
    Next email_iter
    
    'Clear out recordset results
    ScheduleWS.Range("BC4:BD100").ClearContents

    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    
    'Order Info
    BuildHTML = vbNullString
    
    BuildHTML = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    BuildHTML = BuildHTML & "<strong><span style='color: red;'>ATTACH CUSTOMER TRANSMITTAL, DRAWINGS PDF, & DRAWING FILES AS A ZIP FILE AND DELETE THIS LINE.</span></strong><br>"
    BuildHTML = BuildHTML & "<b> Coperion Sales Order: </b>" & vbTab & order_number & "<BR>"
    BuildHTML = BuildHTML & "<b> Ship To: </b>" & vbTab & CustName & "<BR>"
    BuildHTML = BuildHTML & "<b> Current Status: </b>" & vbTab & "Released to Production<BR><BR>"
    
    'Greeting
    BuildHTML = BuildHTML & "Hello,<BR><BR>"

    'Initial Message
    BuildHTML = BuildHTML & "Please find your Coperion " & LCase$(EngType) & " certified drawings attached for the subject job. No response is required.<BR><BR>"
    BuildHTML = BuildHTML & "Please note that any changes to the design or configuration from this point forward will result in a revised ship date, " & _
                                "as well as Change Order fees to cover additional labor, materials, and shipping delays.<BR><BR>"

    'HyperLink
    Dim TextCode As Double
    TextCode = 2
    For CodeIndex = 5 To 10
        TextCode = TextCode + Mid$(order_number, CodeIndex, 1)
    Next CodeIndex

    'Remaining Message
    BuildHTML = BuildHTML & "You will receive a separate email confirming your scheduled ship date. If you have any questions or concerns regarding your ship date, please contact:<BR>"
    BuildHTML = BuildHTML & "&bull; Michelle Young: <a href='mailto:michelle.young@ktron.com'>michelle.young@ktron.com</a><BR>"
    BuildHTML = BuildHTML & "&bull; Staci Voboril: <a href='mailto:staci.voboril@ktron.com'>staci.voboril@ktron.com</a><BR>"
    BuildHTML = BuildHTML & "&bull; Customer Service: 800-821-2476"
        
    
    With OutMail
        .Display
    End With
    
    Dim Signature As String
    Signature = OutMail.HTMLBody
      
    On Error Resume Next
    With OutMail
        .To = "Look up who to send it to via SAP"
        .CC = ENGINEERING_CLERK_EMAIL & "; brian.schmoldt@ktron.com; " & OtherEngEmail
        .BCC = vbNullString
        .Subject = CustName & " PO # " & PONum & " / Coperion Order # " & order_number & " : " & EngType & " Certified Documents"
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body>" & BuildHTML & "<br>" & Signature & "</body></html>"
        .Attachments.Add ("\\uswwq-p-fs01\orders\Common Files\CPE_Schedule\Coperion Food and Performance Materials (FPM) is migrating to Coperion.pdf")
    End With
    On Error GoTo 0
    
    If (Debg = True) Then Exit Sub
    
    Call FindEmailToPerson
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub

Sub SendProductionEmail(prodScheduler As String)
    Dim manufacturingEngineering As String: manufacturingEngineering = "matt.phillips@ktron.com"
    Dim manufacturing_engineering_backup As String: manufacturing_engineering_backup = "carl.johnson@ktron.com"
    Dim manufacturingEngineeringName As String: manufacturingEngineeringName = "Matt"
    Dim manufacturing_engineering_backup_name As String: manufacturing_engineering_backup_name = "Carl"
    
    'TODO: This was commented out because the emails only go to Matt now, but that might not be what we want now

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'Eng Type
    If UCase$(ScheduleWS.Range("AJ2").Value) = "PC" Then
        EngType = "Electrical"
    Else
        EngType = "Mechanical"
    End If
    
    'Order Info
    BuildHTML = vbNullString
    BuildHTML = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    BuildHTML = BuildHTML & "<b> Order #" & Ord & "</b><BR>"
    BuildHTML = BuildHTML & "<b> Customer: " & CustName & "</b><BR><BR>"
    
    
    'Initial Message
    BuildHTML = BuildHTML & "Hello " & manufacturingEngineeringName & ",<BR><BR>"
    BuildHTML = BuildHTML & "Please process the attached release for order #" & Ord & ": " & CustName & ".<BR>"
    BuildHTML = BuildHTML & "Thank You!"
    
    'Get Signature
    OutMail.Display
    Dim Signature As String
    Signature = OutMail.HTMLBody
    
    On Error Resume Next
    
    With OutMail
        .To = manufacturingEngineering
        .CC = "brian.schmoldt@ktron.com"
        .Subject = "Production Release for Order #" & Ord & ": " & CustName
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body>" & BuildHTML & Signature & "</body></html>"
    End With
    
    On Error GoTo 0
    
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Public Sub ChangeOrderEmail()

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'Get Params
    EngineeringType = IIf(UCase$(ScheduleWS.Range("AJ2").Value) = "PC", "ELECTRICAL ", "MECHANICAL ")
    
    'Build HTML String for email
    HTMLString = "<b>Coperion Sales Order: " & UserForm1.orderNumber & ": " & UserForm1.CustomerName & "<br><br>"
    HTMLString = HTMLString & "Could you please enter a Change Order for the following:</b><br><br>"
    
    If (ChangeOrderForm.ChangeOrderTextBox.Visible) Then
        HTMLString = IIf(ChangeOrderForm.ChangeOrderTextBox.Value = "Enter Comments Here.", HTMLString & _
                            "<p style=""color:#FF0000"";>Enter Comments Here.</p>", HTMLString & _
                            ChangeOrderForm.ChangeOrderTextBox.Value & "<br><br>")
    End If
    
    HTMLString = HTMLString & EngineeringType & "APPROVAL DRAWINGS RECEIVED:<br><br>"
    SelectedOption = IIf(ChangeOrderForm.NoChangesButton.Value, "<u>XX</u> ", "__ ")
    HTMLString = HTMLString & SelectedOption & "NO CHANGES, RELEASE TO MANUFACTURING - ALL LINES<br>"
    SelectedOption = IIf(ChangeOrderForm.NoteReleaseButton.Value, "<u>XX</u> ", "__ ")
    HTMLString = HTMLString & SelectedOption & "NOTE CUSTOMER CHANGES AND RELEASE TO MANUFACTURING. - ALL LINES<br>"
    SelectedOption = IIf(ChangeOrderForm.NotApprovedButton.Value, "<u>XX</u> ", "__ ")
    HTMLString = HTMLString & SelectedOption & "NOT APPROVED, MAKE CUSTOMER CHANGES AND ISSUE NEW APPROVALS.<br>"
    SelectedOption = IIf(ChangeOrderForm.NotReceivedButton.Value, "<u>XX</u> ", "__ ")
    HTMLString = HTMLString & SelectedOption & "NOT RECEIVED.<br>"
    SelectedOption = IIf(ChangeOrderForm.NotRequiredButton.Value, "<u>XX</u> ", "__ ")
    HTMLString = HTMLString & SelectedOption & "NOT REQUIRED.<br>"
    SelectedOption = IIf(ChangeOrderForm.ReceivedButton.Value, "<u>XX</u> ", "__ ")
    HTMLString = HTMLString & SelectedOption & "ALREADY RECD - ACTION ALREADY TAKEN"
    
    With OutMail
        .Display
    End With
    
    'Capture personal signature
    Dim Signature As String
    Signature = OutMail.HTMLBody
    
    On Error Resume Next
    With OutMail
        .To = ENGINEERING_CLERK_EMAIL
        .Subject = EngineeringType & "CHANGE ORDER: " & UserForm1.orderNumber & " - " & UserForm1.CustomerName
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body>" & HTMLString & "<br>" & Signature & "</body></html>"
    End With
    
    On Error GoTo 0
    
    'Add attachments
'    If (ChangeOrderForm.ChangeOrderTree.Nodes.Count > 0) Then
'        Dim Node As Variant
'        For Each Node In ChangeOrderForm.ChangeOrderTree.Nodes
'            With OutMail
'                .Attachments.Add (Node.Text)
'            End With
'            Call TransferFiles(Node.Text)
'        Next
'    End If
End Sub

Sub TransferFiles(NodeText As String)

    Dim OrdPrefix As String
    OrdPrefix = "\\USLXA-P-FS01\Sabetha\Orders\Orders\"

    OrdFold = OrdPrefix & Mid$(UserForm1.orderNumber, 1, 7) & "000\" & Mid$(UserForm1.orderNumber, 1, 10) & "*"
    TxtLine = Dir(OrdFold, vbDirectory)
    JSubFold = Mid$(TxtLine, 1, 7) & "000\"
    OrdFold = OrdPrefix & JSubFold & TxtLine & "\Customer-Rep Correspondence\"
    SplitArray = Split(NodeText, "\")
    Dim SplitArrayLength As Long
    SplitArrayLength = UBound(SplitArray) - LBound(SplitArray)
    FileCopy NodeText, OrdFold & SplitArray(SplitArrayLength)
        
End Sub

Sub FindEmailToPerson()

    Call InitiateSAP
    
    Dim sap_main_window As GuiMainWindow: Set sap_main_window = session.FindById("wnd[0]")
    
    OrdStr = Str$(Ord)

    On Error GoTo SkipSAP
    session.FindById("wnd[0]").maximize
    session.StartTransaction ("VA03")
        
    sap_main_window.FindById("usr/ctxtVBAK-VBELN").Text OrdStr
    sap_main_window.SendVKey (0)
        
    If session.Children.Count > 1 Then
        session.FindById("wnd[0]").SendVKey 0
    End If
        
    sap_main_window.FindById("usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press

    Dim SAP_Tabstrip As GuiTabStrip: Set SAP_Tabstrip = sap_main_window.FindById("usr/tabsTAXI_TABSTRIP_HEAD")
    Dim SAP_Texts_Tab_Name As String
    Dim SAP_AddData_Tab_Name As String
        
    Dim Tabstrip_Index As Long
    For Tabstrip_Index = 0 To SAP_Tabstrip.Children.Count - 1
        If SAP_Tabstrip.Children(Tabstrip_Index).Text = "Texts" Then
            SAP_Texts_Tab_Name = SAP_Tabstrip.Children(Tabstrip_Index).Name
        ElseIf SAP_Tabstrip.Children(Tabstrip_Index).Text = "Additional data B" Then
            SAP_AddData_Tab_Name = SAP_Tabstrip.Children(Tabstrip_Index).Name
        End If
    Next Tabstrip_Index
        
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_Texts_Tab_Name).Select
        
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_Texts_Tab_Name & "/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "Z001", "Column1"
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_Texts_Tab_Name & "/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectItem "Z001", "Column1"
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_Texts_Tab_Name & "/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleClickItem "Z001", "Column1"
        
    'Customer and RSM
    EmailPerson = session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_Texts_Tab_Name & "/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
    RsmRepPerson = session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_Texts_Tab_Name & "/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").Text
        
        
    session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_AddData_Tab_Name).Select
        
    'Application Engineer
    ProjResponsible = session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabp" & SAP_AddData_Tab_Name & "/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtZZPJB_NAM").Text
        
    session.FindById("wnd[0]").SendVKey (15)
        
    'Email To
    FindValueStart = InStr(EmailPerson, "EMAIL DWGS TO:") + 14
    FindValueStop = InStr(FindValueStart, EmailPerson, "2.")
        
    If FindValueStop = 0 Then
        FindValueStop = InStr(FindValueStart, EmailPerson, "****")
    End If
        
    If ((FindValueStart <> 14) And (FindValueStop > FindValueStart)) Then
        EmailPerson = Mid$(EmailPerson, FindValueStart, (FindValueStop - FindValueStart))
        OutMail.To = EmailPerson
    End If
        
    'RSM/REP To
    FindValueStart = InStr(RsmRepPerson, "APPLICATION ENGINEER:") + 21
    FindValueStop = InStr(FindValueStart, RsmRepPerson, "SPECIAL ENG NOTES")
        
    If FindValueStop = 0 Then
        FindValueStop = InStr(FindValueStart, RsmRepPerson, "****")
    End If
        
    If ((FindValueStart <> 21) And (FindValueStop > FindValueStart + 1)) Then
        RsmRepPerson = Mid$(RsmRepPerson, FindValueStart, (FindValueStop - FindValueStart))
        OutMail.CC = OutMail.CC & ";" & RsmRepPerson
    End If
        
    For AEIndex = 4 To QueryResults2WS.Cells(QueryResults2WS.Rows.Count, "BR").End(xlUp).Row
        If QueryResults2WS.Cells(AEIndex, "BR").Value = ProjResponsible Then
            OutMail.CC = OutMail.CC & ";" & QueryResults2WS.Cells(AEIndex, "BS").Value
        End If
    Next AEIndex

    Exit Sub
    
SkipSAP:

End Sub

Sub WeeklyNotes()
    WeeklyNotesWS.Activate
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Dim EngInit As String
    
    EngInit = UCase(ScheduleWS.Range("A2").Value)
    EndSpan = Date - WeeklyNotesWS.Cells(5, 3).Value
    Columns("AA:AD").EntireColumn.Hidden = True
    Rows("2:1048576").RowHeight = 15

    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open
  
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    oRs.Source = "SELECT cpe_schedule.datestamp, orderno, comments FROM cpe_schedule WHERE (cpe_schedule.datestamp BETWEEN '" & EndSpan & "' AND '" & Date & "') AND (cpe_schedule.engineer='" & EngInit & "') AND (cpe_schedule.comments <> '" & vbNullString & "') ORDER BY cpe_schedule.orderno, datestamp"

    oRs.Open
    Set ClrRange = WeeklyNotesWS.Range("AA9:AC1048576")
    ClrRange.ClearContents
    ClrRange.CopyFromRecordset oRs
    
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing

    Call FindCustName
    
    Set ClrRange = WeeklyNotesWS.Range("D9:G1048576")
    ClrRange.ClearContents

    lastRow = WeeklyNotesWS.Cells(Rows.Count, 27).End(xlUp).Row

    Columns("AB:AB").NumberFormat = "General"

    For index = 9 To lastRow
        OrdStr = WeeklyNotesWS.Cells(index, 28).Value
        CustName = vbNullString
        If OrdStr <> WeeklyNotesWS.Cells((index - 1), 28).Value Then
            'Call FindCustName
            WeeklyNotesWS.Cells(index, 4).Value = WeeklyNotesWS.Cells(index, 28).Value
            WeeklyNotesWS.Cells(index, 5).Value = WeeklyNotesWS.Cells(index, 30).Value
        End If
        WeeklyNotesWS.Cells(index, 6).Value = WeeklyNotesWS.Cells(index, 27).Value
        WeeklyNotesWS.Cells(index, 7).Value = Chr(149) & " " & WeeklyNotesWS.Cells(index, 29).Value
    Next index

    Columns("D:D").NumberFormat = "General"
    Rows("2:1048576").EntireRow.AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub FindCustName()
    
    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open
  
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    Columns("AB:AB").NumberFormat = "General"
    'Find Customer Name in Active Orders
    Dim LastRow1 As Long
    LastRow1 = WeeklyNotesWS.Cells(Rows.Count, 28).End(xlUp).Row
    Dim Index1 As Long
    For Index1 = 9 To LastRow1

        oRs.Source = "SELECT DISTINCT Prod_Eng.Customer_Name FROM Prod_Eng WHERE Prod_Eng.Order_Num='" & WeeklyNotesWS.Cells(Index1, 28) & "'"
        '(cpe_schedule.datestamp BETWEEN '" & EndSpan & "' AND '" & Date & "') AND (cpe_schedule.engineer='" & EngInit & "') AND (cpe_schedule.comments <> '" & "" & "') ORDER BY cpe_schedule.orderno, datestamp"
    
        oRs.Open
        BuildRange = "AD" & Index1 & ":AD1048576"
        Set ClrRange = WeeklyNotesWS.Range(BuildRange)
        ClrRange.ClearContents
        ClrRange.CopyFromRecordset oRs
        oRs.Close
    Next Index1
    
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing

End Sub

Sub DashboardEmail()
    WeeklyNotes
    
    'Set Send Info
    
    Dim ManagerEmail As String
    Dim ManagerName As String
    
    If ScheduleWS.Range("AJ2").Value = "PC" Then
        ManagerEmail = "nathan.hemman@ktron.com"
        ManagerName = "Nathan"
    Else
        ManagerEmail = "brian.schmoldt@ktron.com"
        ManagerName = "Brian"
    End If
    
    'Create Email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    'Greetings
    BuildHTML = vbNullString
    BuildHTML = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    BuildHTML = BuildHTML & "Hello " & ManagerName & ","
    
    'This Week
    BuildHTML = BuildHTML & "<BR><BR>"
    BuildHTML = BuildHTML & "&emsp;<b><u>This Week</u>:</b><ul>"
    lastRow = WeeklyNotesWS.Cells(Rows.Count, 6).End(xlUp).Row
    'Order Number
    For index = 9 To lastRow
        If WeeklyNotesWS.Cells(index, 4).Value <> vbNullString Then
            If index <> 9 Then
                BuildHTML = BuildHTML & "</ul><BR>"
            End If
            BuildHTML = BuildHTML & "<li style=margin-top:0px;margin-bottom:0px><b>" & WeeklyNotesWS.Cells(index, 4).Value & ":</b> " & WeeklyNotesWS.Cells(index, 5).Value & "</li>"
            'Order Notes
            BuildHTML = BuildHTML & "<ul><li style=margin-top:0px;margin-bottom:0px>[" & WeeklyNotesWS.Cells(index, 6).Value & "] " & WeeklyNotesWS.Cells(index, 29).Value & "</li>"
        Else
            'Order Notes
            BuildHTML = BuildHTML & "<li style=margin-top:0px;margin-bottom:0px>[" & WeeklyNotesWS.Cells(index, 6).Value & "] " & WeeklyNotesWS.Cells(index, 29).Value & "</li>"
        End If
           
    Next index
    BuildHTML = BuildHTML & "</ul></ul>"
           
    'Next Week
    BuildHTML = BuildHTML & "<BR>"
    BuildHTML = BuildHTML & "&emsp;<b><u>Next Week</u>:</b><ul>"
    lastRow = ScheduleWS.Cells(Rows.Count, 2).End(xlUp).Row
    If lastRow > 59 Then lastRow = 59
    If lastRow < 4 Then
    End If
    For index = 4 To lastRow
        If (UCase(ScheduleWS.Cells(index, 13).Value) = "WAITING FOR CUSTOMER APPROVAL") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "HOLD") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "IN QUEUE") _
        Or (UCase(ScheduleWS.Cells(index, 13).Value) = "IN DESIGN") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "CHECKING") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "WORKING ON PRODUCTION RELEASE") _
        Or (UCase(ScheduleWS.Cells(index, 13).Value) = "CHANGE ORDER") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "OTHER") Then
            BuildHTML = BuildHTML & "<li style=margin-top:0px;margin-bottom:0px><b>" & ScheduleWS.Cells(index, 2).Value & ":</b> " & ScheduleWS.Cells(index, 3).Value & " (" & ScheduleWS.Cells(index, 13).Value & ")</li>"
        End If
    Next index
    BuildHTML = BuildHTML & "</ul>"
    
    'Travel/PTO
    BuildHTML = BuildHTML & "<BR>"
    BuildHTML = BuildHTML & "&emsp;<b><u>Travel/PTO</u>:</b>"
    BuildHTML = BuildHTML & "<ul><li style=margin-top:0px;margin-bottom:0px>Travel: None</li>"
    BuildHTML = BuildHTML & "<li style=margin-top:0px;margin-bottom:0px>PTO: None</li></ul>"
    
    'Concerns
    BuildHTML = BuildHTML & "<BR>"
    BuildHTML = BuildHTML & "&emsp;<b><u>Concerns</u>:</b>"
    BuildHTML = BuildHTML & "<ul><li style=margin-top:0px;margin-bottom:0px>None</li></ul>"
    
    'Crisis
    BuildHTML = BuildHTML & "<BR>"
    BuildHTML = BuildHTML & "&emsp;<b><u>Crisis</u>:</b>"
    BuildHTML = BuildHTML & "<ul><li style=margin-top:0px;margin-bottom:0px>None</li></ul>"
    
    With OutMail
        .Display
    End With
    
    'Signature
    Dim Signature As String
    Signature = OutMail.HTMLBody
      
    On Error Resume Next
    With OutMail
        .To = ManagerEmail
        .CC = "brian.schmoldt@ktron.com"
        .BCC = vbNullString
        .Subject = "Dashboard: " & WeeklyNotesWS.Cells(3, 3).Value & " (" & EndSpan & " - " & Date & ")"
        '.Body = strbody
        .HTMLBody = "<html><body>" & BuildHTML & "<br>" & Signature & "</body></html>"
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        '.Display '.Send   'or use .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub



