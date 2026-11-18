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

Sub SendProductionTest()
    Ord = 1100102179
    CustName = LookupCustomerName(Ord)
    SendProductionEmail
End Sub

Sub SendDashboardTest()
    DashboardEmail
End Sub

Sub SendApprovalTest()
    Dim hasReminderText As Boolean: hasReminderText = False
    Dim lastUpdate As Date: lastUpdate = Date
    Dim orderNumber As Long: orderNumber = 1100102179
        
    SendApprovalEmail hasReminderText, orderNumber:=orderNumber
End Sub


Sub SendApprovalEmail(Optional ByVal hasReminderText As Boolean = False, Optional ByVal lastUpdate As Date = 0, Optional ByVal orderNumber As Long = 0, Optional ByVal customerName As String = vbNullString)
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    'If order number not 0
    If orderNumber <> 0 Then Ord = orderNumber
    
    'CustName is only filled in when the schedule loop runs, so reminders and tests supply it
    'or it is looked up here; otherwise "Ship To" would be blank or left over from another order
    If Len(customerName) > 0 Then
        CustName = customerName
    ElseIf orderNumber <> 0 Then
        CustName = LookupCustomerName(Ord)
    End If
    
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
    
    'Build the body and subject from the Email_Templates sheet (EmailTemplates.bas)
    Dim tokens As Object: Set tokens = NewTokens()
    AddToken tokens, "OrderNumber", CStr(Ord)
    AddToken tokens, "ShipTo", CustName
    AddToken tokens, "PONumber", PONum
    AddToken tokens, "EngType", EngType
    AddToken tokens, "EngTypeLower", LCase$(EngType)
    AddToken tokens, "OtherEngType", OthEngType
    AddToken tokens, "OtherEngTypeLower", LCase$(OthEngType)
    AddToken tokens, "LastUpdate", CStr(lastUpdate)
    SetFlag tokens, "Reminder", hasReminderText
    
    Dim emailSubject As String
    BuildHTML = RenderTemplate(TEMPLATE_APPROVAL, tokens, emailSubject)
    
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
        .Subject = emailSubject
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body>" & BuildHTML & "<br>" & Signature & "</body></html>"
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

'Customer name for an order from the Prod_Eng table (empty if not found)
Private Function LookupCustomerName(ByVal orderNo As Long) As String
    Dim con As ADODB.connection: Set con = New ADODB.connection
    con.ConnectionString = databaseConnectionString
    con.Open
    
    Dim rs As ADODB.Recordset
    Set rs = con.Execute("SELECT DISTINCT Customer_Name FROM Prod_Eng WHERE Order_Num=" & orderNo)
    If Not rs.EOF Then LookupCustomerName = CStr(rs.Fields(0).Value & vbNullString)
    
    rs.Close
    con.Close
End Function

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
    
    'Build the body and subject from the Email_Templates sheet (EmailTemplates.bas)
    Dim tokens As Object: Set tokens = NewTokens()
    AddToken tokens, "OrderNumber", order_number
    AddToken tokens, "ShipTo", CustName
    AddToken tokens, "PONumber", PONum
    AddToken tokens, "EngType", EngType
    AddToken tokens, "EngTypeLower", LCase$(EngType)
    
    Dim emailSubject As String
    BuildHTML = RenderTemplate(TEMPLATE_CERTIFIED, tokens, emailSubject)
        
    
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
        .Subject = emailSubject
        .BodyFormat = olFormatHTML
        .HTMLBody = "<html><body>" & BuildHTML & "<br>" & Signature & "</body></html>"
    End With
    On Error GoTo 0
    
    If (Debg = True) Then Exit Sub
    
    Call FindEmailToPerson
    
    Set OutMail = Nothing
    Set OutApp = Nothing
    
End Sub

Sub SendProductionEmail()
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
    
    'Build the body and subject from the Email_Templates sheet (EmailTemplates.bas)
    Dim tokens As Object: Set tokens = NewTokens()
    AddToken tokens, "OrderNumber", CStr(Ord)
    AddToken tokens, "ShipTo", CustName
    AddToken tokens, "RecipientName", manufacturingEngineeringName
    
    Dim emailSubject As String
    BuildHTML = RenderTemplate(TEMPLATE_PRODUCTION, tokens, emailSubject)
    
    'Get Signature
    OutMail.Display
    Dim Signature As String
    Signature = OutMail.HTMLBody
    
    On Error Resume Next
    
    With OutMail
        .To = manufacturingEngineering
        .CC = "brian.schmoldt@ktron.com"
        .Subject = emailSubject
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
    
    'Build the body and subject from the Email_Templates sheet (EmailTemplates.bas)
    Dim tokens As Object: Set tokens = NewTokens()
    AddToken tokens, "OrderNumber", CStr(UserForm1.orderNumber)
    AddToken tokens, "ShipTo", CStr(UserForm1.CustomerName)
    AddToken tokens, "EngTypeUpper", Trim$(EngineeringType)
    
    'Comments typed on the form (text is encoded; line breaks kept)
    Dim commentsHtml As String
    If ChangeOrderForm.ChangeOrderTextBox.Visible Then
        If ChangeOrderForm.ChangeOrderTextBox.Value = "Enter Comments Here." Then
            commentsHtml = "<p style=""color:#FF0000"";>Enter Comments Here.</p>"
        Else
            commentsHtml = Replace(Replace(Replace(HtmlEscape(ChangeOrderForm.ChangeOrderTextBox.Value), vbCrLf, "<br>"), vbCr, "<br>"), vbLf, "<br>") & "<br><br>"
        End If
    End If
    AddHtmlToken tokens, "Comments", commentsHtml
    
    'Selected option shows as underlined XX, others as a blank line
    AddHtmlToken tokens, "NoChanges", IIf(ChangeOrderForm.NoChangesButton.Value, "<u>XX</u>", "__")
    AddHtmlToken tokens, "NoteRelease", IIf(ChangeOrderForm.NoteReleaseButton.Value, "<u>XX</u>", "__")
    AddHtmlToken tokens, "NotApproved", IIf(ChangeOrderForm.NotApprovedButton.Value, "<u>XX</u>", "__")
    AddHtmlToken tokens, "NotReceived", IIf(ChangeOrderForm.NotReceivedButton.Value, "<u>XX</u>", "__")
    AddHtmlToken tokens, "NotRequired", IIf(ChangeOrderForm.NotRequiredButton.Value, "<u>XX</u>", "__")
    AddHtmlToken tokens, "AlreadyReceived", IIf(ChangeOrderForm.ReceivedButton.Value, "<u>XX</u>", "__")
    
    Dim emailSubject As String
    HTMLString = RenderTemplate(TEMPLATE_CHANGE_ORDER, tokens, emailSubject)
    
    With OutMail
        .Display
    End With
    
    'Capture personal signature
    Dim Signature As String
    Signature = OutMail.HTMLBody
    
    On Error Resume Next
    With OutMail
        .To = ENGINEERING_CLERK_EMAIL
        .Subject = emailSubject
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
    
    'InitiateSAP quietly leaves the session empty if SAP is not open or logged in
    If session Is Nothing Then
        MsgBox "Could not connect to SAP, so the customer contact was not looked up for this email." & vbCr & vbCr & _
               "Please open SAP and log in (leave an ""SAP Easy Access"" window available, with fewer than 6 SAP windows open), " & _
               "then fill in the ""To"" line manually or run the email again.", vbExclamation, "SAP not available"
        Exit Sub
    End If
    
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

    'This week: orders with their notes, built from the Weekly_Notes sheet
    Dim thisWeekHtml As String: thisWeekHtml = "<ul>"
    lastRow = WeeklyNotesWS.Cells(Rows.Count, 6).End(xlUp).Row
    'Order Number
    For index = 9 To lastRow
        If WeeklyNotesWS.Cells(index, 4).Value <> vbNullString Then
            If index <> 9 Then
                thisWeekHtml = thisWeekHtml & "</ul><BR>"
            End If
            thisWeekHtml = thisWeekHtml & "<li style=margin-top:0px;margin-bottom:0px><b>" & HtmlEscape(CStr(WeeklyNotesWS.Cells(index, 4).Value)) & ":</b> " & HtmlEscape(CStr(WeeklyNotesWS.Cells(index, 5).Value)) & "</li>"
            'Order Notes
            thisWeekHtml = thisWeekHtml & "<ul><li style=margin-top:0px;margin-bottom:0px>[" & HtmlEscape(CStr(WeeklyNotesWS.Cells(index, 6).Value)) & "] " & HtmlEscape(CStr(WeeklyNotesWS.Cells(index, 29).Value)) & "</li>"
        Else
            'Order Notes
            thisWeekHtml = thisWeekHtml & "<li style=margin-top:0px;margin-bottom:0px>[" & HtmlEscape(CStr(WeeklyNotesWS.Cells(index, 6).Value)) & "] " & HtmlEscape(CStr(WeeklyNotesWS.Cells(index, 29).Value)) & "</li>"
        End If
    Next index
    thisWeekHtml = thisWeekHtml & "</ul></ul>"
    
    'Next week: open orders on the schedule
    Dim nextWeekHtml As String: nextWeekHtml = "<ul>"
    lastRow = ScheduleWS.Cells(Rows.Count, 2).End(xlUp).Row
    If lastRow > 59 Then lastRow = 59
    For index = 4 To lastRow
        If (UCase(ScheduleWS.Cells(index, 13).Value) = "WAITING FOR CUSTOMER APPROVAL") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "HOLD") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "IN QUEUE") _
        Or (UCase(ScheduleWS.Cells(index, 13).Value) = "IN DESIGN") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "CHECKING") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "WORKING ON PRODUCTION RELEASE") _
        Or (UCase(ScheduleWS.Cells(index, 13).Value) = "CHANGE ORDER") Or (UCase(ScheduleWS.Cells(index, 13).Value) = "OTHER") Then
            nextWeekHtml = nextWeekHtml & "<li style=margin-top:0px;margin-bottom:0px><b>" & HtmlEscape(CStr(ScheduleWS.Cells(index, 2).Value)) & ":</b> " & HtmlEscape(CStr(ScheduleWS.Cells(index, 3).Value)) & " (" & HtmlEscape(CStr(ScheduleWS.Cells(index, 13).Value)) & ")</li>"
        End If
    Next index
    nextWeekHtml = nextWeekHtml & "</ul>"
    
    'Build the body and subject from the Email_Templates sheet (EmailTemplates.bas)
    Dim tokens As Object: Set tokens = NewTokens()
    AddToken tokens, "ManagerName", ManagerName
    AddToken tokens, "WeekLabel", CStr(WeeklyNotesWS.Cells(3, 3).Value)
    AddToken tokens, "StartDate", CStr(EndSpan)
    AddToken tokens, "EndDate", CStr(Date)
    AddHtmlToken tokens, "ThisWeek", thisWeekHtml
    AddHtmlToken tokens, "NextWeek", nextWeekHtml
    
    Dim emailSubject As String
    BuildHTML = RenderTemplate(TEMPLATE_DASHBOARD, tokens, emailSubject)
    
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
        .Subject = emailSubject
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



