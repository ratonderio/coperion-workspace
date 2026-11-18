Attribute VB_Name = "EmailTemplates"
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder "Modules"
'
' Editable email templates, stored on a worksheet (TEMPLATE_SHEET) in this workbook.
'
' Sheet layout (one row per line):
'   Column A  Email name (for example "Approval Email")
'   Column B  "Subject" (one row per email) or "Body" (one row per line of HTML)
'   Column C  The text. Body lines are joined in order; line breaks between rows are
'             ignored by email, so use <BR> for a visible line break.
'
' Placeholders in the text:
'   {Token}                     replaced with the value for the order, e.g. {OrderNumber}
'   [[IF Name]] ... [[/IF]]     kept only when the flag Name is on, otherwise removed
'
' The built-in defaults at the bottom of this module are used to create the sheet the first
' time, to add any email that is missing from it, and by RestoreTemplate.

Public Const TEMPLATE_SHEET As String = "Email_Templates"

Public Const TEMPLATE_APPROVAL As String = "Approval Email"
Public Const TEMPLATE_CERTIFIED As String = "Certified Email"
Public Const TEMPLATE_PRODUCTION As String = "Production Release Email"
Public Const TEMPLATE_CHANGE_ORDER As String = "Change Order Email"
Public Const TEMPLATE_DASHBOARD As String = "Dashboard Email"

Private Const COL_NAME As Long = 1
Private Const COL_PART As Long = 2
Private Const COL_TEXT As Long = 3
Private Const COL_HELP_NAME As Long = 5
Private Const COL_HELP_TOKEN As Long = 6
Private Const COL_HELP_TEXT As Long = 7
Private Const HEADER_ROW As Long = 3
Private Const FIRST_DATA_ROW As Long = 4

Private Const PART_SUBJECT As String = "Subject"
Private Const PART_BODY As String = "Body"

' ---------------------------------------------------------------------------
' Tokens
' ---------------------------------------------------------------------------

Public Function NewTokens() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    Set NewTokens = d
End Function

'Adds {Name}; the value is HTML-encoded so names containing & or < cannot break the email.
Public Sub AddToken(ByVal tokens As Object, ByVal tokenName As String, ByVal tokenValue As String)
    tokens("{" & tokenName & "}") = HtmlEncode(tokenValue)
End Sub

'Turns a [[IF Name]] block on or off.
Public Sub SetFlag(ByVal tokens As Object, ByVal flagName As String, ByVal isOn As Boolean)
    tokens("?" & flagName) = isOn
End Sub

'Adds {Name} with a value that is already HTML (a generated list, a bold mark, etc.).
Public Sub AddHtmlToken(ByVal tokens As Object, ByVal tokenName As String, ByVal htmlValue As String)
    tokens("{" & tokenName & "}") = htmlValue
End Sub

'Encodes text for safe use inside HTML.
Public Function HtmlEscape(ByVal text As String) As String
    HtmlEscape = HtmlEncode(text)
End Function

Private Function HtmlEncode(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    HtmlEncode = text
End Function

' ---------------------------------------------------------------------------
' Rendering
' ---------------------------------------------------------------------------

'Returns the finished HTML body and sets emailSubject, using the text on the template sheet.
Public Function RenderTemplate(ByVal templateName As String, ByVal tokens As Object, ByRef emailSubject As String) As String
    Dim html As String
    ReadTemplate templateName, html, emailSubject

    html = ApplyConditionals(html, tokens)

    Dim key As Variant
    For Each key In tokens.Keys
        If Left$(key, 1) = "{" Then
            html = Replace(html, key, tokens(key))
            emailSubject = Replace(emailSubject, key, tokens(key))
        End If
    Next key

    WarnAboutLeftoverTokens html & emailSubject, templateName
    RenderTemplate = html
End Function

Private Function ApplyConditionals(ByVal html As String, ByVal tokens As Object) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\[\[IF (\w+)\]\]([\s\S]*?)\[\[/IF\]\]"

    Dim ms As Object: Set ms = re.Execute(html)
    Dim i As Long, m As Object, flagName As String, keep As Boolean, kept As String
    For i = ms.Count - 1 To 0 Step -1
        Set m = ms(i)
        flagName = "?" & m.SubMatches(0)
        keep = tokens.Exists(flagName)
        If keep Then keep = CBool(tokens(flagName))
        kept = vbNullString
        If keep Then kept = m.SubMatches(1)
        html = Left$(html, m.FirstIndex) & kept & Mid$(html, m.FirstIndex + m.Length + 1)
    Next i
    ApplyConditionals = html
End Function

Private Sub WarnAboutLeftoverTokens(ByVal text As String, ByVal templateName As String)
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "\{[A-Za-z]+\}|\[\[/?IF[^\]]*\]\]"

    Dim ms As Object: Set ms = re.Execute(text)
    If ms.Count = 0 Then Exit Sub

    Dim found As String, i As Long
    For i = 0 To ms.Count - 1
        If InStr(found, ms(i).Value) = 0 Then found = found & vbCr & "  " & ms(i).Value
    Next i
    MsgBox "The """ & templateName & """ template has placeholders that were not recognized:" & found & vbCr & vbCr & _
           "Check the spelling on the " & TEMPLATE_SHEET & " sheet. The email was still created.", vbExclamation, "Email template"
End Sub

' ---------------------------------------------------------------------------
' Template sheet
' ---------------------------------------------------------------------------

'Shows the template sheet (assign to a button).
Public Sub ShowEmailTemplates()
    TemplateSheet().Activate
End Sub

Public Sub RestoreApprovalTemplate()
    RestoreTemplate TEMPLATE_APPROVAL
End Sub

Public Sub RestoreCertifiedTemplate()
    RestoreTemplate TEMPLATE_CERTIFIED
End Sub

Public Sub RestoreProductionTemplate()
    RestoreTemplate TEMPLATE_PRODUCTION
End Sub

Public Sub RestoreChangeOrderTemplate()
    RestoreTemplate TEMPLATE_CHANGE_ORDER
End Sub

Public Sub RestoreDashboardTemplate()
    RestoreTemplate TEMPLATE_DASHBOARD
End Sub

'Replaces one email's rows on the sheet with the built-in default wording.
Public Sub RestoreTemplate(ByVal templateName As String)
    If MsgBox("Replace the """ & templateName & """ rows on the " & TEMPLATE_SHEET & " sheet with the original wording?" & vbCr & vbCr & _
              "Any edits to that email will be lost.", vbYesNo + vbQuestion, "Restore default template") <> vbYes Then Exit Sub

    Dim ws As Worksheet: Set ws = TemplateSheet()
    Dim firstRow As Long: firstRow = FirstRowOf(ws, templateName)

    If firstRow > 0 Then
        Dim r As Long
        For r = LastDataRow(ws) To firstRow Step -1
            If SameName(ws.Cells(r, COL_NAME).Value, templateName) Then ws.Rows(r).Delete
        Next r
    Else
        firstRow = LastDataRow(ws) + 1
    End If

    Dim needed As Long: needed = DefaultLines(templateName).Count + 1
    ws.Rows(firstRow & ":" & (firstRow + needed - 1)).Insert
    WriteTemplateRows ws, templateName, firstRow
    MsgBox """" & templateName & """ restored.", vbInformation, "Restore default template"
End Sub

'Returns the template worksheet, creating it (and its default content) if it is missing.
Private Function TemplateSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TEMPLATE_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Dim previous As Object: Set previous = ActiveSheet
        Dim screenUpdating As Boolean: screenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = TEMPLATE_SHEET
        BuildSheetHeader ws

        If Not previous Is Nothing Then previous.Activate
        Application.ScreenUpdating = screenUpdating
    End If
    Set TemplateSheet = ws
End Function

Private Sub BuildSheetHeader(ByVal ws As Worksheet)
    ws.Columns(COL_TEXT).NumberFormat = "@"
    ws.Cells(1, COL_NAME).Value = "Email templates"
    ws.Cells(1, COL_NAME).Font.Bold = True
    ws.Cells(1, COL_NAME).Font.Size = 14
    ws.Cells(2, COL_NAME).Value = "Edit the text in column C. Body lines are joined in order; line breaks between rows are ignored, so use <BR> for a new line. " & _
                                  "Do not change columns A and B. Placeholders such as {OrderNumber} are listed on the right."

    ws.Cells(HEADER_ROW, COL_NAME).Value = "Email"
    ws.Cells(HEADER_ROW, COL_PART).Value = "Part"
    ws.Cells(HEADER_ROW, COL_TEXT).Value = "Text (HTML)"
    ws.Cells(HEADER_ROW, COL_HELP_NAME).Value = "Email"
    ws.Cells(HEADER_ROW, COL_HELP_TOKEN).Value = "Placeholder"
    ws.Cells(HEADER_ROW, COL_HELP_TEXT).Value = "Meaning"
    ws.Range(ws.Cells(HEADER_ROW, COL_NAME), ws.Cells(HEADER_ROW, COL_TEXT)).Font.Bold = True
    ws.Range(ws.Cells(HEADER_ROW, COL_HELP_NAME), ws.Cells(HEADER_ROW, COL_HELP_TEXT)).Font.Bold = True

    ws.Columns(COL_NAME).ColumnWidth = 20
    ws.Columns(COL_PART).ColumnWidth = 9
    ws.Columns(COL_TEXT).ColumnWidth = 130
    ws.Columns(4).ColumnWidth = 3
    ws.Columns(COL_HELP_NAME).ColumnWidth = 20
    ws.Columns(COL_HELP_TOKEN).ColumnWidth = 22
    ws.Columns(COL_HELP_TEXT).ColumnWidth = 60
End Sub

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    LastDataRow = ws.Cells(ws.Rows.Count, COL_NAME).End(xlUp).Row
    If LastDataRow < HEADER_ROW Then LastDataRow = HEADER_ROW
End Function

Private Function SameName(ByVal a As Variant, ByVal b As String) As Boolean
    SameName = (StrComp(Trim$(CStr(a)), b, vbTextCompare) = 0)
End Function

Private Function FirstRowOf(ByVal ws As Worksheet, ByVal templateName As String) As Long
    Dim r As Long
    For r = FIRST_DATA_ROW To LastDataRow(ws)
        If SameName(ws.Cells(r, COL_NAME).Value, templateName) Then FirstRowOf = r: Exit Function
    Next r
End Function

'Writes the Subject row and Body rows for one email starting at startRow (rows must be empty).
Private Sub WriteTemplateRows(ByVal ws As Worksheet, ByVal templateName As String, ByVal startRow As Long)
    Dim r As Long: r = startRow
    ws.Cells(r, COL_NAME).Value = templateName
    ws.Cells(r, COL_PART).Value = PART_SUBJECT
    ws.Cells(r, COL_TEXT).Value = DefaultSubject(templateName)

    Dim line As Variant
    For Each line In DefaultLines(templateName)
        r = r + 1
        ws.Cells(r, COL_NAME).Value = templateName
        ws.Cells(r, COL_PART).Value = PART_BODY
        ws.Cells(r, COL_TEXT).Value = CStr(line)
    Next line

    WriteHelp ws, templateName
End Sub

Private Sub WriteHelp(ByVal ws As Worksheet, ByVal templateName As String)
    Dim r As Long
    For r = FIRST_DATA_ROW To ws.Cells(ws.Rows.Count, COL_HELP_NAME).End(xlUp).Row
        If SameName(ws.Cells(r, COL_HELP_NAME).Value, templateName) Then Exit Sub
    Next r

    Dim nextRow As Long: nextRow = ws.Cells(ws.Rows.Count, COL_HELP_NAME).End(xlUp).Row + 1
    If nextRow < FIRST_DATA_ROW Then nextRow = FIRST_DATA_ROW

    Dim item As Variant, parts() As String
    For Each item In DefaultHelp(templateName)
        parts = Split(CStr(item), "|")
        ws.Cells(nextRow, COL_HELP_NAME).Value = templateName
        ws.Cells(nextRow, COL_HELP_TOKEN).Value = "'" & parts(0)
        ws.Cells(nextRow, COL_HELP_TEXT).Value = parts(1)
        nextRow = nextRow + 1
    Next item
End Sub

'Reads the Subject and Body for an email from the sheet, adding the default rows first if
'the email is not on the sheet yet.
Private Sub ReadTemplate(ByVal templateName As String, ByRef body As String, ByRef subject As String)
    Dim ws As Worksheet: Set ws = TemplateSheet()

    If FirstRowOf(ws, templateName) = 0 Then WriteTemplateRows ws, templateName, LastDataRow(ws) + 1

    Dim r As Long, part As String, text As String
    body = vbNullString
    subject = vbNullString
    For r = FIRST_DATA_ROW To LastDataRow(ws)
        If SameName(ws.Cells(r, COL_NAME).Value, templateName) Then
            part = Trim$(CStr(ws.Cells(r, COL_PART).Value))
            text = CStr(ws.Cells(r, COL_TEXT).Value)
            If StrComp(part, PART_SUBJECT, vbTextCompare) = 0 Then
                subject = text
            ElseIf StrComp(part, PART_BODY, vbTextCompare) = 0 And Len(text) > 0 Then
                If Len(body) > 0 Then body = body & vbCrLf
                body = body & text
            End If
        End If
    Next r

    'A blank body or subject would send an empty email, so fall back to the default
    If Len(Trim$(body)) = 0 Then body = Join(CollectionToArray(DefaultLines(templateName)), vbCrLf)
    If Len(Trim$(subject)) = 0 Then subject = DefaultSubject(templateName)
End Sub

Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim a() As String
    ReDim a(0 To c.Count - 1)
    Dim i As Long
    For i = 1 To c.Count
        a(i - 1) = c(i)
    Next i
    CollectionToArray = a
End Function

' ---------------------------------------------------------------------------
' Built-in defaults (used to create the sheet, fill in a missing email, and Restore)
' ---------------------------------------------------------------------------

Private Function DefaultSubject(ByVal templateName As String) As String
    Select Case templateName
        Case TEMPLATE_APPROVAL
            DefaultSubject = "{ShipTo} PO # {PONumber} / K-Tron Order # {OrderNumber} : {EngType} Approval Documents"
        Case TEMPLATE_CERTIFIED
            DefaultSubject = "{ShipTo} PO # {PONumber} / K-Tron Order # {OrderNumber} : {EngType} Certified Documents"
        Case TEMPLATE_PRODUCTION
            DefaultSubject = "Production Release for Order #{OrderNumber}: {ShipTo}"
        Case TEMPLATE_CHANGE_ORDER
            DefaultSubject = "{EngTypeUpper} CHANGE ORDER: {OrderNumber} - {ShipTo}"
        Case TEMPLATE_DASHBOARD
            DefaultSubject = "Dashboard: {WeekLabel} ({StartDate} - {EndDate})"
    End Select
End Function

Private Function DefaultLines(ByVal templateName As String) As Collection
    Select Case templateName
        Case TEMPLATE_APPROVAL
            Set DefaultLines = ApprovalLines()
        Case TEMPLATE_CERTIFIED
            Set DefaultLines = CertifiedLines()
        Case TEMPLATE_PRODUCTION
            Set DefaultLines = ProductionLines()
        Case TEMPLATE_CHANGE_ORDER
            Set DefaultLines = ChangeOrderLines()
        Case TEMPLATE_DASHBOARD
            Set DefaultLines = DashboardLines()
        Case Else
            Set DefaultLines = New Collection
    End Select
End Function

Private Function DefaultHelp(ByVal templateName As String) As Variant
    Select Case templateName
        Case TEMPLATE_APPROVAL
            DefaultHelp = Array( _
                "{OrderNumber}|The order number", _
                "{ShipTo}|The customer name", _
                "{PONumber}|The customer PO number", _
                "{EngType}|Electrical or Mechanical (for this engineer)", _
                "{EngTypeLower}|electrical or mechanical, lowercase", _
                "{OtherEngType}|The other discipline (Electrical or Mechanical)", _
                "{OtherEngTypeLower}|The other discipline, lowercase", _
                "{LastUpdate}|Date the approvals were last sent (reminders)", _
                "[[IF Reminder]]...[[/IF]]|Only kept when this is a reminder / follow-up email")
        Case TEMPLATE_CERTIFIED
            DefaultHelp = Array( _
                "{OrderNumber}|The order number", _
                "{ShipTo}|The customer name", _
                "{PONumber}|The customer PO number", _
                "{EngType}|Electrical or Mechanical (for this engineer)", _
                "{EngTypeLower}|electrical or mechanical, lowercase")
        Case TEMPLATE_PRODUCTION
            DefaultHelp = Array( _
                "{OrderNumber}|The order number", _
                "{ShipTo}|The customer name", _
                "{RecipientName}|First name of the person receiving the release")
        Case TEMPLATE_CHANGE_ORDER
            DefaultHelp = Array( _
                "{OrderNumber}|The order number", _
                "{ShipTo}|The customer name", _
                "{EngTypeUpper}|ELECTRICAL or MECHANICAL", _
                "{Comments}|The comments typed on the change order form (nothing if the box is hidden)", _
                "{NoChanges}|XX (underlined) if this choice is picked on the form, otherwise __", _
                "{NoteRelease}|Same, for the next choice", _
                "{NotApproved}|Same, for the next choice", _
                "{NotReceived}|Same, for the next choice", _
                "{NotRequired}|Same, for the next choice", _
                "{AlreadyReceived}|Same, for the last choice")
        Case TEMPLATE_DASHBOARD
            DefaultHelp = Array( _
                "{ManagerName}|First name of the manager receiving the dashboard", _
                "{ThisWeek}|The list of orders and notes from this week (generated)", _
                "{NextWeek}|The list of orders on the schedule for next week (generated)", _
                "{WeekLabel}|Week label from the Weekly_Notes sheet (subject line)", _
                "{StartDate}|First day covered by the notes (subject line)", _
                "{EndDate}|Today's date (subject line)")
        Case Else
            DefaultHelp = Array()
    End Select
End Function

Private Sub AddLine(ByVal lines As Collection, ByVal text As String)
    lines.Add text
End Sub

Private Function ApprovalLines() As Collection
    Dim c As New Collection

    AddLine c, "<strong><span style='color: red;'>ATTACH CUSTOMER TRANSMITTAL, DRAWINGS PDF, &amp; DRAWING FILES AS A ZIP FILE AND DELETE THIS LINE.</span></strong><br>"
    AddLine c, "<b>K-Tron Equipment Order:</b> {OrderNumber}<BR>"
    AddLine c, "<b>Ship To:</b> {ShipTo}<BR>"
    AddLine c, "<b>Current Status:</b> Out for Approval Hold<BR><BR>"
    AddLine c, "Hello,<BR><BR>"
    AddLine c, "[[IF Reminder]]I am following up on the drawing approvals I last sent on {LastUpdate}.  Please let me know if there are any additional questions or concerns.<br><br><hr><br>[[/IF]]"
    AddLine c, "The {EngTypeLower} approval documents and drawings for your equipment order are ready and attached to this email.<BR>"
    AddLine c, "Please review them and reply with any required changes. (See markup guidelines below.)<BR>"
    AddLine c, "Please note that this email only covers your {EngTypeLower} drawings. You will receive a separate email for your {OtherEngTypeLower} approval documents.<BR><BR>"
    AddLine c, "<b>Project Status &amp; Next Steps</b>"
    AddLine c, "<ul>"
    AddLine c, "<li><b>Current Hold:</b> This order is currently on an <b>Out for Approval Hold</b> and is not yet scheduled for shipment. K-Tron will assign an official ship date once the approval process is complete.</li>"
    AddLine c, "<li><b>Change Policy:</b> Any modifications to equipment design or configuration during this review may impact your final ship date and may result in change order fees.</li>"
    AddLine c, "<li><b>How to Submit Changes:</b> We prefer mark-ups on a PDF or scanned file. However, we also accept direct electronic edits as long as every change is clearly highlighted using revision clouds, distinct markings, and notes.</li>"
    AddLine c, "</ul>"
    AddLine c, "If you have questions, comments, or concerns; please do not hesitate to contact me."

    Set ApprovalLines = c
End Function

Private Function CertifiedLines() As Collection
    Dim c As New Collection

    AddLine c, "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    AddLine c, "<strong><span style='color: red;'>ATTACH CUSTOMER TRANSMITTAL, DRAWINGS PDF, &amp; DRAWING FILES AS A ZIP FILE AND DELETE THIS LINE.</span></strong><br>"
    AddLine c, "<b> K-Tron Sales Order: </b> {OrderNumber}<BR>"
    AddLine c, "<b> Ship To: </b> {ShipTo}<BR>"
    AddLine c, "<b> Current Status: </b> Released to Production<BR><BR>"
    AddLine c, "Hello,<BR><BR>"
    AddLine c, "Please find your K-Tron {EngTypeLower} certified drawings attached for the subject job. No response is required.<BR><BR>"
    AddLine c, "Please note that any changes to the design or configuration from this point forward will result in a revised ship date, as well as Change Order fees to cover additional labor, materials, and shipping delays.<BR><BR>"
    AddLine c, "You will receive a separate email confirming your scheduled ship date. If you have any questions or concerns regarding your ship date, please contact:<BR>"
    AddLine c, "&bull; Michelle Young: <a href='mailto:michelle.young@ktron.com'>michelle.young@ktron.com</a><BR>"
    AddLine c, "&bull; Staci Voboril: <a href='mailto:staci.voboril@ktron.com'>staci.voboril@ktron.com</a><BR>"
    AddLine c, "&bull; Customer Service: 800-821-2476"

    Set CertifiedLines = c
End Function

Private Function ProductionLines() As Collection
    Dim c As New Collection

    AddLine c, "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    AddLine c, "<b> Order #{OrderNumber}</b><BR>"
    AddLine c, "<b> Customer: {ShipTo}</b><BR><BR>"
    AddLine c, "Hello {RecipientName},<BR><BR>"
    AddLine c, "Please process the attached release for order #{OrderNumber}: {ShipTo}.<BR>"
    AddLine c, "Thank You!"

    Set ProductionLines = c
End Function

Private Function ChangeOrderLines() As Collection
    Dim c As New Collection

    AddLine c, "<b>K-Tron Sales Order: {OrderNumber}: {ShipTo}<br><br>"
    AddLine c, "Could you please enter a Change Order for the following:</b><br><br>"
    AddLine c, "{Comments}"
    AddLine c, "{EngTypeUpper} APPROVAL DRAWINGS RECEIVED:<br><br>"
    AddLine c, "{NoChanges} NO CHANGES, RELEASE TO MANUFACTURING - ALL LINES<br>"
    AddLine c, "{NoteRelease} NOTE CUSTOMER CHANGES AND RELEASE TO MANUFACTURING. - ALL LINES<br>"
    AddLine c, "{NotApproved} NOT APPROVED, MAKE CUSTOMER CHANGES AND ISSUE NEW APPROVALS.<br>"
    AddLine c, "{NotReceived} NOT RECEIVED.<br>"
    AddLine c, "{NotRequired} NOT REQUIRED.<br>"
    AddLine c, "{AlreadyReceived} ALREADY RECD - ACTION ALREADY TAKEN"

    Set ChangeOrderLines = c
End Function

Private Function DashboardLines() As Collection
    Dim c As New Collection

    AddLine c, "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
    AddLine c, "Hello {ManagerName},<BR><BR>"
    AddLine c, "&emsp;<b><u>This Week</u>:</b>"
    AddLine c, "{ThisWeek}"
    AddLine c, "<BR>"
    AddLine c, "&emsp;<b><u>Next Week</u>:</b>"
    AddLine c, "{NextWeek}"
    AddLine c, "<BR>"
    AddLine c, "&emsp;<b><u>Travel/PTO</u>:</b>"
    AddLine c, "<ul><li style=margin-top:0px;margin-bottom:0px>Travel: None</li>"
    AddLine c, "<li style=margin-top:0px;margin-bottom:0px>PTO: None</li></ul>"
    AddLine c, "<BR>"
    AddLine c, "&emsp;<b><u>Concerns</u>:</b>"
    AddLine c, "<ul><li style=margin-top:0px;margin-bottom:0px>None</li></ul>"
    AddLine c, "<BR>"
    AddLine c, "&emsp;<b><u>Crisis</u>:</b>"
    AddLine c, "<ul><li style=margin-top:0px;margin-bottom:0px>None</li></ul>"

    Set DashboardLines = c
End Function
