Attribute VB_Name = "EmailTemplates"
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder "Modules"
'
' Editable email templates.
'
' Each email can be backed by an Outlook template (.oft) in TEMPLATE_FOLDER. Anyone
' can change the wording by editing that file in Outlook (see EditTemplate), with no
' VBA changes. The built-in default text below is the fallback, and is also what
' RestoreDefaultTemplate writes back out, so there is one source of truth.
'
' Template syntax (plain text typed into the email):
'   {Token}                     replaced with the value for the order, e.g. {OrderNumber}
'   [[IF Name]] ... [[/IF]]     kept only when the flag Name is on, otherwise removed
'
' Do not put your own signature in a template; the sender's signature is added
' automatically when the email is created.

Public Const TEMPLATE_FOLDER As String = "\\cop.hi.corp\fileserver\CFE\CPS\WWQ-FS01\orders\Common Files\CPE_Schedule\Templates\"

Public Const TEMPLATE_APPROVAL As String = "ApprovalEmail"

Private Const OL_TEMPLATE As Long = 5
Private Const OL_DISCARD As Long = 1
Private Const OL_FORMAT_HTML As Long = 2

Private m_Editors As Collection

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

Private Function HtmlEncode(ByVal text As String) As String
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    HtmlEncode = text
End Function

' ---------------------------------------------------------------------------
' Rendering
' ---------------------------------------------------------------------------

'Returns the finished HTML body and sets emailSubject. Uses the .oft when it exists
'and can be read, otherwise the built-in default.
Public Function RenderTemplate(ByVal templateName As String, ByVal tokens As Object, ByRef emailSubject As String) As String
    Dim html As String: html = DefaultBody(templateName)
    emailSubject = DefaultSubject(templateName)

    Dim path As String: path = TemplatePath(templateName)
    Dim readFailed As Boolean

    On Error Resume Next
    Dim exists As Boolean: exists = (Len(Dir(path)) > 0)
    If Err.Number <> 0 Then readFailed = True: Err.Clear
    On Error GoTo 0

    If exists Then
        On Error Resume Next
        Dim app As Object: Set app = CreateObject("Outlook.Application")
        Dim tmpl As Object: Set tmpl = app.CreateItemFromTemplate(path)
        If Err.Number = 0 And Not tmpl Is Nothing Then
            If Len(tmpl.HTMLBody) > 0 Then html = tmpl.HTMLBody Else readFailed = True
            If Len(tmpl.Subject) > 0 Then emailSubject = tmpl.Subject
            tmpl.Close OL_DISCARD
        Else
            readFailed = True
        End If
        Err.Clear
        On Error GoTo 0
    End If

    If readFailed Then
        MsgBox "Could not read the email template:" & vbCr & path & vbCr & vbCr & _
               "The built-in default wording will be used instead.", vbExclamation, "Email template"
    End If

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
    MsgBox "The " & templateName & " template has placeholders that were not recognized:" & found & vbCr & vbCr & _
           "Check the spelling in the template. The email was still created.", vbExclamation, "Email template"
End Sub

' ---------------------------------------------------------------------------
' Editing
' ---------------------------------------------------------------------------

Public Sub EditApprovalTemplate()
    EditTemplate TEMPLATE_APPROVAL
End Sub

Public Sub RestoreApprovalTemplate()
    RestoreDefaultTemplate TEMPLATE_APPROVAL
End Sub

'Opens the template for editing. Closing the window and answering Yes to the save
'prompt writes the changes back to the .oft (see TemplateEditor). If that does not
'happen, use File > Save As > Outlook Template and overwrite the same file.
Public Sub EditTemplate(ByVal templateName As String)
    Dim path As String: path = TemplatePath(templateName)
    If Not TemplateExists(path) Then RestoreDefaultTemplate templateName, True

    Dim app As Object: Set app = CreateObject("Outlook.Application")
    Dim item As Object: Set item = app.CreateItemFromTemplate(path)

    Dim editor As TemplateEditor: Set editor = New TemplateEditor
    editor.Attach item, path
    If m_Editors Is Nothing Then Set m_Editors = New Collection
    m_Editors.Add editor, CStr(ObjPtr(editor))

    item.Display
End Sub

'Writes the built-in default out as the .oft, replacing any existing file.
Public Sub RestoreDefaultTemplate(ByVal templateName As String, Optional ByVal silent As Boolean = False)
    Dim path As String: path = TemplatePath(templateName)
    If Not silent Then
        If MsgBox("Replace the " & templateName & " template with the original wording?" & vbCr & vbCr & _
                  "Any edits to " & path & " will be lost.", vbYesNo + vbQuestion, "Restore default template") <> vbYes Then Exit Sub
    End If

    EnsureTemplateFolder

    Dim app As Object: Set app = CreateObject("Outlook.Application")
    Dim mail As Object: Set mail = app.CreateItem(0)
    mail.BodyFormat = OL_FORMAT_HTML
    mail.Subject = DefaultSubject(templateName)
    mail.HTMLBody = "<html><body>" & DefaultBody(templateName) & "</body></html>"
    mail.SaveAs path, OL_TEMPLATE
    mail.Close OL_DISCARD

    If Not silent Then MsgBox "Template restored:" & vbCr & path, vbInformation, "Restore default template"
End Sub

Public Sub ReleaseEditor(ByVal editor As TemplateEditor)
    On Error Resume Next
    m_Editors.Remove CStr(ObjPtr(editor))
End Sub

' ---------------------------------------------------------------------------
' Files
' ---------------------------------------------------------------------------

Private Function TemplatePath(ByVal templateName As String) As String
    TemplatePath = TEMPLATE_FOLDER & templateName & ".oft"
End Function

Private Function TemplateExists(ByVal path As String) As Boolean
    On Error Resume Next
    TemplateExists = (Len(Dir(path)) > 0)
End Function

Private Sub EnsureTemplateFolder()
    If Len(Dir(TEMPLATE_FOLDER, vbDirectory)) = 0 Then MkDir Left$(TEMPLATE_FOLDER, Len(TEMPLATE_FOLDER) - 1)
End Sub

' ---------------------------------------------------------------------------
' Built-in defaults (fallback, and the source for Restore default)
' ---------------------------------------------------------------------------

Private Function DefaultSubject(ByVal templateName As String) As String
    Select Case templateName
        Case TEMPLATE_APPROVAL
            DefaultSubject = "{ShipTo} PO # {PONumber} / Coperion Order # {OrderNumber} : {EngType} Approval Documents"
    End Select
End Function

Private Function DefaultBody(ByVal templateName As String) As String
    Select Case templateName
        Case TEMPLATE_APPROVAL
            DefaultBody = DefaultApprovalBody()
    End Select
End Function

'Tokens: {OrderNumber} {ShipTo} {PONumber} {EngType} {EngTypeLower} {OtherEngTypeLower} {LastUpdate}
'Flags:  Reminder (follow-up email)
Private Function DefaultApprovalBody() As String
    Dim s As String

    s = "<strong><span style='color: red;'>ATTACH CUSTOMER TRANSMITTAL, DRAWINGS PDF, &amp; DRAWING FILES AS A ZIP FILE AND DELETE THIS LINE.</span></strong><br>"
    s = s & "<b>K-Tron Equipment Order:</b> {OrderNumber}<BR>"
    s = s & "<b>Ship To:</b> {ShipTo}<BR>"
    s = s & "<b>Current Status:</b> Out for Approval Hold<BR><BR>"

    s = s & "Hello,<BR><BR>"

    s = s & "[[IF Reminder]]I am following up on the drawing approvals I last sent on {LastUpdate}.  " & _
            "Please let me know if there are any additional questions or concerns.<br><br><hr><br>[[/IF]]"

    s = s & "The {EngTypeLower} approval documents and drawings for your equipment order are ready and attached to this email.<BR>"
    s = s & "Please review them and reply with any required changes. (See markup guidelines below.)<BR>"
    s = s & "Please note that this email only covers your {EngTypeLower} drawings. You will receive a separate email for your " & _
            "{OtherEngTypeLower} approval documents.<BR><BR>"
    s = s & "<b>Project Status &amp; Next Steps</b>"
    s = s & "<ul>"
    s = s & "<li><b>Current Hold:</b> This order is currently on an <b>Out for Approval Hold</b> and is not yet scheduled for shipment. " & _
            "K-Tron will assign an official ship date once the approval process is complete.</li>"
    s = s & "<li><b>Change Policy:</b> Any modifications to equipment design or configuration during this review may impact your final " & _
            "ship date and may result in change order fees.</li>"
    s = s & "<li><b>How to Submit Changes:</b> We prefer mark-ups on a PDF or scanned file. However, we also accept direct electronic edits " & _
            "as long as every change is clearly highlighted using revision clouds, distinct markings, and notes.</li>"
    s = s & "</ul>"
    s = s & "If you have questions, comments, or concerns; please do not hesitate to contact me."

    DefaultApprovalBody = s
End Function
