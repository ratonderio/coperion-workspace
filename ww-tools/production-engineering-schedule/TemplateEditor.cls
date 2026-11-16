VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@IgnoreModule ProcedureNotUsed
'@Folder "Modules"
'
' Holds an open template while it is being edited so that saving the Outlook window
' writes the changes back to the .oft instead of only saving a draft.

Private WithEvents m_Item As Outlook.MailItem
Private m_Path As String

Public Sub Attach(ByVal item As Object, ByVal templatePath As String)
    Set m_Item = item
    m_Path = templatePath
End Sub

Private Sub m_Item_Write(Cancel As Boolean)
    On Error Resume Next
    m_Item.SaveAs m_Path, 5 'olTemplate
    If Err.Number <> 0 Then
        MsgBox "Could not save the template to:" & vbCr & m_Path & vbCr & vbCr & _
               "Use File > Save As > Outlook Template and overwrite that file.", vbExclamation, "Email template"
    End If
End Sub

Private Sub m_Item_Close(Cancel As Boolean)
    EmailTemplates.ReleaseEditor Me
End Sub
