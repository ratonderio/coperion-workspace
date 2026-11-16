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
    On Error GoTo Fail
    EmailTemplates.SaveTemplateFile m_Item, m_Path
    Exit Sub
Fail:
    MsgBox "Could not save the template:" & vbCr & vbCr & Err.Description & vbCr & vbCr & _
           "You can also use File > Save As > Outlook Template and overwrite the same file.", vbExclamation, "Email template"
End Sub

Private Sub m_Item_Close(Cancel As Boolean)
    EmailTemplates.ReleaseEditor Me
End Sub
