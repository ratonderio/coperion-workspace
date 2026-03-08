VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeOrderForm 
   Caption         =   "Generate Change Order Request"
   ClientHeight    =   5424
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   9765.001
   OleObjectBlob   =   "ChangeOrderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeOrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Forms"
Option Explicit
Public FileLocation As String

Private Sub ChangeOrderTree_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    FileLocation = Data.Files(1)
    ChangeOrderTree.Nodes.Add key:=FileLocation, Text:=FileLocation

    Exit Sub
    
End Sub

Private Sub NoChangesButton_Click()

    ChangeOrderTextBox.Visible = False

End Sub

Private Sub NoteReleaseButton_Click()

    ChangeOrderTextBox.Visible = True

End Sub

Private Sub NotApprovedButton_Click()

    ChangeOrderTextBox.Visible = True

End Sub

Private Sub NotReceivedButton_Click()

    ChangeOrderTextBox.Visible = True
    
End Sub

Private Sub NotRequiredButton_Click()

    ChangeOrderTextBox.Visible = True

End Sub

Private Sub ReceivedButton_Click()

    ChangeOrderTextBox.Visible = True

End Sub

Private Sub ChangeOrderCancelButton_Click()

    ChangeOrderForm.Hide

End Sub

Private Sub ChangeOrderOkButton_Click()

    Call ChangeOrderEmail
    ChangeOrderForm.Hide
    
End Sub

Private Sub UserForm_Initialize()

    Call NoChangesButton_Click

End Sub


