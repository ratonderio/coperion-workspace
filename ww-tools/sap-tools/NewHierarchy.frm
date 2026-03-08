VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewHierarchy 
   Caption         =   "NewHierarchy"
   ClientHeight    =   3150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "NewHierarchy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewHierarchy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Modules"
Private Sub AddCodeButton_Click()

    Dim lastRow As Long
    Dim addCodeName As String
    Dim addCodeValue As String
    
    lastRow = Dropdowns.Cells(Dropdowns.Rows.Count, "H").End(xlUp).Row
    addCodeName = HierNameTextBox1.Value
    addCodeValue = HierCodeTextBox1 & "     " & HierCodeTextBox2 & "    " & HierCodeTextBox3
    
    With Dropdowns
        
        .Unprotect
    
        .Range("H" & lastRow + 1).Value = addCodeName
        .Range("I" & lastRow + 1).Value = addCodeValue
        .Range("H3:I" & lastRow + 1).Sort Key1:=.Range("H3")
                
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True
        
    
    End With
    
    AdjustHierDropdown
    MsgBox "Product Hierarchy Entered" & vbNewLine & addCodeName & ": " & addCodeValue
    CancelCodeEntry_Click

End Sub

Private Sub CancelCodeEntry_Click()

    Unload Me

End Sub

