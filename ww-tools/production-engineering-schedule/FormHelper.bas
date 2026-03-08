Attribute VB_Name = "FormHelper"
'@Folder "__Modules"
'@IgnoreModule
'Code pulled from https://exceloffthegrid.com/hide-or-disable-a-vba-userform-x-close-button/


'Include this code at the top of the module
Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const SC_CLOSE = &HF060

#If VBA7 Then

    Private Declare PtrSafe Function FindWindowA _
    Lib "user32" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function DeleteMenu _
    Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetSystemMenu _
    Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
        
#Else

    Private Declare Function FindWindowA _
                             Lib "user32" (ByVal lpClassName As String, _
                                           ByVal lpWindowName As String) As Long
    Private Declare Function DeleteMenu _
                             Lib "user32" (ByVal hMenu As Long, _
                                           ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Declare Function GetSystemMenu _
                            Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
        
#End If

'Include this code in the same module as the API calls above
Public Sub CloseButtonSettings(frm As Object, show As Boolean)

    Dim windowHandle As Long
    Dim menuHandle As Long
    windowHandle = FindWindowA(vbNullString, frm.Caption)

    If show = True Then

        menuHandle = GetSystemMenu(windowHandle, 1)

    Else

        menuHandle = GetSystemMenu(windowHandle, 0)
        DeleteMenu menuHandle, SC_CLOSE, 0&

    End If

End Sub

'https://www.rondebruin.nl/win/s1/outlook/bmail4.htm
Function GetBoiler(ByVal sFile As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function

