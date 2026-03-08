VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Approvals Received - Continue?"
   ClientHeight    =   4032
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   7155
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'@IgnoreModule MemberNotOnInterface
'@Folder "Forms"
Dim CustomDate As Date
Dim Today As Date
Dim FirstClick As Boolean
Public OkCase As Integer
Public orderNumber As String
Public CustomerName As String
Public RowNumber As String

Private Sub ApprovalReminderButton_Click()
    
    ScheduleWS.CheckBoxes(RowNumber).Value = xlOff
    
    SendApprovalEmail True, ScheduleWS.Cells(RowNumber, "BA").Value, orderNumber
    UserForm_Close

End Sub

Private Sub UpdateApprovalsDateButton_Click()

    ScheduleWS.CheckBoxes(RowNumber).Value = xlOff
    
    UpdateApprovalsDate
    Order_Query
    UserForm_Close

End Sub

Private Sub UpdateApprovalsDate()
    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open
    
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    
    If ScheduleWS.Range("AJ2").Value = "PC" Then
        oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Out='" & Date & "'WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
        oRs.Open
    ElseIf ScheduleWS.Range("AJ2").Value = "ME" Then
        oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Out='" & Date & "'WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
        oRs.Open
    End If
    
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    
End Sub

Private Sub CancelButton_Click()

    ScheduleWS.CheckBoxes(RowNumber).Value = IIf(ScheduleWS.CheckBoxes(RowNumber).Value = 1, xlOff, xlOn)
    UserForm_Close

End Sub

Private Sub ChangeOrderButton_Click()
    
    CloseButtonSettings ChangeOrderForm, False
    ChangeOrderForm.show
    
End Sub

Private Sub ResetButton_Click()

    OkCase = 1

End Sub


Private Sub UpdateButton_Click()

    OkCase = 2

End Sub

Private Sub CustomUpdateButton_Click()
    OkCase = 3
    If FirstClick Then
        MsgBox "Enter Full Year (YYYY) When Submitting Custom Date"
        FirstClick = False
    End If

End Sub

Private Sub AcceptContinueButton_Click()
    
    OkCase = 4
    
End Sub

Private Sub OkButton_Click()

    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open
    
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    
    If Not UserForm1.ResetButton And Not UserForm1.CustomUpdateButton And Not UserForm1.UpdateButton And Not UserForm1.AcceptContinueButton Then
        OkCase = 0
    End If

    Select Case OkCase
        Case 0
            MsgBox "No Choice Selected"
            ScheduleWS.CheckBoxes(RowNumber).Value = IIf(ScheduleWS.CheckBoxes(RowNumber).Value, xlOff, xlOn)
            UserForm_Close
        Case 1
            If UCase$(ScheduleWS.Range("AJ2").Value) = "PC" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Back=0 WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                oRs.Open
            ElseIf UCase$(ScheduleWS.Range("AJ2").Value) = "ME" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Back=0 WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                oRs.Open
            End If
            ScheduleWS.CheckBoxes(RowNumber).Value = xlOff
            UserForm_Close
        Case 2
            If UCase$(ScheduleWS.Range("AJ2").Value) = "PC" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Back='" & Today & "' WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                oRs.Open
            ElseIf UCase$(ScheduleWS.Range("AJ2").Value) = "ME" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Back='" & Today & "' WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                oRs.Open
            End If
            ScheduleWS.CheckBoxes(RowNumber).Value = xlOn
            UserForm_Close
        Case 3
            On Error GoTo ErrorHandler
            CustomDate = UserForm1.CustomDateTextBox.Value
            If UCase$(ScheduleWS.Range("AJ2").Value) = "PC" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Back='" & CustomDate & "' WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                oRs.Open
            ElseIf UCase$(ScheduleWS.Range("AJ2").Value) = "ME" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Back='" & CustomDate & "' WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                oRs.Open
            End If
            ScheduleWS.CheckBoxes(RowNumber).Value = xlOn
            UserForm_Close
        Case 4
            If ScheduleWS.CheckBoxes(RowNumber).Value = 1 Then
                If UCase$(ScheduleWS.Range("AJ2").Value) = "PC" Then
                    oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Back='" & Today & "' WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                    oRs.Open
                ElseIf UCase$(ScheduleWS.Range("AJ2").Value) = "ME" Then
                    oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Back='" & Today & "' WHERE Prod_Eng.Order_Num=" & orderNumber & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                    oRs.Open
                End If
            Else
                ScheduleWS.CheckBoxes(RowNumber).Value = xlOn
            End If
            UserForm_Close
    End Select
    
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Invalid Date or Date Format"
    UserForm1.CustomDateTextBox = Today
    ScheduleWS.CheckBoxes(RowNumber).Value = IIf(ScheduleWS.CheckBoxes(RowNumber).Value, xlOff, xlOn)
    
    UserForm_Close
    
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    Exit Sub

End Sub

Private Sub UserForm_Initialize()

    CloseButtonSettings Me, False
    RowNumber = Application.Caller
    orderNumber = ScheduleWS.Range("B" & RowNumber).Value
    CustomerName = ScheduleWS.Range("C" & RowNumber).Value
    ProdEng = ScheduleWS.Range("AE2").Value
    Today = ScheduleWS.Range("Z1").Value
    UserForm1.TodayLabel = "Today's Date: " & Today
    UserForm1.CustomDateTextBox = Today
    FirstClick = True
    AcceptContinueButton_Click

End Sub

Private Sub UserForm_Close()

    Unload ChangeOrderForm
    Unload Me

End Sub

