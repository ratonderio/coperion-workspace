Attribute VB_Name = "Module1"
Public I As Long, J As Long, K As Long, JJ As Long, KK As Long, DnLn As Long, DB_Col As Long, DB_Row As Long
Public Ord As Long, Ignore As Long, ActiveYr As Long, DescLen As Long, NextLine As Long, NewInteger As Long
Public TotalLines As Long, TotalOrders As Long, SearchOrd As Long, Jmax As Long, Imax As Long, NewLong As Long, JJmax As Long
Public PrevOrd As Long, LineNum As Long, PrevLine As Long, PC_Num As Long, PCnum As Long, ConfMon As Long, ConfYr As Long
Public Index As Long, LastRow As Long
'
Public CExD As Date, PO_Date As Date, CreatedOn As Date, Released As Date, PC_Rel As Date, ME_Rel As Date, PrevDate As Date
Public IntoProdEng As Date, PC_A_Out As Date, ME_A_Out As Date, PC_A_Back As Date, ME_A_Back As Date
Public StartSpan As Date, EndSpan As Date, SchdRel As Date, ActRel As Date, NewDate As Date, SDt As Date, EDt As Date
'
Public Docs As String, PC_Eng As String, ME_Eng As String, PCStatus As String, MEStatus As String, PrjLvl As String, PM As String
Public CustName As String, MatNum As String, Desc As String, SoldTo As String, Customer As String, OtherEng As String
Public ProdCont As String, SU As String, AssyStr As String, ProdEng As String, EType As String, OrdStatus As String
Public MngrOR As String, Net1 As String, Net2 As String, DB_Head(60) As String, U_ID As String, NewString As String, CurStat(50) As String
Public OrdStr As String, OrdFold As String, TxtLine As String, JSubFold As String, EngStatus As String
'
Public LineQuan As Double, EstHrs As Double, NewNumeric As Double
Public LineVal As Currency, PerVal As Currency, OrdVal As Currency, FieldServ As Currency, PrevVal As Currency
Public ClrRange As Range, SAB_Appr As Boolean, Flag1 As Boolean
'
Public Const WkBk = "ProdEng_Manager.xlsm"
Public Const Sht1 = "Query_Results"
Public Const Sht2 = "Schedule"
Public Const Sht3 = "Released"
Public Const Sht4 = "Order_Details"
Public Const Sht5 = "Query_Results2"
Public Const Sht6 = "Updates"
Public Const Sht7 = "Backlog"
Public Const ShtEquip = "Open_Equip"
Public Const ShtParts = "Open_Parts"
Public Const ShtShip = "Shipped"
Public Const ShtAD = "AddDeletes"
Public Const ShtPMC = "PMC_Status"

'\\USWWQ-P-FS01
Public Const Path_SAP = "W:\Manufacturing\Projects\Data"

Sub Order_Details()
'
Set ClrRange = Workbooks(WkBk).Worksheets(Sht5).Range("B3:BF5000")
ClrRange.ClearContents

Ord = Workbooks(WkBk).Worksheets(Sht4).Cells(3, 2).Value

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.Order_Num = " & Ord & ";"
' Retrieve records
oRS.Open
Workbooks(WkBk).Worksheets(Sht5).Range("B3").CopyFromRecordset oRS

Workbooks(WkBk).Worksheets(Sht4).Cells(3, 3).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 4).Value   ' Customer
Workbooks(WkBk).Worksheets(Sht4).Cells(3, 5).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(1, 62).Value  ' Process Control
Workbooks(WkBk).Worksheets(Sht4).Cells(3, 6).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(2, 62).Value  ' Mechanical Designer
Workbooks(WkBk).Worksheets(Sht4).Cells(3, 7).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(47, 62).Value ' Project Manager
Workbooks(WkBk).Worksheets(Sht4).Cells(3, 8).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(48, 62).Value ' Project Engineer
Workbooks(WkBk).Worksheets(Sht4).Cells(3, 9).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 13).Value  ' Document Types

Workbooks(WkBk).Worksheets(Sht4).Cells(5, 5).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 34).Value  ' Project Level
Workbooks(WkBk).Worksheets(Sht4).Cells(5, 6).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 33).Value  ' MultiPlant

If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 47).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(5, 7).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 47).Value  ' Scheduled Pre-Order
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(5, 7).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 49).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(5, 8).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 49).Value  ' Actual ME Pre-Order
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(5, 8).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 48).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(5, 9).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 48).Value  ' Actual PC Pre-Order
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(5, 9).Value = ""
End If

If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 41).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 5).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 41).Value  ' Initial Approvals Out
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 5).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 43).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 6).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 43).Value  ' ME Approvals Out
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 6).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 45).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 7).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 45).Value  ' ME Approvals Back
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 7).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 42).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 8).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 42).Value  ' PC Approvals Out
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 8).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 44).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 9).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 44).Value  ' PC Approvals Back
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(7, 9).Value = ""
End If

If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 57).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 5).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 57).Value  ' Initial Production Release
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 5).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 51).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 6).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 51).Value  ' Final Mechanical Production Release
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 6).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 53).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 7).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 53).Value  ' Actual Mechanical Production Release
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 7).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 50).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 8).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 50).Value  ' Final Control Production Release
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 8).Value = ""
End If
If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 52).Value > 3 Then
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 9).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 52).Value  ' Actual Control Production Release
Else
    Workbooks(WkBk).Worksheets(Sht4).Cells(9, 9).Value = ""
End If

Workbooks(WkBk).Worksheets(Sht4).Cells(11, 5).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 36).Value 'ME Status
Workbooks(WkBk).Worksheets(Sht4).Cells(13, 5).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 35).Value 'ME Status
Workbooks(WkBk).Worksheets(Sht4).Cells(15, 9).Value = Workbooks(WkBk).Worksheets(Sht5).Cells(3, 12).Value 'Industry

AssyStr = ""
For I = 3 To 60
    Ord = Workbooks(WkBk).Worksheets(Sht5).Cells(I, 3).Value
    If Ord < 3 Then Exit For
    AssyStr = AssyStr & "[" & Workbooks(WkBk).Worksheets(Sht5).Cells(I, 5).Value & "] " & Mid(Workbooks(WkBk).Worksheets(Sht5).Cells(I, 7).Value, 1, 40) & Chr(10)
Next I
Workbooks(WkBk).Worksheets(Sht4).Cells(5, 2).Value = AssyStr


End Sub
Sub Eng_AddDeletes()
'
Set ClrRange = Workbooks(WkBk).Worksheets(ShtAD).Range("A5:H5000")
ClrRange.ClearContents

SDt = Workbooks(WkBk).Worksheets(ShtAD).Cells(2, 4).Value
EDt = Workbooks(WkBk).Worksheets(ShtAD).Cells(3, 4).Value
ProdEng = Workbooks(WkBk).Worksheets(Sht2).Cells(2, 1).Value

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
' Retrieve records
oRS.Source = "SELECT cpe_add_delete.engineer, datestamp, orderno, customer, mycount, errorcode FROM cpe_add_delete WHERE cpe_add_delete.datestamp BETWEEN '" & SDt & "' AND '" & EDt & "' AND cpe_add_delete.engineer = '" & ProdEng & "' ORDER BY cpe_add_delete.datestamp"
oRS.Open
Workbooks(WkBk).Worksheets(ShtAD).Range("C5").CopyFromRecordset oRS

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
Sub MissingAddDeletes()
'
'Set ClrRange = Workbooks(WkBk).Worksheets(ShtAD).Range("A5:H5000")
'ClrRange.ClearContents
Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

For I = 2 To 23

    U_ID = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 21).Value
    Ord = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 19).Value
    ProdEng = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 17).Value
    NewDate = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 18).Value
    NewString = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 20).Value
    
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    ' Retrieve records
    oRS.Source = "UPDATE cpe_add_delete SET cpe_add_delete.errorcode = '" & NewString & "' WHERE cpe_add_delete.uniqid = '" & U_ID & "'"
    oRS.Open

Next I

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing
End Sub
Sub Order_Query()
'
Set ClrRange = Workbooks(WkBk).Worksheets(Sht2).Range("B4:N500")
ClrRange.ClearContents

Set ClrRange = Workbooks(WkBk).Worksheets(Sht3).Range("B4:L500")
ClrRange.ClearContents
Set ClrRange = Workbooks(WkBk).Worksheets(Sht3).Range("N4:O500")
ClrRange.ClearContents

'GoTo TEMP1

Set ClrRange = Workbooks(WkBk).Worksheets(Sht1).Range("B3:BF25000")
ClrRange.ClearContents

ProdEng = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 60).Value
EType = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 63).Value
ActiveYr = Workbooks(WkBk).Worksheets(Sht2).Cells(1, 4).Value
Ord = 0
DescLen = 40

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
If EType = "PC" Then
    oRS.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.PC1 = '" & ProdEng & "' AND Active_Year = " & ActiveYr & " ORDER BY Prod_Eng.PC_Rel_F;"
Else
    oRS.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.ME1 = '" & ProdEng & "' AND Active_Year = " & ActiveYr & " ORDER BY Prod_Eng.ME_Rel_F;"
End If
' Retrieve records
oRS.Open
Workbooks(WkBk).Worksheets(Sht1).Range("B3").CopyFromRecordset oRS

TEMP1:

J = 3: JJ = 3
' Released Orders
For I = 3 To 25000
    PrevOrd = Ord
    Ord = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 3).Value
    If Ord < 3 Then Exit For
    
    If Ord = PrevOrd Then
        Flag1 = True
        GoTo DETAILS
    Else
        Flag1 = False
    End If
    
    If EType = "PC" Then
        OrdStatus = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 35).Value
        SchdRel = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 50).Value
        ActRel = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 52).Value
        OtherEng = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 18).Value
        MngrOR = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 37).Value
        EstHrs = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 25).Value + Workbooks(WkBk).Worksheets(Sht1).Cells(I, 26).Value
    Else
        OrdStatus = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 36).Value
        SchdRel = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 51).Value
        ActRel = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 53).Value
        OtherEng = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 16).Value
        MngrOR = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 38).Value
        EstHrs = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 27).Value + Workbooks(WkBk).Worksheets(Sht1).Cells(I, 28).Value
    End If
    
    If Workbooks(WkBk).Worksheets(Sht1).Cells(I, 41).Value > 3 Then
        Apps_Out = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 41).Value
    Else
        Apps_Out = 2
    End If

    If OrdStatus = "RELEASED" Then
        J = J + 1
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 2).Value = Ord
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 3).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 4).Value
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 4).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 8).Value
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 6).Value = OtherEng
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 7).Value = SchdRel
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 8).Value = ActRel
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 9).Value = EstHrs
        If Workbooks(WkBk).Worksheets(Sht1).Cells(I, 54).Value > 3 Then
            Workbooks(WkBk).Worksheets(Sht3).Cells(J, 11).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 54).Value
        Else
        End If
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 14).Value = MngrOR
    Else
        JJ = JJ + 1
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 2).Value = Ord
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 3).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 4).Value
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 4).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 8).Value
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 5).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 13).Value
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 6).Value = OtherEng
        If Apps_Out > 3 Then
            Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 7).Value = Apps_Out
        Else
        End If
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 8).Value = SchdRel
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 9).Value = EstHrs
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 12).Value = OrdStatus
        CurStat(JJ) = OrdStatus
        'Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 11).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 54).Value
        'Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 14).Value = MngrOR
    End If

DETAILS:
    If Flag1 = True And EType = "PC" And OrdStatus <> "RELEASED" Then
        EstHrs = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 25).Value + Workbooks(WkBk).Worksheets(Sht1).Cells(I, 26).Value
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 9).Value = Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 9).Value + EstHrs
    Else
    End If
    If Flag1 = True And EType = "ME" And OrdStatus <> "RELEASED" Then
        EstHrs = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 27).Value + Workbooks(WkBk).Worksheets(Sht1).Cells(I, 28).Value
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 9).Value = Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 9).Value + EstHrs
    Else
    End If
    If Flag1 = True And EType = "PC" And OrdStatus = "RELEASED" Then
        EstHrs = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 25).Value + Workbooks(WkBk).Worksheets(Sht1).Cells(I, 26).Value
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 9).Value = Workbooks(WkBk).Worksheets(Sht3).Cells(J, 9).Value + EstHrs
    Else
    End If
    If Flag1 = True And EType = "ME" And OrdStatus = "RELEASED" Then
        EstHrs = Workbooks(WkBk).Worksheets(Sht1).Cells(I, 27).Value + Workbooks(WkBk).Worksheets(Sht1).Cells(I, 28).Value
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 9).Value = Workbooks(WkBk).Worksheets(Sht3).Cells(J, 9).Value + EstHrs
    Else
    End If
    If Flag1 = False And OrdStatus = "RELEASED" Then
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 5).Value = Workbooks(WkBk).Worksheets(Sht3).Cells(J, 5).Value & "[" & _
            Workbooks(WkBk).Worksheets(Sht1).Cells(I, 5).Value & "] " & Mid$(Workbooks(WkBk).Worksheets(Sht1).Cells(I, 7).Value, 1, DescLen)
    Else
    End If
    If Flag1 = True And OrdStatus = "RELEASED" Then
        Workbooks(WkBk).Worksheets(Sht3).Cells(J, 5).Value = Workbooks(WkBk).Worksheets(Sht3).Cells(J, 5).Value & Chr(10) & "[" & _
            Workbooks(WkBk).Worksheets(Sht1).Cells(I, 5).Value & "] " & Mid$(Workbooks(WkBk).Worksheets(Sht1).Cells(I, 7).Value, 1, DescLen)
    Else
    End If
    If Flag1 = False And OrdStatus <> "RELEASED" Then
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 11).Value = Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 11).Value & "[" & _
            Workbooks(WkBk).Worksheets(Sht1).Cells(I, 5).Value & "] " & Mid$(Workbooks(WkBk).Worksheets(Sht1).Cells(I, 7).Value, 1, DescLen)
    Else
    End If
    If Flag1 = True And OrdStatus <> "RELEASED" Then
        Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 11).Value = Workbooks(WkBk).Worksheets(Sht2).Cells(JJ, 11).Value & Chr(10) & "[" & _
            Workbooks(WkBk).Worksheets(Sht1).Cells(I, 5).Value & "] " & Mid$(Workbooks(WkBk).Worksheets(Sht1).Cells(I, 7).Value, 1, DescLen)
    Else
    End If

Next I
Jmax = J
JJmax = JJ

For I = 4 To Jmax
    Set ClrRange = Workbooks(WkBk).Worksheets(Sht1).Range("BM3:BS5000")
    ClrRange.ClearContents
    
    Net1 = Workbooks(WkBk).Worksheets(Sht3).Cells(I, 4).Value
    oRS.Close
    oRS.Source = "SELECT * FROM cpe_time WHERE cpe_time.jobno='" & Net1 & "'"
    oRS.Open
    Workbooks(WkBk).Worksheets(Sht1).Range("BM3").CopyFromRecordset oRS
    If EType = "PC" Then
        Workbooks(WkBk).Worksheets(Sht3).Cells(I, 10).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 72).Value
    Else
        Workbooks(WkBk).Worksheets(Sht3).Cells(I, 10).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 73).Value
    End If

Next I

For I = 4 To Jmax
    Set ClrRange = Workbooks(WkBk).Worksheets(Sht1).Range("BX3:CD5000")
    ClrRange.ClearContents
    
    Ord = Workbooks(WkBk).Worksheets(Sht3).Cells(I, 2).Value
    oRS.Close
    oRS.Source = "SELECT * FROM cpe_schedule WHERE cpe_schedule.orderno = " & Ord & " ORDER BY cpe_schedule.datestamp"
    oRS.Open
    Workbooks(WkBk).Worksheets(Sht1).Range("BX3").CopyFromRecordset oRS
    For K = 3 To 100
        If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 78).Value < 3 Then Exit For
        If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 77).Value = ProdEng Then
            If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 81).Value > "" Then
                Workbooks(WkBk).Worksheets(Sht3).Cells(I, 15).Value = Workbooks(WkBk).Worksheets(Sht3).Cells(I, 15).Value & Workbooks(WkBk).Worksheets(Sht1).Cells(K, 86) & Workbooks(WkBk).Worksheets(Sht1).Cells(K, 81) & Chr(10)
            Else
            End If
        Else
        End If
    Next K
    
Next I

For I = 4 To JJmax
    Set ClrRange = Workbooks(WkBk).Worksheets(Sht1).Range("BM3:BS5000")
    ClrRange.ClearContents
    
    Net1 = Workbooks(WkBk).Worksheets(Sht2).Cells(I, 4).Value
    oRS.Close
    oRS.Source = "SELECT * FROM cpe_time WHERE cpe_time.jobno='" & Net1 & "'"
    oRS.Open
    Workbooks(WkBk).Worksheets(Sht1).Range("BM3").CopyFromRecordset oRS
    If EType = "PC" Then
        Workbooks(WkBk).Worksheets(Sht2).Cells(I, 10).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 72).Value
    Else
        Workbooks(WkBk).Worksheets(Sht2).Cells(I, 10).Value = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 73).Value
    End If

Next I

For I = 4 To JJmax
    Set ClrRange = Workbooks(WkBk).Worksheets(Sht1).Range("BX3:CD5000")
    ClrRange.ClearContents
    
    Ord = Workbooks(WkBk).Worksheets(Sht2).Cells(I, 2).Value
    oRS.Close
    oRS.Source = "SELECT * FROM cpe_schedule WHERE cpe_schedule.orderno = " & Ord & " ORDER BY cpe_schedule.datestamp"
    oRS.Open
    Workbooks(WkBk).Worksheets(Sht1).Range("BX3").CopyFromRecordset oRS
    For K = 3 To 100
        If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 78).Value < 3 Then Exit For
        If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 77).Value = ProdEng Then
            If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 81).Value > "" Then
                Workbooks(WkBk).Worksheets(Sht2).Cells(I, 14).Value = Workbooks(WkBk).Worksheets(Sht2).Cells(I, 14).Value & Workbooks(WkBk).Worksheets(Sht1).Cells(K, 86) & Workbooks(WkBk).Worksheets(Sht1).Cells(K, 81) & Chr(10)
            Else
            End If
        Else
        End If
    Next K
    
Next I
For I = 4 To 50

Ord = Workbooks(WkBk).Worksheets(Sht2).Cells(I, 2).Value
If Ord < 2 Then Exit For
OrdStr = Str(Ord)
OrdFold = "J:\Orders\" & Mid(OrdStr, 2, 7) & "000\" & Mid(OrdStr, 2, 10) & "*"
TxtLine = Dir(OrdFold, vbDirectory)
Workbooks(WkBk).Worksheets(Sht2).Cells(I, 28).Value = TxtLine
JSubFold = Mid(TxtLine, 1, 7) & "000\"
Workbooks(WkBk).Worksheets(Sht2).Cells(I, 2).Value = Ord
Workbooks(WkBk).Worksheets(Sht2).Range(Cells(I, 2), Cells(I, 2)).Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        "J:\Orders\" & JSubFold & TxtLine
Next I

    Workbooks(WkBk).Worksheets(Sht3).Select
    Rows("4:199").EntireRow.AutoFit

    Workbooks(WkBk).Worksheets(Sht2).Select
    Rows("4:99").EntireRow.AutoFit
    Range("A3").Select

End Sub
Sub Backlog()
' Sht5 = "Query_Results2"
' Sht7 = "Backlog"

Set ClrRange = Workbooks(WkBk).Worksheets(Sht7).Range("B4:H500")
ClrRange.ClearContents
Set ClrRange = Workbooks(WkBk).Worksheets(Sht7).Range("M4:S500")
ClrRange.ClearContents

'EType = Workbooks(WkBk).Worksheets(Sht1).Cells(1, 63).Value

ActiveYr = Workbooks(WkBk).Worksheets(Sht5).Cells(1, 70).Value
Ord = 0
DescLen = 40

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

K = 0

For I = 1 To 16  ' 16
    
    Set ClrRange = Workbooks(WkBk).Worksheets(Sht5).Range("BR3:CB5000")
    ClrRange.ClearContents

    ProdEng = Workbooks(WkBk).Worksheets(Sht7).Cells(1, I + 26).Value
    If I = 8 Then K = 0
    
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    If I < 9 Then
        oRS.Source = "SELECT Prod_Eng.Order_Num, Customer_Name, DocTypes, ME_Rel_F, ME1, ME_Est, Network1, Network2, ME_Status, ME_CO_Est, Line_Num FROM Prod_Eng WHERE Prod_Eng.Active_Year = " & ActiveYr & " AND ME1 = '" & ProdEng & "' ORDER BY Prod_Eng.Order_Num, Line_Num, ME_Rel_F;"
    Else
        oRS.Source = "SELECT Prod_Eng.Order_Num, Customer_Name, DocTypes, PC_Rel_F, PC1, PC_Est, Network1, Network2, PC_Status, PC_CO_Est, Line_Num FROM Prod_Eng WHERE Prod_Eng.Active_Year = " & ActiveYr & " AND PC1 = '" & ProdEng & "' ORDER BY Prod_Eng.Order_Num, Line_Num, PC_Rel_F;"
    End If
    
    oRS.Open
    Workbooks(WkBk).Worksheets(Sht5).Range("BR3").CopyFromRecordset oRS
    
    If Workbooks(WkBk).Worksheets(Sht5).Cells(3, 70).Value < 2 Then GoTo HERE
    
    If I > 8 Then KK = 11 Else KK = 0
    
    EstHrs = 0#
    
    For J = 3 To 1000
        PrevOrd = Ord
        Ord = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 70).Value
       
        If J > 3 And PrevOrd <> Ord And EngStatus <> "RELEASED" And PrevOrd <> 0 Then
            K = K + 1
            Workbooks(WkBk).Worksheets(Sht7).Cells(K + 3, KK + 2).Value = PrevOrd
            Workbooks(WkBk).Worksheets(Sht7).Cells(K + 3, KK + 3).Value = CustName
            Workbooks(WkBk).Worksheets(Sht7).Cells(K + 3, KK + 4).Value = Docs
            Workbooks(WkBk).Worksheets(Sht7).Cells(K + 3, KK + 5).Value = SchdRel
            Workbooks(WkBk).Worksheets(Sht7).Cells(K + 3, KK + 6).Value = ProdEng
            Workbooks(WkBk).Worksheets(Sht7).Cells(K + 3, KK + 7).Value = EstHrs
            EstHrs = 0#
            
            Set oRS = New ADODB.Recordset
            oRS.ActiveConnection = oCon
            If I < 9 Then
                oRS.Source = "SELECT Sum(cpe_time.hours) AS Hrs FROM cpe_time WHERE " & "(((cpe_time.operation=3) OR (cpe_time.operation=30) OR (cpe_time.operation=35) Or (cpe_time.operation=37) Or (cpe_time.operation=38) Or (cpe_time.operation=39)) AND (cpe_time.jobno='" & Net1 & "'))"
                oRS.Open
                Workbooks(WkBk).Worksheets(Sht7).Range(Cells(K + 3, KK + 8), Cells(K + 3, KK + 8)).CopyFromRecordset oRS
            Else
                oRS.Source = "SELECT Sum(cpe_time.hours) AS Hrs FROM cpe_time WHERE " & "(((cpe_time.operation=2) OR (cpe_time.operation=20) OR (cpe_time.operation=25) Or (cpe_time.operation=27) Or (cpe_time.operation=28) Or (cpe_time.operation=29)) AND (cpe_time.jobno='" & Net1 & "'))"
                oRS.Open
                Workbooks(WkBk).Worksheets(Sht7).Range(Cells(K + 3, KK + 8), Cells(K + 3, KK + 8)).CopyFromRecordset oRS
            End If
        Else
            If Ord <> PrevOrd And Ord <> 0 Then EstHrs = 0#
        End If
        
        CustName = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 71).Value
        Docs = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 72).Value
        SchdRel = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 73).Value
        ProdEng = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 74).Value
        Net1 = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 76).Value
        Net2 = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 77).Value
        EngStatus = Workbooks(WkBk).Worksheets(Sht5).Cells(J, 78).Value
        EstHrs = EstHrs + Workbooks(WkBk).Worksheets(Sht5).Cells(J, 75).Value + Workbooks(WkBk).Worksheets(Sht5).Cells(J, 79).Value
    
    Next J

HERE:

Next I

oRS.Close

SDt = Workbooks(WkBk).Worksheets(Sht7).Cells(2, 77).Value
EDt = Workbooks(WkBk).Worksheets(Sht7).Cells(2, 78).Value

Set ClrRange = Workbooks(WkBk).Worksheets(Sht7).Range("BX4:BZ99")
ClrRange.ClearContents
                
oRS.Source = "SELECT cpe_timeoff.usertype, ptodate, hours FROM cpe_timeoff WHERE (cpe_timeoff.ptodate BETWEEN '" & SDt & "' AND '" & EDt & "')"
oRS.Open
Workbooks(WkBk).Worksheets(Sht7).Range("BX4").CopyFromRecordset oRS


End Sub
Sub Order_Updates()
'
Set ClrRange = Workbooks(WkBk).Worksheets(Sht6).Range("W3:CB5000")
ClrRange.ClearContents

Ord = Workbooks(WkBk).Worksheets(Sht6).Cells(3, 2).Value

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.Order_Num = " & Ord & ";"
' Retrieve records
oRS.Open
Workbooks(WkBk).Worksheets(Sht6).Range("X3").CopyFromRecordset oRS

For I = 3 To 100
    For J = 62 To 80
        If Workbooks(WkBk).Worksheets(Sht6).Cells(I, J).Value < 3 Then
            Workbooks(WkBk).Worksheets(Sht6).Cells(I, J).Value = ""
        Else
        End If
    Next J
    If Workbooks(WkBk).Worksheets(Sht6).Cells(I, 25).Value < 3 Then Exit For
Next I

OrdStr = Str(Ord)
CustName = Workbooks(WkBk).Worksheets(Sht6).Cells(3, 26).Value
OrdFold = "J:\Orders\" & Mid(OrdStr, 2, 7) & "000\" & Mid(OrdStr, 2, 10) & "*"
TxtLine = Dir(OrdFold, vbDirectory)
'Workbooks(WkBk).Worksheets(Sht).Cells(12, 7).Value = TxtLine
JSubFold = Mid(TxtLine, 1, 7) & "000\"

Workbooks(WkBk).Worksheets(Sht6).Cells(23, 3).Value = Ord & "_" & CustName
Workbooks(WkBk).Worksheets(Sht6).Range("C23").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
        "J:\Orders\" & JSubFold & TxtLine & "\Engineering Documents\Order_Text\Order_" & Ord & ".pdf"
 
DnLn = 2
Call Next_Line

LastRow = Workbooks(WkBk).Worksheets(Sht6).Cells(Rows.Count, 24).End(xlUp).Row

For Index = 3 To LastRow
    Workbooks(WkBk).Worksheets(Sht6).Cells(Index, 23).Value = "No"
Next Index

End Sub
Sub Next_Line()
'
DnLn = DnLn + 1

Workbooks(WkBk).Worksheets(Sht6).Cells(5, 3).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 26).Value   ' Customer
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 5).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 27).Value   ' Line Number
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 28).Value   ' Material
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 9).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 29).Value   ' Description
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 30).Value  ' Network1
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 31).Value  ' Network2
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 32).Value  ' Network3
Workbooks(WkBk).Worksheets(Sht6).Cells(5, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 33).Value  ' Network4

Workbooks(WkBk).Worksheets(Sht6).Cells(8, 3).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 36).Value   ' File Number
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 5).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 44).Value   ' Scheduler
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 35).Value   ' Document Types
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 9).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 37).Value   ' Region/AE
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 38).Value  ' PC 1
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 40).Value  ' ME 1
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 42).Value  ' PM
Workbooks(WkBk).Worksheets(Sht6).Cells(8, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 43).Value  ' PE

Workbooks(WkBk).Worksheets(Sht6).Cells(11, 5).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 56).Value  ' Project Level
Workbooks(WkBk).Worksheets(Sht6).Cells(11, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 55).Value  ' MultiPlant
Workbooks(WkBk).Worksheets(Sht6).Cells(11, 9).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 60).Value  ' Mechanical Override
Workbooks(WkBk).Worksheets(Sht6).Cells(11, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 47).Value ' PC Estimate
Workbooks(WkBk).Worksheets(Sht6).Cells(11, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 49).Value ' ME Estimate
Workbooks(WkBk).Worksheets(Sht6).Cells(11, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 53).Value ' PM Estimate
Workbooks(WkBk).Worksheets(Sht6).Cells(11, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 51).Value ' Support Estimate

Workbooks(WkBk).Worksheets(Sht6).Cells(14, 5).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 54).Value  ' Engineering Location
Workbooks(WkBk).Worksheets(Sht6).Cells(14, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 62).Value  ' Into Engineering
Workbooks(WkBk).Worksheets(Sht6).Cells(14, 9).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 59).Value  ' Control Override
Workbooks(WkBk).Worksheets(Sht6).Cells(14, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 48).Value ' PC CO Estimate
Workbooks(WkBk).Worksheets(Sht6).Cells(14, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 50).Value ' ME CO Estimate
Workbooks(WkBk).Worksheets(Sht6).Cells(14, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 61).Value ' Active Year
Workbooks(WkBk).Worksheets(Sht6).Cells(14, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 76).Value ' Actual Shipped Date

Workbooks(WkBk).Worksheets(Sht6).Cells(17, 5).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 34).Value  ' Industry
Workbooks(WkBk).Worksheets(Sht6).Cells(17, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 63).Value  ' Initial Apps Out
Workbooks(WkBk).Worksheets(Sht6).Cells(17, 9).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 58).Value  ' ME Status
Workbooks(WkBk).Worksheets(Sht6).Cells(17, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 64).Value ' PC Apps Out
Workbooks(WkBk).Worksheets(Sht6).Cells(17, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 66).Value ' PC Apps Back
Workbooks(WkBk).Worksheets(Sht6).Cells(17, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 65).Value ' ME Apps Out
Workbooks(WkBk).Worksheets(Sht6).Cells(17, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 67).Value ' ME Apps Back

Workbooks(WkBk).Worksheets(Sht6).Cells(20, 5).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 37).Value  ' Region/AE
Workbooks(WkBk).Worksheets(Sht6).Cells(20, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 78).Value  ' Initial Ship Date
Workbooks(WkBk).Worksheets(Sht6).Cells(20, 9).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 57).Value  ' PC Status
Workbooks(WkBk).Worksheets(Sht6).Cells(20, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 68).Value ' PC Pre Order
Workbooks(WkBk).Worksheets(Sht6).Cells(20, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 70).Value ' PC Actual Pre Order
Workbooks(WkBk).Worksheets(Sht6).Cells(20, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 69).Value ' ME Pre Order
Workbooks(WkBk).Worksheets(Sht6).Cells(20, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 71).Value ' ME Actual Pre Order

Workbooks(WkBk).Worksheets(Sht6).Cells(23, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 80).Value  ' Initial Ship
Workbooks(WkBk).Worksheets(Sht6).Cells(23, 11).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 72).Value ' PC Prod Rel
Workbooks(WkBk).Worksheets(Sht6).Cells(23, 13).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 74).Value ' PC Actual Prod Rel
Workbooks(WkBk).Worksheets(Sht6).Cells(23, 15).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 73).Value ' ME Prod Rel
Workbooks(WkBk).Worksheets(Sht6).Cells(23, 17).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 75).Value ' ME Actual Prod Rel

End Sub
Sub PMCStatus()
'
Dim MatPrev As String, PMCIdx As Long

PMCIdx = 4
Set ClrRange = Workbooks(WkBk).Worksheets(ShtPMC).Range("C5:T500")
ClrRange.ClearContents

Open "\\pac5Intra\DepartmentReports\Production_Engineering\Current_Status\_PMC_Overview.HTM" For Output As #1
Print #1, "<HTML>"
Print #1, "<HEAD><TITLE>Operations Overview</TITLE></HEAD>"
Print #1, "<BODY>"
Print #1, "<BASEFONT FACE='arial, helvetica, tahoma, sans-serif'>"
Print #1, "<TABLE Border=True BorderColor='#a0a0a0' BGColor='#ffffff' CellPadding=4 CellSpacing=0>"

Print #1, "<TR><TH Colspan=15 BGColor='00ffff'>Whitewater Equipment Order Execution -  PMC Overview -        Last Update: " & Date$ & "</TH></TR>"
Print #1, "<TR>"
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Order <BR> Number</FONT></TH>"               ' 1
Print #1, "<TH BGColor='#00ffff' Align='Left'><FONT Size=-1>Customer Name</FONT></TH>"      ' 2
Print #1, "<TH BGColor='#00ffff' Align='Right'><FONT Size=-1>Value</FONT></TH>"             ' 3
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Scheduler</FONT></TH>"                       ' 4
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Confirmed<BR>Ship Date</FONT></TH>"          ' 5
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>CExD</FONT></TH>"                            ' 6
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Released</FONT></TH>"                        ' 7
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>ME Sent <BR> Approvals</FONT></TH>"          ' 8
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>ME Approvals <BR> Returned</FONT></TH>"      ' 9
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>ME Status</FONT></TH>"                       '10
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>ME Name</FONT></TH>"                         '11
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>PC Sent <BR> Approvals</FONT></TH>"          '12
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>PC Approvals <BR> Returned</FONT></TH>"      '13
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>PC Status</FONT></TH>"                       '14
Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>PC Name</FONT></TH>"                         '15
Print #1, "</TR>"

For I = 8 To 25000
    Ord = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 3).Value
    PrjLvl = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 38).Value
    LineNum = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 5).Value
    
    If Ord < 1 Then Exit For
    
    If PrjLvl <> "" Then
        MatPrev = MatNum
        MatNum = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 6).Value
        CustName = Customer
        Customer = Mid(Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 10).Value, 1, 40)
        PC_Num = PCnum
        PCnum = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 18).Value
        OrdVal = OrdVal + Round(Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 15).Value, 0)
        SchShip = PrevDate
        If Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 21).Value = "00/00/0000" Then
            PrevDate = 12 / 31 / 2021
        Else
            PrevDate = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 21).Value
        End If
        CExD = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 8).Value
        ME_A_Out = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 42).Value
        ME_A_Back = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 43).Value
        MEStatus = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 53).Value
        ME_Eng = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 49).Value
        PC_A_Out = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 40).Value
        PC_A_Back = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 41).Value
        PCStatus = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 52).Value
        PC_Eng = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 48).Value
        ME_Rel = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 47).Value
        PC_Rel = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 45).Value
        PE = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 50).Value
        PM = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 51).Value
    Else
    End If
        
    If LineNum = 0 And OrdVal <> 0 Then
        If PC_Num = 11100008 Then
            ProdCont = "Betsy"
        ElseIf PC_Num = 11000178 Then
            ProdCont = "Mike"
        End If
        If ME_Eng = "MGM" Then ME_Eng = "Matt"
        If ME_Eng = "BCS" Then ME_Eng = "Brian"
        If ME_Eng = "ACP" Then ME_Eng = "Alex"
        If ME_Eng = "CJT" Then ME_Eng = "Caleb"
        If ME_Eng = "JMA" Then ME_Eng = "James"
        If ME_Eng = "LEH" Then ME_Eng = "Lynsey"
        If ME_Eng = "JC-SAB" Then ME_Eng = "Jonathan C"
        If PC_Eng = "SSK" Then PC_Eng = "Steve"
        If PC_Eng = "REM" Then PC_Eng = "Randy"
        If PC_Eng = "ASD" Then PC_Eng = "Adam"
        If PC_Eng = "RP" Then PC_Eng = "Ryan"
        If PC_Eng = "MEZ" Then PC_Eng = "Mark"
        Released = 1
        If (PCStatus = "RELEASED") And (MEStatus = "RELEASED") And (ME_Rel > PCRel) Then
            Released = ME_Rel
        Else
        End If
        If (PCStatus = "RELEASED") And (MEStatus = "RELEASED") And (PC_Rel > MERel) Then
            Released = PC_Rel
        Else
        End If
        If (PC_Eng = "N/A") And (MEStatus = "RELEASED") Then
            Released = ME_Rel
        Else
        End If
        If (ME_Eng = "N/A") And (PCStatus = "RELEASED") Then
            Released = PC_Rel
        Else
        End If
            
        BGIdx = BGIdx + 1
        If (BGIdx Mod 2) > 0 Then BGColor = "ffffff" Else BGColor = "a0ffa0"
        Print #1, "<TR>"
        Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & Ord & "</FONT></TD>"
        Print #1, "<TD BGColor='#" & BGColor & "' Align='Left'><FONT Size=-1> &nbsp; " & Customer & "</FONT></TD>"
        Print #1, "<TD BGColor='#" & BGColor & "' Align='Right'><FONT Size=-1> $ " & OrdVal & "</FONT></TD>"
        Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & ProdCont & "</FONT></TD>"               ' Planner
        If SchShip > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & SchShip & "</FONT></TD>"           ' Sched Ship
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        If CExD > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & CExD & "</FONT></TD>"               ' CEXD
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        If Released > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & Released & "</FONT></TD>"           ' Released
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        If ME_A_Out > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & ME_A_Out & "</FONT></TD>"           ' ME Appr Out
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        If ME_A_Back > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & ME_A_Back & "</FONT></TD>"          ' ME Appr Back
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & MEStatus & "</FONT></TD>"               ' ME Status
        Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & ME_Eng & "</FONT></TD>"                 ' ME Name
        If PC_A_Out > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & PC_A_Out & "</FONT></TD>"           ' EE Appr Out
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        If PC_A_Back > 40000 Then
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & PC_A_Back & "</FONT></TD>"          ' EE Appr Back
        Else
            Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & "</FONT></TD>"
        End If
        Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & PCStatus & "</FONT></TD>"               ' PC Status
        Print #1, "<TD BGColor='#" & BGColor & "'><FONT Size=-1> &nbsp; " & PC_Eng & "</FONT></TD>"                 ' PC Name
        Print #1, "</TR>"
        
        PMCIdx = PMCIdx + 1
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 3).Value = Ord
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 4).Value = Customer
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 5).Value = OrdVal
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 6).Value = ProdCont
        If SchShip > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 7).Value = SchShip
        If CExD > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 8).Value = CExD
        If Released > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 9).Value = Released
        If ME_A_Out > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 10).Value = ME_A_Out
        If ME_A_Back > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 11).Value = ME_A_Back
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 12).Value = MEStatus
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 13).Value = ME_Eng
        If PC_A_Out > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 14).Value = PC_A_Out
        If PC_A_Back > 40000 Then Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 15).Value = PC_A_Back
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 16).Value = PCStatus
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 17).Value = PC_Eng
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 19).Value = PE
        Workbooks(WkBk).Worksheets(ShtPMC).Cells(PMCIdx, 20).Value = PM
        
        OrdVal = 0
    Else
    End If
    
Next I

Print #1, "</TABLE>"
Print #1, "</BODY>"
Print #1, "</HTML>"
Close #1

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon

Set ClrRange = Workbooks(WkBk).Worksheets(ShtPMC).Range("X5:X99")
ClrRange.ClearContents

For I = 5 To 50
    Set ClrRange = Workbooks(WkBk).Worksheets(Sht1).Range("BX3:CD5000")
    ClrRange.ClearContents
    
    Ord = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 3).Value
    If Ord < 3 Then Exit For
    
    oRS.Source = "SELECT * FROM cpe_schedule WHERE cpe_schedule.orderno = " & Ord & " ORDER BY cpe_schedule.datestamp"
    oRS.Open
    Workbooks(WkBk).Worksheets(Sht1).Range("BX3").CopyFromRecordset oRS
    For K = 3 To 99
        If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 78).Value < 3 Then Exit For
        If Workbooks(WkBk).Worksheets(Sht1).Cells(K, 81).Value > "" Then
            Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 24).Value = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 24).Value & Workbooks(WkBk).Worksheets(Sht1).Cells(K, 86) & _
                Workbooks(WkBk).Worksheets(Sht1).Cells(K, 77) & ": " & Workbooks(WkBk).Worksheets(Sht1).Cells(K, 81) & Chr(10)
        Else
        End If
    Next K
    oRS.Close
    
Next I

Workbooks(WkBk).Worksheets(ShtPMC).Select
Rows("5:99").EntireRow.AutoFit

End Sub
Sub PMC_EMail()
'
Dim EE_W As Long, ME_W As Long
Dim CustName As String, EBody As String, Clerk As String, WW_Team As String
Dim OutApp As Object, OutMail As Object

WW_Team = " - (WW: Betsy, Mike, & Mark)"

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

EBody = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"
Clerk = "J.Summers@schenckprocess.com"
EBody = EBody & "<FONT Color='ff0000' Size=+2><B>WHITEWATER</B></FONT><BR><BR><BR>"

For I = 5 To 100
    EE_W = 0
    ME_W = 0
    Ord = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 3).Value
    If Ord < 1 Then Exit For
    CustName = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 4).Value
    PE = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 21).Value
    PM = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 22).Value
    PCStatus = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 16).Value
    MEStatus = Workbooks(WkBk).Worksheets(ShtPMC).Cells(I, 12).Value
    
    If Mid(MEStatus, 1, 4) = "WAIT" Then ME_W = 1
    If Mid(PCStatus, 1, 4) = "WAIT" Then EE_W = 1
    
    If ME_W = 1 And EE_W = 0 Then
        EBody = EBody & "<OL><U>" & Ord & " - " & CustName & WW_Team & " <B>(PM: " & PM & ", PE: " & PE & ")</B></U><BR>"
        EBody = EBody & "When will mechanical approval designs be returned?</OL><BR><BR>"
    Else
    End If
    
    If ME_W = 0 And EE_W = 1 Then
        EBody = EBody & "<OL><U>" & Ord & " - " & CustName & WW_Team & " <B>(PM: " & PM & ", PE: " & PE & ")</B></U><BR>"
        EBody = EBody & "When will electrical approval designs be returned?</OL><BR><BR>"
    Else
    End If

    If ME_W = 1 And EE_W = 1 Then
        EBody = EBody & "<OL><U>" & Ord & " - " & CustName & WW_Team & " <B>(PM: " & PM & ", PE: " & PE & ")</B></U><BR>"
        EBody = EBody & "When will mechanical and electrical approval designs be returned?</OL><BR><BR>"
    Else
    End If

Next I

EBody = EBody & "Thank you,<BR>Mark<BR><BR>"

    With OutMail
        .To = Clerk
        .CC = ""
        .BCC = ""
        .Subject = "WW, PMC and MFG Weekly Backlog Meeting"
        '.Body = strbody
        .HTMLBody = EBody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .Display '.Send   'or use .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
Sub Chg_CustName()
'
DB_Col = 3
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 3).Value

Call Change_StringValues

End Sub
Sub Chg_LineNum()
'
DB_Col = 4
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewLong = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 5).Value

Call Change_LongValues

End Sub
Sub Chg_Material()
'
DB_Col = 5
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 7).Value

Call Change_StringValues

End Sub
Sub Chg_Description()
'
DB_Col = 6
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 9).Value

Call Change_StringValues

End Sub
Sub Chg_Net1()
'
DB_Col = 7
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 11).Value

Call Change_StringValues

End Sub
Sub Chg_Net2()
'
DB_Col = 8
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 13).Value

Call Change_StringValues

End Sub
Sub Chg_Net3()
'
DB_Col = 9
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 15).Value

Call Change_StringValues

End Sub
Sub Chg_Net4()
'
DB_Col = 10
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(5, 17).Value

Call Change_StringValues

End Sub
Sub Chg_FileNum()
'
DB_Col = 13
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 3).Value

Call Change_StringValues

End Sub
Sub Chg_Scheduler()
'
DB_Col = 21
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 5).Value

Call Change_StringValues

End Sub
Sub Chg_DocTypes()
'
DB_Col = 12
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 7).Value

Call Change_StringValues

End Sub
Sub Chg_Region()
'
DB_Col = 14
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 9).Value

Call Change_StringValues

End Sub
Sub Chg_PC1()
'
DB_Col = 15
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 11).Value

Call Change_StringValues

End Sub
Sub Chg_ME1()
'
DB_Col = 17
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 13).Value

Call Change_StringValues

End Sub
Sub Chg_PM()
'
DB_Col = 19
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 15).Value

Call Change_StringValues

End Sub
Sub Chg_PE()
'
DB_Col = 20
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(8, 17).Value

Call Change_StringValues

End Sub
Sub Chg_PrjLvl()
'
DB_Col = 33
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 5).Value

Call Change_StringValues

End Sub
Sub Chg_MultiPlnt()
'
DB_Col = 32
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 7).Value

Call Change_StringValues

End Sub
Sub Chg_ME_Override()
'
DB_Col = 37
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 9).Value

Call Change_StringValues

End Sub
Sub Chg_PC_Est()
'
DB_Col = 24
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewNumeric = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 11).Value

Call Change_NumericValues

End Sub
Sub Chg_ME_Est()
'
DB_Col = 26
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewNumeric = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 13).Value

Call Change_NumericValues

End Sub
Sub Chg_PM_Est()
'
DB_Col = 30
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewNumeric = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 15).Value

Call Change_NumericValues

End Sub
Sub Chg_Supp_Est()
'
DB_Col = 28
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewNumeric = Workbooks(WkBk).Worksheets(Sht6).Cells(11, 17).Value

Call Change_NumericValues

End Sub
Sub Chg_EngLocation()
'
DB_Col = 31
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 5).Value

Call Change_StringValues

End Sub
Sub Chg_IntoEng()
'
DB_Col = 39
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 7).Value

Call Change_Dates

End Sub
Sub Chg_PC_Override()
'
DB_Col = 36
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 9).Value

Call Change_StringValues

End Sub
Sub Chg_PC_CO_Est()
'
DB_Col = 25
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewNumeric = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 11).Value

Call Change_NumericValues

End Sub
Sub Chg_ME_CO_Est()
'
DB_Col = 27
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewNumeric = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 13).Value

Call Change_NumericValues

End Sub
Sub Chg_ActiveYr()
'
DB_Col = 38
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewLong = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 15).Value

Call Change_LongValues

End Sub
Sub Chg_ShippedDate()
'
DB_Col = 53
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(14, 17).Value

Call Change_Dates

End Sub
Sub Chg_IntialApps()
'
DB_Col = 40
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 7).Value

Call Change_Dates

End Sub
Sub Chg_MEStatus()
'
DB_Col = 35
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 9).Value

Call Change_StringValues

End Sub
Sub Chg_PC_Apps_Out()
'
DB_Col = 41
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 11).Value

Call Change_Dates

End Sub
Sub Chg_PC_Apps_Back()
'
DB_Col = 43
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 13).Value

Call Change_Dates

End Sub
Sub Chg_ME_Apps_Out()
'
DB_Col = 42
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 15).Value

Call Change_Dates

End Sub
Sub Chg_ME_Apps_Back()
'
DB_Col = 44
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 17).Value

Call Change_Dates

End Sub
Sub Chg_InitialRel()
'
DB_Col = 56
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 7).Value

Call Change_Dates

End Sub
Sub Chg_Industry()
'
DB_Col = 11
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(17, 5).Value

Call Change_StringValues

End Sub
Sub Chg_Region_AE()
'
DB_Col = 14
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 5).Value

Call Change_StringValues

End Sub
Sub Chg_PC_Status()
'
DB_Col = 34
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewString = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 9).Value

Call Change_StringValues

End Sub
Sub Chg_PC_PreRel()
'
DB_Col = 45
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 11).Value

Call Change_Dates

End Sub
Sub Chg_PC_Act_Pre()
'
DB_Col = 47
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 13).Value

Call Change_Dates

End Sub
Sub Chg_ME_PreRel()
'
DB_Col = 46
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 15).Value

Call Change_Dates

End Sub
Sub Chg_ME_Act_Pre()
'
DB_Col = 48
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(20, 17).Value

Call Change_Dates

End Sub
Sub Chg_Intial_Ship()
'
DB_Col = 57
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(23, 7).Value

Call Change_Dates

End Sub
Sub Chg_PC_Prod_Rel()
'
DB_Col = 49
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(23, 11).Value

Call Change_Dates

End Sub
Sub Chg_PC_Act_Rel()
'
DB_Col = 51
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(23, 13).Value

Call Change_Dates

End Sub
Sub Chg_ME_Prod_Rel()
'
DB_Col = 50
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(23, 15).Value

Call Change_Dates

End Sub
Sub Chg_ME_Act_Rel()
'
DB_Col = 52
DB_Row = DnLn
U_ID = Workbooks(WkBk).Worksheets(Sht6).Cells(DnLn, 24).Value
NewDate = Workbooks(WkBk).Worksheets(Sht6).Cells(23, 17).Value

Call Change_Dates

End Sub
Sub Change_StringValues()
'DB_Col is the column in the database table
'DB_Row is the row to use in the spreadsheet.

DB_Head(1) = "uniqid"
DB_Head(2) = "OrderNum"
DB_Head(3) = "Customer_Name"
DB_Head(4) = "Line_Num"
DB_Head(5) = "Material"
DB_Head(6) = "Description"
DB_Head(7) = "Network1"
DB_Head(8) = "Network2"
DB_Head(9) = "Network3"
DB_Head(10) = "Network4"
DB_Head(11) = "Industry"
DB_Head(12) = "DocTypes"
DB_Head(13) = "FileNum"
DB_Head(14) = "Region_AE"
DB_Head(15) = "PC1"
DB_Head(16) = "PC2"
DB_Head(17) = "ME1"
DB_Head(18) = "ME2"
DB_Head(19) = "PM"
DB_Head(20) = "PE"
DB_Head(21) = "Scheduler"
DB_Head(22) = "PC_Chkr"
DB_Head(23) = "ME_Chkr"
DB_Head(24) = "PC_Est"
DB_Head(25) = "PC_CO_Est"
DB_Head(26) = "ME_Est"
DB_Head(27) = "ME_CO_Est"
DB_Head(28) = "Supp_Est"
DB_Head(29) = "Soft_Est"
DB_Head(30) = "PM_Est"
DB_Head(31) = "Eng_Loc"
DB_Head(32) = "MultiPlnt"
DB_Head(33) = "Pjt_Lvl"
DB_Head(34) = "PC_Status"
DB_Head(35) = "ME_Status"
DB_Head(36) = "PC_Override"
DB_Head(37) = "ME_Override"
DB_Head(38) = "Active_Year"
DB_Head(39) = "Into_Eng"
DB_Head(40) = "Apps_Out_0"
DB_Head(41) = "PC_Apps_Out"
DB_Head(42) = "ME_Apps_Out"
DB_Head(43) = "PC_Apps_Back"
DB_Head(44) = "ME_Apps_Back"
DB_Head(45) = "PC_PreRel"
DB_Head(46) = "ME_PreRel"
DB_Head(47) = "PC_PreRel_Act"
DB_Head(48) = "ME_PreRel_Act"
DB_Head(49) = "PC_Rel_F"
DB_Head(50) = "ME_Rel_F"
DB_Head(51) = "PC_Act_Rel"
DB_Head(52) = "ME_Act_Rel"
DB_Head(53) = "ShippedDate"
DB_Head(54) = "PC_Rel_0"
DB_Head(55) = "ME_Rel_0"
DB_Head(56) = "Prod_Rel_0"
DB_Head(57) = "Ship_Date_0"

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "UPDATE Prod_Eng SET Prod_Eng." & DB_Head(DB_Col) & " = '" & NewString & "' WHERE Prod_Eng.uniqid = '" & U_ID & "'"
oRS.Open

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

'Call DB_Editor

End Sub
Sub Change_Dates()
'DB_Col is the column in the database table
'DB_Row is the row to use in the spreadsheet.

DB_Head(1) = "uniqid"
DB_Head(2) = "OrderNum"
DB_Head(3) = "Customer_Name"
DB_Head(4) = "Line_Num"
DB_Head(5) = "Material"
DB_Head(6) = "Description"
DB_Head(7) = "Network1"
DB_Head(8) = "Network2"
DB_Head(9) = "Network3"
DB_Head(10) = "Network4"
DB_Head(11) = "Industry"
DB_Head(12) = "DocTypes"
DB_Head(13) = "FileNum"
DB_Head(14) = "Region_AE"
DB_Head(15) = "PC1"
DB_Head(16) = "PC2"
DB_Head(17) = "ME1"
DB_Head(18) = "ME2"
DB_Head(19) = "PM"
DB_Head(20) = "PE"
DB_Head(21) = "Scheduler"
DB_Head(22) = "PC_Chkr"
DB_Head(23) = "ME_Chkr"
DB_Head(24) = "PC_Est"
DB_Head(25) = "PC_CO_Est"
DB_Head(26) = "ME_Est"
DB_Head(27) = "ME_CO_Est"
DB_Head(28) = "Supp_Est"
DB_Head(29) = "Soft_Est"
DB_Head(30) = "PM_Est"
DB_Head(31) = "Eng_Loc"
DB_Head(32) = "MultiPlnt"
DB_Head(33) = "Pjt_Lvl"
DB_Head(34) = "PC_Status"
DB_Head(35) = "ME_Status"
DB_Head(36) = "PC_Override"
DB_Head(37) = "ME_Override"
DB_Head(38) = "Active_Year"
DB_Head(39) = "Into_Eng"
DB_Head(40) = "Apps_Out_0"
DB_Head(41) = "PC_Apps_Out"
DB_Head(42) = "ME_Apps_Out"
DB_Head(43) = "PC_Apps_Back"
DB_Head(44) = "ME_Apps_Back"
DB_Head(45) = "PC_PreRel"
DB_Head(46) = "ME_PreRel"
DB_Head(47) = "PC_Act_PreRel"
DB_Head(48) = "ME_Act_PreRel"
DB_Head(49) = "PC_Rel_F"
DB_Head(50) = "ME_Rel_F"
DB_Head(51) = "PC_Act_Rel"
DB_Head(52) = "ME_Act_Rel"
DB_Head(53) = "ShippedDate"
DB_Head(54) = "PC_Rel_0"
DB_Head(55) = "ME_Rel_0"
DB_Head(56) = "Prod_Rel_0"
DB_Head(57) = "Ship_Date_0"

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "UPDATE Prod_Eng SET Prod_Eng." & DB_Head(DB_Col) & " = '" & NewDate & "' WHERE Prod_Eng.uniqid = '" & U_ID & "'"
oRS.Open

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
Sub Change_NumericValues()
'DB_Col is the column in the database table
'DB_Row is the row to use in the spreadsheet.

DB_Head(1) = "uniqid"
DB_Head(2) = "OrderNum"
DB_Head(3) = "Customer_Name"
DB_Head(4) = "Line_Num"
DB_Head(5) = "Material"
DB_Head(6) = "Description"
DB_Head(7) = "Network1"
DB_Head(8) = "Network2"
DB_Head(9) = "Network3"
DB_Head(10) = "Network4"
DB_Head(11) = "Industry"
DB_Head(12) = "DocTypes"
DB_Head(13) = "FileNum"
DB_Head(14) = "Region_AE"
DB_Head(15) = "PC1"
DB_Head(16) = "PC2"
DB_Head(17) = "ME1"
DB_Head(18) = "ME2"
DB_Head(19) = "PM"
DB_Head(20) = "PE"
DB_Head(21) = "Scheduler"
DB_Head(22) = "PC_Chkr"
DB_Head(23) = "ME_Chkr"
DB_Head(24) = "PC_Est"
DB_Head(25) = "PC_CO_Est"
DB_Head(26) = "ME_Est"
DB_Head(27) = "ME_CO_Est"
DB_Head(28) = "Supp_Est"
DB_Head(29) = "Soft_Est"
DB_Head(30) = "PM_Est"
DB_Head(31) = "Eng_Loc"
DB_Head(32) = "MultiPlnt"
DB_Head(33) = "Pjt_Lvl"
DB_Head(34) = "PC_Status"
DB_Head(35) = "ME_Status"
DB_Head(36) = "PC_Override"
DB_Head(37) = "ME_Override"
DB_Head(38) = "Active_Year"
DB_Head(39) = "Into_Eng"
DB_Head(40) = "Apps_Out_0"
DB_Head(41) = "PC_Apps_Out"
DB_Head(42) = "ME_Apps_Out"
DB_Head(43) = "PC_Apps_Back"
DB_Head(44) = "ME_Apps_Back"
DB_Head(45) = "PC_PreRel"
DB_Head(46) = "ME_PreRel"
DB_Head(47) = "PC_PreRel_Act"
DB_Head(48) = "ME_PreRel_Act"
DB_Head(49) = "PC_Rel_F"
DB_Head(50) = "ME_Rel_F"
DB_Head(51) = "PC_Act_Rel"
DB_Head(52) = "ME_Act_Rel"
DB_Head(53) = "ShippedDate"
DB_Head(54) = "PC_Rel_0"
DB_Head(55) = "ME_Rel_0"
DB_Head(56) = "Prod_Rel_0"
DB_Head(57) = "Ship_Date_0"

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "UPDATE Prod_Eng SET Prod_Eng." & DB_Head(DB_Col) & " = " & NewNumeric & " WHERE Prod_Eng.uniqid = '" & U_ID & "'"
oRS.Open

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

'Call DB_Editor

End Sub
Sub Change_LongValues()
'DB_Col is the column in the database table
'DB_Row is the row to use in the spreadsheet.

DB_Head(1) = "uniqid"
DB_Head(2) = "OrderNum"
DB_Head(3) = "Customer_Name"
DB_Head(4) = "Line_Num"
DB_Head(5) = "Material"
DB_Head(6) = "Description"
DB_Head(7) = "Network1"
DB_Head(8) = "Network2"
DB_Head(9) = "Network3"
DB_Head(10) = "Network4"
DB_Head(11) = "Industry"
DB_Head(12) = "DocTypes"
DB_Head(13) = "FileNum"
DB_Head(14) = "Region_AE"
DB_Head(15) = "PC1"
DB_Head(16) = "PC2"
DB_Head(17) = "ME1"
DB_Head(18) = "ME2"
DB_Head(19) = "PM"
DB_Head(20) = "PE"
DB_Head(21) = "Scheduler"
DB_Head(22) = "PC_Chkr"
DB_Head(23) = "ME_Chkr"
DB_Head(24) = "PC_Est"
DB_Head(25) = "PC_CO_Est"
DB_Head(26) = "ME_Est"
DB_Head(27) = "ME_CO_Est"
DB_Head(28) = "Supp_Est"
DB_Head(29) = "Soft_Est"
DB_Head(30) = "PM_Est"
DB_Head(31) = "Eng_Loc"
DB_Head(32) = "MultiPlnt"
DB_Head(33) = "Pjt_Lvl"
DB_Head(34) = "PC_Status"
DB_Head(35) = "ME_Status"
DB_Head(36) = "PC_Override"
DB_Head(37) = "ME_Override"
DB_Head(38) = "Active_Year"
DB_Head(39) = "Into_Eng"
DB_Head(40) = "Apps_Out_0"
DB_Head(41) = "PC_Apps_Out"
DB_Head(42) = "ME_Apps_Out"
DB_Head(43) = "PC_Apps_Back"
DB_Head(44) = "ME_Apps_Back"
DB_Head(45) = "PC_PreRel"
DB_Head(46) = "ME_PreRel"
DB_Head(47) = "PC_PreRel_Act"
DB_Head(48) = "ME_PreRel_Act"
DB_Head(49) = "PC_Rel_F"
DB_Head(50) = "ME_Rel_F"
DB_Head(51) = "PC_Act_Rel"
DB_Head(52) = "ME_Act_Rel"
DB_Head(53) = "ShippedDate"
DB_Head(54) = "PC_Rel_0"
DB_Head(55) = "ME_Rel_0"
DB_Head(56) = "Prod_Rel_0"
DB_Head(57) = "Ship_Date_0"

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "UPDATE Prod_Eng SET Prod_Eng." & DB_Head(DB_Col) & " = " & NewLong & " WHERE Prod_Eng.uniqid = '" & U_ID & "'"
oRS.Open

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

'Call DB_Editor

End Sub
Sub Special()
'
Set ClrRange = Workbooks(WkBk).Worksheets("Special").Range("B3:J5000")
ClrRange.ClearContents

Net1 = "1100006022"  'Workbooks(WkBk).Worksheets(Sht3).Cells(I, 4).Value

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "SELECT * FROM cpe_time WHERE cpe_time.jobno='" & Net1 & "' ORDER BY cpe_time.datestamp"
' Retrieve records
oRS.Open
Workbooks(WkBk).Worksheets("Special").Range("B3").CopyFromRecordset oRS

End Sub
Sub Special2()
'
Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "SELECT Prod_Eng.Order_Num, Customer_Name, ME1, PC1 FROM Prod_Eng WHERE Prod_Eng.Active_Year = 2019 ORDER BY Prod_Eng.PC1, Order_Num"
oRS.Open

Workbooks(WkBk).Worksheets("Scratch").Select
Set ClrRange = Worksheets("Scratch").Range("W2:AE10000")
ClrRange.ClearContents
Workbooks(WkBk).Worksheets("Scratch").Range("W2").CopyFromRecordset oRS

PrevOrd = 0
J = 1
Ord = 1

For I = 2 To 5000
    If Ord = 0 Then Exit For
    PrevOrd = Ord
    Ord = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 23).Value
    If PrevOrd <> Ord Then
        J = J + 1
        Workbooks(WkBk).Worksheets("Scratch").Cells(J, 28).Value = Ord
        Workbooks(WkBk).Worksheets("Scratch").Cells(J, 29).Value = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 24).Value
        Workbooks(WkBk).Worksheets("Scratch").Cells(J, 30).Value = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 25).Value
        Workbooks(WkBk).Worksheets("Scratch").Cells(J, 31).Value = Workbooks(WkBk).Worksheets("Scratch").Cells(I, 26).Value
    Else
    End If
    
Next I
'
End Sub
Sub CommentFixer()
'
Set ClrRange = Workbooks(WkBk).Worksheets(Sht5).Range("CB3:CB500")
ClrRange.ClearContents

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "UPDATE cpe_schedule SET cpe_schedule.engineer='JRS' WHERE cpe_schedule.engineer = 'jrs'"
oRS.Open

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
Sub Load_SAP_and_Eng_Data()
'
' Insert the SAP exports (Open, Parts and Shipped)
Call Insert_SAP_Data

' Add the Production Engineering Data to the Open_Equip, Open_Parts, and Shipped sheets
Call Add_Prod_Eng_Data

End Sub
Sub Insert_SAP_Data()
' This MACRO opens the data exports from SAP that are stored on the W drive and loads them
' into the spreadsheet for manipulations
'
' Equipment
Set ClrRange = Workbooks(WkBk).Worksheets(ShtEquip).Range("A1:AB65000")
ClrRange.ClearContents

    ChDir Path_SAP
    Workbooks.OpenText Filename:=Path_SAP & "\Open.XLS", Origin:=xlWindows _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1)), TrailingMinusNumbers:=True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("D5").Select
    Columns("D:D").ColumnWidth = 4

    Windows("Open.XLS").Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    Range("A1:AB65000").Select
    Selection.Copy
    Windows(WkBk).Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    Sheets(ShtEquip).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A5").Select
    Windows("Open.XLS").Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.Close

' Parts
Set ClrRange = Workbooks(WkBk).Worksheets(ShtParts).Range("A1:AB65000")
ClrRange.ClearContents

    ChDir Path_SAP
    Workbooks.OpenText Filename:=Path_SAP & "\Parts.XLS", Origin:=xlWindows _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1)), TrailingMinusNumbers:=True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("D5").Select
    Columns("D:D").ColumnWidth = 4

    Windows("Parts.XLS").Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    Range("A1:AB65000").Select
    Selection.Copy
    Windows(WkBk).Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    Sheets(ShtParts).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A5").Select
    Windows("Parts.XLS").Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.Close

' Shipped
Set ClrRange = Workbooks(WkBk).Worksheets(ShtShip).Range("A1:AB65000")
ClrRange.ClearContents

    ChDir Path_SAP
    Workbooks.OpenText Filename:=Path_SAP & "\Shipped.XLS", Origin:=xlWindows _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), _
        Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 1), Array(15, 1), Array( _
        16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), Array(22, 1), _
        Array(23, 1)), TrailingMinusNumbers:=True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("D5").Select
    Columns("D:D").ColumnWidth = 4

    Windows("Shipped.XLS").Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    Range("A1:Y65000").Select
    Selection.Copy
    Windows(WkBk).Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    Sheets(ShtShip).Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A5").Select
    Windows("Shipped.XLS").Activate
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.WindowState = xlNormal
    ActiveWindow.Close
    
    Range("A8:Y65000").Select
    ActiveWorkbook.Worksheets("Shipped").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Shipped").Sort.SortFields.Add Key:=Range( _
        "F8:F65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Shipped").Sort.SortFields.Add Key:=Range( _
        "P8:P65000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Shipped").Sort
        .SetRange Range("A8:X65000")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A7").Select

End Sub
Sub Add_Prod_Eng_Data()
'
Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

'Set oCon2 = New ADODB.Connection
'oCon2.ConnectionString = "Driver={SQl Server};Server=PAC5Intra03\PAC5SQLExpress;Database=PAC1Orders;User ID=sa;Password=manage_ERP;" ' Trusted_Connection=yes;"
'oCon2.Open

Ignore = 0

' Equipment Orders
Workbooks(WkBk).Worksheets(ShtEquip).Select
Set ClrRange = Workbooks(WkBk).Worksheets(ShtEquip).Range("AJ8:BD50000")
ClrRange.ClearContents
For I = 8 To 50000
    Ord = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 3).Value
    LineNum = Workbooks(WkBk).Worksheets(ShtEquip).Cells(I, 5).Value
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT Prod_Eng.Order_Num, Line_Num, Pjt_Lvl, Into_Eng, PC_Apps_Out, PC_Apps_Back, ME_Apps_Out, ME_Apps_Back, PC_Rel_F, PC_Act_Rel, ME_Rel_F, ME_Act_Rel, PC1, ME1, PE, PM, PC_Status, ME_Status FROM Prod_Eng WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.Line_Num=" & LineNum
    ' Retrieve records
    oRS.Open
    Workbooks(WkBk).Worksheets(ShtEquip).Range(Cells(I, 36), Cells(I, 36)).CopyFromRecordset oRS
    If Ord = 0 Then Exit For
Next I

' Parts Orders
Workbooks(WkBk).Worksheets(ShtParts).Select
Set ClrRange = Workbooks(WkBk).Worksheets(ShtParts).Range("AJ8:BD50000")
ClrRange.ClearContents
For I = 8 To 50000
    Ord = Workbooks(WkBk).Worksheets(ShtParts).Cells(I, 3).Value
    LineNum = Workbooks(WkBk).Worksheets(ShtParts).Cells(I, 5).Value
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT Prod_Eng.Order_Num, Line_Num, Pjt_Lvl, Into_Eng, PC_Apps_Out, PC_Apps_Back, ME_Apps_Out, ME_Apps_Back, PC_Rel_F, PC_Act_Rel, ME_Rel_F, ME_Act_Rel, PC1, ME1, PE, PM, PC_Status, ME_Status FROM Prod_Eng WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.Line_Num=" & LineNum
    ' Retrieve records
    oRS.Open
    Workbooks(WkBk).Worksheets(ShtParts).Range(Cells(I, 36), Cells(I, 36)).CopyFromRecordset oRS
    If Ord = 0 Then Exit For
Next I
' Region from Dashboard - NOT USED ANYMORE
'For I = 8 To 50000
'    Ord = Workbooks(WkBk).Worksheets(ShtParts).Cells(I, 3).Value
'    Set oRS2 = New ADODB.Recordset
'    oRS2.ActiveConnection = oCon2
'    oRS2.Source = "SELECT custorder.orderno, region FROM custorder WHERE custorder.orderno = '" & Ord & "'"
    ' Retrieve records
'    oRS2.Open
'    Workbooks(WkBk).Worksheets(ShtParts).Range(Cells(I, 57), Cells(I, 57)).CopyFromRecordset oRS2
'    If Ord = 0 Then Exit For
'Next I

' Shipped Orders
Workbooks(WkBk).Worksheets(ShtShip).Select
Set ClrRange = Workbooks(WkBk).Worksheets(ShtShip).Range("AJ8:BD50000")
ClrRange.ClearContents
For I = 8 To 50000
    Ord = Workbooks(WkBk).Worksheets(ShtShip).Cells(I, 3).Value
    LineNum = Workbooks(WkBk).Worksheets(ShtShip).Cells(I, 5).Value
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT Prod_Eng.Order_Num, Line_Num, Pjt_Lvl, Into_Eng, PC_Apps_Out, PC_Apps_Back, ME_Apps_Out, ME_Apps_Back, PC_Rel_F, PC_Act_Rel, ME_Rel_F, ME_Act_Rel, PC1, ME1, PE, PM, PC_Status, ME_Status FROM Prod_Eng WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.Line_Num=" & LineNum
    ' Retrieve records
    oRS.Open
    Workbooks(WkBk).Worksheets(ShtShip).Range(Cells(I, 36), Cells(I, 36)).CopyFromRecordset oRS
    If Ord = 0 Then Ignore = Ignore + 1
    If Ord = 0 And Ignore > 15 Then Exit For

Next I
    
If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

'If Not oRS2 Is Nothing Then Set oRS2 = Nothing
'If Not oCon2 Is Nothing Then Set oCon2 = Nothing

'
End Sub
Sub ProdCont_Query()
'
Set ClrRange = Workbooks(WkBk).Worksheets(ShtOE).Range("C5:AB5")
ClrRange.ClearContents
Set ClrRange = Workbooks(WkBk).Worksheets(ShtOE).Range("F9:I26")
ClrRange.ClearContents
Set ClrRange = Workbooks(WkBk).Worksheets(ShtOE).Range("B9:D9")
ClrRange.ClearContents

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Ord = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value
' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "SELECT me_netact, scheddate, appdate, reldate, proddate, shipdate, file_num, po_num, scheduler FROM cpe_scheduling WHERE cpe_scheduling.orderno=" & Ord & ";"
' Retrieve records
oRS.Open
Workbooks(WkBk).Worksheets(ShtOE).Range("C5").CopyFromRecordset oRS

J = 0
For I = 9 To 25000
    SearchOrd = Workbooks(WkBk).Worksheets(Sht6).Cells(I, 3).Value
    LineNum = Workbooks(WkBk).Worksheets(Sht6).Cells(I, 5).Value
    If Ord = SearchOrd And LineNum <> 0 Then
        Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(I, 10).Value
        J = J + 1
        Workbooks(WkBk).Worksheets(ShtOE).Cells(J + 8, 6).Value = LineNum
        Workbooks(WkBk).Worksheets(ShtOE).Cells(J + 8, 7).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(I, 7).Value
    Else
    End If
    If Ord = SearchOrd And LineNum = 0 Then Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 3).Value = Workbooks(WkBk).Worksheets(Sht6).Cells(I, 15).Value
    If SearchOrd < 1 Then Exit For

Next I

J = 0
For I = 9 To 25000
    SearchOrd = Workbooks(WkBk).Worksheets(Sht7).Cells(I, 3).Value
    LineNum = Workbooks(WkBk).Worksheets(Sht7).Cells(I, 5).Value
    If Ord = SearchOrd And LineNum <> 0 Then
        Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value = Workbooks(WkBk).Worksheets(Sht7).Cells(I, 10).Value
        J = J + 1
        Workbooks(WkBk).Worksheets(ShtOE).Cells(J + 8, 6).Value = LineNum
        Workbooks(WkBk).Worksheets(ShtOE).Cells(J + 8, 7).Value = Workbooks(WkBk).Worksheets(Sht7).Cells(I, 7).Value
    Else
    End If
    If Ord = SearchOrd And LineNum = 0 Then Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 3).Value = Workbooks(WkBk).Worksheets(Sht7).Cells(I, 15).Value
    If SearchOrd < 1 Then Exit For

Next I


End Sub
Sub Print_SAP_Order()
'
On Error GoTo NoHomeScreen
For I = 0 To 6
    If Left(GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(I + 0).findById("wnd[0]").Text, 15) = "SAP Easy Access" Then
        Set Session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(I + 0)
        On Error GoTo 0
        Exit For
    End If
Next I

GoTo SkipNoHomeScreen:

NoHomeScreen:
If I = 0 Then
    MsgBox "Cannot connect to SAP. Open an SAP Easy Access window to run the program.", Title:="Error!"
    Exit Sub
Else
    MsgBox "An SAP Easy Access window must be open to run the program.", Title:="Error!"
    Exit Sub
End If

SkipNoHomeScreen:

Ord = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value
Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode "F02529"
Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").Text = Ord
Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
Session.findById("wnd[0]/mbar/menu[0]/menu[5]").Select
Session.findById("wnd[1]/usr/tblSAPLVMSGTABCONTROL").getAbsoluteRow(0).Selected = True
Session.findById("wnd[1]/tbar[0]/btn[86]").press
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/titl/shellcont/shell").pressButton "%GOS_TOOLBOX"
Session.findById("wnd[0]/shellcont/shell").pressButton "VIEW_ATTA"
Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = "0"
Session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell
Session.findById("wnd[1]/tbar[0]/btn[12]").press
Session.findById("wnd[0]/shellcont").Close
Session.findById("wnd[0]/tbar[0]/btn[3]").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press

End Sub
Sub NewOrder2DB()
'
Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

Ord = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value

    AssyStr = ""
    AssyStr = Worksheets(ShtOE).Cells(5, 2).Value & ", '"              'OrderNum
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value & "','"    'CustomerName
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(11, 2).Value & "', '"  'DocType
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 7).Value & "', '"   'Prod Date
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 8).Value & "', '"   'Ship Date
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(11, 3).Value & "', '"  'Industry
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 2).Value & "', '"  'EE
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 2).Value & "', '"  'CSE
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 2).Value & "', '"  'ME
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 4).Value & "', '"   'In To CPE Date
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 6).Value & "', '"  'EE Release
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 6).Value & "', '"  'ME Release
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 6).Value & "', '"  'EE Latest
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 6).Value & "', '"  'ME Latest
        
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 5).Value <> "" Then
        AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 5).Value & "', "   'Approval Date
    Else
        AssyStr = AssyStr & "', "
    End If
        
    AssyStr = AssyStr & (0# + Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 3).Value) & ", 0.0, "  'EE Estimate
    AssyStr = AssyStr & (0# + Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 3).Value) & ", 0.0, "  'ME Estimate
    AssyStr = AssyStr & (0# + Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 4).Value) & ", "       'Support Estimate
    AssyStr = AssyStr & (0# + Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 3).Value) & ", "       'Price
    AssyStr = AssyStr & "'', '', 0, '', 0, '', "  ' Act_EE_rel, Act_ME_rel, EE_offset, EE_override, ME_offset, ME_override
    
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 3).Value <> 0 Then
        AssyStr = AssyStr & "'IN QUEUE', "   'EE Status
    Else
        AssyStr = AssyStr & "'', "
    End If
    
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 3).Value <> 0 Then
        AssyStr = AssyStr & "'IN QUEUE', '', '"   'ME Status & Shippeddate
    Else
        AssyStr = AssyStr & "'', '', '"
    End If
        
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 2).Value & "', "                 'Region/AE
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(11, 4).Value & ", "                  'Active Year
    AssyStr = AssyStr & (0# + Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 4).Value) & ", '"          'Programming Estimate
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(19, 2).Value & "', "                 'Project Manager
    AssyStr = AssyStr & (0# + Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 4).Value) & ", '', '', '"  'PE Estimate
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 9).Value & "', '"                'SAP File
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 11).Value & "', '"    'Scheduler
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(2, 13).Value & "', '"     'REP
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(19, 3).Value & "', '"     'Multi-Plant
    AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(19, 4).Value & "'"        'Pre-Release Date
    
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "INSERT INTO cpe_main(uniqid, orderno, customer_name, doc_type, proddate, shipdate, industry, ee, se, me, into_CPE, EE_release, ME_release, EE_latest, ME_latest, Approvals, EE_est, CO_EE, ME_est, CO_ME, Supp_est, Price, Act_EE_rel, Act_ME_rel, EE_offset, EE_Override, ME_offset, ME_Override, EE_status, ME_status, shippeddate, region_AE, Active_Year, SE_est, PE, PE_est, EE_apps_out, ME_apps_out, SAP_file, scheduler, rep, MultiPlnt, ME_PreRel) VALUES(newid(), " & _
                   AssyStr & _
                   ")"
    oRS.Open

    If Worksheets(ShtOE).Cells(5, 3).Value > "" Then
        AssyStr = ""
        AssyStr = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value & ", '"            ' OrderNum
        AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 3).Value & "', '" ' WOs
        AssyStr = AssyStr & Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value & "', "  ' CustomerName
        AssyStr = AssyStr & "0" ' Step

        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "INSERT INTO cpe_order_xref(uniqid, orderno, wo, customer, step) VALUES(newid(), " & _
                    AssyStr & _
                    ")"
        oRS.Open
    Else
    End If
    
If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

' Additional VBA code for creating transmittals

MkDir "R:\Common Files\CPE\Temporary_Order_Files\" & Ord
    
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 2).Value = "N/A" Then
    Else
        MkDir "R:\Common Files\CPE\Temporary_Order_Files\" & Ord & "\" & Ord & "M"
        ShtERS = "R:\Common Files\CPE\Temporary_Order_Files\XXXXXXM.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "XXXXXXM.xlsm"
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(9, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value   'Customer Name
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(10, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value  'Order Number
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(11, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 10).Value 'PO Number
        Workbooks(ShtERS).Worksheets("CUSTOMER RELEASE").Cells(13, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 11).Value 'Designer Name
    End If
    
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 2).Value = "N/A" Then
    Else
        MkDir "R:\Common Files\CPE\Temporary_Order_Files\" & Ord & "\" & Ord & "E"
        ShtERS = "R:\Common Files\CPE\Temporary_Order_Files\XXXXXXE.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "XXXXXXE.xlsm"
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(9, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value    'Customer Name
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(10, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value   'Order Number
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(11, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 10).Value  'PO Number
        Workbooks(ShtERS).Worksheets("Customer Release").Cells(13, 4).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 11).Value 'Designer Name
    End If
    
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 2).Value = "N/A" Then
    Else
        ShtERS = "R:\Common Files\CPE\Temporary_Order_Files\F7-113 Eng Production Transmittal MR.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "F7-113 Eng Production Transmittal MR.xlsm"
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(8, 3).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value    'Customer Name
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(9, 3).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value    'Order Number
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(11, 3).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 11).Value 'Designer Name
    End If
    
    If Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 2).Value = "N/A" Then
    Else
        ShtERS = "R:\Common Files\CPE\Temporary_Order_Files\F7-113 Eng Production Transmittal ER.xlsm"
        Workbooks.Open Filename:=ShtERS
        ShtERS = "F7-113 Eng Production Transmittal ER.xlsm"
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(8, 3).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(9, 2).Value    'Customer Name
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(9, 3).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(5, 2).Value    'Order Number
        Workbooks(ShtERS).Worksheets("Cover_Sheet").Cells(11, 3).Value = Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 11).Value  'Designer Name
    End If

SKIP2:
    ' Clear Order Entry Sheet
Set ClrRange = Workbooks(WkBk).Worksheets(ShtOE).Range("A5:L5")
ClrRange.ClearContents
Set ClrRange = Workbooks(WkBk).Worksheets(ShtOE).Range("F9:I26")
ClrRange.ClearContents
Set ClrRange = Workbooks(WkBk).Worksheets(ShtOE).Range("B9:D9")
ClrRange.ClearContents

Workbooks(WkBk).Worksheets(ShtOE).Cells(11, 2).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(11, 3).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 2).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 2).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 2).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(19, 2).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(19, 3).Value = ""
Workbooks(WkBk).Worksheets(ShtOE).Cells(19, 4).Value = ""
    
Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 3).Value = 0
Workbooks(WkBk).Worksheets(ShtOE).Cells(13, 4).Value = 0
Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 3).Value = 0
Workbooks(WkBk).Worksheets(ShtOE).Cells(15, 4).Value = 0
Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 3).Value = 0
Workbooks(WkBk).Worksheets(ShtOE).Cells(17, 4).Value = 0
'

'
End Sub
Sub SendNew_Email()
'
Dim EBody As String, PC_EMail As String, ME_EMail As String
Dim OutApp As Object
Dim OutMail As Object

Ord = Workbooks(WkBk).Worksheets("Emails").Cells(3, 2).Value
CustName = Workbooks(WkBk).Worksheets("Emails").Cells(3, 3).Value
PC_Eng = Workbooks(WkBk).Worksheets("Emails").Cells(7, 5).Value
ME_Eng = Workbooks(WkBk).Worksheets("Emails").Cells(7, 6).Value
PC_EMail = Workbooks(WkBk).Worksheets("Emails").Cells(8, 7).Value
ME_EMail = Workbooks(WkBk).Worksheets("Emails").Cells(8, 8).Value

'CustName = Workbooks(WkBk).Worksheets("Emails").Cells(5, 7).Value

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

EBody = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"

EBody = EBody & ME_Eng & " and " & PC_Eng & ",<BR><BR>Please process the attached new order notification.<BR><BR>"
EBody = EBody & "Use the following link to review the latest order text:<BR><BR>"
EBody = EBody & Ord & "_" & CustName & "<BR><BR>"
EBody = EBody & "Thank you,<BR>Mark<BR><BR>"

    
    With OutMail
        .To = ME_EMail & "; " & PC_EMail
        .CC = "S.Burger@schenckprocess.com"
        .BCC = ""
        .Subject = "New Order " & Ord
        '.Body = strbody
        .HTMLBody = EBody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .Display '.Send   'or use .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
Sub Send_Change_Order()
'
Dim EBody As String, PC_EMail As String, ME_EMail As String
Dim OutApp As Object
Dim OutMail As Object

Ord = Workbooks(WkBk).Worksheets("Emails").Cells(3, 2).Value
CustName = Workbooks(WkBk).Worksheets("Emails").Cells(3, 3).Value
PC_Eng = Workbooks(WkBk).Worksheets("Emails").Cells(7, 5).Value
ME_Eng = Workbooks(WkBk).Worksheets("Emails").Cells(7, 6).Value
PC_EMail = Workbooks(WkBk).Worksheets("Emails").Cells(8, 7).Value
ME_EMail = Workbooks(WkBk).Worksheets("Emails").Cells(8, 8).Value

'CustName = Workbooks(WkBk).Worksheets("Emails").Cells(5, 7).Value

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

EBody = "<BASEFONT FACE='Calibri, arial, MS Sans Serif'>"

EBody = EBody & ME_Eng & " and " & PC_Eng & ",<BR><BR>Please process the attached change order notification.<BR><BR>"
EBody = EBody & "Use the following link to review the latest order text:<BR><BR>"
EBody = EBody & Ord & "_" & CustName & "<BR><BR>"
EBody = EBody & "Thank you,<BR>Mark<BR><BR>"

    
    With OutMail
        .To = ME_EMail & "; " & PC_EMail
        .CC = "S.Burger@schenckprocess.com"
        .BCC = ""
        .Subject = "Change Order for " & Ord
        '.Body = strbody
        .HTMLBody = EBody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .Display '.Send   'or use .Display
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing


End Sub
Sub WhoHasOrder()
' Enter Order Number to See CPE Order Info
'
Dim E_eng As String, M_eng As String, EngComments As String

Ord = Workbooks(WkBk).Worksheets("Emails").Cells(3, 2).Value

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

' Databdase Query
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "SELECT cpe_main.orderno, customer_name, into_CPE, ee, me, EE_status, ME_status, doc_type, approvals, EE_apps_out, ME_apps_out, EE_apps_back, ME_apps_back, MultiPlnt, Industry FROM cpe_main WHERE cpe_main.orderno=" & Ord & ";"
oRS.Open

Workbooks(WkBk).Worksheets("Emails").Select
Set ClrRange = Worksheets("Emails").Range("B3:O3")
ClrRange.ClearContents
Workbooks(WkBk).Worksheets("Emails").Range("B3").CopyFromRecordset oRS
    
Set ClrRange = Worksheets("Check").Range("A1:B500")
ClrRange.ClearContents

E_eng = Workbooks(WkBk).Worksheets("Emails").Cells(3, 5).Value
M_eng = Workbooks(WkBk).Worksheets("Emails").Cells(3, 6).Value

' Databdase Query
    If E_eng <> "N/A" Then
        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "SELECT cpe_schedule.datestamp, comments FROM cpe_schedule WHERE cpe_schedule.engineer='" & E_eng & "' AND comments>' ' AND orderno=" & Ord & " ORDER BY cpe_schedule.datestamp;"
        oRS.Open
    
    ' Retrieve EE Notes
        'Set R = DB.OpenRecordset(qsql, dbOpenDynaset)
        'If R.RecordCount > 0 Then
            Set ClrRange = Workbooks(WkBk).Worksheets("Check").Range("A1:B202")
            ClrRange.ClearContents
            Workbooks(WkBk).Worksheets("Check").Range("A3").CopyFromRecordset oRS
        'Else
        'End If
    Else
    End If
    
' Databdase Query
    If M_eng <> "N/A" Then
        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "SELECT cpe_Schedule.datestamp, comments FROM cpe_schedule WHERE cpe_schedule.engineer='" & M_eng & "' AND comments>' ' AND orderno=" & Ord & " ORDER BY cpe_schedule.datestamp;"
        oRS.Open
    ' Retrieve ME Notes
        'Set R = DB.OpenRecordset(qsql, dbOpenDynaset)
        'If R.RecordCount > 0 Then
            Set ClrRange = Workbooks(WkBk).Worksheets("Check").Range("A203:B500")
            ClrRange.ClearContents
            Workbooks(WkBk).Worksheets("Check").Range("A203").CopyFromRecordset oRS
        'Else
        'End If
    Else
    End If

oRS.Close
oCon.Close

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing
    
    EngComments = ""
    
    For I = 3 To 202
    If Workbooks(WkBk).Worksheets("Check").Cells(I, 2).Value <> "" Then
        EngComments = EngComments & Chr(10) & "[" & Workbooks(WkBk).Worksheets("Check").Cells(I, 5).Value & "/" & Workbooks(WkBk).Worksheets("Check").Cells(I, 6).Value & "] " & Workbooks(WkBk).Worksheets("Check").Cells(I, 2).Value
    Else
        Exit For
    End If
    Next I
    
    Workbooks(WkBk).Worksheets("Emails").Cells(5, 7).Value = EngComments
    EngComments = ""
    
    For I = 203 To 500
    If Workbooks(WkBk).Worksheets("Check").Cells(I, 2).Value <> "" Then
        EngComments = EngComments & Chr(10) & "[" & Workbooks(WkBk).Worksheets("Check").Cells(I, 5).Value & "/" & Workbooks(WkBk).Worksheets("Check").Cells(I, 6).Value & "] " & Workbooks(WkBk).Worksheets("Check").Cells(I, 2).Value
    Else
        Exit For
    End If
    Next I
    
    Workbooks(WkBk).Worksheets("Emails").Cells(5, 8).Value = EngComments
    EngComments = ""

'oRS.Close
'DB.Close
End Sub
Sub Publish2Intranet()
'
Dim HTMPath As String, HtmFn As String
Dim StDate As String, EnDate As String, Grp As String

HTMPath = "\\pac5Intra\Intranet\Documents\DepartmentReports\Production_Engineering\Current_Status\"
HtmFn = HTMPath & "_Active_Orders.HTM"
WkSht = "Check"

StDate = Workbooks(WkBk).Worksheets("Check").Cells(1, 11).Value
EnDate = Workbooks(WkBk).Worksheets("Check").Cells(2, 11).Value

Set ClrRange = Workbooks(WkBk).Worksheets("Check").Range("M3:Z500")
ClrRange.ClearContents

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open

    ' Retrieve Database Information
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    ' Retrieve records
    oRS.Source = "SELECT cpe_main.orderno, customer_name, into_CPE, EE_latest, ME_latest, ee, me, Act_EE_rel, Act_ME_rel, EE_status, ME_status, Approvals, EE_Apps_out, ME_apps_out FROM cpe_main WHERE (cpe_main.EE_Latest>'1/1/2015' AND Act_EE_rel<'1/3/1900') OR (ME_Latest>'1/1/2015' AND Act_ME_rel<'1/3/2000') ORDER BY cpe_main.orderno"
    oRS.Open
    ' Retrieve records
    Set ClrRange = Workbooks(WkBk).Worksheets(WkSht).Range("M3:Z500")
    ClrRange.ClearContents
    Workbooks(WkBk).Worksheets(WkSht).Range("M3").CopyFromRecordset oRS

Open HtmFn For Output As #1

    Print #1, "<HTML>"
    Print #1, "<HEAD>"
    Print #1, "<TITLE>Production Engineering Active Schedule</TITLE>"
    Print #1, "</HEAD>"
    Print #1, ""
    Print #1, "<BODY>"
    Print #1, "<BASEFONT FACE='arial, helvetica, tahoma, sans-serif'>"
    Print #1, "<TABLE Border=True BorderColor='#a0a0a0' BGColor='#ffffff' CellPadding=4 CellSpacing=0>"
    Print #1, ""
    Print #1, "<TR><TH Colspan=12 BGColor='00ffff'>Production Engineering - Active Orders - Update: "; Date$; "</TH></TR>"
    Print #1, "<TR>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Order Number</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Customer Name</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Into CPE</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Approvals Out</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Planned EE Release</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Planned ME Release</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Elec Eng</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Mech Eng</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>EE Status</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>ME Status</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>EE Appv Out</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>ME Appv Out</FONT></TH>"
    Print #1, "</TR>"

For I = 3 To 200

    If Worksheets(WkSht).Cells(I, 13).Value > 0 Then
        Print #1, ""
        Print #1, "<TR>"
        Print #1, "<TD Align=Center><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 13).Value; "</FONT></TD>"  ' Order Number
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 14).Value; "</FONT></TD>"        ' Customer Name
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 15).Value; "</FONT></TD>"        ' Into CPE
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 24).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 24).Value; "</FONT></TD>"    ' Approvals Out
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 16).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 16).Value; "</FONT></TD>"    ' Latest EE Rel
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 17).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 17).Value; "</FONT></TD>"    ' Latest ME Rel
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 18).Value; "</FONT></TD>"        ' EE
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 19).Value; "</FONT></TD>"        ' ME
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 22).Value; "</FONT></TD>"        ' EE Status
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 23).Value; "</FONT></TD>"        ' ME Status
        
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 25).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 25).Value; "</FONT></TD>"        ' EE Approvals Out
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 26).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 26).Value; "</FONT></TD>"        ' ME Approvals Out
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        
        Print #1, "</TR>"
    Else
        Exit For
    End If

Next I
   
   Print #1, ""
   Print #1, "</TABLE>"
   Print #1, ""
   Print #1, "</BODY>"
   Print #1, "</HTML>"

Close

Set ClrRange = Workbooks(WkBk).Worksheets("Check").Range("M3:Z500")
ClrRange.ClearContents

' Retrieve Recent Releases Information
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT cpe_main.orderno, customer_name, into_CPE, EE_latest, ME_latest, ee, me, Act_EE_rel, Act_ME_rel FROM cpe_main WHERE (cpe_main.Act_EE_rel BETWEEN '" & StDate & "' AND '" & EnDate & "') OR (Act_ME_rel BETWEEN '" & StDate & "' AND '" & EnDate & "') ORDER BY cpe_main.orderno"
    ' Retrieve records
    oRS.Open
    Set ClrRange = Workbooks(WkBk).Worksheets(WkSht).Range("M3:W500")
    ClrRange.ClearContents
    Workbooks(WkBk).Worksheets(WkSht).Range("M3").CopyFromRecordset oRS

'HtmFn = HTMPath & "_CPE2_Schedule.HTM"
HtmFn = HTMPath & "_Recent_Releases.HTM"

Open HtmFn For Output As #1

    Print #1, "<HTML>"
    Print #1, "<HEAD>"
    Print #1, "<TITLE>Production Engineering Recent Releases</TITLE>"
    Print #1, "</HEAD>"
    Print #1, ""
    Print #1, "<BODY>"
    Print #1, "<BASEFONT FACE='arial, helvetica, tahoma, sans-serif'>"
    Print #1, "<TABLE Border=True BorderColor='#a0a0a0' BGColor='#ffffff' CellPadding=4 CellSpacing=0>"
    Print #1, ""
    Print #1, "<TR><TH Colspan=11 BGColor='00ffff'>Production Engineering - Recent Releases - Update: "; Date$; "</TH></TR>"
    Print #1, "<TR>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Order Number</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Customer Name</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Into CPE</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Planned EE Release</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Planned ME Release</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Elec Eng</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Mech Eng</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Actual EE Release</FONT></TH>"
    Print #1, "<TH BGColor='#00ffff'><FONT Size=-1>Actual ME Release</FONT></TH>"
    Print #1, "</TR>"

For I = 3 To 200

    If Worksheets(WkSht).Cells(I, 13).Value > 0 Then
        Print #1, ""
        Print #1, "<TR>"
        Print #1, "<TD Align=Center><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 13).Value; "</FONT></TD>"  ' Order Number
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 14).Value; "</FONT></TD>"        ' Customer Name
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 15).Value; "</FONT></TD>"        ' Into CPE
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 16).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 16).Value; "</FONT></TD>"        ' Latest EE Rel
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 17).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 17).Value; "</FONT></TD>"        ' Latest ME Rel
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 18).Value; "</FONT></TD>"        ' EE
        Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 19).Value; "</FONT></TD>"        ' ME
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 20).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 20).Value; "</FONT></TD>"        ' Actual EE Rel
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        If Workbooks(WkBk).Worksheets(WkSht).Cells(I, 21).Value > 40000 Then
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; "; Workbooks(WkBk).Worksheets(WkSht).Cells(I, 21).Value; "</FONT></TD>"        ' Actual ME Rel
        Else
            Print #1, "<TD NoWrap><FONT Size=-1>&nbsp; </FONT></TD>"
        End If
        Print #1, "</TR>"
    Else
        Exit For
    End If

Next I
   
   Print #1, ""
   Print #1, "</TABLE>"
   Print #1, ""
   Print #1, "</BODY>"
   Print #1, "</HTML>"

Close

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
Sub DeleteARecord()
'Delete a record

Set oCon = New ADODB.Connection
oCon.ConnectionString = "Driver={SQl Server};Server=USLXA-P-SQL01;Database=PAC1CPE;User ID=dash;Password=manage_DB;" ' Trusted_Connection=yes;"
oCon.Open
 
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon

LastRow = Workbooks(WkBk).Worksheets(Sht6).Cells(Rows.Count, 24).End(xlUp).Row

For Index = 3 To LastRow
    If UCase(Workbooks(WkBk).Worksheets(Sht6).Cells(Index, 23).Value) = "YES" Then
        oRS.Source = "DELETE FROM Prod_Eng WHERE Prod_Eng.uniqID = '" & Workbooks(WkBk).Worksheets(Sht6).Cells(Index, 24).Value & "'"
        oRS.Open
    End If
Next Index

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

Call Order_Updates
End Sub
