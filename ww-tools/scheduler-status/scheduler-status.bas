Attribute VB_Name = "Module1"
' Schedulers Status Tool
' Original - Version 1.0 - Dated 12/1/2012
' SQL Server Express version - Dated 7/19/2013
'
Public Const DBPath = "\\pac5Intra\intranet\Database\CPE\"
Public Const SourcePath = "W:\Manufacturing\Projects\Data\"
Public Row_E_Idx As Long
Public Row_P_Idx As Long
Public SAP_ID As Long
Public Row_OTD_Idx As Long
Public MPCancel As Boolean
Public ClrRange As Range
Public Scheduler As String, PC_Confirm As String, NewPCConf As String, PC_ConfirmPrev As String
Public Tech_Name As String
Public WBook As String
Public WkSp As Workspace, DB As Database, Q As QueryDef, R As Recordset

Public Const serverName As String = "USLXA-P-SQL01"
Public Const CPEdatabaseName As String = "PAC1CPE"
Public Const OrdersDatabaseName As String = "Dashboards"
Public Const databaseUserID As String = "dash"
Public Const databasePassword As String = "manage_DB"

'Public Const serverName As String = "PAC5Intra03\PAC5SQLExpress"
'Public Const CPEdatabaseName As String = "PAC1CPE"
'Public Const OrdersDatabaseName As String = "Dashboards"
'Public Const databaseUserID As String = "sa"
'Public Const databasePassword As String = "manage_ERP"

Public Const PAC1CPEdbConnectionString As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & CPEdatabaseName & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"

Public Const PAC1OrdersDbConnectionString As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & OrdersDatabaseName & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"


Sub Scheduler_Update()
Attribute Scheduler_Update.VB_ProcData.VB_Invoke_Func = " \n14"
'
Dim I As Long, PrevOrd As Long, Ord As Long, Search_ID As Long, Flag1 As Long, J As Long, K As Long, OrdNum As Long
Dim CPE_EE_Status As String, CPE_ME_Status As String, Stuff As String, Reasons(10) As String, Firmed As String, LnDesc As String
Dim OrdNotes As String, AccumNotes As String, Ord_Status As String, Act_P_Date As Date, ColorRng As Range, BlnkLn As String
Dim oCon As ADODB.Connection
Dim oRS As ADODB.Recordset

WBook = "Scheduler_Status.xlsm"

Scheduler = UCase(Sheet1.Cells(2, 1).Value)
PrevOrd = 0: Ord = 0
Row_E_Idx = 4
Row_P_Idx = 10
Flag1 = 0
Reasons(1) = "": Reasons(2) = "Customer postponed Delivery"
Reasons(3) = "Error in Schedule Planning": Reasons(4) = "Missing Material"
Reasons(5) = "Engineering Late": Reasons(6) = "Production short of works capacity"
Reasons(7) = "Late delivery caused by customer": Reasons(8) = "Delivery time confirmed by sales too short"
Reasons(9) = "New product development / R&D Changes": Reasons(10) = ""

Set ClrRange = Sheet1.Range("A5:T510")
ClrRange.ClearContents
Set ClrRange = Sheet2.Range("A1:Z50000")
ClrRange.ClearContents
Set ClrRange = Sheet5.Range("A1:Z50000")
ClrRange.ClearContents

Set ColorRng = Sheet1.Range("E5:E510")
ColorRng.Select
    With Selection.Interior
        .Pattern = xlNone ' Clear Background Color
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

Set ColorRng = Sheet1.Range("O4:O110")
ColorRng.Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With

Set ColorRng = Sheet1.Range("A2:A2")
ColorRng.Select

Open SourcePath & "Open.XLS" For Input As #1
I = 2: J = 1

Do Until EOF(1)
    InpStr = Input$(1, #1)
    If InpStr = Chr(13) Then InpStr = Input$(1, #1): Sheet2.Cells(I, J).Value = Stuff: Stuff = "": I = I + 1: J = 1: InpStr = ""
    If InpStr = Chr(9) Then Sheet2.Cells(I, J).Value = Stuff: Stuff = "": J = J + 1: InpStr = ""
    Stuff = Stuff & InpStr
Loop

Close #1

Open SourcePath & "Parts.XLS" For Input As #1
I = 2: J = 1

Do Until EOF(1)
    InpStr = Input$(1, #1)
    If InpStr = Chr(13) Then InpStr = Input$(1, #1): Sheet5.Cells(I, J).Value = Stuff: Stuff = "": I = I + 1: J = 1: InpStr = ""
    If InpStr = Chr(9) Then Sheet5.Cells(I, J).Value = Stuff: Stuff = "": J = J + 1: InpStr = ""
    Stuff = Stuff & InpStr
Loop

Close #1
'==============================================================================================
SAP_ID = 11100132
Application.ScreenUpdating = False

For I = 2 To 200
    If Scheduler = UCase(Sheet3.Cells(I, 3).Value) Then
        SAP_ID = Sheet3.Cells(I, 1).Value
        Sheet1.Cells(1, 2).Value = "Orders For " & Sheet3.Cells(I, 4).Value
        Tech_Name = Sheet3.Cells(I, 4).Value
        Exit For
    Else
    End If
Next I

For J = 8 To 5000
    PrevOrd = Ord: Ord = Sheet2.Cells(J, 3).Value
    Search_ID = Sheet2.Cells(J, 18).Value
    BlnkLn = Sheet2.Cells(J, 2).Value
    If (Ord = 0 And BlnkLn <> "*") Then Exit For
    If SAP_ID = Search_ID Then
        Flag1 = 1
        If Ord <> PrevOrd Then
            LnDesc = Sheet2.Cells(J, 7).Value
            If (LnDesc = "Field Service Start Up (MTO)" Or LnDesc = "CUSTOM ENGINEERING, TRAVEL, PROGRAMMING,") Then
            ' Do Nothing
            Else
                Row_E_Idx = Row_E_Idx + 1
                Sheet1.Cells(Row_E_Idx, 2).Value = Ord                                                    ' Order Number
                Sheet1.Cells(Row_E_Idx, 3).Value = Sheet2.Cells(J, 11).Value ' Customer
            
                If Sheet2.Cells(J, 21).Value = "00/00/0000" Or Sheet2.Cells(J, 21).Value = "" Then
                    Sheet1.Cells(Row_E_Idx, 4).Value = 44196 ' Committed Ship Date
                Else
                    Sheet1.Cells(Row_E_Idx, 4).Value = Sheet2.Cells(J, 21).Value ' Committed Ship Date
                End If
            
                Sheet1.Cells(Row_E_Idx, 5).Value = Sheet2.Cells(J, 9).Value  ' Scheduled Ship Date
                Sheet1.Cells(Row_E_Idx, 16).Value = Reasons(Val(Sheet2.Cells(J, 23).Value) + 1)
                If Sheet2.Cells(J, 21).Value = "00/00/0000" Then Firmed = "" Else Firmed = "X"
                Sheet1.Cells(Row_E_Idx, 1).Value = Firmed
            End If
        Else
        End If
    Else
    End If
    
    If (Search_ID = 0 And Flag1 = 1) Then
        Sheet1.Cells(Row_E_Idx, 7).Value = Sheet2.Cells(J, 15).Value
        Flag1 = 0
    Else
    End If
    

Next J

Row_P_Idx = Row_E_Idx + 2
Sheet1.Cells(Row_P_Idx, 2).Value = "P A R T S               "
Row_P_Idx = Row_P_Idx + 0

Ord = 1: Flag1 = 0

For J = 8 To 5000
    PrevOrd = Ord: Ord = Sheet5.Cells(J, 3).Value
    Search_ID = Sheet5.Cells(J, 18).Value
    BlnkLn = Sheet5.Cells(J, 2).Value
    If (Ord = 0 And BlnkLn <> "*") Then Exit For
        
    If SAP_ID = Search_ID Then
        Flag1 = 1
        If Ord <> PrevOrd Then
            Row_P_Idx = Row_P_Idx + 1
            Sheet1.Cells(Row_P_Idx, 2).Value = Ord                                                    ' Order Number
            Sheet1.Cells(Row_P_Idx, 3).Value = Sheet5.Cells(J, 11).Value ' Customer
            
            If Sheet5.Cells(J, 21).Value = "00/00/0000" Or Sheet5.Cells(J, 21).Value = "" Then
                Sheet1.Cells(Row_P_Idx, 4).Value = 44196 ' Committed Ship Date
            Else
                Sheet1.Cells(Row_P_Idx, 4).Value = Sheet5.Cells(J, 21).Value ' Committed Ship Date
            End If

            Sheet1.Cells(Row_P_Idx, 5).Value = Sheet5.Cells(J, 9).Value  ' Scheduled Ship Date
            Sheet1.Cells(Row_P_Idx, 16).Value = Reasons(Val(Sheet5.Cells(J, 23).Value) + 1)
            If Sheet5.Cells(J, 21).Value = "00/00/0000" Then Firmed = "" Else Firmed = "X"
            Sheet1.Cells(Row_P_Idx, 1).Value = Firmed
        Else
        End If
    Else
        Flag1 = 0
    End If
    
    If (Search_ID = 0 And Flag1 = 1) Then
        Sheet1.Cells(Row_P_Idx, 7).Value = Sheet5.Cells(J, 15).Value
        Flag1 = 0
    Else
    End If
Next J

' == Make Connections to CPE Database and Dashboard ==========================

Set oCon = New ADODB.Connection
oCon.ConnectionString = PAC1CPEdbConnectionString ' Trusted_Connection=yes;"
oCon.Open

For I = 5 To Row_P_Idx

    ' == Retrieve Data From CPE Database ==========================
    If I = (Row_E_Idx + 2) Then I = I + 2
    
    Ord = Sheet1.Cells(I, 2).Value
    If Ord < 1 Then Ord = 1
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    oRS.Source = "SELECT Prod_Eng.Line_Num, Order_Num, Customer_Name, DocTypes, Prod_Rel_0, Ship_Date_0, Industry, PC1, PC2, ME1, Into_Eng, PC_Rel_0, ME_Rel_0, PC_Rel_F, ME_Rel_F, Apps_Out_0, PC_Est, PC_CO_Est, ME_Est, ME_CO_Est, Supp_Est, PM_Est, PC_Act_Rel, ME_Act_Rel, Line_Num, PC_Override, Line_Num, ME_Override, PC_Status, ME_Status, ShippedDate, Region_AE, Active_Year, Soft_Est, Pjt_Lvl, PM_Est, PC_Apps_Out, ME_Apps_Out, FileNum, Scheduler  FROM Prod_Eng WHERE Prod_Eng.Order_Num=" & Ord & ";"
    ' Retrieve records
    oRS.Open
    Set ClrRange = Sheet4.Range("A53:AN99")
    ClrRange.ClearContents
    Sheet4.Range("A53").CopyFromRecordset oRS
    
    
    If Len(Sheet4.Cells(53, 29).Value) = 0 Then
        CPE_EE_Status = "-"
    Else
        CPE_EE_Status = Mid(Sheet4.Cells(53, 29).Value, 1, 1)
    End If
        
    If Len(Sheet4.Cells(53, 30).Value) = 0 Then
        CPE_ME_Status = "-"
    Else
        CPE_ME_Status = Mid(Sheet4.Cells(53, 30).Value, 1, 1)
    End If
   
    If Ord = 1 Then GoTo NO_UPDATE_1
    
    Sheet1.Cells(I, 8).Value = CPE_EE_Status & " " & CPE_ME_Status
        
    If Sheet4.Cells(53, 35).Value <> "" Then
        Sheet1.Cells(I, 9).Value = Sheet4.Cells(53, 8).Value & "/" & Sheet4.Cells(53, 35).Value
        Sheet1.Cells(I, 10).Value = Sheet4.Cells(53, 10).Value & "/" & Sheet4.Cells(53, 35).Value
    Else
        Sheet1.Cells(I, 9).Value = Sheet4.Cells(53, 8).Value
        Sheet1.Cells(I, 10).Value = Sheet4.Cells(53, 10).Value
    End If
    If Sheet4.Cells(53, 23).Value > 40909 Then
        Sheet1.Cells(I, 11).Value = Sheet4.Cells(53, 23).Value
    Else
    End If
    If Sheet4.Cells(53, 24).Value > 40909 Then
        Sheet1.Cells(I, 12).Value = Sheet4.Cells(53, 24).Value
    Else
    End If
    If Sheet4.Cells(1, 22).Value > 40909 Then
        Sheet1.Cells(I, 13).Value = Sheet4.Cells(1, 22).Value + 2
    Else
        If Sheet4.Cells(53, 5).Value > 40909 Then
            Sheet1.Cells(I, 13).Value = Sheet4.Cells(53, 5).Value + 2
        Else
        End If
    End If

' ========
NO_UPDATE_1:

    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    ' Retrieve records
    oRS.Source = "SELECT cpe_shipmeet.orderno, comments, status, act_prod_date, PC_Confirmed FROM cpe_shipmeet WHERE cpe_shipmeet.orderno=" & Ord & " ORDER BY cpe_shipmeet.datestamp;"
    oRS.Open

    Set ClrRange = Sheet4.Range("B15:F51")
    ClrRange.ClearContents
        
    Sheet4.Range("B15").CopyFromRecordset oRS
    AccumNotes = ""

    For K = 15 To 35
        
        OrdNum = Sheet4.Cells(K, 2).Value
        OrdNotes = Sheet4.Cells(K, 3).Value
        If OrdNum <> 0 Then
            AccumNotes = AccumNotes + " " + Sheet4.Cells(K, 3).Value
            If Sheet4.Cells(K, 6).Value > "" Then
                NewPCConf = Sheet4.Cells(K, 6).Value
            Else
            End If
            'Ord_Status = Sheet4.Cells(K, 4).Value
            'Act_P_Date = Sheet4.Cells(K, 5).Value
        Else
            Sheet1.Cells(I, 19).Value = AccumNotes
            Sheet1.Cells(I, 14).Value = NewPCConf
            Sheet1.Cells(I, 20).Value = NewPCConf
            'If Act_P_Date > 0 Then Sheet1.Cells(I, 14).Value = Act_P_Date
            Exit For
        End If
    
    Next K
    NewPCConf = ""
' ======
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    ' Databdase Query
    oRS.Source = "SELECT cpe_materials.orderno, prod_date, status, override FROM cpe_materials WHERE cpe_materials.orderno=" & Ord & ";"
    oRS.Open

    Set ClrRange = Sheet4.Range("D10:G10")
    ClrRange.ClearContents
        
    Sheet4.Range("D10").CopyFromRecordset oRS
    If Sheet4.Cells(10, 5).Value > 40908 Then
        Sheet1.Cells(I, 15).Value = Sheet4.Cells(10, 5).Value
    Else
    End If
    Sheet1.Cells(I, 17).Value = Sheet4.Cells(10, 6).Value
    
' == Retrieve Data From Dashboard ==========================

    ActiveWorkbook.Worksheets("Scratch").Select
    Range("A10").Select

    Set ClrRange = Sheet4.Range("A10:B11")
    ClrRange.ClearContents
    
    Set oCon2 = New ADODB.Connection
    oCon2.ConnectionString = PAC1OrdersDbConnectionString ' Trusted_Connection=yes;"
    oCon2.Open
    Set oRS2 = New ADODB.Recordset

    oRS2.ActiveConnection = oCon2 'HERE
    oRS2.Source = "Select wwo_custorder.orderno, hold_code, reviewdate From wwo_custorder Where wwo_custorder.orderno='" & Ord & "'"
    oRS2.Open
 
    Range("A10").CopyFromRecordset oRS2
    
    If Ord = 1 Then GoTo NO_UPDATE_2
    
    Sheet1.Cells(I, 6).Value = Sheet4.Cells(10, 2).Value

    If Sheet1.Cells(I, 13).Value < 40000 Then
        Sheet1.Cells(I, 13).Value = Sheet4.Cells(10, 3).Value + 2
    Else
    End If
    
NO_UPDATE_2:
    
Next I

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

If Not oRS2 Is Nothing Then Set oRS2 = Nothing
If Not oCon2 Is Nothing Then Set oCon2 = Nothing

Application.ScreenUpdating = True
    
ActiveWorkbook.Worksheets("OTD").Select
Range("B4").Select

Call OTD_Metric

Call Format
'
End Sub
Sub OTD_Metric()
'
' On Time Delivery Macro
'
Dim I As Long, PrevOrd As Long, Ord As Long, Search_ID As Long, Flag1 As Long
Dim CPE_EE_Status As String, CPE_ME_Status As String, Stuff As String, Reasons(10) As String
Dim SchedShip As Date
Dim oCon As ADODB.Connection
Dim oRS As ADODB.Recordset

WBook = "Scheduler_Status.xlsm"

Scheduler = Sheet1.Cells(2, 1).Value
PrevOrd = 0: Ord = 0
Row_OTD_Idx = 3
Reasons(1) = "": Reasons(2) = "Customer postponed Delivery"
Reasons(3) = "Error in Schedule Planning": Reasons(4) = "Missing Material"
Reasons(5) = "Engineering Late": Reasons(6) = "Production short of works capacity"
Reasons(7) = "Late delivery caused by customer": Reasons(8) = "Delivery time confirmed by sales too short"
Reasons(9) = "New product development / R&D Changes": Reasons(10) = ""

Set ClrRange = Sheet7.Range("A1:Z10000")
ClrRange.ClearContents
Set ClrRange = Sheet6.Range("A4:I5000")
ClrRange.ClearContents
Set ClrRange = Sheet6.Range("K4:K5000")
ClrRange.ClearContents

Open SourcePath & "Shipped.XLS" For Input As #1
I = 2: J = 1

Do Until EOF(1)
    InpStr = Input$(1, #1)
    If InpStr = Chr(13) Then InpStr = Input$(1, #1): Sheet7.Cells(I, J).Value = Stuff: Stuff = "": I = I + 1: J = 1: InpStr = ""
    If InpStr = Chr(9) Then Sheet7.Cells(I, J).Value = Stuff: Stuff = "": J = J + 1: InpStr = ""
    Stuff = Stuff & InpStr
Loop

Close #1
Application.ScreenUpdating = False

Set oCon = New ADODB.Connection
oCon.ConnectionString = PAC1CPEdbConnectionString ' Trusted_Connection=yes;"
oCon.Open

Flag1 = 0

For J = 8 To 5000
    PrevOrd = Ord: Ord = Sheet7.Cells(J, 3).Value
    Search_ID = Sheet7.Cells(J, 21).Value
    
    If Sheet7.Cells(J, 19).Value = "00/00/0000" Or Sheet7.Cells(J, 19).Value = "" Then
        SchedShip = 44196
    Else
        SchedShip = Sheet7.Cells(J, 19).Value
    End If
    
    If Ord = 0 Then Exit For
        
    If SAP_ID = Search_ID Then
        Flag1 = 1
        If Ord <> PrevOrd Then
            Row_OTD_Idx = Row_OTD_Idx + 1
            Sheet6.Cells(Row_OTD_Idx, 2).Value = Ord                                                    ' Order Number
            Sheet6.Cells(Row_OTD_Idx, 3).Value = Sheet7.Cells(J, 10).Value ' Customer
            Sheet6.Cells(Row_OTD_Idx, 4).Value = Tech_Name
            Sheet6.Cells(Row_OTD_Idx, 7).Value = SchedShip  ' Scheduled Ship Date
            Sheet6.Cells(Row_OTD_Idx, 8).Value = Sheet7.Cells(J, 16).Value  ' Actual Ship Date
            Sheet6.Cells(Row_OTD_Idx, 9).Value = Reasons(Val(Sheet7.Cells(J, 20).Value) + 1)
        
            ' Databdase Query
            Set oRS = New ADODB.Recordset
            oRS.ActiveConnection = oCon
            oRS.Source = "SELECT cpe_materials.orderno, prod_date, status, override FROM cpe_materials WHERE cpe_materials.orderno=" & Ord & ";"
            ' Retrieve records
            oRS.Open
            Set ClrRange = Sheet4.Range("D10:G10")
            ClrRange.ClearContents
            Sheet4.Range("D10").CopyFromRecordset oRS
            If Sheet4.Cells(10, 5).Value > 3200 Then
                Sheet6.Cells(Row_OTD_Idx, 6).Value = Sheet4.Cells(10, 5).Value ' Scheduled Release
            Else
            End If
            Sheet6.Cells(Row_OTD_Idx, 11).Value = Sheet4.Cells(10, 7).Value ' Override
        Else
        End If
    Else
        Flag1 = 0
    End If

Next J

Application.ScreenUpdating = True

ActiveWorkbook.Worksheets("Summary").Select
Range("B4").Select

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing
'
End Sub
Sub New_Updates()
'
Dim I As Long, J As Long, Imax As Long, Flg1 As Long, ONUM As Long, qdf As QueryDef
Dim Act_Prd_Date As Date, NewStatus As String, NewComment As String, ShipDate As Date

WBook = "Scheduler_Status.xlsm"
Flg1 = 0

Set oCon = New ADODB.Connection
oCon.ConnectionString = PAC1CPEdbConnectionString ' Trusted_Connection=yes;"
oCon.Open

For I = 5 To 5000
    
    ONUM = Val(Sheet1.Cells(I, 2).Value)
    Act_Prd_Date = Sheet1.Cells(I, 15).Value
    Ship_Date = Sheet1.Cells(I, 5).Value
    PC_Confirm = Sheet1.Cells(I, 14).Value
    PC_ConfirmPrev = Sheet1.Cells(I, 20).Value
    NewStatus = Sheet1.Cells(I, 17).Value
    NewComment = Sheet1.Cells(I, 18).Value
    FiltNotes = ""
    For K = 1 To Len(NewComment)
        If (Mid(NewComment, K, 1) = Chr(34) Or Mid(NewComment, K, 1) = Chr(39) Or Mid(NewComment, K, 1) = Chr(10)) Then
            FiltNotes = FiltNotes & " "
        Else
            FiltNotes = FiltNotes & Mid(NewComment, K, 1)
        End If
    Next K
    
    If NewComment <> "" Or PC_Confirm <> PC_ConfirmPrev Then
        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "INSERT INTO cpe_shipmeet(uniqid, orderno, datestamp, expectship, comments, act_prod_date, status, PC_Confirmed) VALUES(newid(), " & _
                      ONUM & ", '" & Date$ & "', '" & ShipDate & "', '" & FiltNotes & "', '', '', '" & PC_Confirm & "')"
        oRS.Open
        Sheet1.Cells(I, 19).Value = Sheet1.Cells(I, 19).Value + NewComment
        Sheet1.Cells(I, 18).Value = ""
        Sheet1.Cells(I, 14).Value = ""
    Else
        If ONUM = 0 Then Flg1 = Flg1 + 1
        If Flg1 > 9 Then Exit For
    End If
 
    If Act_Prd_Date > 40909 Or NewStatus <> "" Then
        ' See if order is in table
        ' Databdase Query
        
        Set oRS = New ADODB.Recordset
        oRS.ActiveConnection = oCon
        oRS.Source = "SELECT cpe_materials.orderno FROM cpe_materials WHERE cpe_materials.orderno=" & ONUM & ";"
        ' Retrieve records
        oRS.Open
        Set ClrRange = Sheet4.Range("Q15:Q15")
        ClrRange.ClearContents
        ClrRange.CopyFromRecordset oRS
        
        If Sheet4.Cells(15, 17).Value = 0 Then
            'Can Not Find Order in Database - so add it
            Set oRS = New ADODB.Recordset
            oRS.ActiveConnection = oCon
            oRS.Source = "INSERT INTO cpe_materials(uniqid, orderno, prod_date, status, override) VALUES(newid(), " & _
                          ONUM & ", '" & Act_Prd_Date & "', '" & NewStatus & "', '')"
            oRS.Open
        Else
            Set oRS = New ADODB.Recordset
            oRS.ActiveConnection = oCon
            oRS.Source = "UPDATE cpe_materials SET cpe_materials.prod_date='" & Act_Prd_Date & "', status='" & NewStatus & "' WHERE cpe_materials.orderno=" & ONUM
            oRS.Open
        End If
    Else
        If ONUM = 0 Then Flg1 = Flg1 + 1
        If Flg1 > 9 Then Exit For
        
    End If

Next I
    
If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
Sub Mfg_Eng()
'
Dim I As Long, J As Long, Imax As Long, ONUM As Long
Dim Act_Prd_Date As Date, NewStatus As String, NewComment As String, ShipDate As Date
Dim oCon As ADODB.Connection
Dim oRS As ADODB.Recordset

WBook = "Scheduler_Status.xlsm"

Set oCon = New ADODB.Connection
oCon.ConnectionString = PAC1CPEdbConnectionString ' Trusted_Connection=yes;"
oCon.Open

'Set R = DB.OpenRecordset("Mfg_Eng", dbOpenDynaset)  ' Open Access Database Table 'CPE_Sched'

For I = 5 To 10

' == Retrieve Data From CPE Database ==========================
  
    Ord = Sheet1.Cells(I, 2).Value
    ' Databdase Query
    Set oRS = New ADODB.Recordset
    oRS.ActiveConnection = oCon
    ' Retrieve records
    oRS.Source = "SELECT cpe_mfg_eng.orderno, in_date, datatype, out_date, comments, hot, large FROM cpe_mfg_eng WHERE (cpe_mfg_eng.orderno=" & Ord & " AND datatype = " & Chr(34) & "MECH. PROD. RELEASE" & Chr(34) & ");"
    oRS.Open
   
    Set ClrRange = Sheet4.Range("W10:AE10")
    ClrRange.ClearContents
    Sheet4.Range("W10").CopyFromRecordset oRS
    Sheet1.Cells(I, 12).Value = Sheet4.Cells(10, 27).Value
    
Next I

oRS.Close
oCon.Close

' == Retrieve Data From Dashboard ==========================

    ActiveWorkbook.Worksheets("Scratch").Select
    Range("A10").Select

    Set ClrRange = Sheet4.Range("A10:B11")
    ClrRange.ClearContents
    
    Set oCon = New ADODB.Connection
    oCon.ConnectionString = PAC1CPEdbConnectionString ' Trusted_Connection=yes;"
    oCon.Open
    Set oRS = New ADODB.Recordset

    oRS.ActiveConnection = oCon 'HERE
    oRS.Source = "Select wwo_custorder.orderno, hold_code From custorder Where wwo_custorder.orderno=" & Ord
    oRS.Open
 
    Range("A10").CopyFromRecordset oRS
    
    Sheet1.Cells(I, 6).Value = Sheet4.Cells(10, 2).Value

' ==========================


oRS.Close
oCon.Close

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
Sub Format()
    
    Dim cll As Range
    Dim rng As Range
    
    tdy = Date
    Set rng = Range("E5:E110")
    
    For Each cll In rng
        If cll.Value < tdy Then
            cll.Interior.Color = 255
        End If
        If cll.Value = "" Then
            cll.Interior.Pattern = xlNone
        End If
        If cll.Value = tdy Then
            cll.Interior.Color = 65535
        End If
        If cll.Value > tdy Then
            cll.Interior.Pattern = xlNone
        End If
    Next
    
      Set rng = Range("O5:O110")
    
    For Each cll In rng
        
        If cll.Value = "" Then
            If cll.Offset(0, -1).Value < tdy Then
                cll.Interior.Color = 255
            End If
        End If
        
        If cll.Value = "" Then
            If cll.Offset(0, -1).Value = tdy Then
                cll.Interior.Color = 65535
            End If
        End If
        
        If cll.Value = "" Then
            If cll.Offset(0, -1).Value = "" Then
                cll.Interior.Color = 14281213
            End If
        End If
        
        If cll <> "" Then
            cll.Interior.Color = 14281213
        End If

    Next

End Sub
Sub DeleteARecord()
'Delete a record

Set oCon = New ADODB.Connection
oCon.ConnectionString = PAC1CPEdbConnectionString ' Trusted_Connection=yes;"
oCon.Open
 
Set oRS = New ADODB.Recordset
oRS.ActiveConnection = oCon
oRS.Source = "DELETE FROM cpe_shipmeet WHERE cpe_shipmeet.orderno = 1100388247 AND cpe_shipmeet.datestamp = '4/29/2019'"
oRS.Open

If Not oRS Is Nothing Then Set oRS = Nothing
If Not oCon Is Nothing Then Set oCon = Nothing

End Sub
