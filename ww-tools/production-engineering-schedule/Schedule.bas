Attribute VB_Name = "Schedule"
Option Explicit
'@IgnoreModule ConstantNotUsed, ProcedureNotUsed, SetAssignmentWithIncompatibleObjectType, MemberNotOnInterface, EmptyWhileWendBlock
'@Folder "Modules"
Public i As Long
Public j As Long
Public K As Long
Public JJ As Long
Public index As Long
Public CodeIndex As Long
Public Ord As Long
Public ActiveYr As Long
Public DescLen As Long
Public lastRow As Long
Public Imax As Long
Public PrevOrd As Long
Public FindValueStart As Long
Public FindValueStop As Long

Public SchdRel As Date
Public ActRel As Date
Public Apps_Out As Date
Public Apps_Back As Date
Public PreRelease As Date
Public EndSpan As Date

Public PC_Eng As String
Public ME_Eng As String
Public PCStatus As String
Public MEStatus As String
Public EngType As String
Public OthEngType As String
Public CustName As String
Public Desc As String
Public OtherEng As String
Public PrevStatus As String
Public AssyStr As String
Public ProdEng As String
Public EType As String
Public OrdStatus As String
Public PONum As String
Public MngrOR As String
Public NewString As String
Public CurStat(50) As String
Public OrdStr As String
Public OrdFold As String
Public TxtLine As String
Public JSubFold As String
Public BuildHTML As String
Public BuildRange As String
Public EmailPerson As String
Public RsmRepPerson As String

Public EstHrs As Double
Public ClrRange As Range
Public Flag1 As Boolean
Public OutApp As Outlook.Application
Public OutMail As Outlook.MailItem
'
Public session As GuiSession
Public SAPconnection As GuiConnection
Public oCon As ADODB.connection
Public oRs As ADODB.Recordset

Public Const Path_SAP As String = "W:\Manufacturing\Projects\Data"

Public Const serverName As String = "USLXA-P-SQL01.cps.local"
Public Const databaseName As String = "PAC1CPE"
Public Const dashboardsDatabase As String = "Dashboards"
Public Const databaseUserID As String = "dash"
Public Const databasePassword As String = "manage_DB"

Public Const databaseConnectionString As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & databaseName & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"

Public Const dashboardsDatabaseConnectionString As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & dashboardsDatabase & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"

'Delay function
Function Delay(ms As Long)
    Delay = Timer + ms / 1000
    Do While Timer < Delay: DoEvents: Loop
End Function

Sub Order_Query()
    Dim StartTime As Variant
    StartTime = Timer
    
    Application.ScreenUpdating = False
    
    ScheduleWS.Select
    ScheduleWS.Range("B4:P500").ClearContents
    ScheduleWS.Range("AC4:AD500").ClearContents
    ScheduleWS.Range("BA4:BA500").ClearContents
    ScheduleWS.CheckBoxes.Delete
    
    ReleasedWS.Range("B4:O10000").ClearContents
    
    QueryResultsWS.Range("B3:BF25000").ClearContents
    
    'Clear Screen
    Application.ScreenUpdating = True
    Delay (10)
    Application.ScreenUpdating = False
    
    ScheduleWS.Range("A2").Value = UCase$(ScheduleWS.Range("A2").Value)
    ProdEng = ScheduleWS.Range("AE2").Value
    EType = ScheduleWS.Range("AJ2").Value
    ActiveYr = ScheduleWS.Range("D1").Value
    Ord = 0
    DescLen = 40
    
    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open
    
    ' Database Query
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    If EType = "PC" Then
        If UCase$(ScheduleWS.Cells(3, 26).Value) = "NO" Then
            oRs.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.PC1 = '" & ProdEng & "' ORDER BY Prod_Eng.PC_Rel_F, Order_Num, Line_Num;"
        Else
            oRs.Source = _
                "SELECT * FROM Prod_Eng " & _
                "WHERE Prod_Eng.PC1 = '" & ProdEng & _
                "' AND PC_Act_Rel > '" & CStr(ActiveYr) & _
                "' AND PC_Act_Rel < '" & CStr(ActiveYr + 1) & _
                "' OR Prod_Eng.PC1 = '" & ProdEng & _
                "' AND PC_Act_Rel < 3" & _
                " ORDER BY Prod_Eng.PC_Rel_F, Order_Num, Line_Num;"
        End If
    Else
        If UCase$(ScheduleWS.Cells(3, 26).Value) = "NO" Then
            oRs.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.ME1 = '" & ProdEng & "' ORDER BY Prod_Eng.ME_Rel_F, Order_Num, Line_Num;"
        Else
            oRs.Source = _
                "SELECT * FROM Prod_Eng " & _
                "WHERE Prod_Eng.ME1 = '" & ProdEng & _
                "' AND ME_Act_Rel > '" & CStr(ActiveYr) & _
                "' AND ME_Act_Rel < '" & CStr(ActiveYr + 1) & _
                "' OR Prod_Eng.ME1 = '" & ProdEng & _
                "' AND ME_Act_Rel < 3" & _
                " ORDER BY Prod_Eng.ME_Rel_F, Order_Num, Line_Num;"
        End If
    End If
    
    ' Retrieve records
    oRs.Open
    QueryResultsWS.Range("B3").CopyFromRecordset oRs
    oRs.Close
    
    j = 3: JJ = 3
    ' Released Orders
    For i = 3 To 25000
        PrevOrd = Ord
        Ord = QueryResultsWS.Cells(i, 3).Value
        If Ord < 3 Then Exit For
        
        If Ord = PrevOrd Then
            Flag1 = True
            GoTo DETAILS
        Else
            Flag1 = False
        End If
        
        Dim Eng_Apps_Out As Date
        
        If EType = "PC" Then
            OrdStatus = QueryResultsWS.Cells(i, 35).Value
            SchdRel = QueryResultsWS.Cells(i, 50).Value
            ActRel = QueryResultsWS.Cells(i, 52).Value
            OtherEng = QueryResultsWS.Cells(i, 18).Value
            MngrOR = QueryResultsWS.Cells(i, 37).Value
            EstHrs = QueryResultsWS.Cells(i, 25).Value + QueryResultsWS.Cells(i, 26).Value
            Apps_Back = QueryResultsWS.Cells(i, 44).Value
            PreRelease = QueryResultsWS.Cells(i, 46).Value
            Eng_Apps_Out = QueryResultsWS.Cells(i, 42).Value
        Else
            OrdStatus = QueryResultsWS.Cells(i, 36).Value
            SchdRel = QueryResultsWS.Cells(i, 51).Value
            ActRel = QueryResultsWS.Cells(i, 53).Value
            OtherEng = QueryResultsWS.Cells(i, 16).Value
            MngrOR = QueryResultsWS.Cells(i, 38).Value
            EstHrs = QueryResultsWS.Cells(i, 27).Value + QueryResultsWS.Cells(i, 28).Value
            Apps_Back = QueryResultsWS.Cells(i, 45).Value
            PreRelease = QueryResultsWS.Cells(i, 47).Value
            Eng_Apps_Out = QueryResultsWS.Cells(i, 43).Value
        End If
        
        Dim productionScheduler As String: productionScheduler = QueryResultsWS.Cells(i, 22).Value
        
        
        If QueryResultsWS.Cells(i, 41).Value > 3 Then
            Apps_Out = QueryResultsWS.Cells(i, 41).Value
        Else
            Apps_Out = 2
        End If
    
        If OrdStatus = "RELEASED" Then
            j = j + 1
            ReleasedWS.Cells(j, 2).Value = Ord
            ReleasedWS.Cells(j, 3).Value = QueryResultsWS.Cells(i, 4).Value
            ReleasedWS.Cells(j, 6).Value = OtherEng
            ReleasedWS.Cells(j, 7).Value = SchdRel
            ReleasedWS.Cells(j, 8).Value = ActRel
            ReleasedWS.Cells(j, 9).Value = EstHrs
            If QueryResultsWS.Cells(i, 54).Value > 3 Then
                ReleasedWS.Cells(j, 11).Value = QueryResultsWS.Cells(i, 54).Value
            End If
            ReleasedWS.Cells(j, 14).Value = MngrOR
            ReleasedWS.Cells(j, 13).Value = "=IF(H" & j & ">G" & j & ",""LATE"",""OK"")"
            ReleasedWS.Cells(j, 26).Value = "=IF(M" & j & "=""LATE"",IF(N" & j & "="""",1,0),0)"
        Else                                     'Current Orders
            Application.ScreenUpdating = True
            Delay (10)
            Application.ScreenUpdating = False
            JJ = JJ + 1
            ScheduleWS.Cells(JJ, 2).Value = Ord
            ScheduleWS.Cells(JJ, 3).Value = QueryResultsWS.Cells(i, 4).Value
            ScheduleWS.Cells(JJ, 5).Value = QueryResultsWS.Cells(i, 13).Value
            ScheduleWS.Cells(JJ, 6).Value = OtherEng
            
            ScheduleWS.Cells(JJ, 7).Value = IIf(PreRelease > 3, PreRelease, "-")
            ScheduleWS.Cells(JJ, "BA").Value = IIf(Eng_Apps_Out > 3, Eng_Apps_Out, vbNullString)
            
            ScheduleWS.Cells(JJ, "BB").Value = UCase$(productionScheduler)
            
            If Apps_Out > 3 Then
                Dim CBX As CheckBox
                ScheduleWS.Cells(JJ, 8).Value = Apps_Out
                Set CBX = ScheduleWS.CheckBoxes.Add(Cells(JJ, 16).Left, Cells(JJ, 16).Top, 50, 17.25)
                CBX.Caption = vbNullString
                CBX.LinkedCell = "P" & JJ
                CBX.Display3DShading = False
                CBX.Name = JJ
                CBX.OnAction = "UpdateAppsBackDate"
                If Apps_Back > 3 Then
                    CBX.Value = xlOn
                Else
                    CBX.Value = xlOff
                End If
            End If
            
            ScheduleWS.Cells(JJ, 9).Value = SchdRel
            ScheduleWS.Cells(JJ, 10).Value = EstHrs
            ScheduleWS.Cells(JJ, 13).Value = OrdStatus
            ScheduleWS.Cells(JJ, 30).Value = OrdStatus
            ScheduleWS.Cells(JJ, 15).Value = GetOrdComments(Ord)
            
            'Create Hyperlink
            OrdStr = Str$(Ord)
            OrdFold = "J:\Orders\" & Mid$(OrdStr, 2, 7) & "000\" & Mid$(OrdStr, 2, 10) & "*"
            TxtLine = Dir(OrdFold, vbDirectory)
            ScheduleWS.Cells(JJ, 29).Value = TxtLine
            JSubFold = Mid$(TxtLine, 1, 7) & "000\"
            ScheduleWS.Range(Cells(JJ, 2), Cells(JJ, 2)).Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                                       "J:\Orders\" & JSubFold & TxtLine
    
        End If
    
DETAILS:
        If Flag1 = True And EType = "PC" And OrdStatus <> "RELEASED" Then
            EstHrs = QueryResultsWS.Cells(i, 25).Value + QueryResultsWS.Cells(i, 26).Value
            ScheduleWS.Cells(JJ, 10).Value = ScheduleWS.Cells(JJ, 10).Value + EstHrs
        End If
        
        If Flag1 = True And EType = "ME" And OrdStatus <> "RELEASED" Then
            EstHrs = QueryResultsWS.Cells(i, 27).Value + QueryResultsWS.Cells(i, 28).Value
            ScheduleWS.Cells(JJ, 10).Value = ScheduleWS.Cells(JJ, 10).Value + EstHrs
        End If
        
        If Flag1 = True And EType = "PC" And OrdStatus = "RELEASED" Then
            EstHrs = QueryResultsWS.Cells(i, 25).Value + QueryResultsWS.Cells(i, 26).Value
            ReleasedWS.Cells(j, 9).Value = ReleasedWS.Cells(j, 9).Value + EstHrs
        End If
        
        If Flag1 = True And EType = "ME" And OrdStatus = "RELEASED" Then
            EstHrs = QueryResultsWS.Cells(i, 27).Value + QueryResultsWS.Cells(i, 28).Value
            ReleasedWS.Cells(j, 9).Value = ReleasedWS.Cells(j, 9).Value + EstHrs
        End If
        
        If Flag1 = False And OrdStatus = "RELEASED" Then
            ReleasedWS.Cells(j, 5).Value = ReleasedWS.Cells(j, 5).Value & "[" & _
                                        QueryResultsWS.Cells(i, 5).Value & "] " & Mid$(QueryResultsWS.Cells(i, 7).Value, 1, DescLen)
        End If
        
        If Flag1 = True And OrdStatus = "RELEASED" Then
            ReleasedWS.Cells(j, 5).Value = ReleasedWS.Cells(j, 5).Value & Chr$(10) & "[" & _
                                        QueryResultsWS.Cells(i, 5).Value & "] " & Mid$(QueryResultsWS.Cells(i, 7).Value, 1, DescLen)
        End If
        
        If Flag1 = False And OrdStatus <> "RELEASED" Then
            ScheduleWS.Cells(JJ, 12).Value = ScheduleWS.Cells(JJ, 12).Value & "[" & _
                                            QueryResultsWS.Cells(i, 5).Value & "] " & Mid$(QueryResultsWS.Cells(i, 7).Value, 1, DescLen)
        End If
    
        If Flag1 = True And OrdStatus <> "RELEASED" Then
            ScheduleWS.Cells(JJ, 12).Value = ScheduleWS.Cells(JJ, 12).Value & Chr$(10) & "[" & _
                                            QueryResultsWS.Cells(i, 5).Value & "] " & Mid$(QueryResultsWS.Cells(i, 7).Value, 1, DescLen)
        End If
    
        'Build Kronos Activity
        If InStr(QueryResultsWS.Cells(i, 6).Value, "V065022.A01") > 0 Or InStr(QueryResultsWS.Cells(i, 7).Value, "WW ENG") > 0 Then
            Dim kronosNetwork As Variant
            kronosNetwork = IIf(QueryResultsWS.Cells(i, 3).Value < 1100109999, _
                                "VK-" & Trim$(QueryResultsWS.Cells(i, 3).Value) & "/1.1.1.3.1/" & Trim$(QueryResultsWS.Cells(i, 8).Value), _
                                Trim$(QueryResultsWS.Cells(i, 3).Value) & "/" & Format$(Trim$(QueryResultsWS.Cells(i, 5).Value), "000000") & "/" & Trim$(QueryResultsWS.Cells(i, 8).Value))
    
            kronosNetwork = IIf(EType = "PC", kronosNetwork & "/0020", kronosNetwork & "/0030")
            
            If QueryResultsWS.Cells(i, 8).Value = vbNullString Then
                kronosNetwork = vbNullString
            End If
            
            If OrdStatus <> "RELEASED" Then
                ScheduleWS.Cells(JJ, 4).Value = kronosNetwork
            Else
                ReleasedWS.Cells(j, 4).Value = kronosNetwork
            End If
        End If
        
    Next i
    Application.ScreenUpdating = True
    Delay (10)
    Application.ScreenUpdating = False

    Call Release_Count

    ReleasedWS.Activate
    ReleasedWS.Rows.EntireRow.AutoFit
    ReleasedWS.Range("C2").Select
    ActiveWindow.ScrollRow = 4
    
    ScheduleWS.Activate
    ScheduleWS.Rows.EntireRow.AutoFit
    ScheduleWS.Range("A2").Select
    ActiveWindow.ScrollRow = 4
    
    ScheduleWS.Range("AA3").Value = Round((Timer - StartTime), 1) & " Seconds"
End Sub

Sub Order_Details()
    '
    QueryResults2WS.Range("B3:BF5000").ClearContents
    OrderDetailsWS.Range("B16:Y1000").ClearContents

    Ord = OrderDetailsWS.Cells(3, 2).Value

    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open

    ' Database Query
    Set oRs = New ADODB.Recordset
    oRs.ActiveConnection = oCon
    oRs.Source = "SELECT * FROM Prod_Eng WHERE Prod_Eng.Order_Num = " & Ord & ";"
    ' Retrieve records
    oRs.Open
    QueryResults2WS.Range("B3").CopyFromRecordset oRs

    OrderDetailsWS.Range("C3").Value = QueryResults2WS.Cells(3, 4).Value ' Customer
    OrderDetailsWS.Range("E3").Value = QueryResults2WS.Cells(1, 62).Value ' Process Control
    OrderDetailsWS.Range("F3").Value = QueryResults2WS.Cells(2, 62).Value ' Mechanical Designer
    OrderDetailsWS.Range("G3").Value = QueryResults2WS.Cells(47, 62).Value ' Project Manager
    OrderDetailsWS.Range("H3").Value = QueryResults2WS.Cells(48, 62).Value ' Project Engineer
    OrderDetailsWS.Range("I3").Value = QueryResults2WS.Cells(3, 13).Value ' Document Types
    OrderDetailsWS.Range("J3").Value = UCase$(QueryResults2WS.Cells(3, "V").Value) ' Scheduler
    
    OrderDetailsWS.Range("E5").Value = QueryResults2WS.Cells(3, 34).Value ' Project Level
    OrderDetailsWS.Range("F5").Value = QueryResults2WS.Cells(3, 33).Value ' MultiPlant
    OrderDetailsWS.Range("G5").Value = copy_if_date(QueryResults2WS.Cells(3, 46)) 'Scheduled Pre-Order
    OrderDetailsWS.Range("H5").Value = copy_if_date(QueryResults2WS.Cells(3, 49)) 'ME Actual Pre-Order
    OrderDetailsWS.Range("I5").Value = copy_if_date(QueryResults2WS.Cells(3, 48)) 'PC Actual Pre-Order
    
    OrderDetailsWS.Range("E7").Value = copy_if_date(QueryResults2WS.Cells(3, 41)) 'Scheduled Approvals
    OrderDetailsWS.Range("F7").Value = copy_if_date(QueryResults2WS.Cells(3, 43)) 'Mechanical Approvals Out
    OrderDetailsWS.Range("G7").Value = copy_if_date(QueryResults2WS.Cells(3, 45)) 'Mechanical Approvals Back
    OrderDetailsWS.Range("H7").Value = copy_if_date(QueryResults2WS.Cells(3, 42)) 'Control Approvals Out
    OrderDetailsWS.Range("I7").Value = copy_if_date(QueryResults2WS.Cells(3, 44)) 'Control Approvals Back
    
    OrderDetailsWS.Range("E9").Value = copy_if_date(QueryResults2WS.Cells(3, 57)) 'Production Release
    OrderDetailsWS.Range("F9").Value = copy_if_date(QueryResults2WS.Cells(3, 51)) 'Final ME Production Release
    OrderDetailsWS.Range("G9").Value = copy_if_date(QueryResults2WS.Cells(3, 53)) 'Actual ME Production Release
    OrderDetailsWS.Range("H9").Value = copy_if_date(QueryResults2WS.Cells(3, 50)) 'Final PC Production Release
    OrderDetailsWS.Range("I9").Value = copy_if_date(QueryResults2WS.Cells(3, 52)) 'Actual PC Production Release
    

    'HyperLink Checksum Value
    'ex. 110517206, 1+1+5+1+7+2+0+6=23
    Dim TextCode As Variant
    TextCode = 2
    For CodeIndex = 5 To 10
        TextCode = TextCode + Mid$(Ord, CodeIndex, 1)
    Next CodeIndex
    
    'ME Status
    OrderDetailsWS.Range("E11").Value = QueryResults2WS.Cells(3, 36).Value
    'ME Link
    If OrderDetailsWS.Range("F3").Value <> "N/A" Then
        Dim URL_Text As Variant
        URL_Text = "http://pacdwg.schenckprocess.com/fdl.asp?sn=2&ano=" & Right$(Ord, 6) & "M" & TextCode
        OrderDetailsWS.Range("I11").Hyperlinks.Add Range("I11"), Address:=URL_Text, TextToDisplay:=URL_Text
    Else
        OrderDetailsWS.Range("I11").ClearContents
    End If

    'PE Status
    OrderDetailsWS.Range("E13").Value = QueryResults2WS.Cells(3, 35).Value
    'PE Link
    If OrderDetailsWS.Range("E3").Value <> "N/A" Then
        URL_Text = "http://pacdwg.schenckprocess.com/fdl.asp?sn=2&ano=" & Right$(Ord, 6) & "E" & TextCode
        OrderDetailsWS.Range("I13").Hyperlinks.Add Range("I13"), Address:=URL_Text, TextToDisplay:=URL_Text
    Else
        OrderDetailsWS.Range("I13").ClearContents
    End If

    AssyStr = vbNullString
    
    Dim row_iter As Long
    For row_iter = 3 To 60
        Ord = QueryResults2WS.Cells(row_iter, 3).Value2
        If Ord < 3 Then Exit For
        AssyStr = AssyStr & "[" & QueryResults2WS.Cells(row_iter, 5).Value & "] " & Mid$(QueryResults2WS.Cells(row_iter, 7).Value, 1, 40) & Chr$(10)
    Next row_iter
    
    OrderDetailsWS.Cells(5, 2).Value = AssyStr
    
    get_all_order_information OrderDetailsWS.Range("B3").Value
    
    For row_iter = 16 To 116
    
        OrderDetailsWS.Range("B" & row_iter).Value = FullDatabaseDumpWS.Range("BR" & row_iter - 13).Value
        OrderDetailsWS.Range("F" & row_iter).Value = FullDatabaseDumpWS.Range("CY" & row_iter - 13).Value
        OrderDetailsWS.Range("K" & row_iter).Value = FullDatabaseDumpWS.Range("E" & row_iter - 13).Value
        OrderDetailsWS.Range("P" & row_iter).Value = FullDatabaseDumpWS.Range("CA" & row_iter - 13).Value
        OrderDetailsWS.Range("U" & row_iter).Value = FullDatabaseDumpWS.Range("DJ" & row_iter - 13).Value
    
    Next row_iter

    
End Sub

Function copy_if_date(input_range As Range) As String

    copy_if_date = IIf(input_range.Value > 3, input_range.Value, vbNullString)

End Function

Sub Update_Sched()
    ' Update for Status and/or Notes Subroutine
    '
    Dim NewSts As String
    Dim Nots As String
    Dim EngInit As String
    Dim AssyStr As String
    Dim FiltNotes As String
    Dim Tday As Date

    EngInit = ScheduleWS.Range("A2").Value
    Tday = ScheduleWS.Range("Z1").Value

    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open

    For i = 4 To 50
    
        Ord = ScheduleWS.Cells(i, 2).Value
        OrdStatus = ScheduleWS.Cells(i, 13).Value
        NewSts = ScheduleWS.Cells(i, 13).Value
        PrevStatus = ScheduleWS.Cells(i, 30).Value
        Nots = ScheduleWS.Cells(i, 14).Value
        FiltNotes = vbNullString
    
        For K = 1 To Len(Nots)
            If (Mid$(Nots, K, 1) = Chr$(34) Or Mid$(Nots, K, 1) = Chr$(39) Or Mid$(Nots, K, 1) = Chr$(10)) Then
                FiltNotes = FiltNotes & " "
            Else
                FiltNotes = FiltNotes & Mid$(Nots, K, 1)
            End If
        Next K
        
        If Ord = 0 Then Exit For
        If (ScheduleWS.Cells(i, 14).Value > vbNullString Or OrdStatus <> CurStat(i)) Then
            AssyStr = vbNullString
            AssyStr = "'" & EngInit & "', "      ' Engineer
            AssyStr = AssyStr & Ord & ", '"      ' Order Number
            AssyStr = AssyStr & PrevStatus & "', '" ' Previous Status
            AssyStr = AssyStr & NewSts & "', '"  ' New Status
            AssyStr = AssyStr & Mid$(FiltNotes, 1, 254) & "', '" ' Comments
            AssyStr = AssyStr & Date$ & "'"      ' Date Stamp
            Set oRs = New ADODB.Recordset
            oRs.ActiveConnection = oCon
            oRs.Source = "INSERT INTO cpe_schedule(uniqid, engineer, orderno, prev_status, new_status, comments, datestamp) VALUES(newid(), " & _
                         AssyStr & ")"
            oRs.Open
            
        End If

    Next i

    ' Check / Change Status
    For i = 4 To 50
        Ord = ScheduleWS.Cells(i, 2).Value
        OrdStatus = ScheduleWS.Cells(i, 13).Value
        ProdEng = ScheduleWS.Range("AE2").Value
        EType = ScheduleWS.Range("AJ2").Value
    
        If Ord = 0 Then Exit For
        CurStat(i) = ScheduleWS.Cells(i, 30).Value
    
        If (OrdStatus <> CurStat(i)) Then
        
            Set oRs = New ADODB.Recordset
            oRs.ActiveConnection = oCon

            If OrdStatus = "RELEASED" And EType = "PC" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Status='" & OrdStatus & "', PC_Act_Rel='" & Tday & "' WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                oRs.Open
            End If
            If OrdStatus <> "RELEASED" And EType = "PC" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Status='" & OrdStatus & "'WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                oRs.Open
            End If
        
            If OrdStatus = "RELEASED" And EType = "ME" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Status='" & OrdStatus & "', ME_Act_Rel='" & Tday & "' WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                oRs.Open
            End If
            
            If OrdStatus <> "RELEASED" And EType = "ME" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Status='" & OrdStatus & "'WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                oRs.Open
            End If
            
            If OrdStatus = "WAITING FOR CUSTOMER APPROVAL" And CurStat(i) <> "WAITING FOR CUSTOMER APPROVAL" Then
                ScheduleWS.Cells(i, 24).Value = 1
                PrevOrd = Ord
            ElseIf OrdStatus = "RELEASED" Then
                ScheduleWS.Cells(i, 24).Value = 2
                PrevOrd = Ord
            Else
                ScheduleWS.Cells(i, 24).Value = 0
            End If
        End If
    Next i

    For i = 4 To 50
        Ord = ScheduleWS.Cells(i, 2).Value
        CustName = ScheduleWS.Cells(i, 3).Value
        If ScheduleWS.Cells(i, 24).Value = 1 Then
            ScheduleWS.Cells(i, 24).Value = 0
            'Database
            If EType = "PC" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.PC_Apps_Out='" & Tday & "'WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.PC1 = '" & ProdEng & "'"
                oRs.Open
            ElseIf EType = "ME" Then
                oRs.Source = "UPDATE Prod_Eng SET Prod_Eng.ME_Apps_Out='" & Tday & "'WHERE Prod_Eng.Order_Num=" & Ord & " AND Prod_Eng.ME1 = '" & ProdEng & "'"
                oRs.Open
            End If
            SendApprovalEmail
        ElseIf ScheduleWS.Cells(i, 24).Value = 2 Then
            ScheduleWS.Cells(i, 24).Value = 0
            
            SendCertifiedEmail Ord
        
            SendProductionEmail (ScheduleWS.Cells(i, "BB"))

        End If
    Next i

    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing

    Order_Query
    
End Sub

Sub LaterNotes()

    Dim EngNote As String
    Dim FiltNotes As String
    Dim EngInit As String
    Dim AssyStr As String
    
    EngInit = ScheduleWS.Range("A2").Value
    Ord = LaterNotesWS.Cells(4, 2).Value
    EngNote = LaterNotesWS.Cells(4, 3).Value

    For K = 1 To Len(EngNote)
        If (Mid$(EngNote, K, 1) = Chr$(34) Or Mid$(EngNote, K, 1) = Chr$(39) Or Mid$(EngNote, K, 1) = Chr$(10)) Then
            FiltNotes = FiltNotes & " "
        Else
            FiltNotes = FiltNotes & Mid$(EngNote, K, 1)
        End If
    Next K

    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open

    If Ord <> 0 Then
        AssyStr = vbNullString
        AssyStr = "'" & EngInit & "', "          ' Engineer
        AssyStr = AssyStr & Ord & ", '"          ' Order Number
        AssyStr = AssyStr & "RELEASED', '"       ' Previous Status
        AssyStr = AssyStr & "RELEASED', '"       ' New Status
        AssyStr = AssyStr & Mid$(FiltNotes, 1, 254) & "', '" ' Comments
        AssyStr = AssyStr & Date$ & "'"          ' Date Stamp
    
        Set oRs = New ADODB.Recordset
        oRs.ActiveConnection = oCon
        oRs.Source = "INSERT INTO cpe_schedule(uniqid, engineer, orderno, prev_status, new_status, comments, datestamp) VALUES(newid(), " & _
                     AssyStr & ")"
        oRs.Open
    End If

    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing

End Sub

Sub Notes_By_Order()
    '
    Dim EngInit As String

    EngInit = LaterNotesWS.Cells(20, 1).Value

    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString: oCon.Open

    Ord = LaterNotesWS.Cells(20, 2).Value

    If Ord > 1 Then
    
        Set oRs = New ADODB.Recordset
        oRs.ActiveConnection = oCon
    
        If UCase$(LaterNotesWS.Cells(17, 2).Value) = "YES" Then
            oRs.Source = "SELECT cpe_schedule.comments, engineer, datestamp, prev_status, new_status FROM cpe_schedule WHERE (cpe_schedule.engineer='" & EngInit & "' AND orderno=" & Ord & " AND cpe_schedule.comments<>'') ORDER BY  cpe_schedule.datestamp"
        Else
            oRs.Source = "SELECT cpe_schedule.comments, engineer, datestamp, prev_status, new_status FROM cpe_schedule WHERE (cpe_schedule.orderno=" & Ord & " AND cpe_schedule.comments<>'') ORDER BY  cpe_schedule.engineer, datestamp"
        End If
        
        oRs.Open
        
        Set ClrRange = LaterNotesWS.Range("C20:GF500000")
        ClrRange.ClearContents
        ClrRange.CopyFromRecordset oRs
        
        'Scheduler Notes
        LaterNotesWS.Range("O20").CopyFromRecordset oRs
        
        oRs.Close
        oCon.Close
    End If

    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing

End Sub

Sub Release_Count()
    Dim NumOfRels As Long
    Dim PCMnth As Long
    Dim MEMnth As Long
    Dim PCYear As Long
    Dim MEYear As Long

    ' By Engineer
    Set ClrRange = QueryResultsWS.Range("CJ3:DI50000")
    ClrRange.ClearContents
    Ord = 0
    Flag1 = False

    For i = 3 To 25000
        PrevOrd = Ord
        Ord = QueryResultsWS.Cells(i, 3).Value
        Dim Line_Num As Variant
        Line_Num = QueryResultsWS.Cells(i, 5).Value

        If PrevOrd <> Ord And Flag1 = True Then
            If NumOfRels = 0 Then NumOfRels = 1
            If PC_Eng <> vbNullString And PC_Eng <> "N/A" And PCStatus = "RELEASED" Then
                PCYear = Year(QueryResultsWS.Cells(Imax, 52).Value)
                PCMnth = Month(QueryResultsWS.Cells(Imax, 52).Value)
                If PCYear = ScheduleWS.Cells(1, 4).Value Then
                    QueryResultsWS.Cells(Imax, 88).Value = NumOfRels
                    QueryResultsWS.Cells(Imax, 89 + PCMnth).Value = NumOfRels
                End If
            End If
            If ME_Eng <> vbNullString And ME_Eng <> "N/A" And MEStatus = "RELEASED" Then
                MEYear = Year(QueryResultsWS.Cells(Imax, 53).Value)
                MEMnth = Month(QueryResultsWS.Cells(Imax, 53).Value)
                If MEYear = ScheduleWS.Cells(1, 4).Value Then
                    QueryResultsWS.Cells(Imax, 89).Value = NumOfRels
                    QueryResultsWS.Cells(Imax, 101 + MEMnth).Value = NumOfRels
                End If
            End If
        
            NumOfRels = 0
            Flag1 = False
        End If
        
        If Line_Num > 0 And PrevOrd <> Ord And Flag1 = False Then
            Desc = QueryResultsWS.Cells(i, 7).Value
            PC_Eng = QueryResultsWS.Cells(i, 16).Value
            ME_Eng = QueryResultsWS.Cells(i, 18).Value
            PCStatus = QueryResultsWS.Cells(i, 35).Value
            MEStatus = QueryResultsWS.Cells(i, 36).Value
            If Mid$(Desc, 1, 4) = "MECH" Then NumOfRels = NumOfRels + 1
            If Mid$(Desc, 1, 5) = "MULTI" Then NumOfRels = NumOfRels + 1
            If Mid$(Desc, 1, 3) = "DMO" Then NumOfRels = NumOfRels + 1
            If Mid$(Desc, 1, 3) = "SYS" Then NumOfRels = NumOfRels + 1
            K = InStr(1, Desc, ",")
            If K > 7 Then NewString = Mid$(Desc, K - 6, 6) Else NewString = vbNullString
            If NewString = "FEEDER" Then NumOfRels = NumOfRels + 1
            Flag1 = True
            Imax = i
        End If
    
        If Line_Num > 0 And PrevOrd = Ord And Flag1 = True Then
            Desc = QueryResultsWS.Cells(i, 7).Value
            PC_Eng = QueryResultsWS.Cells(i, 16).Value
            ME_Eng = QueryResultsWS.Cells(i, 18).Value
            PCStatus = QueryResultsWS.Cells(i, 35).Value
            MEStatus = QueryResultsWS.Cells(i, 36).Value
            If Mid$(Desc, 1, 4) = "MECH" Then NumOfRels = NumOfRels + 1
            If Mid$(Desc, 1, 5) = "MULTI" Then NumOfRels = NumOfRels + 1
            If Mid$(Desc, 1, 3) = "DMO" Then NumOfRels = NumOfRels + 1
            If Mid$(Desc, 1, 3) = "SYS" Then NumOfRels = NumOfRels + 1
            K = InStr(1, Desc, ",")
            If K > 7 Then NewString = Mid$(Desc, K - 6, 6) Else NewString = vbNullString
            If NewString = "FEEDER" Then NumOfRels = NumOfRels + 1
        End If

        If Ord = 0 Then Exit For

    Next i

End Sub

Sub InitiateSAP()

    'Reset the session Connections
    On Error Resume Next
    If Not SAPconnection Is Nothing Then Set SAPconnection = Nothing
    If Not session Is Nothing Then Set session = Nothing
    On Error GoTo 0
    
    'Check if logged on and initiate the connection
    On Error GoTo SAP_NotLoggedIn
    Set SAPconnection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
    On Error GoTo 0
    
StartLookingForEasyAccess:
    'Find Number of instances (children) open
    Dim NumOfWindowsSAP As Variant
    NumOfWindowsSAP = SAPconnection.Sessions.Count
    'Seach for an initial "SAP Easy Access" Screen
    Dim SAPIndex As Variant
    For SAPIndex = 0 To NumOfWindowsSAP - 1
        If Left$(SAPconnection.Children(CInt(SAPIndex)).FindById("wnd[0]").Text, 15) = "SAP Easy Access" Then
            Set session = SAPconnection.Children(CInt(SAPIndex))
            Exit For
        End If
    Next SAPIndex
     
    'If a session is not established (None of the instances are "SAP Easy Access")
    If session Is Nothing Then
        'If a new instance can be open
        If NumOfWindowsSAP < 6 Then
            SAPconnection.Sessions.Item(0).CreateSession
            
            Do While NumOfWindowsSAP = SAPconnection.Sessions.Count: Loop
            
            GoTo StartLookingForEasyAccess
            'At max instances, cause an error
        Else
            Exit Sub
        End If
    End If

    Exit Sub
    
SAP_NotLoggedIn:
End Sub

Sub UpdateAppsBackDate()

    UserForm1.show
    
End Sub

Function GetOrdComments(OrderCom)
    Set ClrRange = QueryResultsWS.Range("BX3:CD5000")
    ClrRange.ClearContents
    GetOrdComments = vbNullString


    On Error GoTo oConoRS_NotSet
    oRs.Source = "SELECT * FROM cpe_schedule WHERE cpe_schedule.orderno = " & OrderCom & " ORDER BY cpe_schedule.datestamp"
    On Error GoTo 0
    oRs.Open
    QueryResultsWS.Range("BX3").CopyFromRecordset oRs
    lastRow = QueryResultsWS.Cells(Rows.Count, 78).End(xlUp).Row
    oRs.Close
    For K = 3 To lastRow
        If ((QueryResultsWS.Cells(K, 77).Value = ProdEng) And (QueryResultsWS.Cells(K, 81).Value <> vbNullString)) Then
            If GetOrdComments <> vbNullString Then
                GetOrdComments = GetOrdComments & Chr$(10)
            End If
            GetOrdComments = GetOrdComments & QueryResultsWS.Cells(K, 86) & QueryResultsWS.Cells(K, 81)
        End If
    Next K
    Exit Function
    
oConoRS_NotSet:
    If Not oRs Is Nothing Then Set oRs = Nothing
    If Not oCon Is Nothing Then Set oCon = Nothing
    Set oCon = New ADODB.connection
    oCon.ConnectionString = databaseConnectionString ' Trusted_Connection=yes;"
    oCon.Open
    Dim oRS_Comm As Variant
    Set oRS_Comm = New ADODB.Recordset
    oRS_Comm.ActiveConnection = oCon
    Resume
    
End Function


Sub get_and_paste_sql_query(ByRef sql_query As String, ByRef ws_range As Range, Optional ByVal connection_string As String = databaseConnectionString)

    Dim queryConnection As ADODB.connection: Set queryConnection = New ADODB.connection
    queryConnection.ConnectionString = connection_string
    queryConnection.Open
    
    Dim queryRecordset As ADODB.Recordset
    
    Set queryRecordset = queryConnection.Execute(sql_query)
    
    ws_range.CopyFromRecordset queryRecordset
    
    If Not queryRecordset Is Nothing Then Set queryRecordset = Nothing
    If Not queryConnection Is Nothing Then Set queryConnection = Nothing
    

End Sub


'Get documents for order details sheet
Public Sub get_documents_order_details()
    
    Dim order_number As String: order_number = OrderDetailsWS.Range("B3")
    Order_Details
    get_documents order_number

End Sub

Public Sub get_ae_users()

    
    'get_sql_recordset PAC1CPE_CONNECTION_STRING, "SELECT WW_Email FROM WW_Prod_Eng_Names WHERE WW_Active = 'Y' AND WW_Initials <> 'N/A' AND WW_Department = 'AE'", QueryResults2WS.Range("BS4")
    
    get_ww_users "AE", QueryResults2WS.Range("BU4")

End Sub
