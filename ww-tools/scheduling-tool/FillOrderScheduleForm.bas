Attribute VB_Name = "FillOrderScheduleForm"
'@IgnoreModule SetAssignmentWithIncompatibleObjectType
Option Explicit
'@Folder("VBAProject")
'@EntryPoint
Sub Main()
    Dim orderNumber As String
    orderNumber = [Order_Number]
    
    Dim SAP_Session As GuiSession
    Dim SAP_Window As GuiMainWindow

    Set SAP_Session = InitiateSAP
    Set SAP_Window = SAP_Session.FindById("wnd[0]")
    
    On Error GoTo ErrorDisplay
    
    SAP_Window.Maximize
    SAP_Session.StartTransaction ("VA03")

    SAP_Window.FindById("usr/ctxtVBAK-VBELN").Text = orderNumber
    SAP_Window.SendVKey (5)
    
    If SAP_Session.Children.Count > 1 Then
        SAP_Window.SendVKey (0)
    End If

    Dim poNumber As String
    poNumber = SAP_Session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").Text
    
    On Error GoTo NetworkAndFileNumberError
    'MTO Order
    If (orderNumber > 1100109999) Then
        Dim projectOrder As Boolean
        projectOrder = False
        
        'Search for WW ENG LABOR LINE
        SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4402/subSUBSCREEN_TC:SAPMV45A:4900/subSUBSCREEN_BUTTONS:SAPMV45A:4050/btnBT_POPO").press
        SAP_Session.FindById("wnd[1]/usr/txtRV45A-PO_ARKTX").Text = "WHITEWATER ENGINEERING*"

        SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        
        Dim Child
        'Get line number and format
        For Each Child In SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4402/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").Columns
            If (Child.Title = "Item") Then
                Dim lineNumber As String
                lineNumber = Child.Item(0).Text
                
                Dim engineeringItemPath As String
                engineeringItemPath = Child.Item(0).ID
                Exit For
            End If
        Next Child
        lineNumber = Format$(lineNumber, "000000")
        
        'Get and format file number
        SAP_Session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12").Select
        Dim fileNumber As String
        fileNumber = SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZ1AK").Text
        fileNumber = fileNumber & " " & SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZ2AK").Text
        fileNumber = fileNumber & " " & SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtVBAK-ZZ3AK").Text
        [File_Number] = fileNumber
        SAP_Session.FindById("wnd[0]/tbar[0]/btn[3]").press
        'Get network number
        SAP_Session.FindById(engineeringItemPath).SetFocus
        SAP_Session.FindById("wnd[0]").SendVKey (2)
        SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07").Select
        SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4500/btnRV45A-TEXT_BESCHAFFUNG").press
        Dim networkNumber As String
        networkNumber = SAP_Session.FindById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text
        SAP_Session.FindById("wnd[0]").SendVKey (15)
        SAP_Session.FindById("wnd[0]").SendVKey (15)
    Else
        projectOrder = True
        'Get and format file number
        SAP_Session.FindById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
        SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13").Select
        fileNumber = SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZ1AK").Text
        fileNumber = fileNumber & " " & SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/ctxtVBAK-ZZ2AK").Text
        fileNumber = fileNumber & " " & SAP_Session.FindById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\13/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/txtVBAK-ZZ3AK").Text
        [File_Number] = fileNumber
        SAP_Session.FindById("wnd[0]").SendVKey (15)
        SAP_Session.FindById("wnd[0]/tbar[0]/okcd").Text = "CJ20N"
        SAP_Session.FindById("wnd[0]/tbar[0]/btn[0]").press
        SAP_Session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton ("OPEN")
        SAP_Session.FindById("wnd[1]/usr/ctxtCNPB_W_ADD_OBJ_DYN-PROJ_EXT").Text = "VK-" & orderNumber & "-1"
        SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        SAP_Session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").nodeContextMenu ("000001")
        SAP_Session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectContextMenuItem ("EBLM")
        SAP_Session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell").pressButton ("SEARCH_TREE")
        SAP_Session.FindById("wnd[1]/usr/radG_SEARCH_PATTERN-FLG_ELE").Select
        SAP_Session.FindById("wnd[1]/usr/txtG_SEARCH_PATTERN-TEXT").Text = "VK-" & orderNumber & "-1.1.1.3"
        SAP_Session.FindById("wnd[1]/tbar[0]/btn[0]").press
        Dim currentNode As String
        currentNode = SAP_Session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode
        Dim networkNode As String
        networkNode = currentNode + 1
        networkNode = Format$(networkNode, "000000")
        SAP_Session.FindById("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell/shellcont[1]/shell").selectedNode = networkNode
        
        If (SAP_Session.FindById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/txtCAUFVD-KTEXT").Text = "WW MFG NETWORK") Then
            networkNumber = SAP_Session.FindById("wnd[0]/usr/subDETAIL_AREA:SAPLCNPB_M:1010/subIDENTIFICATION:SAPLCOKO:2816/ctxtCAUFVD-AUFNR").Text
        End If
        
        lineNumber = "1.1.1.3.1"
        SAP_Session.FindById("wnd[0]").SendVKey (15)
    End If

    [Customer_PO] = poNumber
    [E_Network] = networkNumber
    [M_Network] = networkNumber
    
    If (projectOrder) Then
        [E_Kronos] = "VK-" & orderNumber & "/" & lineNumber & "/" & networkNumber & "/" & "0020"
        [M_Kronos] = "VK-" & orderNumber & "/" & lineNumber & "/" & networkNumber & "/" & "0030"
    Else
        [E_Kronos] = orderNumber & "/" & lineNumber & "/" & networkNumber & "/" & "0020"
        [M_Kronos] = orderNumber & "/" & lineNumber & "/" & networkNumber & "/" & "0030"
    End If
    
    [E_Step] = "0020"
    [M_Step] = "0030"
    
    MsgBox "Script Complete"
    
    End
    
    
NetworkAndFileNumberError:
    MsgBox "Error retriving File Number/Network Number." & vbNewLine & "Err: " & Err.Description
    Exit Sub
    
ErrorDisplay:
    MsgBox "Err: " & Err.Description
    
End Sub

Public Function InitiateSAP()
    Dim SAP_Connection As GuiConnection
    Dim SAP_Session As GuiSession
    Dim NumOfWindowsSAP As Long
    
    'Reset the SAP_Session Connections
    On Error Resume Next
    If Not SAP_Connection Is Nothing Then Set SAP_Connection = Nothing
    If Not SAP_Session Is Nothing Then Set SAP_Session = Nothing
    On Error GoTo 0
    
    'Check if logged on and initiate the connection
    On Error GoTo SAP_NotLoggedIn
    Set SAP_Connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
    On Error GoTo 0
    
StartLookingForEasyAccess:
    'Find Number of instances (children) open
    NumOfWindowsSAP = SAP_Connection.Sessions.Count
    
    Dim SAPIndex As Long
    'Seach for an initial "SAP Easy Access" Screen
    For SAPIndex = 0 To NumOfWindowsSAP - 1
        If Left$(SAP_Connection.Children(CInt(SAPIndex)).FindById("wnd[0]/titl").Text, 15) = "SAP Easy Access" Then
            Set SAP_Session = SAP_Connection.Children(CInt(SAPIndex))
            Exit For
        End If
    Next SAPIndex
     
    'If a SAP_Session is not established (None of the instances are "SAP Easy Access")
    If SAP_Session Is Nothing Then
        'If a new instance can be open
        If NumOfWindowsSAP < 6 Then
            SAP_Connection.Sessions.Item(0).CreateSession
            While NumOfWindowsSAP = SAP_Connection.Sessions.Count
            Wend
            GoTo StartLookingForEasyAccess
            'At max instances, cause an error
        Else
            MsgBox "Cannot connect to SAP. Open an ""SAP Easy Access"" window to run the program.", Title:="Error!"
            End
        End If
    End If
    
    Set InitiateSAP = SAP_Session
    
    Exit Function
    
SAP_NotLoggedIn:
    MsgBox "Please sign into SAP to continue", Title:="Error!"
    End
    
End Function

