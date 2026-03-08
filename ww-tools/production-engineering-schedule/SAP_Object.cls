VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SAP_Object"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@IgnoreModule ProcedureNotUsed, UnderscoreInPublicClassModuleMember, SetAssignmentWithIncompatibleObjectType, EmptyWhileWendBlock
'@Folder "__Modules"
Private mSAP_GUI As Object
Private mSAP_App As GuiApplication
Private mSAP_Connection As GuiConnection
Private mSAP_Session As GuiSession
Private mSAP_MainWindow As GuiMainWindow

'Initialize connection object
Private Sub Class_Initialize()
    On Error GoTo SAP_NotLoggedIn
    Set mSAP_GUI = GetObject("SAPGUI")
    Set mSAP_App = mSAP_GUI.GetScriptingEngine
    Set mSAP_Connection = mSAP_App.Connections(0)
    On Error GoTo 0
    
StartLookingForEasyAccess:
    'Find Number of instances (children) open
    Dim NumOfWindowsSAP As Long: NumOfWindowsSAP = mSAP_Connection.Sessions.Count
    Dim SAPIndex As Long

    'Seach for an initial "SAP Easy Access" Screen, instantiate session and mainwindow
    For SAPIndex = 0 To NumOfWindowsSAP - 1
        If Left(mSAP_Connection.Children(CInt(SAPIndex)).FindById("wnd[0]").Text, 15) = "SAP Easy Access" Then
            Set mSAP_Session = mSAP_Connection.Children(CInt(SAPIndex))
            Set mSAP_MainWindow = mSAP_Session.FindById("wnd[0]")
            
            Exit For
        End If
    Next SAPIndex
     
    'If a session is not established (None of the instances are "SAP Easy Access")
    If mSAP_Session Is Nothing Then
        'If a new instance can be open
        If NumOfWindowsSAP < 6 Then
            mSAP_Connection.Sessions.Item(0).CreateSession
            
            Do While NumOfWindowsSAP = mSAP_Connection.Sessions.Count: Loop
            
            GoTo StartLookingForEasyAccess
            'At max instances, cause an error
        Else
            MsgBox "Cannot connect to SAP. Open an ""SAP Easy Access"" window to run the program.", Title:="Error!"
            Exit Sub
        End If
    End If
    
    Exit Sub
SAP_NotLoggedIn:
    MsgBox "Please sign into SAP to continue", Title:="Error!"
    End
End Sub


Public Property Get SAP_GUI() As Object

    Set SAP_GUI = mSAP_GUI
    
End Property

Public Property Let SAP_GUI(ByVal vNewValue As Object)

    Set mSAP_GUI = vNewValue
    
End Property

Public Property Get sap_app() As GuiApplication

    Set sap_app = mSAP_App
    
End Property

Public Property Let sap_app(ByVal vNewValue As GuiApplication)

    Set mSAP_App = vNewValue
    
End Property

Public Property Get SAP_Connection() As GuiConnection

    Set SAP_Connection = mSAP_Connection

End Property

Public Property Let SAP_Connection(ByVal vNewValue As GuiConnection)

    Set mSAP_Connection = vNewValue

End Property

Public Property Get sap_session() As GuiSession

    Set sap_session = mSAP_Session

End Property

Public Property Let sap_session(ByVal vNewValue As GuiSession)

    Set mSAP_Session = vNewValue

End Property

Public Property Get SAP_MainWindow() As GuiMainWindow

    Set SAP_MainWindow = mSAP_MainWindow

End Property

Public Property Let SAP_MainWindow(ByVal vNewValue As GuiMainWindow)

    Set mSAP_MainWindow = vNewValue

End Property
