VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WW_User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "__Modules"
Option Explicit
'Define constants
Const m_CPE_Tools = "\\uswwq-p-fs01\orders\Common Files\CPE_Schedule\"
Const m_Local_Cache = "\WW_Tools\Tools Cache\"
Const m_Local_Log = "\WW_Tools\Logs\"
Const m_WDrive_Data = "\\uswwq-p-fs01\PAC5-W\Manufacturing\Projects\Data\"
Const m_Temp_Order_Files = "\\USWWQ-P-FS01\orders\Common Files\CPE\Temporary_Order_Files\"
Const m_Open_Parts_Workbook = "Parts.XLS"
Const m_Open_Orders_Workbook = "Open.XLS"
Const m_Shipped_Orders_Workbook = "Shipped.XLS"

'Define properties
Private m_username As String
Private m_full_name As String
Private m_local_user_appdata As String
Private m_local_user_cache As String
Private m_local_user_log As String

Private Sub Class_Initialize()
    
    m_username = Environ$("username")
    m_full_name = ""
    m_local_user_appdata = Environ$("localappdata")
    m_local_user_cache = m_local_user_appdata & m_Local_Cache
    m_local_user_log = m_local_user_appdata & m_Local_Log

End Sub

Sub ww_registry_update(update_cache As Boolean)
    
    If (Dir(m_local_user_appdata & "\WW_Tools", vbDirectory) = vbNullString) Then
        'Make WW Tools Directory & Tools Cache
        MkDir (m_local_user_appdata & "\WW_Tools\")
        MkDir (m_local_user_cache)
        MkDir (m_local_user_log)
    End If
    
    If (Dir(m_local_user_cache, vbDirectory) = vbNullString) Then MkDir (m_local_user_cache)
    If (Dir(m_local_user_log, vbDirectory) = vbNullString) Then MkDir (m_local_user_log)

    SaveSetting AppName:="WW_Tools", Section:="Directories", key:="CPE Tools", Setting:=m_CPE_Tools
    SaveSetting AppName:="WW_Tools", Section:="Directories", key:="Local WW Cache", Setting:=m_local_user_cache
    SaveSetting AppName:="WW_Tools", Section:="Directories", key:="Local WW Log", Setting:=m_local_user_log
    SaveSetting AppName:="WW_Tools", Section:="Directories", key:="WDrive Data", Setting:=m_WDrive_Data
    SaveSetting AppName:="WW_Tools", Section:="Directories", key:="Temp Order Files", Setting:=m_Temp_Order_Files
    
    If update_cache Then update_local_cache
    
End Sub

Sub update_local_cache()

    FileCopy m_WDrive_Data & m_Open_Orders_Workbook, m_local_user_cache & m_Open_Orders_Workbook
    FileCopy m_WDrive_Data & m_Open_Parts_Workbook, m_local_user_cache & m_Open_Parts_Workbook
    FileCopy m_WDrive_Data & m_Shipped_Orders_Workbook, m_local_user_cache & m_Shipped_Orders_Workbook

End Sub

'username
Public Property Get username() As String
    username = m_username
End Property

Public Property Let username(ByVal vNewValue As String)
    m_username = vNewValue
End Property

'full name
Public Property Get full_name() As String
    full_name = m_full_name
End Property

Public Property Let full_name(ByVal vNewValue As String)
    m_full_name = vNewValue
End Property

'local user appdata
Public Property Get local_user_appdata() As String
    local_user_appdata = m_local_user_appdata
End Property

Public Property Let local_user_appdata(ByVal vNewValue As String)
    m_local_user_appdata = vNewValue
End Property

'local user cache
Public Property Get local_user_cache() As String
    local_user_cache = m_local_user_cache
End Property

Public Property Let local_user_cache(ByVal vNewValue As String)
    m_local_user_cache = vNewValue
End Property

'local user log
Public Property Get local_user_log() As String
    local_user_log = m_local_user_log
End Property

Public Property Let local_user_log(ByVal vNewValue As String)
    m_local_user_log = vNewValue
End Property
