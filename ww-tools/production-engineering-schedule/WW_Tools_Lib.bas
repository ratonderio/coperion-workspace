Attribute VB_Name = "WW_Tools_Lib"
'@Folder "__Modules"
Option Explicit
Option Base 1

'Enum for SAP V Keys
Public Enum V_Keys
    V_Enter = 0
    V_F1 = 1
    V_F2 = 2
    V_Back = 3
    V_F4 = 4
    V_F5 = 5
    V_F6 = 6
    V_F7 = 7
    V_F8 = 8
    V_F9 = 9
    V_F10 = 10
    V_Save = 11
    V_Cancel = 12
    V_Shift_And_F1 = 13
    V_Shift_And_F2 = 14
    V_Exit = 15
    V_Shift_And_F4 = 16
    V_Shift_And_F5 = 17
    V_Shift_And_F6 = 18
    V_Shift_And_F7 = 19
    V_Shift_And_F8 = 20
    V_Shift_And_F9 = 21
    V_Shift_And_Ctrl_And_0 = 22
    V_Shift_And_F11 = 23
    V_Shift_And_F12 = 24
    V_Ctrl_And_F1 = 25
    V_Ctrl_And_F2 = 26
    V_Ctrl_And_F3 = 27 'Select All
    V_Ctrl_And_F4 = 28 'Deselect All
    V_Ctrl_And_F5 = 29
    V_Ctrl_And_F6 = 30
    V_Ctrl_And_F7 = 31
    V_Ctrl_And_F8 = 32
    V_Ctrl_And_F9 = 33
    V_Ctrl_And_F10 = 34
    V_Ctrl_And_F11 = 35
    V_Ctrl_And_F12 = 36
    V_Ctrl_And_Shift_And_F1 = 37
    V_Ctrl_And_Shift_And_F2 = 38
    V_Ctrl_And_Shift_And_F3 = 39
    V_Ctrl_And_Shift_And_F4 = 40
    V_Ctrl_And_Shift_And_F5 = 41
    V_Ctrl_And_Shift_And_F6 = 42
    V_Ctrl_And_Shift_And_F7 = 43
    V_Ctrl_And_Shift_And_F8 = 44
    V_Ctrl_And_Shift_And_F9 = 45
    V_Ctrl_And_Shift_And_F10 = 46
    V_Ctrl_And_Shift_And_F11 = 47
    V_Ctrl_And_Shift_And_F12 = 48
    V_Ctrl_And_E = 70
    V_Ctrl_And_F = 71
    V_Ctrl_And_Forward_Slash = 72
    V_Ctrl_And_Backward_Slash = 73
    V_Ctrl_And_N = 74
    V_Ctrl_And_O = 75
    V_Ctrl_And_X = 76
    V_Ctrl_And_C = 77
    V_Ctrl_And_V = 78
    V_Ctrl_And_Z = 79
    V_First_Page = 80
    V_Previous_Page = 81
    V_Next_Page = 82
    V_Last_Page = 83
    V_Ctrl_And_G = 84
    V_Ctrl_And_R = 85
    V_Ctrl_And_P = 86
End Enum

'Declare SQL constants
Public Const serverName As String = "USLXA-P-SQL01.cps.local"
Public Const CPEdatabaseName As String = "PAC1CPE"
Public Const OrdersDatabaseName As String = "Dashboards"
Public Const databaseUserID As String = "dash"
Public Const databasePassword As String = "manage_DB"

'Default PAC1CPE connection for Prod_Eng table and several other key tables
Public Const PAC1CPE_CONNECTION_STRING As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & CPEdatabaseName & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"

'Database connection to the PAC1Orders
Public Const PAC1OrdersDbConnectionString As String = "Driver={SQL Server};Server=" _
& serverName & ";Database=" & OrdersDatabaseName & ";User ID=" & databaseUserID _
& ";Password=" & databasePassword & ";"

'Declare globals for current user
Global download_directory As String
'Requires importing the WW_User class from WW reference libraries
Global current_ww_user As WW_User

'REQUIRES REFERENCES:
'Microsoft ActiveX Data Objects 6.1 Library
'Provide sql connection string (See constants at top of module), sql query, and the range to copy to
Sub get_sql_recordset(sql_connection_string As String, sql_query As String, sql_copy_range As Range)
    
    'Dimension and set new ADODB Connection
    Dim sql_connection As ADODB.connection: Set sql_connection = New ADODB.connection
    'Assign sql connection string from method input
    sql_connection.ConnectionString = sql_connection_string
    'Open the connection defined by the connection string
    sql_connection.Open
    'Dimension and set sql recordset to the value returned when executing the
    'method query input on the connection defined above
    Dim sql_recordset As ADODB.Recordset: Set sql_recordset = sql_connection.Execute(sql_query)
    'Copy the resulting recordset to the method input range
    sql_copy_range.CopyFromRecordset sql_recordset
    
End Sub

'Test get sql recordset subroutine
Sub test_get_sql_recordset()

    get_sql_recordset PAC1CPE_CONNECTION_STRING, "SELECT * FROM WW_Tools", AddDeleteSupplementWS.Range("BB2")
    
End Sub

'TODO
'Pass in a collection of sql query strings and the target worksheet to paste
Public Sub get_SQL_by_query_list(ByVal sql_string_collection As Collection, ByRef target_worksheet As Worksheet)

    Dim dbConDash As ADODB.connection: Set dbConDash = New ADODB.connection
    dbConDash.ConnectionString = PAC1CPE_CONNECTION_STRING: dbConDash.Open
    
    Dim utilRS As ADODB.Recordset: Set utilRS = New Recordset
    Dim currentColIter As Long: currentColIter = 1
    Dim sqlStringIter As Long
    'for each sql query in the sql query collection
    For sqlStringIter = 1 To sql_string_collection.Count
        
        'Internal iterator for recordset fields
        Dim tempIter As Long: tempIter = 0
        'set utility recordset to the recordset returned when executing this sql query
        Set utilRS = dbConDash.Execute(sql_string_collection.Item(sqlStringIter))
        'row 3, column current column iterator, copy recordset values
        target_worksheet.Cells(3, currentColIter).CopyFromRecordset utilRS
            
        'for each item in recordset fields
        For tempIter = 0 To utilRS.Fields.Count - 1
            'row 2, add the current item position to the current column
            'this is just iterating through all of the field names and labeling them on the sheet
            target_worksheet.Cells(2, currentColIter + tempIter) = utilRS.Fields(tempIter).Name
        Next tempIter
            
        'current column iterator = current column iterator + the number of fields returned since we
        'are doing them all at once with the values and have already iterated through the field names
        target_worksheet.Cells(1, currentColIter).Value = sql_string_collection.Item(sqlStringIter)
        
        'add current column and the number of fields in the table to continue adding columns from additional tables
        currentColIter = currentColIter + utilRS.Fields.Count
            
    Next sqlStringIter

End Sub

'TODO: Finish log method
Sub log_message(message As String, Optional level As String = "INFO")

    On Error Resume Next
    Dim logPath As String
    Dim msg As String
    logPath = "C:\Logs\Log_" & Format(Now(), "YYYYMMDD") & ".txt"
    
    Dim fileNo As Integer: fileNo = FreeFile
    Open logPath For Append As #fileNo
    Print #fileNo, Format(Now(), "yyyy-mm-dd HH:MM:ss") & " [" & level & "] - " & msg
    Close #fileNo

End Sub


'TODO: Determine if the individual subroutines for each form is how we want to do it
'Still mulling over how to make more generic
'Get all documents related to an order
Public Sub get_documents(order_number As String)
    Dim prod_release_success As Boolean
    Dim cust_release_success As Boolean
    Dim prod_eng_order_success As Boolean
    Dim eng_order_review_success As Boolean
    Dim quote_success As Boolean
    Dim order_text_success As Boolean
    
    download_release_forms order_number, ScheduleWS.Range("AJ2"), prod_release_success, cust_release_success
    download_production_order order_number, prod_eng_order_success
    download_order_review order_number, eng_order_review_success
    download_order_entry order_number, quote_success
    download_order_text order_number, order_text_success
    
    Dim messageText As String
    messageText = "Production Release: " & IIf(prod_release_success, "Success", "Failed") & vbNewLine
    messageText = messageText & "Customer Release: " & IIf(cust_release_success, "Success", "Failed") & vbNewLine
    messageText = messageText & "Order Text: " & IIf(order_text_success, "Success", "Failed") & vbNewLine
    messageText = messageText & "Sales Quote: " & IIf(quote_success, "Success", "Failed") & vbNewLine
    messageText = messageText & "Production Order: " & IIf(prod_eng_order_success, "Success", "Failed") & vbNewLine
    messageText = messageText & "Engineering Order Review: " & IIf(eng_order_review_success, "Success", "Failed") & vbNewLine
    
    MsgBox messageText

End Sub

'Update user directory by creating a new ww_user for the current user
'and running "ww_registry_update" where passing in true will also
'download the 3 orders workbooks to user cache
Sub update_user_registry(update_cache As Boolean)
    
    'Set new global ww_user if one doesn't exist
    If current_ww_user Is Nothing Then Set current_ww_user = New WW_User
    
    'Update current user's registry
    current_ww_user.ww_registry_update update_cache

End Sub

'Download Production Order from \\uswwq-p-fs01\orders\Common Files\CPE_Schedule\Processed\
'TODO: Use registry entries or sql entries
Public Sub download_production_order(order_number As String, ByRef was_successful As Boolean)
    Dim source_address As String
    
    On Error GoTo inError
    source_address = "\\uswwq-p-fs01\orders\Common Files\CPE_Schedule\Processed\" & Left$(order_number, 7) & "xxx\" & order_number & "-0.pdf"
    FileCopy source_address, (download_directory & order_number & "-0.pdf")
    
    was_successful = True
    Exit Sub

inError:
    was_successful = False
    
End Sub

'Download
Public Sub download_order_review(order_number As String, ByRef was_successful As Boolean)
    Dim source_address As String
    Dim file_found As String
    
    On Error GoTo inError
    source_address = "\\USLXA-P-FS01\sabetha\Orders\Orders\" & Left$(order_number, 7) & "000\"
    file_found = Dir((source_address & order_number & "*"), vbDirectory)
    
    If (ScheduleWS.Range("AJ2") = "PC") Then
        source_address = source_address & file_found & "\Sales\Internal_Communication\OrderReview  " & order_number & "_EE.pdf"
        If Dir(source_address, vbDirectory) <> vbNullString Then
            FileCopy source_address, (download_directory & "OrderReview " & order_number & "_EE.pdf")
        Else: Exit Sub
        End If
    ElseIf (ScheduleWS.Range("AJ2") = "ME") Then
        source_address = source_address & file_found & "\Sales\Internal_Communication\OrderReview  " & order_number & "_ME.pdf"
        If Dir(source_address, vbDirectory) <> vbNullString Then
            FileCopy source_address, (download_directory & "OrderReview " & order_number & "_ME.pdf")
        End If
    Else: Exit Sub
    End If
    
    was_successful = True
    Exit Sub

inError:
    was_successful = False
    
End Sub

Public Sub download_order_entry(order_number As String, ByRef was_successful As Boolean)
    Dim source_address_ordentry As String
    Dim source_address_ihco As String
    Dim file_found_ordentry As String
    Dim file_found_ihco As String
    
    On Error GoTo inError
    source_address_ordentry = "\\USLXA-P-FS01\sabetha\Orders\Orders\" & Left$(order_number, 7) & "000\"
    file_found_ordentry = Dir((source_address_ordentry & order_number & "*"), vbDirectory)
    
    source_address_ordentry = source_address_ordentry & file_found_ordentry & "\Sales\Quote Cost Sheet\"
    file_found_ordentry = Dir((source_address_ordentry & "*OrderEntry.pdf"), vbDirectory)
    file_found_ihco = Dir((source_address_ordentry & "*IHCO*.pdf"), vbDirectory)
    
    source_address_ordentry = source_address_ordentry & file_found_ordentry
    source_address_ihco = source_address_ordentry & file_found_ihco
    
    If Dir(source_address_ordentry, vbDirectory) <> vbNullString And Dir(source_address_ordentry, vbDirectory) <> "." Then
        FileCopy source_address_ordentry, (download_directory & order_number & "_Quote Cost Sheet.pdf")
    ElseIf Dir(source_address_ihco, vbDirectory) <> vbNullString And Dir(source_address_ihco, vbDirectory) <> "." Then
        FileCopy source_address_ihco, (download_directory & order_number & "_Quote Cost Sheet.pdf")
    Else: Exit Sub
    End If
    
    was_successful = True
    Exit Sub

inError:
    was_successful = False
    
End Sub

'Download order text
Public Sub download_order_text(order_number As String, ByRef was_successful As Boolean)
    Dim source_address As String
    Dim file_found As String
    
    On Error GoTo inError
    source_address = "\\USLXA-P-FS01\sabetha\Orders\Orders\" & Left$(order_number, 7) & "000\"
    file_found = Dir((source_address & order_number & "*"), vbDirectory)
    
    source_address = source_address & file_found & "\Engineering Documents\Order_Text\"
    file_found = Dir((source_address & "Order_" & order_number & ".pdf"), vbDirectory)
    
    source_address = source_address & file_found
    
    If Dir(source_address, vbDirectory) <> vbNullString Then
        FileCopy source_address, (download_directory & "Order_" & order_number & ".pdf")
    Else: Exit Sub
    End If
    
    was_successful = True
    Exit Sub
inError:
    was_successful = False
    
End Sub

'Provide column letter and worksheet to get last row in the column
Function get_last_row_in_column(column_letter As String, worksheet_ As Worksheet) As Long

    get_last_row_in_column = worksheet_.Cells(worksheet_.Rows.Count, column_letter).End(xlUp).Row

End Function

'Provide worksheet to get last used column
Function get_last_column_on_worksheet(worksheet_ As Worksheet) As Long

    get_last_column_on_worksheet = worksheet_.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

End Function '




'=================================================================
'WINDCHILL METHODS
'=================================================================


'Download customer and production release forms from windchill to WW_User local cache
'Supply order number and engineer type as either "PC" or "ME"
Public Sub download_release_forms(ByVal order_number As String, _
                                        ByVal eng_type As String, _
                                        Optional ByRef prod_release_success As Boolean = False, _
                                        Optional ByRef cust_release_success As Boolean)
    
    Dim customer_release_filename As String
    Dim production_release_filename As String

    If (eng_type = "PC") Then
        customer_release_filename = order_number & "E.xlsm"
        production_release_filename = order_number & "ER.xlsm"
    ElseIf (eng_type = "ME") Then
        customer_release_filename = order_number & "M.xlsm"
        production_release_filename = order_number & "MR.xlsm"
    End If
    
    Dim http_request_string As String
    Dim windchill_order As WindchillObject: Set windchill_order = New WindchillObject

    download_directory = GetSetting("WW_Tools", "Directories", "Local WW Cache")
    
    'TODO:This needs to be more robust; probably this will now reference the registry update from WW_User
    If (Dir(download_directory, vbDirectory) = vbNullString) Then
        update_user_registry True
    End If

    'Get OID info for the windchill object
    windchill_order.get_windchill_data order_number, eng_type

    'Format request string for CUSTOMER RELEASE form
    http_request_string = windchill_order.URL_WINDCHILL_DOCUMENT_DOWNLOAD & windchill_order.customer_release_OID & "/" & customer_release_filename
    'Download CUSTOMER RELEASE file from Windchill
    download_windchill_file http_request_string, customer_release_filename, cust_release_success
    
    'Format request string for PRODUCTION RELEASE form
    http_request_string = windchill_order.URL_WINDCHILL_DOCUMENT_DOWNLOAD & windchill_order.production_release_OID & "/" & production_release_filename
    'Download PRODUCTION RELEASE file from Windchill
    download_windchill_file http_request_string, production_release_filename, prod_release_success
    
End Sub

'TODO: Make sure we're good with what we want these to do and make more generic if needed

Public Sub download_windchill_file(http_request_string As String, fileName As String, ByRef was_successful As Boolean)
    'HTTP Actions Related
    Dim http_request As Object: Set http_request = CreateObject("MSXML2.XMLHTTP")
    Dim http_response_stream As ADODB.Stream: Set http_response_stream = New ADODB.Stream
    
    On Error GoTo inError
    
    With http_request
        .Open "GET", http_request_string, False
        .Send
    End With

    With http_response_stream
        .Open
        .Type = 1
        .Write http_request.ResponseBody
        .SaveToFile download_directory & fileName, adSaveCreateOverWrite
        .Close
    End With
    
    was_successful = True
    Exit Sub

inError:
    was_successful = False

End Sub


'==========================================================
'SAP METHODS
'==========================================================
Public Sub Set_SAP_Table_Position_And_Refresh(ByRef sap_session As GuiSession, ByRef SAP_Window As GuiMainWindow, ByRef SAP_Table As GuiTableControl, ByVal SAP_Table_ID As String, ByVal SAP_Scroll_Location As Long)
    
    Set SAP_Window = sap_session.FindById("wnd[0]")
    Set SAP_Table = SAP_Window.FindById(SAP_Table_ID)
    
    SAP_Table.VerticalScrollbar.Position = SAP_Scroll_Location
    
    Set SAP_Window = sap_session.FindById("wnd[0]")
    Set SAP_Table = SAP_Window.FindById(SAP_Table_ID)

End Sub

Public Function renew_main_window(ByRef sap_session As GuiSession) As GuiMainWindow

    Set renew_main_window = sap_session.FindById("wnd[0]")

End Function

Public Function Renew_Table(ByRef sap_session As GuiSession, ByVal SAP_Table_ID As String) As GuiTableControl

    Set Renew_Table = sap_session.FindById("wnd[0]/" & SAP_Table_ID)
    
End Function

Sub get_sap_parameters_for_login()
Attribute get_sap_parameters_for_login.VB_ProcData.VB_Invoke_Func = "g\n14"
    
    Dim sap_username As String: sap_username = GetSetting("WW_Tools", "Directories", "SAP Username")
    Dim sap_password As String: sap_password = GetSetting("WW_Tools", "Directories", "SAP Password")
    
    open_sap_and_logon sap_username, sap_password

End Sub

'Take username and password as input to login to SAOP
Sub open_sap_and_logon(sap_username As String, sap_password As String)

    Dim sap_com_dispatch
    Dim sap_application As GuiApplication
    Dim connection As GuiConnection
    Dim session As GuiSession
    Dim WSHShell

    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus

    Set WSHShell = CreateObject("WScript.Shell")

    Do Until WSHShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop

    Set WSHShell = Nothing

    Set sap_com_dispatch = GetObject("SAPGUI")
    Set sap_application = sap_com_dispatch.GetScriptingEngine
    Set connection = sap_application.OpenConnection("SFP - (Production)", True)
    Set session = connection.Children(0)

    session.FindById("wnd[0]/usr/txtRSYST-MANDT").Text = "001"
    session.FindById("wnd[0]/usr/txtRSYST-BNAME").Text = sap_username
    session.FindById("wnd[0]/usr/pwdRSYST-BCODE").Text = sap_password
    session.FindById("wnd[0]").SendVKey 0

    Set session = Nothing
    Set connection = Nothing
    Set sap_com_dispatch = Nothing

End Sub

'Quick tool to change descriptions
'TODO
Sub get_sap_object()

    Dim sap_object As New sap_object
    sap_object.sap_session.StartTransaction "MM02"
    
    Dim sap_main_window As GuiMainWindow: Set sap_main_window = sap_object.SAP_MainWindow
    
    Dim parts_iter As Long
    Dim last_row As Long: last_row = get_last_row_in_column("B", SAP_DescriptionsUpdateWS)
    
    For parts_iter = 2 To last_row
    
        sap_main_window.FindByName("RMMG1-MATNR", "GuiCTextField").Text = SAP_DescriptionsUpdateWS.Range("B" & parts_iter).Value
        sap_main_window.SendVKey V_F7
        sap_main_window.FindByName("MAKT-MAKTX", "GuiTextField").Text = SAP_DescriptionsUpdateWS.Range("C" & parts_iter).Value
        sap_main_window.SendVKey V_Save
        
        Stop
        
    Next parts_iter
    
    MsgBox "Done.  Probably?"

End Sub


'====================================
'NOTE:TEMP TEST
'====================================
Sub testgetbom()
    AddDeleteInfoWS.Range("O2:Q100").ClearContents
    get_BOM "W139854.B20", AddDeleteInfoWS, "O", 2
End Sub

Sub collect_boms()

    Dim order_lines As Collection: Set order_lines = New Collection
    Dim ord_line As Variant
    
    For Each ord_line In Split(AddDeleteEntryWS.Range("G3"), ",")
        order_lines.Add (Trim$(ord_line))
    Next ord_line
    
    Dim iter As Long
    For iter = 1 To order_lines.Count
        get_sql_recordset PAC1CPE_CONNECTION_STRING, _
                            "SELECT Material from Prod_Eng where Order_Num = '" & _
                                AddDeleteEntryWS.Range("B2") & _
                                "' AND Line_Num = '" & order_lines(iter) & "'", _
                            AddDeleteInfoWS.Range("D" & iter)
        
        get_BOM AddDeleteInfoWS.Range("D" & iter), AddDeleteInfoWS, "E", 1
        Stop
    Next iter
    Stop
    'order_lines

End Sub

Public Sub get_BOM(material_number As String, worksheet_input As Worksheet, start_column As String, start_row As Long)
    
    Dim bom_sap_object As sap_object: Set bom_sap_object = New sap_object
    Dim bom_sap_session As GuiSession: Set bom_sap_session = bom_sap_object.sap_session
    Dim bom_sap_main_window As GuiMainWindow: Set bom_sap_main_window = bom_sap_object.SAP_MainWindow
        
    'SAP Function - Start the BOM
    bom_sap_session.StartTransaction "CS03"
    bom_sap_main_window.FindByName("RC29N-MATNR", "GuiCTextField").Text = material_number
    bom_sap_main_window.SendVKey (V_Enter)
    
    Dim max_rows As Long: max_rows = find_max_sap_bom_row(bom_sap_session)
    Dim visible_rows As Long: visible_rows = find_visible_sap_rows(bom_sap_session)

    Dim row_iter As Long: row_iter = 0
    For row_iter = 0 To max_rows
    
        Dim cs03_table_control As GuiTableControl: Set cs03_table_control = get_table_control(bom_sap_session)
    
        If (row_iter > 0 And row_iter Mod visible_rows = 0) Then
            cs03_table_control.VerticalScrollbar.Position = row_iter
            Set cs03_table_control = get_table_control(bom_sap_session)
        End If
        
        Dim cs03_table_row As GuiTableRow: Set cs03_table_row = cs03_table_control.GetAbsoluteRow(row_iter)
        
        On Error GoTo ErrorHandler
        
        worksheet_input.Range(start_column & row_iter + start_row).Value = cs03_table_row.Item(2).Text
        worksheet_input.Range(Chr$(Asc(start_column) + 1) & row_iter + start_row).Value = cs03_table_row.Item(3).Text
        worksheet_input.Range(Chr$(Asc(start_column) + 2) & row_iter + start_row).Value = cs03_table_row.Item(4).Text
    
    Next row_iter
    
EndTable:

    bom_sap_session.EndTransaction
    Exit Sub

ErrorHandler:
    If Err.Number = 614 Then GoTo EndTable Else: Resume Next
    
End Sub


'Find the available SAP BOM rows
Function find_max_sap_bom_row(ByRef sap_session As GuiSession) As Long

    find_max_sap_bom_row = sap_session.FindById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").RowCount
    
End Function

'Find the visible rows available in the GuiSession object
Function find_visible_sap_rows(ByVal sap_session As GuiSession) As Long

    find_visible_sap_rows = sap_session.FindById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").VisibleRowCount

End Function


Function get_table_control(ByRef sap_session As GuiSession) As GuiTableControl

    Set get_table_control = sap_session.FindById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT")

End Function


'====================================
'NOTE:TEMP TEST
'====================================
Sub test_bom_import()
    
    import_bom_from_workbook "C:\Users\j.purdon\Desktop\Shortcuts\WIP\BOM Examples\W147585.B01.xls", AddDeleteInfoWS, "S", 2

End Sub


'This sub assumes AutoCAD Electrical BOM to file export format, or solidworks export format
Sub import_bom_from_workbook(workbook_name As String, write_worksheet As Worksheet, start_write_column As String, start_write_row As Long)
    
    'Dimension and set new RegExp object
    Dim Material_Regex As RegExp: Set Material_Regex = New RegExp
    'Regex matching pattern - WW Material Numbers
    Material_Regex.Pattern = "[vVwW]\d{6}\.?[aAbB]\d{2}"
    'Match multiple when executing and add to collection
    Material_Regex.Global = True
    Material_Regex.IgnoreCase = True
    
    'Prevent application from performing any graphical updates
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
      
    'Workbook to be opened; Dimension and assign upon opening the workbook
    Dim import_workbook As Workbook: Set import_workbook = Application.Workbooks.Open(workbook_name)
    'Grab the first active sheet for object assignment
    Dim import_worksheet As Worksheet: Set import_worksheet = import_workbook.ActiveSheet
    'What row has the header fields, it is different between AutoCAD and Solidworks
    Dim header_row As Long: header_row = IIf(import_worksheet.Range("A1") = "PART LIST", 2, 1)
    'The material number column that will be written to
    Dim material_column As String: material_column = start_write_column
    'The description column that will be written to
    Dim description_column As String: description_column = Chr$(Asc(start_write_column) + 1)
    'The quantity column that will be written to
    Dim quantity_column As String: quantity_column = Chr$(Asc(start_write_column) + 2)
    'Last column in the workbook that has any values
    Dim last_column As Long: last_column = import_worksheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
            
    'Iterate over all the columns with values in the imported notebook
    Dim current_column As Long
    For current_column = 1 To last_column
        'MatchCollection to hold the results of RegExp execution
        Dim material_number_collection As MatchCollection
        Dim material_number As String
        Dim material_number_iter As Long
        Dim last_row As Long
        
        'Select the column header field, trim and convert to Ucase for select case
        Select Case Trim(UCase(import_worksheet.Cells(header_row, current_column).Value))
            'If the column header name is part number (solidworks) or material (autocad) then
            Case "PART NUMBER", "MATERIAL"
                'find the last row of the imported notebook
                last_row = import_worksheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
                'Iterate through each material number in the column and execute regex on the value
                'to ensure the part number is formatted correctly.
                For material_number_iter = 1 To last_row - header_row
                    material_number = import_worksheet.Cells(material_number_iter + header_row, current_column)
                    
                    'If there is an S - and W - for different part numbers per facility Sabetha/Whitewater
                    'Only grab the last match for a properly formatted material number
                    Set material_number_collection = Material_Regex.Execute(material_number)
                    write_worksheet.Cells(start_write_row + material_number_iter - 1, material_column) = _
                                                material_number_collection(material_number_collection.Count - 1)
                Next material_number_iter
            'Both AutoCAD and SolidwWorks have the description field labeled "description"
            Case "DESCRIPTION"
                With import_worksheet
                    .Range(.Cells(header_row + 1, current_column), _
                           .Cells(header_row + 501, current_column)).Copy
                End With
                With write_worksheet
                    .Range(.Cells(start_write_row, description_column), _
                           .Cells(start_write_row + 500, material_column)).PasteSpecial Paste:=xlPasteValues
                End With
            'Slight variation, QTY. for solidworks, QTY for autocad
            Case "QTY.", "QTY"
                With import_worksheet
                    .Range(.Cells(header_row + 1, current_column), _
                           .Cells(header_row + 501, current_column)).Copy
                End With
                With write_worksheet
                    .Range(.Cells(start_write_row, quantity_column), _
                           .Cells(start_write_row + 500, material_column)).PasteSpecial Paste:=xlPasteValues
                End With
        End Select
        
    Next
    
    'Close the imported workbook
    import_workbook.Close
    
    'Allow graphical updates to continue
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
End Sub

'Write to AddDeleteSupplementWS.Range("AA1")
Sub check_cache_button_press()
    
    AddDeleteSupplementWS.Range("AA1:AA10000").ClearContents
    check_cache AddDeleteSupplementWS, "AA", 1

End Sub

Sub check_cache(write_worksheet As Worksheet, start_column As String, start_row As Long)
    
    Dim cache_folder_path As String
    Dim file_name As String
    Dim row_iter As Integer

    ' Use GetSetting to retrieve the path
    ' Replace "YourDefaultPath" with a default path if the setting might not exist
    cache_folder_path = GetSetting("WW_Tools", "Directories", "Local WW Cache", Environ$("localappdata") & "\WW_Tools\Tools Cache")

    ' Check if the path is valid
    If Dir(cache_folder_path, vbDirectory) = "" Then
        MsgBox "Cache folder path not found.", vbExclamation
        Exit Sub
    End If
    
    write_worksheet.Cells(start_row, start_column).Value = cache_folder_path
    
    row_iter = start_row + 1 ' Starting row

    ' List all files in the directory
    file_name = Dir(cache_folder_path & "\*.*") ' Change *.* to a specific file type if needed
    Do While file_name <> ""
        ' Create hyperlinks in Excel
        write_worksheet.Hyperlinks.Add Anchor:=write_worksheet.Cells(row_iter, start_column), _
                                        Address:=cache_folder_path & "\" & file_name, TextToDisplay:=file_name
        row_iter = row_iter + 1
        file_name = Dir
    Loop

End Sub

'Clear out local user cache
'Open, Parts, and Shipped xls files will be the only remaining files
Sub clear_cache()
    Dim cache_folder_path As String
    Dim file_name As String

    ' Retrieve the cache folder path
    cache_folder_path = GetSetting("WW_Tools", "Directories", "Local WW Cache", Environ$("localappdata") & "\WW_Tools\Tools Cache")

    ' Check if the path is valid
    If Dir(cache_folder_path, vbDirectory) = "" Then
        MsgBox "Cache folder path not found.", vbExclamation
        Exit Sub
    End If

    ' Iterate through all files in the directory
    file_name = Dir(cache_folder_path & "\*.*")
    Do While file_name <> ""
        ' Check if the file is not one of the files to keep
        If file_name <> "Open.XLS" And file_name <> "Parts.XLS" And file_name <> "Shipped.XLS" Then
            ' Delete the file
            Kill cache_folder_path & "\" & file_name
        End If
        file_name = Dir ' Get next file
    Loop

    MsgBox "Tools Cache Cleared." & vbNewLine & "Files Not Deleted When Clearing Cache:" & vbNewLine & _
                                        "Open.XLS" & vbNewLine & _
                                        "Parts.XLS" & vbNewLine & _
                                        "Shipped.XLS"
End Sub

Sub clear_cache_button_press()

    clear_cache
    check_cache_button_press

End Sub

Sub hide_worksheets()

    'Another thing that might be doable is just marking some cell on every
    'worksheet that indicates whether it should typically be hidden or not
    hide_worksheet AddDeleteInfoWS
    hide_worksheet AddDeleteKPIsWS
    hide_worksheet AddDeleteSupplementWS
    hide_worksheet FullDatabaseDumpWS
    hide_worksheet QueryResultsWS
    hide_worksheet QueryResults2WS
    hide_worksheet DevToolsWS
    hide_worksheet SAP_DescriptionsUpdateWS


End Sub

'Hide input worksheet
Private Sub hide_worksheet(input_worksheet As Worksheet)

    If input_worksheet.Visible = xlSheetVisible Then input_worksheet.Visible = xlSheetHidden

End Sub

'Unhide all worksheets in workbook
Sub unhide_worksheets()

    Dim o_worksheet As Worksheet
    For Each o_worksheet In ThisWorkbook.Worksheets
        o_worksheet.Visible = xlSheetVisible
    Next o_worksheet

End Sub

'Update SAP Parameters stored in registry, this is dev utility
Public Sub update_sap_parameters(sap_username As String, sap_password As String)
    ' Update the SAP Username
    SaveSetting "WW_Tools", "Directories", "SAP Username", sap_username
    ' Update the SAP Password
    SaveSetting "WW_Tools", "Directories", "SAP Password", sap_password
End Sub




Sub get_ww_users(department As String, target_range As Range)

    Dim sql_string As String
    sql_string = "SELECT * FROM WW_Prod_Eng_Names WHERE WW_Active = 'Y' AND WW_Initials <> 'N/A' AND WW_Department = '" & department & "'"

    get_sql_recordset PAC1CPE_CONNECTION_STRING, sql_string, target_range

End Sub
