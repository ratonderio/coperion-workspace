Attribute VB_Name = "AddDeletes"
'@IgnoreModule ProcedureNotUsed, ExcelMemberMayReturnNothing, SetAssignmentWithIncompatibleObjectType, MemberNotOnInterface
'@Folder "Modules"
Option Explicit

Dim add_del_order_number As Long
Dim add_del_date As String
Dim add_del_engineer As String
Dim add_del_lines_affected As String
Dim add_del_description As String
Dim add_del_reason_code As String
Dim add_del_action As String
Dim add_del_quantity As Integer
Dim add_del_revision As String
Dim add_del_material As String

Sub reset_add_delete_form()

    'Clear the contents of the FullDatabaseDump worksheet
    FullDatabaseDumpWS.Range("A1:MA9999").ClearContents
    AddDeleteEntryWS.Range("R4:AF100").ClearContents
    
    Dim iter As Long
    Dim current_range As Range
    For iter = 2 To 8
    
        Set current_range = AddDeleteEntryWS.Range("G" & iter)
        Select Case iter
            Case 2: current_range = "A"
            Case 3: current_range = "INSERT LINES"
            Case 4: current_range = "NOTE"
            Case 7: current_range = "INSERT DESCRIPTION"
            Case 8: current_range = "E001: Production Engineering Error"
            Case Else: current_range = vbNullString
            
        End Select
    
    Next iter
    
    get_all_order_information AddDeleteEntryWS.Range("B2").Value
    fill_listbox_items
    
End Sub

Sub get_all_add_delete_order_information()
    get_all_order_information AddDeleteEntryWS.Range("B2").Value
    fill_listbox_items
End Sub

'Use with the button 'Get Order Info' on AddDeleteEntry for now
'This might end up getting used with the order details sheet as well
Sub get_all_order_information(full_order_details_order_number As String)

    FullDatabaseDumpWS.Range("A1:ZZ1000").ClearContents
    
    'Create a collection of SQL commands to be run with the 'get sql' routine
    Dim sqlCommands As Collection: Set sqlCommands = New Collection
    
    '@Ignore VariableNotUsed
    Dim line_items_listbox As ControlFormat: Set line_items_listbox = AddDeleteEntryWS.Shapes.Item("Line Items Listbox").ControlFormat
    
    'Dash @ USLXA-P-FS01/PAC1CPE
    'PAC1CPE - prod_status
    sqlCommands.Add ("SELECT * FROM PAC1CPE.dbo.prod_status WHERE orderno = '" & full_order_details_order_number & "'") '1 - Dash
    'PAC1CPE - Prod_Eng
    sqlCommands.Add ("SELECT * FROM PAC1CPE.dbo.Prod_Eng WHERE Order_Num = '" & full_order_details_order_number & "' ORDER BY Line_Num") '2 - Dash
    'PAC1CPE - cpe_shipmeet
    sqlCommands.Add ("SELECT * FROM PAC1CPE.dbo.cpe_shipmeet WHERE orderno = '" & full_order_details_order_number & "'") '3 - Dash
    'PAC1CPE - cpe_schedule
    sqlCommands.Add ("SELECT * FROM PAC1CPE.dbo.cpe_schedule WHERE orderno = '" & full_order_details_order_number & _
                     "' AND NOT comments = ''")  '4 - Dash
    'PAC1CPE -  cpe_scheduling
    sqlCommands.Add ("SELECT * FROM PAC1CPE.dbo.cpe_scheduling WHERE orderno = '" & full_order_details_order_number & "'") '5 - Dash

    'Dash @ USLXA-p-FS01/Dashboards
    'Dashboards - wwo_issues
    sqlCommands.Add ("SELECT * FROM Dashboards.dbo.wwo_issues WHERE orderno = '" & full_order_details_order_number & "'") '6 - Dash
    
    get_SQL_by_query_list sqlCommands, FullDatabaseDumpWS
    
End Sub

Public Sub insert_add_delete()

    Dim order_kpi_connection As ADODB.connection: Set order_kpi_connection = New ADODB.connection
    order_kpi_connection.ConnectionString = databaseConnectionString
    
    'TODO: Update with loop
    add_del_order_number = AddDeleteEntryWS.Range("B2")
    add_del_date = AddDeleteEntryWS.Range("B1")
    add_del_engineer = AddDeleteEntryWS.Range("A1")
    add_del_lines_affected = AddDeleteEntryWS.Range("G3")
    add_del_description = Replace(AddDeleteEntryWS.Range("G7"), "'", "''")
    add_del_reason_code = AddDeleteEntryWS.Range("G8")
    add_del_action = AddDeleteEntryWS.Range("G4")
    add_del_material = AddDeleteEntryWS.Range("G6")
    add_del_quantity = AddDeleteEntryWS.Range("G5")
    add_del_revision = AddDeleteEntryWS.Range("G2")
    
    Dim skip_insert As Boolean
    skip_insert = is_duplicate(add_del_order_number, add_del_revision, add_del_description)
    
    If (Not (skip_insert)) Then
    
        Dim sql_string As String: sql_string = "INSERT INTO Order_KPI(OrderNumber, AddDeleteDate, Engineer, " & _
            "LinesAffected, Description, ReasonCode, ActionPerformed, MaterialNumber, Quantity, Revision) " & _
            "VALUES(" & _
            add_del_order_number & ", '" & _
            add_del_date & "', '" & _
            add_del_engineer & "', '" & _
            add_del_lines_affected & "', '" & _
            add_del_description & "', '" & _
            Left$(add_del_reason_code, 4) & "', '" & _
            add_del_action & "', '" & _
            add_del_material & "', " & _
            add_del_quantity & ", '" & _
            add_del_revision & "')"

        order_kpi_connection.Open
        order_kpi_connection.Execute (sql_string)

        MsgBox "Add Delete Entered For Order: " & add_del_order_number & vbNewLine & "Revision: " & add_del_revision
    Else
        MsgBox "Duplicate Entry"
    End If

End Sub

Private Function is_duplicate(order_number As Long, revision_letter As String, revision_description As String) As Boolean

    AddDeleteInfoWS.Range("AA1:AZ10").ClearContents
    get_order_add_deletes
    get_sql_recordset PAC1CPE_CONNECTION_STRING, "SELECT * FROM Order_KPI WHERE OrderNumber = '" & order_number & "' AND Revision = '" & revision_letter & "' AND Description = '" & revision_description & "'", AddDeleteInfoWS.Range("AA1")
    
    is_duplicate = IIf((AddDeleteInfoWS.Range("AA1") = vbNullString), False, True)

End Function

'Probably want to change this to a generic function and pass in the order number
'TODO
Public Sub get_order_add_deletes()

    AddDeleteEntryWS.Range("R10:AB100").ClearContents

    Dim order_kpi_connection As ADODB.connection: Set order_kpi_connection = New ADODB.connection
    order_kpi_connection.ConnectionString = databaseConnectionString
    
    Dim order_kpi_recordset As ADODB.Recordset
    Dim sql_string As String: sql_string = "SELECT * FROM dbo.Order_KPI WHERE OrderNumber = " & _
        AddDeleteEntryWS.Range("B2") & " ORDER BY Revision DESC"
                
    order_kpi_connection.Open
    Set order_kpi_recordset = order_kpi_connection.Execute(sql_string)
    
    AddDeleteEntryWS.Range("R10").CopyFromRecordset order_kpi_recordset
    
    If (AddDeleteEntryWS.Range("R10").Value = vbNullString) Then
        MsgBox "No previous history found."
    Else
        MsgBox "Revision history found, Latest Revision: " & AddDeleteEntryWS.Range("AB10")
    End If
    
    order_kpi_connection.Close
    

    
    
End Sub

Sub fill_listbox_items()

    Dim line_item_listbox As ListBox: Set line_item_listbox = AddDeleteEntryWS.ListBoxes("List Box 16")
    Dim listbox_last_row As Long: listbox_last_row = FullDatabaseDumpWS.Cells(Rows.Count, "K").End(xlUp).Row
    
    line_item_listbox.RemoveAllItems
    
    Dim rowIter As Long
    For rowIter = 3 To listbox_last_row

        line_item_listbox.AddItem (AddDeleteInfoWS.Range("B" & rowIter))

    Next rowIter
     
End Sub

Sub ListBox16_Change()

    Dim line_item_listbox As ListBox: Set line_item_listbox = AddDeleteEntryWS.ListBoxes("List Box 16")
    Dim line_item As Variant
    Dim line_item_index As Long: line_item_index = 1
    
    Dim print_string As String
     
    For Each line_item In line_item_listbox.List

        If line_item_listbox.Selected(line_item_index) Then
            print_string = IIf(print_string = vbNullString, _
                               Left$(line_item, InStr(1, line_item, ":") - 1), _
                               print_string & ", " & Left$(line_item, InStr(1, line_item, ":") - 1))
            
        End If
        
        line_item_index = line_item_index + 1
     
    Next line_item
        
    AddDeleteEntryWS.Range("G3").Value = print_string
End Sub

'=================================================
'UPDATE ADD DELETE FORMS
'=================================================
Sub update_production_release_add_delete(order_number As String, eng_type As String)
    'Update the user registry but not the cache.  Default cache is WW_Tools
    update_user_registry False
    
    'Download both release forms for the order number to cache
    download_release_forms order_number, eng_type
    
    Dim production_release_filename As String
    Dim production_release As Workbook
    
    'Filename is the full local cache name (directory) and order number and either ER or MR .xlsm
    production_release_filename = IIf(eng_type = "PC", _
                                      current_ww_user.local_user_cache & order_number & "ER.xlsm", _
                                      IIf(eng_type = "ME", _
                                          current_ww_user.local_user_cache & order_number & "MR.xlsm", _
                                          "INVALID"))
                  
    'Set the production release workbook
    Set production_release = Application.Workbooks.Open(production_release_filename)
    
    'Set the production release cover sheet, sheet name should always be Cover_Sheet
    Dim release_cover_sheet As Worksheet: Set release_cover_sheet = production_release.Worksheets("Cover_Sheet")
    Dim current_revision_letter As String: current_revision_letter = "A"
    Dim last_revision_row As Long: last_revision_row = get_last_row_in_column("A", release_cover_sheet)
    
    Dim row_iter As Long
    For row_iter = 18 To last_revision_row
        
        If release_cover_sheet.Cells(row_iter, "A").Value > current_revision_letter Then
            current_revision_letter = release_cover_sheet.Cells(row_iter, "A").Value
        End If
        
    Next row_iter
    
    'Convert current revision letter to ASCII value to add 1; then convert result to character
    'Place on release cover in the newest revision field
    release_cover_sheet.Cells(last_revision_row + 1, "A") = Chr(Asc(current_revision_letter) + 1)
    'Newest revision date equals today
    release_cover_sheet.Cells(last_revision_row + 1, "B") = Date
    'Newest revision lines affected, these are all pulled from AddDeleteEntryWS
    release_cover_sheet.Cells(last_revision_row + 1, "C") = AddDeleteEntryWS.Range("G3")
    'Newest revision description
    release_cover_sheet.Cells(last_revision_row + 1, "D") = AddDeleteEntryWS.Range("G7")
    'Newest revision reason code
    release_cover_sheet.Cells(last_revision_row + 1, "E") = Left$(AddDeleteEntryWS.Range("G8"), 4)

    production_release.SaveAs Left$(production_release.FullName, Len(production_release.FullName) - 5) & "_REV_" & Chr(Asc(current_revision_letter) + 1)
    'Stop
    

End Sub

Sub update_production_release_button_click()

    update_production_release_add_delete AddDeleteEntryWS.Range("B2"), ScheduleWS.Range("AJ2")

End Sub


