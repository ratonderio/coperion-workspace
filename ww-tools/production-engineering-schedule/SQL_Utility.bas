Attribute VB_Name = "SQL_Utility"
'@Folder("__Modules")
Option Explicit
Sub get_database_tables_and_columns(update_sql_tables As Boolean)
    DevToolsWS.Range("AA2:AZ5000").ClearContents
    Dim database_tables_connection As ADODB.connection
    Set database_tables_connection = New ADODB.connection
    database_tables_connection.ConnectionString = PAC1CPE_CONNECTION_STRING
    database_tables_connection.Open
    
    Dim database_tables_recordset As ADODB.Recordset
    Set database_tables_recordset = database_tables_connection.Execute("SELECT " & _
                                                                       "TABLE_CATALOG, TABLE_SCHEMA, TABLE_NAME, COLUMN_NAME, ORDINAL_POSITION, DATA_TYPE " & _
                                                                       "FROM PAC1CPE.INFORMATION_SCHEMA.COLUMNS ORDER BY TABLE_NAME, ORDINAL_POSITION")
    
    DevToolsWS.Range("AA2").CopyFromRecordset database_tables_recordset
    
    If update_sql_tables Then generate_table_dictionary

End Sub

Sub generate_table_dictionary()
    Dim table_dictionary As Scripting.Dictionary: Set table_dictionary = New Dictionary
    Dim last_row As Long: last_row = get_last_row_in_column("AC", DevToolsWS)
    
    Dim row_iter As Long
    Dim tableName As String
    Dim column_name As Variant
    
    SQL_TablesWS.Range("A1:AZ500").ClearContents
    
    For row_iter = 2 To last_row
        tableName = DevToolsWS.Cells(row_iter, "AC").Value
        column_name = DevToolsWS.Cells(row_iter, "AD").Value
        
        If Not table_dictionary.Exists(tableName) Then
            ' If the table name doesn't exist in the dictionary, add it with a new collection
            table_dictionary.Add tableName, New Collection
        End If
        
        ' Add the column name to the collection associated with the table
        table_dictionary(tableName).Add column_name
    Next row_iter
    
    
    Dim key As Variant
    Dim row_counter As Long
    Dim column_counter As Long: column_counter = 1 'Start on Column 1
    
    For Each key In table_dictionary.Keys
        'Column names start on row 2
        row_counter = 2
        'Set first row value to table name
        SQL_TablesWS.Cells(1, column_counter).Value = key
        
        'For each column name in the dictionary for that key
        For Each column_name In table_dictionary(key)
            'SQL Tables cell current row counter, currnt column counter = column name
            SQL_TablesWS.Cells(row_counter, column_counter).Value = column_name
            'Increment row counter
            row_counter = row_counter + 1
        Next column_name
        'increment column counter
        column_counter = column_counter + 1
    Next key

End Sub

Sub run_update()
    
    get_database_tables_and_columns True

End Sub

Sub update_tables_dropdown()

    create_dropdown DevToolsWS

End Sub


Sub create_dropdown(input_worksheet As Worksheet)
    'Get last column on SQL Tables worksheet
    Dim last_column As Long: last_column = get_last_column_on_worksheet(SQL_TablesWS)
    
    Dim tableNames As String
    'column iterator
    Dim column_iter As Long
    
    'for column 1 to last column
    For column_iter = 1 To last_column
        'format a comma separated list of strings
        tableNames = tableNames & SQL_TablesWS.Cells(1, column_iter).Value & ","
    Next column_iter
    
    tableNames = Left(tableNames, Len(tableNames) - 1) ' Remove the trailing comma
    
    cell_validation_using_string input_worksheet.Range("C2"), tableNames
    'cell_validation_using_string input_worksheet.Range("F2"), tableNames
    
End Sub


Sub cell_validation_using_collection(input_range As Range, input_collection As Collection)
    Dim collection_string As String
    Dim collection_iter As Long
    
    For collection_iter = 1 To input_collection.Count
    
        collection_string = collection_string & input_collection(collection_iter) & ","
    
    Next collection_iter
    
    collection_string = Left$(collection_string, Len(collection_string) - 1)
    input_range.Validation.Delete
    input_range.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, collection_string
    
    Stop
    

End Sub

Sub cell_validation_using_string(input_range As Range, input_string As String)

    input_range.Validation.Delete
    input_range.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, input_string

End Sub


