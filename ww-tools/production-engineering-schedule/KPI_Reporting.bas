Attribute VB_Name = "KPI_Reporting"
'@IgnoreModule ProcedureNotUsed
'@Folder("Modules")
Option Explicit

Sub CompareBOMs()

    Dim last_row_sap As Long
    Dim last_row_import As Long
    Dim iter_sap_bom As Long
    Dim iter_import_bom As Long
    Dim materialNumber As String
    
    Dim temp_dict As Dictionary: Set temp_dict = New Dictionary
    Dim sap_bom_dict As Dictionary: Set sap_bom_dict = New Dictionary
    Dim import_bom_dict As Dictionary: Set import_bom_dict = New Dictionary
    Dim desc_dict As Dictionary: Set desc_dict = New Dictionary
    
    AddDeleteSupplementWS.Range("A2:S1000").ClearContents
    
    'last row of SAP BOM and IMPORT BOM
    last_row_sap = AddDeleteInfoWS.Cells(AddDeleteInfoWS.Rows.Count, "O").End(xlUp).Row
    last_row_import = AddDeleteInfoWS.Cells(AddDeleteInfoWS.Rows.Count, "S").End(xlUp).Row
    
    
    'iterate through SAP BOM numbers and add to dictionary; if exists sum quantity
    For iter_sap_bom = 2 To last_row_sap
        materialNumber = AddDeleteInfoWS.Cells(iter_sap_bom, "O").Value
        
        If Not sap_bom_dict.Exists(materialNumber) Then
            'Add material number and quantity if material number is not in SAP BOM dict
            sap_bom_dict.Add materialNumber, AddDeleteInfoWS.Cells(iter_sap_bom, "Q").Value
            'Add material description to DESCRIPTION DICT
            desc_dict.Add materialNumber, AddDeleteInfoWS.Cells(iter_sap_bom, "P").Value
            'Add material number and quantity to TEMP DICT
            temp_dict.Add materialNumber, AddDeleteInfoWS.Cells(iter_sap_bom, "Q").Value
        Else
            'add the value in the SAP BOM dict to the quantity of this material number and insert into dict
            sap_bom_dict(materialNumber) = sap_bom_dict(materialNumber) + AddDeleteInfoWS.Cells(iter_sap_bom, "Q").Value
            'add the value in the TEMP DICT to the quantity of this material number and insert into dict
            temp_dict(materialNumber) = temp_dict(materialNumber) + AddDeleteInfoWS.Cells(iter_sap_bom, "Q").Value
        End If
    Next iter_sap_bom
    
    
    'iterate through IMPORT BOM numbers and add to dictionary; if exists sum quantity
    For iter_import_bom = 2 To last_row_import
        materialNumber = AddDeleteInfoWS.Cells(iter_import_bom, "S").Value
        
        If Not import_bom_dict.Exists(materialNumber) Then
            'Add material number and quantity if material number is not in IMPORT BOM DICT
            import_bom_dict.Add materialNumber, AddDeleteInfoWS.Cells(iter_import_bom, "U").Value
        Else
            'add the value in the IMPORT BOM DICT to the quantity of this material number and insert into dict
            import_bom_dict(materialNumber) = import_bom_dict(materialNumber) + AddDeleteInfoWS.Cells(iter_import_bom, "U").Value
        End If
        
        If Not temp_dict.Exists(materialNumber) Then
            'add material number and quantity if material is not in TEMP DICT
            temp_dict.Add materialNumber, AddDeleteInfoWS.Cells(iter_import_bom, "U").Value
        Else
            'subtract the current quantity from the TEMP DICT quantity to get the difference in quantities for difference bom
            temp_dict(materialNumber) = temp_dict(materialNumber) - AddDeleteInfoWS.Cells(iter_import_bom, "U").Value
        End If
        
        If Not desc_dict.Exists(materialNumber) Then
            'add material number and description to DESCRIPTION DICT if material is not in DESCRIPTION DICT
            desc_dict.Add materialNumber, AddDeleteInfoWS.Cells(iter_import_bom, "T").Value
        End If
        
    Next iter_import_bom

    'print material number and quantity from SAP BOM dict to AddDeleteSupplement Worksheet
    Dim iter_key As Variant
    
    'iterator starts at 2
    Dim iter_row As Long: iter_row = 2
    
    'copy SAP BOM dict to AddDeleteSupplement Worksheet
    For Each iter_key In sap_bom_dict
        
        AddDeleteSupplementWS.Cells(iter_row, 2) = iter_key
        AddDeleteSupplementWS.Cells(iter_row, 3) = desc_dict(iter_key)
        AddDeleteSupplementWS.Cells(iter_row, 4) = sap_bom_dict(iter_key)
        
        iter_row = iter_row + 1
        
    Next iter_key
    
    'reset iter row
    iter_row = 2
    'copy IMPORT BOM DICT to AddDeleteSupplement Worksheet
    For Each iter_key In import_bom_dict
        
        AddDeleteSupplementWS.Cells(iter_row, 6) = iter_key
        AddDeleteSupplementWS.Cells(iter_row, 7) = desc_dict(iter_key)
        AddDeleteSupplementWS.Cells(iter_row, 8) = import_bom_dict(iter_key)
        
        iter_row = iter_row + 1
        
    Next iter_key
    
    'reset iter row
    iter_row = 2
    'print material number and quantity from TEMP DICT to AddDeleteSupplement Worksheet
    For Each iter_key In temp_dict
        
        AddDeleteSupplementWS.Cells(iter_row, 10) = iter_key
        AddDeleteSupplementWS.Cells(iter_row, 11) = desc_dict(iter_key)
        AddDeleteSupplementWS.Cells(iter_row, 12) = temp_dict(iter_key)
        
        iter_row = iter_row + 1
        
    Next iter_key


End Sub

Sub get_yearly_eng_orders()

    get_yearly_engineering_orders ("2023")

End Sub

Sub color_code_diff_bom()
    
    Dim anotherIter As Long
    For anotherIter = 2 To 100
    
        Select Case AddDeleteSupplementWS.Range("L" & anotherIter)
            Case Is >= 1: 'Add Parts (RED)
                AddDeleteSupplementWS.Range("J" & anotherIter & ":L" & anotherIter).Interior.color = RGB(30, 90, 220)
            Case Is < 0: 'Delete Parts (BLUE)
                With AddDeleteSupplementWS.Range("J" & anotherIter & ":L" & anotherIter)
                    .Interior.color = RGB(158, 14, 14)
                    .Font.color = vbWhite
                End With
            Case 0: 'No difference (GREEN)
                AddDeleteSupplementWS.Range("J" & anotherIter & ":L" & anotherIter).Interior.color = RGB(0, 220, 0)
        End Select
    
    Next anotherIter
End Sub


Sub get_yearly_engineering_orders(ByVal report_year As String)

    AddDeleteSupplementWS.Range("A2:Z1000").ClearContents

    Dim sql_string As String
    Dim sql_copy_range As Range: Set sql_copy_range = AddDeleteSupplementWS.Range("AA2")
    
    sql_string = "SELECT * FROM (SELECT *, ROW_NUMBER() OVER (PARTITION BY Order_Num ORDER BY Order_Num) as rn " & _
    "FROM Prod_Eng WHERE PC_Act_Rel >= '" & report_year & "' OR ME_Act_Rel >= '" & report_year & "') AS SUBQUERY WHERE rn = 1;"
    
    get_and_paste_sql_query sql_string, sql_copy_range

End Sub


