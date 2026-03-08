VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BOMCreation 
   Caption         =   "BOM_Creation"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10455
   OleObjectBlob   =   "BOMCreation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BOMCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Modules"
Const F_Item_Number As String = "RC29P-POSNR" 'Text Field
Const F_Material_Number As String = "RC29P-IDNRK" 'CText Field
Const F_Material_Description As String = "RC29P-KTEXT" 'Text Field
Const F_Material_Quantity As String = "RC29P-MENGE" 'Text Field
Const F_Item_ID As String = "RC29P-IDENT" 'Text Field
Const F_Table_Valid_Rows As String = "RC29P-ENTRY" 'Text Field

Const F_Material_Number_Entry As String = "RC29N-MATNR" 'CText Field
Const F_Plant_Number_Entry As String = "RC29N-WERKS" 'CText Field
Const F_BOM_Usage_Entry As String = "RC29N-STLAN" 'CText Field
Const F_Valid_Date_Entry As String = "RC29N-DATUV"

Const F_Text_Field As String = "GuiTextField"
Const F_CText_Field As String = "GuiCTextField"
Const T_CS_Table As String = "usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT"

Const CS01_Title As String = "Create material BOM: General Item Overview"
Const CS02_Title As String = "Change material BOM: General Item Overview"

Dim DNU_Dict As Dictionary

Dim Electrical_B_Numbers(9) As String
Dim Mechanical_B_Numbers(9) As String


'
'!!!!! INITIALIZE !!!!!
'
Private Sub UserForm_Initialize()

    Dim B_Index As Long

    For B_Index = 0 To 9
    
        Electrical_B_Numbers(B_Index) = ".B2" & B_Index
        Mechanical_B_Numbers(B_Index) = ".B3" & B_Index
    
    Next

    Username_Label.Caption = UCase(Environ("username"))
    Department_ListBox.List = Array("ELECTRICAL", "MECHANICAL")
    Reporting_ListBox.List = Array("EMAIL", "EXCEL -> Default Location", "EXCEL -> Custom Location")
    Routing_ListBox.List = Array("NO", "YES")
    B_Number_ListBox.List = Electrical_B_Numbers
    
    
    Department_ListBox.ListIndex = 0
    Reporting_ListBox.ListIndex = 0
    Routing_ListBox.ListIndex = 0
    B_Number_ListBox.ListIndex = 0
    
    B_Number_Label.Visible = False
    B_Number_ListBox.Visible = False
    Custom_Location_Label.Visible = False
    Pick_Location_Button.Visible = False
    
    If Not [BOM_Creation_Material_Number] = vbNullString Then
        Material_Number_TextBox.Value = [BOM_Creation_Material_Number]
    End If
    
    Generate_Material_Number

End Sub

'
'!!!!! CHANGE PROCEDURES !!!!!
'
Private Sub B_Number_ListBox_Change()

    Generate_Material_Number

End Sub

Private Sub B_Numbers_CheckBox_Change()

    B_Number_Label.Visible = B_Numbers_CheckBox.Value
    B_Number_ListBox.Visible = B_Numbers_CheckBox.Value
    B_Number_ListBox.ListIndex = 0
    
End Sub

Private Sub Department_ListBox_Change()

    Select Case Department_ListBox.ListIndex
    
        Case 0
            B_Number_ListBox.List = Electrical_B_Numbers
            Material_Description_Label.Visible = True
            Material_Description_TextBox.Visible = True
            Default_Options_Checkbox.Visible = True
            Part_Type_Label.Visible = True
            Storage_Location_Label.Visible = True
            Product_Hierarchy_Label.Visible = True
        Case 1
            B_Number_ListBox.List = Mechanical_B_Numbers
            Material_Description_Label.Visible = False
            Material_Description_TextBox.Visible = False
            Default_Options_Checkbox.Visible = False
            Part_Type_Label.Visible = False
            Storage_Location_Label.Visible = False
            Product_Hierarchy_Label.Visible = False
        Case Else
            Debug.Print "WHY"
    End Select
    
    B_Number_ListBox.ListIndex = 0
    Generate_Material_Number
    

End Sub

Private Sub Reporting_ListBox_Change()

    Select Case Reporting_ListBox.ListIndex
    
        Case 0, 1
            Custom_Location_Label.Visible = False
            Pick_Location_Button.Visible = False
        Case 2
            Custom_Location_Label.Visible = True
            Pick_Location_Button.Visible = True
        Case Else
            Debug.Print "WHY"
    
    End Select

End Sub

'
'!!!!! EXIT PROCEDURES !!!!!
'
Private Sub Material_Number_TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    Generate_Material_Number

End Sub

'
'!!!!! CLICK PROCEDURES !!!!!
'
Private Sub Pick_Location_Button_Click()

    Dim Folder_Picker As Office.FileDialog: Set Folder_Picker = Application.FileDialog(msoFileDialogFolderPicker)
    Dim Custom_Folder As String
    
    With Folder_Picker
    
        .InitialView = msoFileDialogViewDetails
        .InitialFileName = Environ("userprofile") & "\Downloads\"

        If .Show = True Then
            
            Custom_Folder = .SelectedItems(1)
        
        End If
    End With
    
    Custom_Location_Label.Caption = Custom_Folder
    
End Sub

Private Sub Cancel_Button_Click()

    Unload Me

End Sub

Private Sub Create_BOM_Button_Click()

    Create_BOM_Init

End Sub
'
'!!!!! CUSTOM PROCEDURES !!!!!
'
Private Sub Create_BOM_Init()
    'Stop
    If Material_Number_TextBox.Value = vbNullString Then
        MsgBox "No material number was entered.", vbExclamation, "No Material Number"
        Exit Sub
    ElseIf [BOM_Creation_Material_Description].Text = "#N/A" Then
        MsgBox "Invalid material number, material not found, or material has not populated in the database yet.", vbExclamation, "Invalid Material Number"
        Exit Sub
    ElseIf Drawing_Number_TextBox.Value = vbNullString Then
        MsgBox "No drawing number was entered.", vbExclamation, "No Drawing Number"
        Exit Sub
    ElseIf Department_ListBox.ListIndex = 0 And Material_Description_TextBox = vbNullString Then
        MsgBox "No Material Description Listed" & vbNewLine & "Electrical must enter a material description", vbExclamation, "No Material Description"
        Exit Sub
    ElseIf Department_ListBox.ListIndex = 0 And Len(Material_Description_TextBox.Value) > 40 Then
        MsgBox "Material Description is too long" & vbNewLine & "Reduce Material Description length to 40 characters or less", vbExclamation, "Material Description Length"
        Exit Sub
    ElseIf Department_ListBox.ListIndex = 0 And Not Default_Options_Checkbox.Value Then
        MsgBox "Default Options Not Checked" & vbNewLine & "Manually create the material number and BOM", vbExclamation, "Default Options Unchecked"
        Exit Sub
    End If
    
    If Department_ListBox.ListIndex = 0 Then
    
        ClearTableCPN
        CreatePartNumbers.Range("C5").Value = Generated_Number_TextBox.Value 'Generated Material Number
        CreatePartNumbers.Range("E5").Value = Drawing_Number_TextBox.Value 'Drawing Number
        CreatePartNumbers.Range("G5").Value = Material_Description_TextBox.Value 'Description
        CreatePartNumbers.Range("I5").Value = "Make To Order (Ind)" 'MRP Controller/Part Type
        CreatePartNumbers.Range("J5").Value = "Electrical Assembly (A119)" 'Storage Location
        CreatePartNumbers.Range("K5").Value = "Controller PE" 'Product Hierarchy
        CreatePartNumbers.Range("L5").Value = "Yes" 'Create Info Record
        Get_Material_Details (True)
        
    Else
    
        Get_Material_Details (False)
    
    End If
    
End Sub

Private Sub Get_Material_Details(Optional ByVal IsElectrical As Boolean = False)

    BOM_Creation.Protect Contents:=False
    ProcessDataBOM.Protect Contents:=False
    
    'Check if logged on and initiate the connection
    On Error GoTo SAP_NotLoggedIn
    Dim SAP_GUI As Object: Set SAP_GUI = GetObject("SAPGUI")
    Dim SAP_App As GuiApplication: Set SAP_App = SAP_GUI.GetScriptingEngine
    Dim SAP_Connection As GuiConnection: Set SAP_Connection = SAP_App.Children(0)
    On Error GoTo 0
    
StartLookingForEasyAccess:
    'Find Number of instances (children) open
    Dim SAP_Session As GuiSession
    Dim NumOfWindowsSAP As Long: NumOfWindowsSAP = SAP_Connection.Sessions.Count
    Dim SAPIndex As Long

    'Seach for an initial "SAP Easy Access" Screen
    For SAPIndex = 0 To NumOfWindowsSAP - 1
        If Left(SAP_Connection.Children(CInt(SAPIndex)).FindById("wnd[0]").Text, 15) = "SAP Easy Access" Then
            Set SAP_Session = SAP_Connection.Children(CInt(SAPIndex))
            Exit For
        End If
    Next SAPIndex
     
    'If a session is not established (None of the instances are "SAP Easy Access")
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
            Exit Sub
        End If
    End If
    
    If IsElectrical Then GoTo Skip_Material_Details

    Dim SAP_Window As GuiMainWindow: Set SAP_Window = SAP_Session.FindById("wnd[0]")
    
    
    '--------------------------------------------------------------------------------------
    'THIS IS A PRIME TARGET FOR REFACTORING TO A LOOP ONCE ALL THE PROD HIERARCHY IS PUT IN
    '----------------------------------------TODO------------------------------------------
    Dim Dropdown_Dict As Dictionary: Set Dropdown_Dict = New Dictionary
    
    If Not Dropdown_Dict.Exists("11M     K    BXN") Then Dropdown_Dict.Add "11M     K    BXN", "Belt Weigher BXN"
    If Not Dropdown_Dict.Exists("11M     K    BXO") Then Dropdown_Dict.Add "11M     K    BXO", "Belt Weigher BXO"
    If Not Dropdown_Dict.Exists("11M     K    BEM") Then Dropdown_Dict.Add "11M     K    BEM", "BEMP Single Idler"
    If Not Dropdown_Dict.Exists("11M     K    BMP") Then Dropdown_Dict.Add "11M     K    BMP", "BEMP Dual Idler"
    If Not Dropdown_Dict.Exists("11M     D    BHW") Then Dropdown_Dict.Add "11M     D    BHW", "Bin Weighing Systems"
    If Not Dropdown_Dict.Exists("11E     Z    BKP") Then Dropdown_Dict.Add "11E     Z    BKP", "Baker Perkins (Elec.)"
    If Not Dropdown_Dict.Exists("11M     Z    BKP") Then Dropdown_Dict.Add "11M     Z    BKP", "Baker Perkins (Mech.)"
    If Not Dropdown_Dict.Exists("11E     K    CPE") Then Dropdown_Dict.Add "11E     K    CPE", "Controller PE"
    If Not Dropdown_Dict.Exists("11E     K    DIT") Then Dropdown_Dict.Add "11E     K    DIT", "DISOCONT Tersus"
    If Not Dropdown_Dict.Exists("11E     K    ZEK") Then Dropdown_Dict.Add "11E     K    ZEK", "Electrical Accessory"
    If Not Dropdown_Dict.Exists("11M     K    HLX") Then Dropdown_Dict.Add "11M     K    HLX", "Helixes"
    If Not Dropdown_Dict.Exists("14M     F    HMS") Then Dropdown_Dict.Add "14M     F    HMS", "Horizontal Material Separator"
    If Not Dropdown_Dict.Exists("14M     F    HBD") Then Dropdown_Dict.Add "14M     F    HBD", "Hygienic Bag Dump Station"
    If Not Dropdown_Dict.Exists("14M     F    HCT") Then Dropdown_Dict.Add "14M     F    HCT", "Hygienic Collector (HCT)"
    If Not Dropdown_Dict.Exists("14M     F    HSER") Then Dropdown_Dict.Add "14M     F    HSER", "Hygienic Filter (HSER)"
    If Not Dropdown_Dict.Exists("14M     F    HRT") Then Dropdown_Dict.Add "14M     F    HRT", "Hygienic Round Top (HRT)"
    If Not Dropdown_Dict.Exists("11E     K    ISA") Then Dropdown_Dict.Add "11E     K    ISA", "Intecont Satus"
    If Not Dropdown_Dict.Exists("11E     K    ITE") Then Dropdown_Dict.Add "11E     K    ITE", "Intecont Tersus"
    If Not Dropdown_Dict.Exists("15M     K    KKS") Then Dropdown_Dict.Add "15M     K    KKS", "Kemutec Spare"
    If Not Dropdown_Dict.Exists("15M     K    KKU") Then Dropdown_Dict.Add "15M     K    KKU", "Kemutec Units"
    If Not Dropdown_Dict.Exists("11M     K    ZMK") Then Dropdown_Dict.Add "11M     K    ZMK", "Mechanic Feeding Accessories"
    If Not Dropdown_Dict.Exists("11M     K    MCT") Then Dropdown_Dict.Add "11M     K    MCT", "MechaTron"
    If Not Dropdown_Dict.Exists("11M     K    MCS") Then Dropdown_Dict.Add "11M     K    MCS", "Multicor"
    If Not Dropdown_Dict.Exists("11M     K    MSG") Then Dropdown_Dict.Add "11M     K    MSG", "Multistream Chute"
    If Not Dropdown_Dict.Exists("11M     K    NZL") Then Dropdown_Dict.Add "11M     K    NZL", "Nozzles"
    If Not Dropdown_Dict.Exists("11M     K    PFA") Then Dropdown_Dict.Add "11M     K    PFA", "PureFeed Auger"
    If Not Dropdown_Dict.Exists("11M     K    SAM") Then Dropdown_Dict.Add "11M     K    SAM", "SacMaster"
    If Not Dropdown_Dict.Exists("11M     K    SAV") Then Dropdown_Dict.Add "11M     K    SAV", "Series Feeder Volumetric"
    If Not Dropdown_Dict.Exists("11M     K    SFT") Then Dropdown_Dict.Add "11M     K    SFT", "Solidsflow Fiber"
    If Not Dropdown_Dict.Exists("11M     K    SFG") Then Dropdown_Dict.Add "11M     K    SFG", "Solidsflow Gravimetric"
    If Not Dropdown_Dict.Exists("11M     K    SFS") Then Dropdown_Dict.Add "11M     K    SFS", "Solidsflow Streamout"
    If Not Dropdown_Dict.Exists("11M     K    SFV") Then Dropdown_Dict.Add "11M     K    SFV", "Solidsflow Volumetric"
    If Not Dropdown_Dict.Exists("11M     F    SAEH") Then Dropdown_Dict.Add "11M     F    SAEH", "Supplied Air Extruder Hood"
    If Not Dropdown_Dict.Exists("14M     F    TUS") Then Dropdown_Dict.Add "14M     F    TUS", "Truck Unload System"
    If Not Dropdown_Dict.Exists("11M     K    DEA") Then Dropdown_Dict.Add "11M     K    DEA", "Weighfeeder DEA"
    If Not Dropdown_Dict.Exists("11M     K    DMO") Then Dropdown_Dict.Add "11M     K    DMO", "Weighfeeder DMO"
    If Not Dropdown_Dict.Exists("11Z     Z    ZUB") Then Dropdown_Dict.Add "11Z     Z    ZUB", "Others"
    
    If Not Dropdown_Dict.Exists("A111") Then Dropdown_Dict.Add "A111", "Warehouse (A111)"
    If Not Dropdown_Dict.Exists("A112") Then Dropdown_Dict.Add "A112", "System Assembly (A112)"
    If Not Dropdown_Dict.Exists("A113") Then Dropdown_Dict.Add "A113", "Feeder Assembly (A113)"
    If Not Dropdown_Dict.Exists("A114") Then Dropdown_Dict.Add "A114", "MS/Wind/Saw (A114)"
    If Not Dropdown_Dict.Exists("A115") Then Dropdown_Dict.Add "A115", "Sheet Metal (A115)"
    If Not Dropdown_Dict.Exists("A116") Then Dropdown_Dict.Add "A116", "Helix/Nozzle (A116)"
    If Not Dropdown_Dict.Exists("A117") Then Dropdown_Dict.Add "A117", "Shipping (A117)"
    If Not Dropdown_Dict.Exists("A118") Then Dropdown_Dict.Add "A118", "North (A118)"
    If Not Dropdown_Dict.Exists("A119") Then Dropdown_Dict.Add "A119", "Electrical Assembly (A119)"
    If Not Dropdown_Dict.Exists("A120") Then Dropdown_Dict.Add "A120", "Test Lab (A120)"
    If Not Dropdown_Dict.Exists("A121") Then Dropdown_Dict.Add "A121", "Hopper Molding (A121)"
    If Not Dropdown_Dict.Exists("A122") Then Dropdown_Dict.Add "A122", "SS Weld/Grind (A122)"
    If Not Dropdown_Dict.Exists("A123") Then Dropdown_Dict.Add "A123", "CS Weld/Painting (A123)"
    If Not Dropdown_Dict.Exists("A212") Then Dropdown_Dict.Add "A212", "Repair Center (A212)"
    
    If Not Dropdown_Dict.Exists("11A") Then Dropdown_Dict.Add "11A", "Purchased To Order (Ind)"
    If Not Dropdown_Dict.Exists("11B") Then Dropdown_Dict.Add "11B", "Make To Order (Ind)"
    
    ClearTableCPN
    SAP_Session.StartTransaction ("ZJ61")

    SAP_Window.FindById("usr/ctxtMARA-MATNR").Text = Material_Number_TextBox.Value
    SAP_Window.SendVKey (6)
    
    CreatePartNumbers.Range("C5").Value = Generated_Number_TextBox.Value 'Generated Material Number
    CreatePartNumbers.Range("E5").Value = Drawing_Number_TextBox.Value 'Drawing Number
    CreatePartNumbers.Range("G5").Value = SAP_Window.FindById("usr/txtZJ61_D1100-MAKTX1").Text 'Description
    CreatePartNumbers.Range("I5").Value = Dropdown_Dict.Item(SAP_Window.FindById("usr/ctxtZJ61_D1100-DISPO").Text) 'MRP Controller/Part Type
    CreatePartNumbers.Range("J5").Value = Dropdown_Dict.Item(SAP_Window.FindById("usr/ctxtZJ61_D1100-LGPRO").Text) 'Storage Location
    CreatePartNumbers.Range("K5").Value = Dropdown_Dict.Item(SAP_Window.FindById("usr/ctxtZJ61_D1100-PRODH").Text) 'Product Hierarchy
    CreatePartNumbers.Range("L5").Value = "Yes" 'Create Info Record
    
    SAP_Session.StartTransaction ("S000")
    
Skip_Material_Details:
    
    Sequence
    
    ClearTableCPN
    
    BOM_Creation.Activate
    
    New_Material_BOM_Creation SAP_Session
    
    BOM_Creation.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    ProcessDataBOM.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True

    End
    
SAP_NotLoggedIn:
    Debug.Print "NOT LOGGED IN ERROR"
    Debug.Print "Error Source: " & Err.Source & vbNewLine & _
                "Error Number: " & Err.Number & vbNewLine & _
                "Error Description: " & Err.Description & vbNewLine & _
                "Error Help Context: " & Err.HelpContext & vbNewLine & _
                "Error Last DLL Err: " & Err.LastDllError
    MsgBox "Please sign into SAP to continue", Title:="Error!"
    BOM_Creation.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    ProcessDataBOM.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFiltering:=True
    End
    
Press_Enter:
    Debug.Print "PRESS ENTER ERROR"
    Debug.Print "Error Source: " & Err.Source & vbNewLine & _
                "Error Number: " & Err.Number & vbNewLine & _
                "Error Description: " & Err.Description & vbNewLine & _
                "Error Help Context: " & Err.HelpContext & vbNewLine & _
                "Error Last DLL Err: " & Err.LastDllError
                
    SAP_Session.FindById("wnd[0]").SendVKey (0)
    Resume

End Sub

Private Sub Generate_Material_Number()

    Dim Material_Regex As RegExp: Set Material_Regex = New RegExp
    
    Material_Regex.Pattern = "[vVwW]\d{6}\.?[aAbB]\d{2}$"
    
    If Material_Regex.Test(Material_Number_TextBox) Then
        
        Dim Mechanical_Material_Number As String: Mechanical_Material_Number = Left(Material_Number_TextBox.Value, InStr(1, Material_Number_TextBox.Value, ".") - 1)
        Mechanical_Material_Number = Mechanical_Material_Number & Mechanical_B_Numbers(B_Number_ListBox.ListIndex)
        
        Dim Electrical_Material_Number As String: Electrical_Material_Number = Left(Material_Number_TextBox.Value, InStr(1, Material_Number_TextBox.Value, ".") - 1)
        Electrical_Material_Number = Electrical_Material_Number & Electrical_B_Numbers(B_Number_ListBox.ListIndex)
        
        Generated_Number_TextBox.Value = IIf(Department_ListBox.ListIndex, Mechanical_Material_Number, Electrical_Material_Number)
    
    Else
    
        Generated_Number_TextBox.Value = "INVALID"
        
    End If
End Sub


Sub New_Material_BOM_Creation(SAP_Session As GuiSession)
    Init_DNU_Dict
    
    SAP_Session.StartTransaction ("CS02")
    
    Dim SAP_Window As GuiMainWindow: Set SAP_Window = SAP_Session.FindById("wnd[0]")
    
    SAP_Window.FindById("usr/ctxtRC29N-MATNR").Text = Material_Number_TextBox.Value
    SAP_Window.SendVKey (V_F5)
    
    Dim SAP_Table As GuiTableControl: Set SAP_Table = SAP_Window.FindById(T_CS_Table)
    
    'Dictionaries for previous and new BOM
    Dim Existing_BOM_Dict As Dictionary: Set Existing_BOM_Dict = New Dictionary
    Dim New_BOM_Dict As Dictionary: Set New_BOM_Dict = New Dictionary
    
    'Total number of rows in the previous BOM table
    Dim Existing_Num_Rows As Long: Existing_Num_Rows = SAP_Window.FindByName(F_Table_Valid_Rows, F_Text_Field).Text
    
    'Visible rows on the screen
    Dim Visible_Row_Count As Long: Visible_Row_Count = SAP_Table.VisibleRowCount
    
    'Index for For loop
    Dim Existing_BOM_Index As Long
    Dim SAP_Table_Index As Long
    
    'SAP Component Collection to pull all visible rows of a column into
    Dim SAP_Material_Numbers_Collection As GuiComponentCollection
    Dim SAP_Material_Quantity_Collection As GuiComponentCollection
    
    'Add all new BOM items and quantity to dict
    For SheetRow = 7 To BOM_Creation.Cells(BOM_Creation.Rows.Count, "C").End(xlUp).Row
    
        If BOM_Creation.Cells(SheetRow, "C").Text = vbNullString Then GoTo Next_Sheet_Row

        BOM_Creation.Range("A7:E506").NumberFormat = "General"
    
        If Not New_BOM_Dict.Exists(BOM_Creation.Cells(SheetRow, "C").Text) Then
            New_BOM_Dict.Add BOM_Creation.Cells(SheetRow, "C").Text, BOM_Creation.Cells(SheetRow, "E").Text
        Else
            Dim NewValue2 As String: NewValue2 = CLng(New_BOM_Dict.Item(BOM_Creation.Cells(SheetRow, "C").Text)) + _
                                                CLng(BOM_Creation.Cells(SheetRow, "E").Text)
            New_BOM_Dict.Item(BOM_Creation.Cells(SheetRow, "C").Text) = NewValue2
        End If
Next_Sheet_Row:
    Next SheetRow

    'Set SAP Collections
    Set SAP_Material_Numbers_Collection = SAP_Table.FindAllByName(F_Material_Number, F_CText_Field)
    Set SAP_Material_Quantity_Collection = SAP_Table.FindAllByName(F_Material_Quantity, F_Text_Field)
    
    Dim Deletion_Required As Boolean: Deletion_Required = False
    
    'Add Existing BOM to a dictionary
    For Existing_BOM_Index = 0 To Existing_Num_Rows - 1
        'If Index is on first loop, reset table index
        If Existing_BOM_Index = 0 Then
        
            SAP_Table_Index = 0
            'Grab all Material Numbers and Quantity from rows visible on SAP Screen
            Set SAP_Material_Numbers_Collection = SAP_Table.FindAllByName(F_Material_Number, F_CText_Field)
            Set SAP_Material_Quantity_Collection = SAP_Table.FindAllByName(F_Material_Quantity, F_Text_Field)
            
        'If index has reached the last visible row, reset table index
        ElseIf Existing_BOM_Index = Visible_Row_Count Then
        
            SAP_Table_Index = 0
            Set_SAP_Table_Position_And_Refresh SAP_Session, SAP_Window, SAP_Table, T_CS_Table, SAP_Table.VerticalScrollbar.Position + Visible_Row_Count
            
            'Grab all Material Numbers and Quantity from rows visible on SAP Screen
            Set SAP_Material_Numbers_Collection = SAP_Table.FindAllByName(F_Material_Number, F_CText_Field)
            Set SAP_Material_Quantity_Collection = SAP_Table.FindAllByName(F_Material_Quantity, F_Text_Field)
        
        End If
        
        'If material number is a placeholder
        If DNU_Dict.Exists(SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text) Then GoTo Next_SAP_Line
        
        
        If Not Existing_BOM_Dict.Exists(SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text) Then
            Existing_BOM_Dict.Add SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text, SAP_Material_Quantity_Collection.ElementAt(SAP_Table_Index).Text
        Else
            Dim NewValue As String: NewValue = CLng(Existing_BOM_Dict.Item(SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text)) + _
                                                CLng(SAP_Material_Quantity_Collection.ElementAt(SAP_Table_Index).Text)
            Existing_BOM_Dict.Item(SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text) = NewValue
        End If
        
        If New_BOM_Dict.Exists(SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text) Then
            
            Deletion_Required = True
            SAP_Table.Rows(SAP_Table_Index).Selected = True
        
        End If
Next_SAP_Line:
        SAP_Table_Index = SAP_Table_Index + 1
    Next Existing_BOM_Index

    'Remove entries from existing bom dict if the new BOM material and quantity are the same
    For Each DictKey In New_BOM_Dict.Keys
        If Existing_BOM_Dict.Exists(DictKey) Then
            If Existing_BOM_Dict.Item(DictKey) = New_BOM_Dict.Item(DictKey) Then
                Existing_BOM_Dict.Remove (DictKey)
            Else
                Existing_BOM_Dict.Item(DictKey) = New_BOM_Dict.Item(DictKey) - Existing_BOM_Dict.Item(DictKey)
                If Existing_BOM_Dict.Item(DictKey) = 0 Then Existing_BOM_Dict.Remove (DictKey)
                
                
            End If
        End If
    Next
    
    'Delete .B01 entries that were present on the new BOM
    Set SAP_Window = Renew_Main_Window(SAP_Session)
    
    If Deletion_Required Then
    
        SAP_Window.SendVKey (14)
        Dim SAP_Popup As GuiModalWindow: Set SAP_Popup = SAP_Session.FindById("wnd[1]")
        SAP_Popup.FindById("usr/btnSPOP-OPTION1").Press
        
    End If

    Set_SAP_Table_Position_And_Refresh SAP_Session, SAP_Window, SAP_Table, T_CS_Table, CLng(SAP_Window.FindByName(F_Table_Valid_Rows, F_Text_Field).Text)
    Set SAP_Material_Numbers_Collection = SAP_Table.FindAllByName(F_Material_Number, F_CText_Field)
    Set SAP_Material_Quantity_Collection = SAP_Table.FindAllByName(F_Material_Quantity, F_Text_Field)
    
    SAP_Material_Numbers_Collection.ElementAt(0).Text = Generated_Number_TextBox.Value
    SAP_Material_Quantity_Collection.ElementAt(0).Text = "1"
    SAP_Window.SendVKey (V_Enter)
    
    'SAVE AND CS01
    SAP_Window.SendVKey (V_Save)
    SAP_Session.StartTransaction ("CS01")
    
    'Fill out CS01 Fields
    SAP_Window.FindByName(F_Material_Number_Entry, F_CText_Field).Text = Generated_Number_TextBox.Value
    SAP_Window.FindByName(F_Plant_Number_Entry, F_CText_Field).Text = "1111"
    SAP_Window.FindByName(F_BOM_Usage_Entry, F_CText_Field).Text = "1"
    SAP_Window.SendVKey (V_Enter)
    
    Set SAP_Table = Renew_Table(SAP_Session, T_CS_Table)
    
    Existing_Num_Rows = SAP_Window.FindByName(F_Table_Valid_Rows, F_Text_Field).Text
    Set_SAP_Table_Position_And_Refresh SAP_Session, SAP_Window, SAP_Table, T_CS_Table, Existing_Num_Rows
    
    Dim New_BOM_Index As Long
    Dim Total_Entered_BOM_Count As Long: Total_Entered_BOM_Count = Get_Last_Row(BOM_Creation, "C")
    Dim SAP_Material_Item_Numbers_Collection As GuiComponentCollection
    
    For New_BOM_Index = 7 To Total_Entered_BOM_Count
    
        If New_BOM_Index = 7 Then
            SAP_Table_Index = 0
            Set SAP_Material_Item_Numbers_Collection = SAP_Table.FindAllByName(F_Item_Number, F_Text_Field)
            Set SAP_Material_Numbers_Collection = SAP_Table.FindAllByName(F_Material_Number, F_CText_Field)
            Set SAP_Material_Quantity_Collection = SAP_Table.FindAllByName(F_Material_Quantity, F_Text_Field)
        ElseIf SAP_Table_Index = Visible_Row_Count Then
            SAP_Table_Index = 0
            
            SAP_Window.SendVKey (V_Enter)
            Set SAP_Window = Renew_Main_Window(SAP_Session)
            'Set SAP_Table = Renew_Table(SAP_Session, T_CS_Table)
            
            For i = 0 To Visible_Row_Count - 1
            
                If SAP_Window.FindById("titl").Text = CS01_Title Then Exit For
                SAP_Window.SendVKey (V_Enter)
                
            Next i
            
            Set SAP_Table = Renew_Table(SAP_Session, T_CS_Table)
            
            Set_SAP_Table_Position_And_Refresh SAP_Session, SAP_Window, SAP_Table, T_CS_Table, SAP_Table.VerticalScrollbar.Position + Visible_Row_Count
            
            Set SAP_Material_Item_Numbers_Collection = SAP_Table.FindAllByName(F_Item_Number, F_Text_Field)
            Set SAP_Material_Numbers_Collection = SAP_Table.FindAllByName(F_Material_Number, F_CText_Field)
            Set SAP_Material_Quantity_Collection = SAP_Table.FindAllByName(F_Material_Quantity, F_Text_Field)
        End If
        
        SAP_Material_Item_Numbers_Collection.ElementAt(SAP_Table_Index).Text = BOM_Creation.Cells(New_BOM_Index, "B").Text
        SAP_Material_Numbers_Collection.ElementAt(SAP_Table_Index).Text = BOM_Creation.Cells(New_BOM_Index, "C").Text
        SAP_Material_Quantity_Collection.ElementAt(SAP_Table_Index).Text = BOM_Creation.Cells(New_BOM_Index, "E").Text
        
        SAP_Table_Index = SAP_Table_Index + 1

    Next New_BOM_Index
    
    SAP_Window.SendVKey (V_Enter)
    
    For i = 0 To Visible_Row_Count - 1
    
        If SAP_Window.FindById("titl").Text = CS01_Title Then Exit For
        SAP_Window.SendVKey (V_Enter)
        
    Next i

    SAP_Window.SendVKey (V_Save)
    SAP_Window.SendVKey (V_Exit)

    MsgBox "Material Number: " & Generated_Number_TextBox.Text & " has been created." & vbNewLine & _
                "BOM Created for " & Generated_Number_TextBox.Text & "." & vbNewLine & _
                Generated_Number_TextBox.Text & " has been added to " & Material_Number_TextBox.Text & "'s BOM.", vbInformation, "Complete"
    
    BOM_Options Existing_BOM_Dict
End Sub

Sub BOM_Options(Report_Dict As Dictionary)
    
    Reporting_Process_Data Report_Dict
    
    If Report_Dict.Count = 0 Then Exit Sub
    
    Select Case Reporting_ListBox.ListIndex
    Case 0
        Reporting_Option_Email
    Case 1
        Custom_Location_Label.Caption = Environ("userprofile") & "\Downloads"
        Reporting_Option_Excel
    Case 2
        Reporting_Option_Excel
    Case Else
        Debug.Print "Error Number: " & Err.Number & vbNewLine & _
                    "Error Source: " & Err.Source & vbNewLine & _
                    "Error Description: " & Err.Description & vbNewLine & _
                    "Error Help Context: " & Err.HelpContext & vbNewLine & _
                    "Error Help File: " & Err.HelpFile
    End Select
    
    If Routing_ListBox.ListIndex Then
    
        Verify_Routing_Status CStr(Reporting_ListBox.ListIndex), Custom_Location_Label.Caption
        
    End If

End Sub

Sub Init_DNU_Dict()

    Set DNU_Dict = New Dictionary

    If Not DNU_Dict.Exists(vbNullString) Then DNU_Dict.Add vbNullString, vbNullString
    If Not DNU_Dict.Exists("V383639.B01") Then DNU_Dict.Add "V383639.B01", "(OBS)Freight - Prepaid and Add"
    If Not DNU_Dict.Exists("V383639.B02") Then DNU_Dict.Add "V383639.B02", "(OBS)FREIGHT - Prepaid / Included *"
    If Not DNU_Dict.Exists("V383639.B03") Then DNU_Dict.Add "V383639.B03", "Extended Warranty"
    If Not DNU_Dict.Exists("V383639.B04") Then DNU_Dict.Add "V383639.B04", "EXPORT"
    If Not DNU_Dict.Exists("V383639.B05") Then DNU_Dict.Add "V383639.B05", "INSURANCE"
    If Not DNU_Dict.Exists("V383639.B06") Then DNU_Dict.Add "V383639.B06", "RENTAL PURCHASE"
    If Not DNU_Dict.Exists("V383639.B07") Then DNU_Dict.Add "V383639.B07", "MATERIAL CERTIFICATION"
    If Not DNU_Dict.Exists("V383639.B08") Then DNU_Dict.Add "V383639.B08", "EXPEDITE FEE"
    If Not DNU_Dict.Exists("V383639.B09") Then DNU_Dict.Add "V383639.B09", "MATERIAL TESTING"
    If Not DNU_Dict.Exists("V383639.B10") Then DNU_Dict.Add "V383639.B10", "INSPECTION UL/CUL/CSA"
    If Not DNU_Dict.Exists("V383639.B11") Then DNU_Dict.Add "V383639.B11", "HANDLING CHARGE"
    If Not DNU_Dict.Exists("V383639.B12") Then DNU_Dict.Add "V383639.B12", "RESTOCKING FEE"
    If Not DNU_Dict.Exists("V383639.B13") Then DNU_Dict.Add "V383639.B13", "FACTORY ACCEPTANCE TESTING"
    If Not DNU_Dict.Exists("V383639.B14") Then DNU_Dict.Add "V383639.B14", "DISCOUNT ADJUSTMENT"
    If Not DNU_Dict.Exists("V383639.B15") Then DNU_Dict.Add "V383639.B15", "IQ / OQ DOCUMENTATION"
    If Not DNU_Dict.Exists("V383639.B16") Then DNU_Dict.Add "V383639.B16", "Manuals-add'l copies per equip line item"
    If Not DNU_Dict.Exists("V383639.B17") Then DNU_Dict.Add "V383639.B17", "Processing Fee"
    If Not DNU_Dict.Exists("V383639.B18") Then DNU_Dict.Add "V383639.B18", "Evaluation Fee"
    If Not DNU_Dict.Exists("V383639.B19") Then DNU_Dict.Add "V383639.B19", "TO BE SPECIFIED"
    If Not DNU_Dict.Exists("V383639.B20") Then DNU_Dict.Add "V383639.B20", "APPROVAL DRAWING"
    If Not DNU_Dict.Exists("V383639.B21") Then DNU_Dict.Add "V383639.B21", "CERTIFIED DRAWINGS"
    If Not DNU_Dict.Exists("V383639.B22") Then DNU_Dict.Add "V383639.B22", "STORAGE FEE"
    If Not DNU_Dict.Exists("V383639.B23") Then DNU_Dict.Add "V383639.B23", "DISPOSAL FEE"
    If Not DNU_Dict.Exists("V383639.B24") Then DNU_Dict.Add "V383639.B24", "CANCELLATION FEE"
    If Not DNU_Dict.Exists("V383639.B25") Then DNU_Dict.Add "V383639.B25", "SPECIAL PAINT"
    If Not DNU_Dict.Exists("V383639.B26") Then DNU_Dict.Add "V383639.B26", "Extended Payment Terms"
    If Not DNU_Dict.Exists("V383639.B27") Then DNU_Dict.Add "V383639.B27", "CONTROL NOT REQUIRED"

End Sub

Sub Reporting_Process_Data(Report_Dict As Dictionary)
    
    If Report_Dict.Count = 0 Then Exit Sub
    ProcessDataBOM.Range("M7:N506").ClearContents
    
    Dim Process_Data_BOM_Row As Long: Process_Data_BOM_Row = 7
    
    For Each DictKey In Report_Dict
    
        ProcessDataBOM.Cells(Process_Data_BOM_Row, "M").Value = DictKey
        ProcessDataBOM.Cells(Process_Data_BOM_Row, "N").Value = Report_Dict.Item(DictKey)
        Process_Data_BOM_Row = Process_Data_BOM_Row + 1
    
    Next DictKey

End Sub

Sub Reporting_Option_Email()

    Dim EmailApp As Outlook.Application: Set EmailApp = New Outlook.Application
    Dim EmailItem As Outlook.MailItem: Set EmailItem = EmailApp.CreateItem(olMailItem)
    
    Dim Email_Text As String: Email_Text = "<!DOCTYPE html><html><head><style>table{border-collapse:collapse;}" & _
        "tr{border-bottom:1px solid #ddd;}th{padding-right:1em;}td{padding-right:1em;}</style></head>" & _
        "<body><table><tr><th>MATERIAL NUMBER</th>" & _
        "<th>MATERIAL DESCRIPTION</th>" & "<th>QUANTITY DIFFERENCE</th></tr>"
                                    
    Dim Planner_Index As Long
    For Planner_Index = 7 To ProcessDataBOM.Cells(ProcessDataBOM.Rows.Count, "M").End(xlUp).Row

        Dim MaterialNumber As String: MaterialNumber = ProcessDataBOM.Cells(Planner_Index, "M").Value
        Dim MaterialDescription As String: MaterialDescription = ProcessDataBOM.Cells(Planner_Index, "O").Value
        Dim QuantityValue As String: QuantityValue = ProcessDataBOM.Cells(Planner_Index, "N").Value
        Dim TableRow As String: TableRow = "<tr><td>" & MaterialNumber & "</td>" & _
            "<td>" & MaterialDescription & "</td>" & _
            "<td>" & QuantityValue & "</td></tr>"
                                               
        Email_Text = Email_Text & TableRow
    
    Next Planner_Index
    
    Email_Text = Email_Text & "</table></body></html>"
    
    With EmailItem
        .To = Environ("Username") & "@schenckprocess.com"
        .Subject = "BOM DISCREPANCY REPORT: " & BOM_Creation.Range("C4")
        .HTMLBody = Email_Text
        .Send
    End With

End Sub

Sub Reporting_Option_Excel()

    Dim Planner_Report_WB As Workbook: Set Planner_Report_WB = Workbooks.Add
    Dim Planner_Report_WS As Worksheet: Set Planner_Report_WS = Planner_Report_WB.ActiveSheet
    
    Dim Report_Name As String: Report_Name = Material_Number_TextBox.Value & " - Material Discrepancy Report.xlsx"
    Dim Report_Location As String: Report_Location = Custom_Location_Label.Caption
    Dim Planner_WS_Current_Index As Long: Planner_WS_Current_Index = 2
    
    With Planner_Report_WS
    
        .Range("A1").Value = "MATERIAL NUMBER"
        .Range("B1").Value = "MATERIAL DESCRIPTION"
        .Range("C1").Value = "QUANTITY DIFFERENCE"
        
        Dim Planner_Index As Long
        For Planner_Index = 7 To ProcessDataBOM.Cells(ProcessDataBOM.Rows.Count, "M").End(xlUp).Row
        
            .Cells(Planner_WS_Current_Index, "A").Value = ProcessDataBOM.Cells(Planner_Index, "M").Value
            .Cells(Planner_WS_Current_Index, "B").Value = ProcessDataBOM.Cells(Planner_Index, "O").Value
            .Cells(Planner_WS_Current_Index, "C").Value = ProcessDataBOM.Cells(Planner_Index, "N").Value
            
            Planner_WS_Current_Index = Planner_WS_Current_Index + 1

        Next Planner_Index
        
        .Range("A1:C100").Borders.LineStyle = xlContinuous
        .Range("A1:C100").Columns.AutoFit
    
    End With
    
    Planner_Report_WB.SaveAs Custom_Location_Label.Caption & "\" & Report_Name, xlWorkbookDefault
    Planner_Report_WB.Close

End Sub


Function Get_Last_Row(ByRef Entered_Sheet As Worksheet, ByVal Entered_Column As String) As Long

    Get_Last_Row = Entered_Sheet.Cells(Entered_Sheet.Rows.Count, Entered_Column).End(xlUp).Row

End Function
