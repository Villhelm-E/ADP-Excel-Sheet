Option Explicit

'UserForm GLobals
Public ListingMode As String
Public Incomplete As Boolean

Private Sub UserForm_Initialize()

    'position the userform
    Call CenterForm(ListAmazon)
    
    'count number of pages and save to global variable PageCount
    PageCount = MultiPage1.Pages.count - 1  'subtract 1 to account for first page being 0
    
    'rename userform caption
    ListAmazon.Caption = "List on Amazon (" & MultiPage1.SelectedItem.Caption & ")"
    
    'Font size
    Me.NextPage.Font.Size = 11

    'disable back button
    Me.PreviousPage.Enabled = False
    
    'disable buttons for unfinished code
    Me.SetListing.Enabled = False
    Me.ExistingSingleListing.Enabled = False
    Me.UpdateListing.Enabled = False
    
    'go to first page
    MultiPage1.value = 0
    
    'hide tabs from multipage, to look cleaner
    MultiPage1.Style = fmTabStyleNone
    
    'populate combo boxes
    ComboBoxes
    
    'Assume the part is New
    Me.condition_type.value = "New"
    
    'Assume part type is not spark plug
    Me.PlugType.Enabled = False
    Me.PlugTypeLabel.Enabled = False
    
    'Assume the part is vehicle specific
    Me.fit_type.value = "Vehicle Specific"
    
'    'Add controls to test page
'    AddControls
    
    'Adds asterisk to required fields
    RequiredFields

End Sub

Private Sub MultiPage1_Change()

    'decide when to enable or disable the Previous page button
    If MultiPage1.value <= 0 Then
        Me.PreviousPage.Enabled = False
    Else
        Me.PreviousPage.Enabled = True
    End If
    
    'change Next button to List button on the last page
    If MultiPage1.value >= PageCount Or MultiPage1.value = 0 Then   'PageCount is global variable
        Me.NextPage.Enabled = False
    Else
        Me.NextPage.Enabled = True
    End If
    
'    'Determine which pages to show based on the type of listing the user chose
'    Select Case ListingMode
'        Case "Bundle"
'            Dim i As Integer
'
'            'hide all pages
'            For i = 1 To PageCount - 1
'                MultiPage1.Pages(i).Visible = False
'            Next i
'
'            'make the Bundle page the only visible page
'            MultiPage1.Pages(4).Visible = True
'            MultiPage1.Value = 4
'
'        Case Else
'            'hide all pages
'            For i = 1 To PageCount
'                MultiPage1.Pages(i).Visible = True
'            Next i
'
'            MultiPage1.Pages(4).Visible = False
'    End Select
    
    'rename userform caption based on page caption
    ListAmazon.Caption = "List on Amazon (" & MultiPage1.SelectedItem.Caption & ")"
    
    'If user is on Part page run PartInfoSub
    If ListAmazon.Caption Like "*Part*" Then
        PartInfoSub
    End If
    
    'If user is on Parentage page run
    If ListAmazon.Caption Like "*Parentage*" Then
        Me.parent_sku.value = GenSKU(True)
    End If
    
    'If user is on Review page run ReviewPageSub
    If ListAmazon.Caption Like "*Review*" Then
        ReviewPageSub
    End If
    
    'If user is on second page change the caption of the back/cancel button
    If MultiPage1.value = 1 Then
        Me.PreviousPage.Caption = "Start Over"
    Else
        Me.PreviousPage.Caption = "Back"
    End If
    
    'check if parentage is selected
    ParentagePageCheck

End Sub

Private Sub brand_name_Change()

    If Me.brand_name.value <> "" Then
        Me.BrandLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

'Private Sub AddControls()
'
'    'ProgIDs for adding controls with VBA
'    'CheckBox           Forms.CheckBox.1
'    'ComboBox           Forms.ComboBox.1
'    'CommandButton      Forms.CommandButton.1
'    'Frame              Forms.Frame.1
'    'Image              Forms.Image.1
'    'Label              Forms.Label.1
'    'ListBox            Forms.ListBox.1
'    'MultiPage          Forms.MultiPage.1
'    'OptionButton       Forms.OptionButton.1
'    'ScrollBar          Forms.ScrollBar.1
'    'SpinButton         Forms.SpinButton.1
'    'TabStrip           Forms.TabStrip.1
'    'TextBox            Forms.TextBox.1
'    'ToggleButton       Forms.ToggleButton.1
'
'    'Control Variable
'    Dim AddedControl As control
'
'    'Variables for Adding Controls
'    Dim TextBoxes As msforms.TextBox
'    Dim Labels As msforms.LABEL
'    Dim CheckBoxes As msforms.CheckBox
'    Dim ComboBoxes As msforms.ComboBox
'    Dim CommandButtons As msforms.CommandButton
'
'    'Add Controls
'    Set TextBoxes = MultiPage1.Pages(3).Controls.Add("Forms.TextBox.1", "Part1", True)
'    Set Labels = MultiPage1.Pages(3).Controls.Add("Forms.Label.1", "Part1Label", True)
'    Set TextBoxes = MultiPage1.Pages(3).Controls.Add("Forms.TextBox.1", "Part2", True)
'
'    MultiPage1.Pages(3).Part1.left = 132
'    MultiPage1.Pages(3).Part1.top = 18
'    MultiPage1.Pages(3).Part1.width = 96
'    MultiPage1.Pages(3).Part1.height = 18
'
'
'
'    '
'    For Each AddedControl In MultiPage1.Pages(3).Controls
'        If AddedControl.name = "Part1" Then
'            AddedControl.Left = 18
'            AddedControl.Top = 18
'            AddedControl.Width = 96
'            AddedControl.Height = 18
'        End If
'    Next AddedControl
'
'    For Each AddedControl In MultiPage1.Pages(3).Controls
'        If AddedControl.name = "Part1Label" Then
'            AddedControl.Caption = "Label"
'            AddedControl.Left = 132
'            AddedControl.Top = 18
'            AddedControl.Width = 96
'            AddedControl.Height = 18
'        End If
'    Next AddedControl
'
'    For Each AddedControl In MultiPage1.Pages(3).Controls
'        If AddedControl.name = "Part2" Then
'            AddedControl.Left = 18
'            AddedControl.Top = 54
'            AddedControl.Width = 96
'            AddedControl.Height = 18
'        End If
'    Next AddedControl
'
'End Sub

Private Sub condition_type_Change()

    'check if condition is blank
    If Me.condition_type.value = "" Then
        Me.condition_note.value = ""
        Me.condition_note.Enabled = False
        Me.NoteLabel.Enabled = False
        
        Me.ConditionLabel.ForeColor = RGB(255, 0, 0)
    Else
        If Me.condition_type.value = "New" Then
            Me.condition_note.value = ""
            Me.condition_note.Enabled = False
            Me.NoteLabel.Enabled = False
        Else
            Me.condition_note.Enabled = True
            Me.NoteLabel.Enabled = True
        End If
        
        Me.ConditionLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub external_product_id_AfterUpdate()
    
    If IsNumeric(Me.external_product_id.value) = True And Len(Me.external_product_id.value) = 12 Then
        Me.ProductIDLabel.ForeColor = RGB(0, 0, 0)
    Else
        If Me.external_product_id.value <> "" Then
            MsgBox ("Not a valid product ID")
        End If
        Me.ProductIDLabel.ForeColor = RGB(255, 0, 0)
    End If

End Sub

Private Sub fit_type_Change()

    If Me.fit_type.value = "" Then
        Me.FitmentTypeLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.FitmentTypeLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub ParentageCheckBox_Change()

    ParentagePageCheck

End Sub

Private Sub ParentagePageCheck()

    Dim parentagePage As Integer
    
    parentagePage = 2   '2nd page of multipage
    
    'Parentage page in Multipage should only appear if user checked the variations checkbox
    If Me.ParentageCheckBox = True Then
        Me.MultiPage1.Pages(parentagePage).Visible = True
        Me.item_package_quantity.value = ""
        Me.item_package_quantity.Enabled = False
        Me.QuantityLabel.Enabled = False
    Else
        Me.MultiPage1.Pages(parentagePage).Visible = False
        Me.item_package_quantity.Enabled = True
        Me.QuantityLabel.Enabled = True
    End If

End Sub

Private Sub part_number_Change()

    If Me.part_number.value <> "" Then
        Me.PartNumLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub Manufacturer_Change()

    If Me.Manufacturer.value <> "" Then
        Me.ManufacturerLabel.ForeColor = RGB(0, 0, 0)
        If Me.Manufacturer.value <> "AD Auto Parts" Then
            Me.reboxed.Enabled = True
        Else
            Me.reboxed = False
            Me.reboxed.Enabled = False
        End If
    End If

End Sub

Private Sub item_package_quantity_Change()

    If Me.item_package_quantity.value = "" Or IsNumeric(Me.item_package_quantity.value) = False Then
        Me.QuantityLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.QuantityLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub part_type_id_Change()

    'if part type is Spark Plug, then enable the spark plug type combo box, otherwise disable
    If Me.part_type_id.value = "Spark Plug" Then
        Me.PlugType.Enabled = True
        Me.PlugTypeLabel.Enabled = True
    Else
        If Me.part_type_id.value = "" Then
            Me.PartTypeLabel.ForeColor = RGB(255, 0, 0)
        Else
            Me.PartTypeLabel.ForeColor = RGB(0, 0, 0)
        End If
        
        Me.PlugType.value = ""
        Me.PlugType.Enabled = False
        Me.PlugTypeLabel.Enabled = False
    End If

End Sub

Private Sub reboxed_Change()

    If Me.reboxed = True Then
        Me.legal_disclaimer_description.value = "Bulk packed. Not packaged in original manufacturer packaging. Packaged in AD Auto Parts packaging."
    ElseIf Me.reboxed = False Then
        Me.legal_disclaimer_description.value = Replace(Me.legal_disclaimer_description.value, "Bulk packed. Not packaged in original manufacturer packaging. Packaged in AD Auto Parts packaging.", "")
        If left(Me.legal_disclaimer_description.value, 1) = " " Then Me.legal_disclaimer_description.value = Right(Me.condition_note.value, Len(Me.condition_note.value) - 1)
    End If

End Sub

Private Sub standard_price_Change()

    If Me.standard_price.value = "" Or IsNumeric(Me.standard_price.value) = False Then
        Me.PriceLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.PriceLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub merchant_shipping_group_name_Change()

    If Me.merchant_shipping_group_name.value = "" Then
        Me.ShippingTemplateLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.ShippingTemplateLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub ListButton_Click()
    
    Dim AmazonFields() As Variant
    Dim lastcolumnletter As String
    ReDim AmazonFields(0) As Variant
    
    'Headers
    If CheckAmazonTemplate = False Then Call AmazonHeaders     'Amazon module
    'Save the number and letter of the last column
    lastcolumnletter = NumberToColumn(CountColumns(range("2:2")))
    
    'check if set
    Select Case ListingMode
        Case "ADP"
            If Me.ParentageCheckBox = False Then
                Call ListSingle(CountRows("A:A") + 1, lastcolumnletter, False)
            Else
                Call ListSets(lastcolumnletter)
            End If
        
        Case "Bundle"
            'Call ListBundle(lastcolumnletter)
        
        Case "Brand"
            
        
        Case "Update"
            
        
    End Select
    
    Unload Me

End Sub

Private Sub NewSingleListing_Click()

    ListingMode = "ADP"
    MultiPage1.value = MultiPage1.value + 1

End Sub

Private Sub SetListing_Click()

    ListingMode = "Bundle"
    MultiPage1.value = MultiPage1.value + 1

End Sub

Private Sub ExistingSingleListing_Click()

    ListingMode = "Brand"

End Sub

Private Sub UpdateListing_Click()

    ListingMode = "Update"

End Sub

Private Sub PartInfoSub()

    If ListingMode = "ADP" Then
        
        'disable Brand becuase it has to be AD Auto Parts and enable Product ID if single
        Me.brand_name.value = ""
        Me.brand_name.Enabled = False
        Me.BrandLabel.Enabled = False
        Me.external_product_id.Enabled = True
        
        Me.update_delete.value = ""
        
        'Access the Master Database and query the next available GTIN/UPC
        Dim UPC As Recordset
        Set UPC = MstrDb.Execute("SELECT * FROM GTINs WHERE SKU Is Null AND User Is Null AND DateReserved Is Null ORDER BY GTIN")
        UPC.MoveFirst
        
        'fill in the next available GTIN as a UPC in external_product_id box
        Me.external_product_id.value = Right(UPC.fields("GTIN").value, 12)
        
        UPC.Close
        
    Else
        Me.brand_name.Enabled = True
        Me.BrandLabel.Enabled = True
        Me.external_product_id.Enabled = False
    End If
    
End Sub

Private Sub ReviewPageSub()

    'Generate SKU
    If Me.part_number.value <> "" And Me.part_type_id.value <> "" And Me.Manufacturer.value <> "" Then
        'Generate SKU
        GenSKU
        
        'Generate Description
        GenDesc
        
        'Generate Title
        Me.item_name = GenTitle(Me.item_package_quantity.value, Me.Manufacturer.value, Me.part_number.value, Me.part_type_id.value, Me.PlugType.value, "", Me.oem_equivalent_part_number1.value, "", Me.ParentageCheckBox.value)
    End If

End Sub

Private Function GenSKU(Optional parent As Boolean) As String

    'Get the Prefix Code from Part Type
    Dim prefix As String
    Set rst = MstrDb.Execute("SELECT DISTINCT * FROM PartTypes WHERE ACESPartType=" & Chr(34) & Me.part_type_id.value & Chr(34))
    rst.MoveFirst
    prefix = rst.fields("PrefixCode").value
    rst.Close
    
    'Get the Suffix Code from Manufacturer
    Dim suffix As String
    Set rst = MstrDb.Execute("SELECT DISTINCT * FROM Manufacturers WHERE ManufacturerFull=" & Chr(34) & Me.Manufacturer.value & Chr(34))
    suffix = rst.fields("SuffixCode").value
    rst.Close
    
    If parent = True Then suffix = suffix & "-P"
    
    'Generate SKU
    GenSKU = prefix & "-" & Me.part_number & "-" & suffix
    Me.item_sku.value = GenSKU

End Function

Private Sub GenDesc()

    Select Case Me.ParentageCheckBox
        Case False
            Me.product_description.value = "This item is designed to be an exact replacement that meets or exceeds original specifications. Please ensure correct part fitment before purchasing this product. Contact the seller directly for additional product information and availability."
        Case True
            Me.product_description.value = "This item is designed to be an exact replacement that meets or exceeds original specifications. Please ensure correct part fitment before purchasing this product. Contact the seller directly for additional product information and availability."
    End Select

End Sub

Private Function GenTitle(quantity As String, Manufacturer As String, partNum As String, partType As String, PlugType As String, fits As String, equivPart As String, equivBrand As String, isSet As Boolean, Optional SetArr, Optional listingrow As Integer) As String

    Dim name As String
    
    'Single Listing
    If Me.ParentageCheckBox = False Then
        'non single title
        If quantity > 1 Then
            name = Manufacturer & " " & partType & " For _ Compatible with _ " & Me.oem_equivalent_part_number1.value
        Else
            'sets
            If quantity > 1 Then
                name = "Set of " & quantity & " " & partType & "s For _ Compatible with _ " & Me.oem_equivalent_part_number1.value
            Else
                'single
                name = partType & " For _ Compatible with _ " & Me.oem_equivalent_part_number1.value
            End If
        End If
    Else
        'parent title
        If IsMissing(SetArr) = False Then quantity = Replace(SetArr(listingrow), "Setof", "")
        
        If quantity <> "" Then
            'sets
            If quantity > 1 Then
                name = "Set of " & quantity & " " & partType & "s For _ Compatible with _ " & Me.oem_equivalent_part_number1.value
            Else
                'single
                name = partType & " For _ Compatible with _ " & Me.oem_equivalent_part_number1.value
            End If
        End If
    End If
    
    GenTitle = name

End Function

Private Sub ComboBoxes()

    FillPartTypes
    
    FillPlugTypes
    
    FillManufs
    
    FillBrands
    
    FillOrientations
    
    FillFitmentTypes
    
    FillConditions
    
    FillShippingTemplates
    
    FillUpdateDelete
    
    FillWeightUnits

End Sub

Private Sub FillPartTypes()

    Dim rst As Recordset
    
    'open Part Types table in Master Database
    'If the part you want to list is not showing up, you need to add it in the Master Database into the PartTypes table
    Set rst = MstrDb.Execute("SELECT * FROM PartTypes ORDER BY ACESPartType")
    
    'populate combobox with part types
    With Me.part_type_id
        .Clear
        Do
            .AddItem rst.fields("ACESPartType").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillPlugTypes()

    Dim rst As Recordset
    
    'open Spark Plug Type table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM SparkPlugTypes ORDER BY Type")
    
    'populate combobox with part types
    With Me.PlugType
        .Clear
        Do
            .AddItem rst.fields("Type").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillManufs()

    Dim rst As Recordset
    
    'open manufacturers table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM Manufacturers")
    
    'populate combobox with manufacturers
    With Me.Manufacturer
        .Clear
        Do
            .AddItem rst.fields("ManufacturerFull").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillBrands()

    Dim rst As Recordset
    
    'open manufacturers table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM Manufacturers")
    
    'populate combobox with manufacturers
    With Me.brand_name
        .Clear
        Do
            .AddItem rst.fields("ManufacturerFull").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillOrientations()

    Dim rst As Recordset
    
    'open orientations table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM Orientations")
    
    'populate combobox with orientations
    With Me.orientation
        .Clear
        Do
            .AddItem rst.fields("Orientation").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillFitmentTypes()

    Dim rst As Recordset
    
    'open fitmenttypes table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM FitmentTypes")
    
    'populate combobox with fitment types
    With Me.fit_type
        .Clear
        Do
            .AddItem rst.fields("FitmentType").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillConditions()

    Dim rst As Recordset
    
    'open conditions table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM Conditions")
    
    'populate combobox with conditions
    With Me.condition_type
        .Clear
        Do
            .AddItem rst.fields("Condition").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillShippingTemplates()

    Dim rst As Recordset
    
    'open shipping templates table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM ShippingTemplates")
    
    'populate combobox with shipping templates
    With Me.merchant_shipping_group_name
        .Clear
        Do
            .AddItem rst.fields("Shipping Template").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillUpdateDelete()

    Dim rst As Recordset
    
    'open update/delete table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM [Update/Delete]")
    
    'populate combobox with update/delete
    With Me.update_delete
        .Clear
        Do
            .AddItem rst.fields("Update/Delete").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub FillWeightUnits()

    Dim rst As Recordset
    
    'open weightunits table in Master Database
    Set rst = MstrDb.Execute("SELECT * FROM WeightUnits")
    
    'populate combobox with units
    With Me.website_shipping_weight_unit_of_measure
        .Clear
        Do
            .AddItem rst.fields("WeightUnit").value
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    rst.Close

End Sub

Private Sub NextPage_Click()

    'check for blank entries
    CheckEmptyRequiredFields
    
    If Incomplete = False Then
        'move to next page
        MultiPage1.value = MultiPage1.value + 1     'automatically skips disabled pages
    Else
        MsgBox ("Please fill out required fields.")
    End If

End Sub

Private Sub PreviousPage_Click()

    Dim ctr As control
    
    'if user goes back to the first page, clear all values, user chose to cancel listing
    If MultiPage1.value = 1 Then
        For Each ctr In Me.Controls
            If TypeName(ctr) = "ComboBox" Or TypeName(ctr) = "TextBox" Then
                ctr.value = ""
            End If
        Next ctr
    End If
    
    'check to see if any pages are disabled so that the back button works properly to skip them
    If MultiPage1.Pages(MultiPage1.value - 1).Visible = False Then
        MultiPage1.value = MultiPage1.value - 2     'need to go back 2 if page disabled
    Else
        'otherwise go back 1
        MultiPage1.value = MultiPage1.value - 1
    End If

End Sub

Private Sub RequiredFields()

    Dim cCont As control
    
    'Query the required fields from the AmazonTemplateFields table in Master Database
    Set rst = MstrDb.Execute("SELECT Label_Name FROM AmazonTemplateFields Where Required = True ORDER BY TemplateOrder")
    rst.MoveFirst
    
    'go through each field in the table
    Do While Not rst.EOF
        'for each field we're going to loop through every control in the userform
        For Each cCont In Me.Controls
            'only focus on label controls
            If TypeName(cCont) = "Label" Then
                'if the field name is found in the label in the userform
                If InStr(1, cCont.Caption, rst.fields("Label_Name").value) > 0 Then
                    'if the field and the label are equal
                    If cCont.Caption = rst.fields("Label_Name").value Then
                        'add asterisk to the end to denote required field
                        cCont.Caption = cCont.Caption & "*"
                    End If
                    'end loop as soon as a match is found so no time is wasted looping through the rest of the controls
                    GoTo End_Loop
                End If
            End If
        Next cCont
End_Loop:
        rst.MoveNext
    Loop
    
    'close the recordset
    rst.Close

End Sub

Private Sub CheckEmptyRequiredFields()

    Dim cCont As control
    Dim savedCont As control
    Dim currentPage As Integer
    Dim currentControl As String
    
    'use current page in multi page
    currentPage = Me.MultiPage1.value
    
    'start by assuming all fields are filled
    Incomplete = False
    
    'Query the required fields from the AmazonTemplateFields table in Master Database
    Set rst = MstrDb.Execute("SELECT Label_Name, Field_Name FROM AmazonTemplateFields Where Required = True ORDER BY TemplateOrder")
    
    For Each cCont In Me.MultiPage1.Pages(currentPage).Controls
        If Not TypeName(cCont) = "Label" And Not TypeName(cCont) = "MultiPage" And Not TypeName(cCont) = "CommandButton" And Not TypeName(cCont) = "CheckBox" And Not TypeName(cCont) = "Nothing" Then
            rst.MoveFirst
            Do While Not rst.EOF
                If LCase(cCont.name) = LCase(rst.fields("Field_Name").value) Then
                    If cCont.Enabled = True Then
                        If cCont = "" Then
                            For Each savedCont In Me.MultiPage1.Pages(currentPage).Controls
                                If TypeName(savedCont) = "Label" Then
                                    If ListingMode = "Single" Then
                                        If left(savedCont.Caption, Len(savedCont.Caption) - 1) = rst.fields("Label_Name").value And savedCont.Caption <> "Brand Name" Then
                                            'if any field is not filled, return incomplete
                                            Incomplete = True
                                            savedCont.ForeColor = RGB(255, 0, 0)
                                            GoTo Exit_Loop
                                        End If
                                    End If
                                End If
                            Next savedCont
                        Else
                            'should this line of code do anything?
                            GoTo Exit_Loop
                        End If
                    Else
                        GoTo Exit_Loop
                    End If
                End If
                rst.MoveNext
            Loop
        End If
Exit_Loop:
    Next cCont
    
    'close the recordset
    rst.Close

End Sub

'listingRow is the row that will be populated in the sheet for that listing
'IsSet is used to list a single part number, if false, only a single listing is being listed, if true then multiple sets of that part number are being listed
'setArr and i are only needed for listing multiple sets of a single part number
Private Sub ListSingle(listingrow As Integer, lastcolumnletter As String, isSet As Boolean, Optional SetArr, Optional i As Integer)
    
    'Match the control name in the Form to the field name in the sheet and put the user-entered value into the correct cell
    Call EnterControls(lastcolumnletter, listingrow)
    
    'sku
    Call EnterSKU(lastcolumnletter, listingrow, isSet, SetArr, i)
    
    'Product Id type
    Call EnterProductIDType(lastcolumnletter, listingrow)
    
    Call EnterProductID(lastcolumnletter, listingrow)
    
    'Feed product type
    Call EnterFeedProductType(lastcolumnletter, listingrow)
    
    'Item Type Keyword
    Call EnterItemType(lastcolumnletter, listingrow)

    'Brand
    Call EnterBrand(lastcolumnletter, listingrow)
    
    'Manufacturer
    Call EnterManufacturer(lastcolumnletter, listingrow)
    
    'Price
    Call EnterPrice(lastcolumnletter, listingrow, isSet, SetArr, i)
    
    'Package Quantity
    Call EnterPackageQauntity(lastcolumnletter, listingrow)
    
    'Shipping Template
    Call EnterShippingTemplate(lastcolumnletter, listingrow, isSet, SetArr, i)
    
    'Is dicontinued by manufacturer
    Call EnterDiscontinued(lastcolumnletter, listingrow)
    
    'number of items
    Call EnterNumberofItems(lastcolumnletter, listingrow, isSet, SetArr, i)

    'quantity
    Call EnterQuantity(lastcolumnletter, listingrow)

    'product tax code
    Call EnterTaxCode(lastcolumnletter, listingrow)

    'handling time
    Call EnterHandlingTime(lastcolumnletter, listingrow)
    
    'Item Dimensions Unit of Measure
    Call EnterDimensionsUnitOfMeasure(lastcolumnletter, listingrow)
    
    'California Prop 65 Warning
    Call EnterProp65(lastcolumnletter, listingrow)

    'warranty
    Call EnterWarranty(lastcolumnletter, listingrow)
    
    'Overwrite fields (populated by EnterControls sub) that change for sets
    If isSet = True Then
        'Weight
        Call EnterWeight(lastcolumnletter, listingrow, SetArr, i)
        
        'package quantity
        If AmazonColumn(lastcolumnletter, "item_package_quantity") > 0 Then Cells(listingrow, AmazonColumn(lastcolumnletter, "item_package_quantity")).value = Replace(SetArr(i), "Setof", "")
        
        'Size Name
        Call EnterSizeName(lastcolumnletter, listingrow, SetArr, i)
        
        'overwrite part number
        If AmazonColumn(lastcolumnletter, "part_number") > 0 Then
            If SetArr(i) = "Setof1" Then
                Cells(listingrow, AmazonColumn(lastcolumnletter, "part_number")).value = Me.part_number.value
            Else
                Cells(listingrow, AmazonColumn(lastcolumnletter, "part_number")).value = Me.part_number.value & "-" & _
                Replace(SetArr(i), "Setof", "")
            End If
        End If
        
        'parentage
        If AmazonColumn(lastcolumnletter, "parent_child") > 0 Then Cells(listingrow, AmazonColumn(lastcolumnletter, "parent_child")).value = "child"
        
        'relationship type
        If AmazonColumn(lastcolumnletter, "relationship_type") > 0 Then Cells(listingrow, AmazonColumn(lastcolumnletter, "relationship_type")).value = "Variation"
        
        'variation theme
        If AmazonColumn(lastcolumnletter, "variation_theme") > 0 Then Cells(listingrow, AmazonColumn(lastcolumnletter, "variation_theme")).value = "sizeName"
        
        'title
        Call EnterTitle(lastcolumnletter, listingrow, SetArr, i)
        
        'description
        
        
        'price
        
        
        'retail price
        
        
        'item dimensions
        
        
    End If

End Sub

Private Sub ListSets(lastcolumnletter As String)

    Dim listingrow As Integer
    Dim SetArr()
    ReDim SetArr(0)
    
    'put sets user wants to list in an array
    Call SetsArray(SetArr)
    
    'count number of rows and add 1 to find the next blank row
    listingrow = CountRows("A:A") + 1
    
    DefineAmazonVariables
    
    Dim i As Integer
    
    'add parent and parentage field info after doing each set
    Call ListParent(lastcolumnletter, listingrow)
    
    'recount number of rows and add 1 to find the next blank row
    listingrow = CountRows("A:A") + 1
    
    'loop through every set, listing each like single
    For i = 0 To UBound(SetArr)
        Call ListSingle(listingrow + i, lastcolumnletter, True, SetArr, i)
    Next i

End Sub

Private Sub ListParent(lastcolumnletter As String, listingrow As Integer)

    'variable to match field to control
    Dim foundcolumn As Integer
    
    'look for SKU field and populate it with parent SKU user provided
    foundcolumn = AmazonColumn(lastcolumnletter, "item_sku")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Me.parent_sku.value
    
    'populate brand
    Call EnterBrand(lastcolumnletter, listingrow)
    
    'look for title field and populate it with autogenerated title
    foundcolumn = AmazonColumn(lastcolumnletter, "item_name")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "Parent title"
    
    'look for manufacturer field and populate it with manufacturer user provided
    foundcolumn = AmazonColumn(lastcolumnletter, "manufacturer")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Me.Manufacturer.value
    
    'look for part number field and populate it with autogenerated part number
    foundcolumn = AmazonColumn(lastcolumnletter, "part_number")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Me.part_number & "-P"
    
    'look for product type field and populate it with "autopart"
    Call EnterFeedProductType(lastcolumnletter, listingrow)
    
    'look for part type field and populate it with part type user provided
    foundcolumn = AmazonColumn(lastcolumnletter, "part_type_id")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Me.part_type_id.value
    
    'look for item type field and populate it with item type
    Call EnterItemType(lastcolumnletter, listingrow)
    
    'look for condition field and populate it with condition user provided
    foundcolumn = AmazonColumn(lastcolumnletter, "condition_type")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Me.condition_type.value
    
    'look for parentage field and populate it with "parent"
    foundcolumn = AmazonColumn(lastcolumnletter, "parent_child")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "parent"
    
    'look for variation theme field and populate it with "SizeName"
    foundcolumn = AmazonColumn(lastcolumnletter, "variation_theme")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "SizeName"

End Sub

Private Sub ListBundle(lastcolumnletter As String, listingrow As Integer, Optional SetArr)

    ListAmazon.Show

End Sub

Private Sub ListBrand(lastcolumnletter As String, listingrow As Integer, Optional SetArr)

    

End Sub

Private Sub SetsArray(SetArr)
    
    Dim cCont As control
    
    For Each cCont In Me.MultiPage1.Pages(2).Controls   '2 is the page number of Parentage page
        If TypeName(cCont) = "CheckBox" Then
            If cCont = True Then
                'add every checked box in Parentage page to array
                If UBound(SetArr) > 0 Or SetArr(0) <> "" Then ReDim Preserve SetArr(UBound(SetArr) + 1)
                SetArr(UBound(SetArr)) = cCont.name
            End If
        End If
    Next
    
End Sub

Private Sub EnterControls(lastcolumnletter As String, listingrow As Integer, Optional SetArr)

    Dim cCont As control
    Dim foundcolumn As Integer
    
    'loop through controls and populate header with values user entered
    For Each cCont In Me.Controls
        'filter out the type of controls we need
        If Not TypeName(cCont) = "Label" And Not TypeName(cCont) = "MultiPage" And Not TypeName(cCont) = "CommandButton" And Not TypeName(cCont) = "Nothing" Then
            'Find the column number of the Control
            
            foundcolumn = AmazonColumn(lastcolumnletter, cCont.name)
            
            'some items on the userform are not in the Amazon Template, like the Reboxed checkbox
            'if field name is not found, FoundColumn will return 0, and can't have a 0th column
            'add exception to external_product_id because it needs to be pulled from Master Database
            'add exception to manufacturer because it usually needs to be overridden
            If foundcolumn > 0 And cCont.name <> "external_product_id" And cCont.name <> "manufacturer" Then Cells(listingrow, foundcolumn).value = cCont.value
        End If
    Next cCont

End Sub

Private Sub EnterSKU(lastcolumnletter As String, listingrow As Integer, isSet As Boolean, SetArr, i As Integer)

    Dim foundcolumn As Integer
    foundcolumn = AmazonColumn(lastcolumnletter, "item_sku")
    
    'add set size to end of sku
    If isSet = True Then
        If SetArr(i) = "Setof1" Then
            Cells(listingrow, foundcolumn).value = Me.item_sku.value
        Else
            Cells(listingrow, foundcolumn).value = Me.item_sku.value & "-" & Replace(SetArr(i), "Setof", "")
        End If
    Else
        Cells(listingrow, foundcolumn).value = Me.item_sku.value
    End If

End Sub

Private Sub EnterProductIDType(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "external_product_id_type")
    
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "UPC"

End Sub

Private Sub EnterProductID(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    Dim setSKU
    
    'find column where product Id goes
    foundcolumn = AmazonColumn(lastcolumnletter, "external_product_id")
    
    Dim GTIN As String
    Dim ProductID As Recordset
    
    'get next availble UPC from GTINs table
    Set ProductID = MstrDb.Execute("SELECT * FROM GTINs WHERE SKU Is Null AND User Is Null and DateReserved Is Null ORDER BY GTIN")
    ProductID.MoveFirst
    
    'save GTIN to variable
    GTIN = ProductID.fields("GTIN").value
    ProductID.Close
    
    'remove the first towo digits from the GTIN to get UPC
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Right(GTIN, 12)
    
    'find the column where SKU is located
    foundcolumn = AmazonColumn(lastcolumnletter, "item_sku")
    If foundcolumn > 0 Then setSKU = Cells(listingrow, foundcolumn).value
    
    'find the user
    Dim User As String
    'grab user from ComputerUsersTable based on current computer being used
    Set ProductID = MstrDb.Execute("SELECT UserName FROM ComputerUsersTable WHERE ComputerName = " & Chr(34) & Environ$("computername") & Chr(34))
    User = ProductID.fields("UserName")
    ProductID.Close
    
    'update the GTINs table to reserve UPCs if not in debugging mode
    If DebugMode = False Then
        Set rst = MstrDb.Execute("UPDATE [GTINs] Set GTINs.SKU = " & Chr(34) & setSKU & Chr(34) & ", GTINs.User = " & Chr(34) & User & Chr(34) & ", GTINs.DateReserved = Now WHERE GTINs.GTIN = " & Chr(34) & GTIN & Chr(34))
    End If
    
End Sub

Private Sub EnterFeedProductType(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "feed_product_type")
    
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "autopart"

End Sub

Private Sub EnterItemType(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    'lookup the BTG value of the part type
    Set rst = MstrDb.Execute("SELECT BTGValue FROM AAIAPartTypes WHERE AAIAPartType=" & Chr(34) & Me.part_type_id.value & Chr(34))
    rst.MoveFirst
    
    'find the column in the excel sheet with "item_type" in the second row to enter the BTG value
    foundcolumn = AmazonColumn(lastcolumnletter, "item_type")
    
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = rst.fields("BTGValue").value
    
    'close the query
    rst.Close

End Sub

Private Sub EnterTitle(lastcolumnletter As String, listingrow As Integer, SetArr, i As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "item_name")
    
    If foundcolumn > 0 Then
        If Me.ParentageCheckBox.value = True Then
            'Cells(listingrow, foundcolumn).Value = GenTitle(Me.item_package_quantity.Value, Me.Manufacturer.Value, Me.part_number.Value, Me.part_type_id.Value, "", Me.oem_equivalent_part_number1.Value, "", Me.ParentageCheckBox.Value, SetArr, i)
            Cells(listingrow, foundcolumn).value = GenTitle(Me.item_package_quantity.value, Me.Manufacturer.value, Me.part_number.value, Me.part_type_id.value, Me.PlugType.value, "", Me.oem_equivalent_part_number1.value, "", Me.ParentageCheckBox, SetArr, i)
        Else
            Cells(listingrow, foundcolumn).value = GenTitle(Me.item_package_quantity.value, Me.Manufacturer.value, Me.part_number.value, Me.part_type_id.value, Me.PlugType.value, "", Me.oem_equivalent_part_number1.value, "", Me.ParentageCheckBox, SetArr, i)
        End If
    End If

End Sub

Private Sub EnterBrand(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    'brand
    Select Case ListingMode
        Case "Single", "ADP"
            foundcolumn = AmazonColumn(lastcolumnletter, "brand_name")
            
            If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "AD Auto Parts"
            
        Case "Brand"
            
            
        Case "Update"
            
            
    End Select

End Sub

Private Sub EnterManufacturer(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    'brand
    If ListingMode = "Single" Then
        foundcolumn = AmazonColumn(lastcolumnletter, "brand_name")
        
        If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "AD Auto Parts"
    Else
        'to be determined
        
    End If

End Sub

Private Sub EnterPrice(lastcolumnletter As String, listingrow As Integer, isSet As Boolean, SetArr, i As Integer)

    Dim foundcolumn As Integer
    foundcolumn = AmazonColumn(lastcolumnletter, "standard_price")
    
    If isSet = True Then
        Cells(listingrow, foundcolumn).value = Replace(SetArr(i), "Setof", "") * Me.standard_price.value
    Else
        Cells(listingrow, foundcolumn).value = Me.standard_price.value
    End If

End Sub

Private Sub EnterPackageQauntity(lastcolumnletter As String, listingrow As Integer, Optional SetListing As Boolean, Optional SetArr, Optional i As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "item_package_quantity")
    
    If ListingMode = "Single" Then
        'if SetListing is True, then user is listing sets. Calculate package quantity differently
        If SetListing = True Then
            'find out how to multiply this to find the number of items
            Cells(listingrow, foundcolumn).value = Replace(SetArr(i), "Setof", "")
        Else
            Cells(listingrow, foundcolumn).value = Me.item_package_quantity
        End If
    Else
        'placeholder
        
    End If

End Sub

Private Sub EnterShippingTemplate(lastcolumnletter As String, listingrow As Integer, isSet As Boolean, Optional SetArr, Optional i As Integer)
    
    Dim weight_oz As Double
    Dim shiptempcol As Integer
    
    'first find the column where the shipping template will be
    shiptempcol = AmazonColumn(lastcolumnletter, "merchant_shipping_group_name")
    
    'Enter the shipping template the user chose
    Cells(listingrow, shiptempcol).value = Me.merchant_shipping_group_name.value
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Code not needed anymore because it's impossible to predict what shipping template needs to be used
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    'if user entered "Weight-Based" for the shipping tmeplate, calculate and replace it with the correct template
'    If Cells(listingrow, shiptempcol).Value = "Weight-Based" Then
'        'convert weight to ounces and multiply by package qauntity to find total weight of listing
'        If Me.website_shipping_weight_unit_of_measure = "LB" Then
'            weight_oz = Me.website_shipping_weight.Value * 16 * Me.item_package_quantity.Value
'        Else
'            If IsSet = True Then
'                weight_oz = Me.website_shipping_weight.Value * Replace(SetArr(i), "Setof", "")
'            Else
'                weight_oz = Me.website_shipping_weight.Value * Me.item_package_quantity.Value
'            End If
'        End If
'
'        'choose appropriate Amazon Shipping Template based on weight
'        Select Case weight_oz
'        Case Is <= 13   '13 ounces or less
'            '13 ounce template
'            Cells(listingrow, shiptempcol).Value = "13 oz. Template"
'
'        Case Is <= 128  'between 13 oz and 8 lb.
'            '1-8 lb. Template
'            'round weight up to the next pound
'            If weight_oz > 16 Then
'                weight_oz = weight_oz / 16
'            Else
'                weight_oz = 1  'if weight is over 13 ounces but less than a pound, calculate template based on 1 pound
'            End If
'
'            'Concatenate the shipping template name
'            Cells(listingrow, shiptempcol).Value = RoundUp(weight_oz) & " lb. Template"
'
'        Case Is <= 160  'between 8 lb. and 10 lb.
'            '9-10 lb. Template
'            Cells(listingrow, shiptempcol).Value = "9-10 lb. Template"
'
'        Case Is <= 192  'between 10 and 12 lb.
'            '11-12 lb. Template
'            Cells(listingrow, shiptempcol).Value = "11-12 lb. Template"
'
'        Case Is <= 224  'between 12 and 14 lb.
'            '13-14 lb. Template
'            Cells(listingrow, shiptempcol).Value = "13-14 lb. Template"
'
'        Case Is <= 288  'between 14 and 18 lb.
'            '15-18 lb. Template
'            Cells(listingrow, shiptempcol).Value = "15-18 lb. Template"
'
'        Case Is <= 304  'between 18 and 19 lb.
'            '19 lb. Template
'            Cells(listingrow, shiptempcol).Value = "19 lb. Template"
'
'        Case Else   'over 19 pounds
'            '20-45 lb. Template
'            Cells(listingrow, shiptempcol).Value = "20-45 lb. Template"
'        End Select
'    Else
'        'If user entered Prime for shipping template, replace with Prime shipping template
'        Cells(listingrow, shiptempcol).Value = Me.merchant_shipping_group_name
'    End If

End Sub

Private Sub EnterDiscontinued(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "is_discontinued_by_manufacturer")
        
    'if field was found and the user left is_discontinued_by_manufacturer checkbox blank then change to null
    If Me.is_discontinued_by_manufacturer = False Then
        If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = ""
    End If

End Sub

Private Sub EnterNumberofItems(lastcolumnletter As String, listingrow As Integer, Optional SetListing As Boolean, Optional SetArr, Optional i As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "number_of_items")
    
    'if SetListing is True, then user is listing sets. Calculate number of items differently
    If SetListing = True Then
        'find out how to multiply this to find the number of items
        Cells(listingrow, foundcolumn).value = Replace(SetArr(i), "Setof", "") 'find column letter first
    Else
        foundcolumn = AmazonColumn(lastcolumnletter, "number_of_items")
        If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Me.item_package_quantity
    End If

End Sub

Private Sub EnterQuantity(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "quantity")
        
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = 1    'Finale takes care of quantity listed

End Sub

Private Sub EnterTaxCode(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "product_tax_code")
        
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "A_GEN_TAX"

End Sub

Private Sub EnterHandlingTime(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "fulfillment_latency")
        
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = 1   'should always be 1 to meet Amazon's standards

End Sub

Private Sub KeyProductFeatures(lastcolumnletter As String, listingrow As Integer)

    Dim i As Integer
    Set rst = MstrDb.Execute("SELECT * FROM KeyProductFeatures WHERE ([Manufacturer]=" & Chr(34) & Me.Manufacturer.value & Chr(34) & " AND [PartType]=" & Chr(34) & Me.part_type_id.value & Chr(34) & ")")
    rst.MoveFirst
    
    Dim KPFColumn As Integer
    
    KPFColumn = AmazonColumn(lastcolumnletter, "bullet_point1")
    
    For i = KPFColumn To KPFColumn + 4
        Cells(listingrow, i).value = rst.fields("KeyProductFeature" & i).value
    Next i

End Sub

Private Sub EnterDimensionsUnitOfMeasure(lastcolumnletter As String, listingrow As Integer)

    Dim LengthColumn As Integer
    Dim HeightColumn As Integer
    Dim WidthColumn As Integer
    Dim MeasureColumn As Integer
    
    'Find Item Dimension columns
    LengthColumn = AmazonColumn(lastcolumnletter, "item_length")
    HeightColumn = AmazonColumn(lastcolumnletter, "item_height")
    WidthColumn = AmazonColumn(lastcolumnletter, "item_width")
    MeasureColumn = AmazonColumn(lastcolumnletter, "item_dimensions_unit_of_measure")
    
    'If any item dimension is not null, the enter IN into the unit of measure field
    If Cells(listingrow, LengthColumn).value <> "" Or Cells(listingrow, HeightColumn).value <> "" Or Cells(listingrow, WidthColumn).value <> "" Then
        Cells(listingrow, MeasureColumn).value = "IN"
    End If

End Sub

Private Sub EnterWeight(lastcolumnletter As String, listingrow As Integer, SetArr, i As Integer)
    
    Dim SetSize As Integer
    SetSize = Replace(SetArr(i), "Setof", "")
    
    Dim WeightOz As Double
    Dim WeightLb As Double
    
    'find weight of a single
    If Me.website_shipping_weight_unit_of_measure.value = "LB" Then
        WeightOz = Me.website_shipping_weight.value * 16
        WeightLb = Me.website_shipping_weight.value
    Else
        WeightOz = Me.website_shipping_weight.value
        WeightLb = Me.website_shipping_weight.value / 16
    End If
    
    'enter the value
    Dim weightfield As Integer
    Dim unitfield As Integer
    weightfield = AmazonColumn(lastcolumnletter, "website_shipping_weight")
    unitfield = AmazonColumn(lastcolumnletter, "website_shipping_weight_unit_of_measure")
    
    'Enter the weight amount
    If weightfield > 0 Then
        If WeightLb * SetSize >= 1 Then
            'if weight is over a pound, use LB
            Cells(listingrow, weightfield).value = Round(WeightLb * SetSize, 2)
            Cells(listingrow, unitfield).value = "LB"
        Else
            'if weight is under a pound, use OZ
            Cells(listingrow, weightfield).value = Round(WeightOz * SetSize, 2)
            Cells(listingrow, unitfield).value = "OZ"
        End If
    End If

End Sub

Private Sub EnterProp65(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    foundcolumn = AmazonColumn(lastcolumnletter, "california_proposition_65_compliance_type")
        
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "Passenger or Off Road Vehicle"   'All our items are for Passenger or Off Road Vehicle

End Sub

Private Sub EnterWarranty(lastcolumnletter As String, listingrow As Integer)

    Dim foundcolumn As Integer
    
    'enter the warranty type
    foundcolumn = AmazonColumn(lastcolumnletter, "mfg_warranty_description_type")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "Parts"
    
    'enter the warranty description
    foundcolumn = AmazonColumn(lastcolumnletter, "warranty_description")
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = "Manufacturer warranty for 180 days from date of purchase, covers exchange of defective part while supplies last or a prorated return of defective part."

End Sub

Private Sub EnterSizeName(lastcolumnletter As String, listingrow As Integer, SetArr, i As Integer)

    Dim foundcolumn As Integer
    foundcolumn = AmazonColumn(lastcolumnletter, "size_name")
        
    If foundcolumn > 0 Then Cells(listingrow, foundcolumn).value = Replace(SetArr(i), "Setof", "Set of ")

End Sub

Private Sub website_shipping_weight_Change()

    If IsNumeric(Me.website_shipping_weight.value) = False Then
        Me.ShippingWeightLabel.ForeColor = RGB(255, 0, 0)
    End If

End Sub
