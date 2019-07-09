Option Explicit

'UserForm GLobals
Public ListingMode As String
Public Incomplete As Boolean

Private Sub brand_name_Change()

    If Me.brand_name.Value <> "" Then
        Me.BrandLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub condition_type_Change()

    'check if condition is blank
    If Me.condition_type.Value = "" Then
        Me.condition_note.Value = ""
        Me.condition_note.Enabled = False
        Me.NoteLabel.Enabled = False
        
        Me.ConditionLabel.ForeColor = RGB(255, 0, 0)
    Else
        If Me.condition_type.Value = "New" Then
            Me.condition_note.Value = ""
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
    
    If IsNumeric(Me.external_product_id.Value) = True And Len(Me.external_product_id.Value) = 12 Then
        Me.ProductIDLabel.ForeColor = RGB(0, 0, 0)
    Else
        If Me.external_product_id.Value <> "" Then
            MsgBox "Not a valid product ID"
        End If
        Me.ProductIDLabel.ForeColor = RGB(255, 0, 0)
    End If

End Sub

Private Sub fit_type_Change()

    If Me.fit_type.Value = "" Then
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
    
    parentagePage = 2
    
    'Parentage page in Multipage should only appear if user checked the variations checkbox
    If Me.ParentageCheckBox = True Then
        Me.MultiPage1.Pages(parentagePage).Visible = True
        Me.item_package_quantity.Value = ""
        Me.item_package_quantity.Enabled = False
        Me.QuantityLabel.Enabled = False
    Else
        Me.MultiPage1.Pages(parentagePage).Visible = False
        Me.item_package_quantity.Enabled = True
        Me.QuantityLabel.Enabled = True
    End If

End Sub

Private Sub part_number_Change()

    If Me.part_number.Value <> "" Then
        Me.PartNumLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub Manufacturer_Change()

    If Me.Manufacturer.Value <> "" Then
        Me.ManufacturerLabel.ForeColor = RGB(0, 0, 0)
        If Me.Manufacturer.Value <> "AD Auto Parts" Then
            Me.reboxed.Enabled = True
        Else
            Me.reboxed = False
            Me.reboxed.Enabled = False
        End If
    End If

End Sub

Private Sub item_package_quantity_Change()

    If Me.item_package_quantity.Value = "" Or IsNumeric(Me.item_package_quantity.Value) = False Then
        Me.QuantityLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.QuantityLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub part_type_id_Change()

    If Me.part_type_id.Value = "" Then
        Me.PartTypeLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.PartTypeLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub reboxed_Change()

    If Me.reboxed = True Then
        Me.legal_disclaimer_description.Value = "Bulk packed. Not packaged in original manufacturer packaging. Packaged in AD Auto Parts packaging."
    ElseIf Me.reboxed = False Then
        Me.legal_disclaimer_description.Value = Replace(Me.legal_disclaimer_description.Value, "Bulk packed. Not packaged in original manufacturer packaging. Packaged in AD Auto Parts packaging.", "")
        If Left(Me.legal_disclaimer_description.Value, 1) = " " Then Me.legal_disclaimer_description.Value = Right(Me.condition_note.Value, Len(Me.condition_note.Value) - 1)
    End If

End Sub

Private Sub standard_price_Change()

    If Me.standard_price.Value = "" Then
        Me.PriceLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.PriceLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub merchant_shipping_group_name_Change()

    If Me.merchant_shipping_group_name.Value = "" Then
        Me.ShippingTemplateLabel.ForeColor = RGB(255, 0, 0)
    Else
        Me.ShippingTemplateLabel.ForeColor = RGB(0, 0, 0)
    End If

End Sub

Private Sub ListButton_Click()
    
    Dim AmazonFields() As Variant
    ReDim AmazonFields(0) As Variant
    
    'Headers
    If CheckAmazonTemplate = False Then Call AmazonHeaders     'Amazon module
    
    'check if set
    Select Case ListingMode
        Case "Single"
            If Me.ParentageCheckBox = False Then
                ListSingle
            Else
                ListSets
            End If
        
        Case "Set"
            
        
        Case "Existing"
            
        
        Case "Update"
            
        
    End Select
    
    Unload Me

End Sub

Private Sub NewSingleListing_Click()

    ListingMode = "Single"
    MultiPage1.Value = MultiPage1.Value + 1

End Sub

Private Sub SetListing_Click()

    ListingMode = "Set"

End Sub

Private Sub ExistingSingleListing_Click()

    ListingMode = "Existing"

End Sub

Private Sub UpdateListing_Click()

    ListingMode = "Update"

End Sub

Private Sub UserForm_Initialize()

    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'count number of pages and save to global variable PageCount
    PageCount = MultiPage1.Pages.count - 1  'subtract 1 to account for first page being 0
    
    'rename userform caption
    ListAmazon.Caption = "List on Amazon (" & MultiPage1.SelectedItem.Caption & ")"
    
    'Font size
    Me.NextPage.Font.Size = 11

    'disable back button
    Me.PreviousPage.Enabled = False
    
    'go to first page
    MultiPage1.Value = 0
    
    'hide tabs from multipage, to look cleaner
    MultiPage1.Style = fmTabStyleNone
    
    'populate combo boxes
    comboBoxes
    
    'Assume the part is New
    Me.condition_type.Value = "New"
    
    'Assume the part is vehicle specific
    Me.fit_type.Value = "Vehicle Specific"
    
    RequiredFields

End Sub

Private Sub MultiPage1_Change()

    'decide when to enable or disable the Previous page button
    If MultiPage1.Value <= 0 Then
        Me.PreviousPage.Enabled = False
    Else
        Me.PreviousPage.Enabled = True
    End If
    
    'change Next button to List button on the last page
    If MultiPage1.Value >= PageCount Or MultiPage1.Value = 0 Then   'PageCount is global variable
        Me.NextPage.Enabled = False
    Else
        Me.NextPage.Enabled = True
    End If
    
    'rename userform caption based on page caption
    ListAmazon.Caption = "List on Amazon (" & MultiPage1.SelectedItem.Caption & ")"
    
    If ListAmazon.Caption Like "*Part*" Then
        PartInfoSub
    End If
    
    If ListAmazon.Caption Like "*Review*" Then
        ReviewPageSub
    End If
    
    If MultiPage1.Value = 1 Then
        Me.PreviousPage.Caption = "Cancel Listing"
    Else
        Me.PreviousPage.Caption = "Back"
    End If
    
    ParentagePageCheck

End Sub

Private Sub PartInfoSub()

    Dim UPC As Recordset
    Set UPC = MstrDb.Execute("SELECT * FROM Barcodes WHERE SKU Is Null AND User Is Null ORDER BY UPC")
    UPC.MoveFirst
    
    Me.external_product_id.Value = UPC.Fields("UPC").Value
    
    UPC.Close
    
    If ListingMode = "Single" Then
        Me.brand_name.Value = ""
        Me.brand_name.Enabled = False
        Me.BrandLabel.Enabled = False
        
        Me.update_delete.Value = ""
        Me.update_delete.Enabled = False
        Me.UpdateDeleteLabel.Enabled = False
    Else
        Me.brand_name.Enabled = True
        Me.BrandLabel.Enabled = True
    End If

End Sub

Private Sub ReviewPageSub()

    'Generate SKU
    If Me.part_number.Value <> "" And Me.part_type_id.Value <> "" And Me.Manufacturer.Value <> "" Then
        'Generate SKU
        GenSKU
        
        'Generate Description
        GenDesc
        
        'Generate Title
        GenTitle
    End If

End Sub

Private Sub GenSKU()

    'Get the Prefix Code from Part Type
    Dim prefix As String
    Set rst = MstrDb.Execute("SELECT DISTINCT * FROM PartTypes WHERE ACESPartType=" & Chr(34) & Me.part_type_id.Value & Chr(34))
    prefix = rst.Fields("PrefixCode").Value
    rst.Close
    
    'Get the Suffix Code from Manufacturer
    Dim suffix As String
    Set rst = MstrDb.Execute("SELECT DISTINCT * FROM Manufacturers WHERE ManufacturerFull=" & Chr(34) & Me.Manufacturer.Value & Chr(34))
    suffix = rst.Fields("SuffixCode").Value
    rst.Close
    
    'Generate SKU
    Me.item_sku.Value = prefix & "-" & Me.part_number & "-" & suffix

End Sub

Private Sub GenDesc()

    Select Case Me.ParentageCheckBox
        Case False
            Me.product_description.Value = Me.Manufacturer.Value & " " & Me.part_number.Value & " " & Me.part_type_id.Value
        Case True
            
    End Select

End Sub

Private Sub GenTitle()

    Me.item_name.Value = Me.item_package_quantity.Value & " " & Me.Manufacturer.Value & " " & Me.part_number.Value & " " & Me.part_type_id.Value

End Sub

Private Sub comboBoxes()

    FillPartTypes
    
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
            .AddItem rst.Fields("ACESPartType").Value
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
            .AddItem rst.Fields("ManufacturerFull").Value
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
            .AddItem rst.Fields("ManufacturerFull").Value
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
            .AddItem rst.Fields("Orientation").Value
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
            .AddItem rst.Fields("FitmentType").Value
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
            .AddItem rst.Fields("Condition").Value
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
            .AddItem rst.Fields("Shipping Template").Value
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
            .AddItem rst.Fields("Update/Delete").Value
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
            .AddItem rst.Fields("WeightUnit").Value
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
        MultiPage1.Value = MultiPage1.Value + 1     'automatically skips disabled pages
    Else
        MsgBox "Please fill out required fields."
    End If

End Sub

Private Sub PreviousPage_Click()

    Dim ctr As control
    
    'if user goes back to the first page, clear all values, user chose to cancel listing
    If MultiPage1.Value = 1 Then
        For Each ctr In Me.Controls
            If TypeName(ctr) = "ComboBox" Or TypeName(ctr) = "TextBox" Then
                ctr.Value = ""
            End If
        Next ctr
    End If
    
    'check to see if any pages are disabled so that the back button works properly to skip them
    If MultiPage1.Pages(MultiPage1.Value - 1).Visible = False Then
        MultiPage1.Value = MultiPage1.Value - 2     'need to go back 2 if page disabled
    Else
        'otherwise go back 1
        MultiPage1.Value = MultiPage1.Value - 1
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
                If InStr(1, cCont.Caption, rst.Fields("Label_Name").Value) > 0 Then
                    'if the field and the label are equal
                    If cCont.Caption = rst.Fields("Label_Name").Value Then
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
    currentPage = Me.MultiPage1.Value
    
    'start by assuming all fields are filled
    Incomplete = False
    
    'Query the required fields from the AmazonTemplateFields table in Master Database
    Set rst = MstrDb.Execute("SELECT Label_Name, Field_Name FROM AmazonTemplateFields Where Required = True ORDER BY TemplateOrder")
    
    For Each cCont In Me.MultiPage1.Pages(currentPage).Controls
        If Not TypeName(cCont) = "Label" And Not TypeName(cCont) = "MultiPage" And Not TypeName(cCont) = "CommandButton" And Not TypeName(cCont) = "CheckBox" And Not TypeName(cCont) = "Nothing" Then
            rst.MoveFirst
            Do While Not rst.EOF
                If LCase(cCont.Name) = LCase(rst.Fields("Field_Name").Value) Then
                    If cCont.Enabled = True Then
                        If cCont = "" Then
                            For Each savedCont In Me.MultiPage1.Pages(currentPage).Controls
                                If TypeName(savedCont) = "Label" Then
                                    If ListingMode = "Single" Then
                                        If Left(savedCont.Caption, Len(savedCont.Caption) - 1) = rst.Fields("Label_Name").Value And savedCont.Caption <> "Brand Name" Then
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

Private Sub ListSingle()

    Dim LastColumnLetter As String
    Dim ListingRow As Integer
    
    'count number of rows and add 1 to find the next blank row
    ListingRow = CountRows("A:A") + 1
    
    'Save the number and letter of the last column
    LastColumnLetter = NumberToColumn(CountColumns(Range("2:2")))
    
    'loop through controls and populate header with values user entered
    Call EnterControls(LastColumnLetter, ListingRow)
    
    'Product Id type
    Call EnterProductIDType(LastColumnLetter, ListingRow)
    
    'Feed product type
    Call EnterFeedProductType(LastColumnLetter, ListingRow)
    
    'Item Type Keyword
    Call EnterItemType(LastColumnLetter, ListingRow)

    'Brand
    Call EnterBrand(LastColumnLetter, ListingRow)
    
    'Shipping Template
    Call EnterShippingTemplate(LastColumnLetter, ListingRow)
    
    'Is dicontinued by manufacturer
    Call EnterDiscontinued(LastColumnLetter, ListingRow)
    
    'number of items
    Call EnterNumberofItems(LastColumnLetter, ListingRow)

    'quantity
    Call EnterQuantity(LastColumnLetter, ListingRow)

    'product tax code
    Call EnterTaxCode(LastColumnLetter, ListingRow)

    'handling time
    Call EnterHandlingTime(LastColumnLetter, ListingRow)
    
    'Item Dimensions Unit of Measure
    Call EnterDimensionsUnitOfMeasure(LastColumnLetter, ListingRow)
    
    'California Prop 65 Warning
    Call EnterProp65(LastColumnLetter, ListingRow)

    'warranty
    Call EnterWarranty(LastColumnLetter, ListingRow)

End Sub

Private Sub ListSets()

    Dim LastColumnLetter As String
    Dim ListingRow As Integer
    Dim SetArr()
    ReDim SetArr(0)
    
    'put sets user wants to list in an array
    Call SetsArray(SetArr)
    
    'count number of rows and add 1 to find the next blank row
    ListingRow = CountRows("A:A") + 1
    
    'Save the number and letter of the last column
    LastColumnLetter = NumberToColumn(CountColumns(Range("2:2")))
    
    DefineAmazonVariables
    
    Dim i As Integer
    
    For i = 1 To UBound(SetArr)
        Call EnterControls(LastColumnLetter, ListingRow + i)
    Next i

End Sub

Private Sub SetsArray(SetArr)
    
    Dim cCont As control
    
    For Each cCont In Me.MultiPage1.Pages(2).Controls   '2 is the page number of Parentage page
        If TypeName(cCont) = "CheckBox" Then
            If cCont = True Then
                If UBound(SetArr) > 0 Or SetArr(0) <> "" Then ReDim Preserve SetArr(UBound(SetArr) + 1)
                SetArr(UBound(SetArr)) = cCont.Name
            End If
        End If
    Next
    
End Sub

Private Sub EnterControls(LastColumnLetter As String, ListingRow As Integer)

    Dim cCont As control
    Dim FoundColumn As Integer
    
    'loop through controls and populate header with values user entered
    For Each cCont In Me.Controls
        'filter out the type of controls we need
        If Not TypeName(cCont) = "Label" And Not TypeName(cCont) = "MultiPage" And Not TypeName(cCont) = "CommandButton" And Not TypeName(cCont) = "Nothing" Then
            'Find the column number of the Control
            
            FoundColumn = AmazonColumn(LastColumnLetter, cCont.Name)
            
            'some items on the userform are not in the Amazon Template, like the Reboxed checkbox
            'if field name is not found, FoundColumn will return 0, and can't have a 0th column
            If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = cCont.Value
        End If
    Next cCont

End Sub

Private Sub EnterProductIDType(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "external_product_id_type")
    
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = "UPC"

End Sub

Private Sub EnterFeedProductType(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "feed_product_type")
    
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = "autopart"

End Sub

Private Sub EnterItemType(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    'lookup the BTG value of the part type
    Set rst = MstrDb.Execute("SELECT BTGValue FROM AAIAPartTypes WHERE AAIAPartType=" & Chr(34) & Me.part_type_id.Value & Chr(34))
    rst.MoveFirst
    
    'find the column in the excel sheet with "item_type" in the second row to enter the BTG value
    FoundColumn = AmazonColumn(LastColumnLetter, "item_type")
    
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = rst.Fields("BTGValue").Value
    
    'close the query
    rst.Close

End Sub

Private Sub EnterBrand(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    'brand
    If ListingMode = "Single" Then
        FoundColumn = AmazonColumn(LastColumnLetter, "brand_name")
        
        If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = "AD Auto Parts"
    Else
        'to be determined
        
    End If

End Sub

Private Sub EnterShippingTemplate(LastColumnLetter As String, ListingRow As Integer)

    Dim weight_oz As Double
    Dim shiptempcol As Integer
    
    'first find the column where the shipping template will be
    shiptempcol = AmazonColumn(LastColumnLetter, "merchant_shipping_group_name")
    
    'if user entered "Weight-Based" for the shipping tmeplate, calculate and replace it with the correct template
    If Cells(ListingRow, shiptempcol).Value = "Weight-Based" Then
        'convert weight to ounces and multiply by package qauntity to find total weight of listing
        If Me.website_shipping_weight_unit_of_measure = "LB" Then
            weight_oz = Me.website_shipping_weight.Value * 16 * Me.item_package_quantity.Value
        Else
            weight_oz = Me.website_shipping_weight.Value * Me.item_package_quantity.Value
        End If
        
        'choose appropriate Amazon Shipping Template based on weight
        Select Case weight_oz
        Case Is <= 13   '13 ounces or less
            '13 ounce template
            Cells(ListingRow, shiptempcol).Value = "13 oz. Template"
        
        Case Is <= 128  'between 13 oz and 8 lb.
            '1-8 lb. Template
            'round weight up to the next pound
            If weight_oz > 16 Then
                weight_oz = weight_oz / 16
            Else
                weight_oz = 16  'if weight is over 13 ounces but less than a pound, calculate template based on 1 pound
            End If
            
            'Concatenate the shipping template name
            Cells(ListingRow, shiptempcol).Value = RoundUp(weight_oz) & " lb. Template"
        
        Case Is <= 160  'between 8 lb. and 10 lb.
            '9-10 lb. Template
            Cells(ListingRow, shiptempcol).Value = "9-10 lb. Template"
        
        Case Is <= 192  'between 10 and 12 lb.
            '11-12 lb. Template
            Cells(ListingRow, shiptempcol).Value = "11-12 lb. Template"
        
        Case Is <= 224  'between 12 and 14 lb.
            '13-14 lb. Template
            Cells(ListingRow, shiptempcol).Value = "13-14 lb. Template"
        
        Case Is <= 288  'between 14 and 18 lb.
            '15-18 lb. Template
            Cells(ListingRow, shiptempcol).Value = "15-18 lb. Template"
        
        Case Is <= 304  'between 18 and 19 lb.
            '19 lb. Template
            Cells(ListingRow, shiptempcol).Value = "19 lb. Template"
            
        Case Else   'over 19 pounds
            '20-45 lb. Template
            Cells(ListingRow, shiptempcol).Value = "20-45 lb. Template"
        End Select
    Else
        'If user entered Prime for shipping template, replace with Prime shipping template
        Cells(ListingRow, shiptempcol).Value = Me.merchant_shipping_group_name
    End If
    
End Sub

Private Sub EnterDiscontinued(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "is_discontinued_by_manufacturer")
        
    'if field was found and the user left is_discontinued_by_manufacturer checkbox blank then change to null
    If Me.is_discontinued_by_manufacturer = False Then
        If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = ""
    End If

End Sub

Private Sub EnterNumberofItems(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "number_of_items")
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = Me.item_package_quantity

End Sub

Private Sub EnterQuantity(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "quantity")
        
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = 1    'Finale takes care of quantity listed

End Sub

Private Sub EnterTaxCode(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "product_tax_code")
        
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = "A_GEN_TAX"

End Sub

Private Sub EnterHandlingTime(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "fulfillment_latency")
        
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = 1   'should always be 1 to meet Amazon's standards

End Sub

Private Sub KeyProductFeatures(LastColumnLetter As String, ListingRow As Integer)

    Dim i As Integer
    Set rst = MstrDb.Execute("SELECT * FROM KeyProductFeatures WHERE ([Manufacturer]=" & Chr(34) & Me.Manufacturer.Value & Chr(34) & " AND [PartType]=" & Chr(34) & Me.part_type_id.Value & Chr(34) & ")")
    rst.MoveFirst
    
    Dim KPFColumn As Integer
    
    KPFColumn = AmazonColumn(LastColumnLetter, "bullet_point1")
    
    For i = KPFColumn To KPFColumn + 4
        Cells(ListingRow, i).Value = rst.Fields("KeyProductFeature" & i).Value
    Next i

End Sub

Private Sub EnterDimensionsUnitOfMeasure(LastColumnLetter As String, ListingRow As Integer)

    Dim LengthColumn As Integer
    Dim HeightColumn As Integer
    Dim WidthColumn As Integer
    Dim MeasureColumn As Integer
    
    'Find Item Dimension columns
    LengthColumn = AmazonColumn(LastColumnLetter, "item_length")
    HeightColumn = AmazonColumn(LastColumnLetter, "item_height")
    WidthColumn = AmazonColumn(LastColumnLetter, "item_width")
    MeasureColumn = AmazonColumn(LastColumnLetter, "item_dimensions_unit_of_measure")
    
    'If any item dimension is not null, the enter IN into the unit of measure field
    If Cells(ListingRow, LengthColumn).Value <> "" Or Cells(ListingRow, HeightColumn).Value <> "" Or Cells(ListingRow, WidthColumn).Value <> "" Then
        Cells(ListingRow, MeasureColumn).Value = "IN"
    End If

End Sub

Private Sub EnterProp65(LastColumnLetter As String, ListingRow As Integer)

    Dim FoundColumn As Integer
    
    FoundColumn = AmazonColumn(LastColumnLetter, "california_proposition_65_compliance_type")
        
    If FoundColumn > 0 Then Cells(ListingRow, FoundColumn).Value = "Passenger or Off Road Vehicle"   'All our items are for Passenger or Off Road Vehicle

End Sub

Private Sub EnterWarranty(LastColumnLetter As String, ListingRow As Integer)

    Dim rFindwar As Range
    Dim rFindwarTyp As Range
    
    

End Sub

Private Sub website_shipping_weight_Change()

    If IsNumeric(Me.website_shipping_weight.Value) = False Then
        Me.ShippingWeightLabel.ForeColor = RGB(255, 0, 0)
    End If

End Sub