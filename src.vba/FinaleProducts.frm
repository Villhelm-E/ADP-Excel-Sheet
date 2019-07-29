Option Explicit

Dim CatFieldArr()

Private Sub UserForm_Initialize()

    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'fill in the Finale categories in the Categories List Box
    PopulateCategoriesListBox
    
    'Initialize the Categories and Field array
    InitializeCatField

End Sub

Private Sub PopulateCategoriesListBox()

    'query Categories
    Set rst = MstrDb.Execute("SELECT Category FROM FinaleProductFields WHERE Not Category =" & Chr(34) & "Default" & Chr(34) & " GROUP BY Category;")
    rst.MoveFirst
    
    'populate the categories list box
    Dim i As Integer
    With rst
        While Not .EOF
            Me.CategoryListBox.AddItem rst.Fields("Category").Value
            rst.MoveNext
        Wend
    End With
    
    rst.Close

End Sub

Private Sub PopulateFieldListBox(ChosenCat As String)

    'query Fields
    Set rst = MstrDb.Execute("SELECT Field FROM FinaleProductFields WHERE Category = " & Chr(34) & ChosenCat & Chr(34) & ";")
    rst.MoveFirst
    
    'reset the lsit box
    Me.FieldListBox.Clear
    
    'populate the Field list box
    Dim i As Integer
    With rst
        While Not .EOF
            Me.FieldListBox.AddItem rst.Fields("Field").Value
            rst.MoveNext
        Wend
    End With
    
    rst.Close

End Sub

Private Sub UpdateCatFieldArr(fieldsCount As Integer)

    'i is looping through the values in field list box instead of fields in CatFieldArr so the index isn't matching
    Dim i As Integer
    Dim j As Integer
    For i = 0 To Me.FieldListBox.ListCount - 1
        For j = 0 To UBound(CatFieldArr())
            If CatFieldArr(j, 0) = Me.FieldListBox.List(i) Then
                CatFieldArr(j, 2) = Me.FieldListBox.Selected(i)
                GoTo stop_loop
            End If
        Next j
stop_loop:
    Next i

End Sub

Private Sub InitializeCatField()
    
    'Count Fields
    Set rst = MstrDb.Execute("SELECT Field, Category FROM FinaleProductFields WHERE Category <> " & Chr(34) & "Default" & Chr(34) & ";")
    
    Dim Fields As Integer
    Fields = rst.RecordCount
    
    'array is going to be built like countoffields,countofcategories
    ReDim CatFieldArr(Fields - 1, 2)
    
    'populate CatFieldArr array
    Dim i As Integer
    rst.MoveFirst
    For i = 0 To Fields - 1
        'Field Name
        CatFieldArr(i, 0) = rst.Fields("Field").Value
        
        'Field Category
        CatFieldArr(i, 1) = rst.Fields("Category").Value
        
        'Selection
        CatFieldArr(i, 2) = False
        
        rst.MoveNext
    Next i
    
    'close recordset
    rst.Close

End Sub

Private Sub CategoryListBox_AfterUpdate()

    Dim i As Integer
    Dim Cat As String
    Dim x As Integer
    
    SaveSelectedFields
    
    'find the selected Category
    For i = 0 To Me.CategoryListBox.ListCount - 1
        If Me.CategoryListBox.Selected(i) = True Then
            'when loop finds selected option, save to variable
            Cat = Me.CategoryListBox.List(i)
            'and end the loop
            i = Me.CategoryListBox.ListCount - 1
        End If
    Next i
    
    'populate the fields in fieldlistbox based on the category the user chose
    Call PopulateFieldListBox(Cat)
    
    'update the checkboxes in fieldlistbox with the fields the user selected
    LoadSelectedFields

End Sub

Private Sub SaveSelectedFields()

    Dim CurrCat As Integer
    Dim UBFields As Integer
    Dim i As Long
    
    'loop through listbox to find the selected option
    For i = 0 To Me.CategoryListBox.ListCount - 1
        If Me.CategoryListBox.Selected(i) = True Then
            'when loop finds selected option, save to variable
            CurrCat = Me.CategoryListBox.ListIndex
            'and end the loop
            i = Me.CategoryListBox.ListCount - 1
        End If
    Next i
    
    UBFields = Me.FieldListBox.ListCount - 1
    
    If UBFields <> -1 Then ReDim Preserve CategoriesArr(Me.CategoryListBox.ListCount - 1, UBFields) 'maybe rewrite this module to not use CategoriesArr global variable
    
    'loop through listbox to find the selected option
    For i = 0 To UBFields
        CategoriesArr(CurrCat, i) = Me.FieldListBox.Selected(i)
    Next i
    
    'Count Fields
    Set rst = MstrDb.Execute("SELECT Field FROM FinaleProductFields WHERE Category <> " & Chr(34) & "Default" & Chr(34) & ";")
    Dim Fields As Integer
    Fields = rst.RecordCount
    rst.Close
    
    Call UpdateCatFieldArr(Fields)

End Sub

Private Sub LoadSelectedFields()

    Dim i As Integer
    Dim j As Integer
    For i = 0 To Me.FieldListBox.ListCount - 1
        For j = 0 To UBound(CatFieldArr())
            If CatFieldArr(j, 0) = Me.FieldListBox.List(i) Then
                Me.FieldListBox.Selected(i) = CatFieldArr(j, 2)
                GoTo Exit_Loop
            End If
        Next j
Exit_Loop:
    Next i

End Sub

Private Sub SubmitBtn_Click()
    
    'add Product id first
    Range("A1").Value = "Product id"    'hard coded for now, will update to use default fields in the future
    
    'loop through CatFieldArr and add each item with value True as header
    Dim i As Integer
    Dim c As Integer
    c = 2
    For i = 0 To UBound(CatFieldArr()) - 1
        If CatFieldArr(i, 2) = True Then
            Cells(1, c).Value = CatFieldArr(i, 0)
            c = c + 1
        End If
    Next i
    
    'clean up
    Dim lastcolumn As String
    lastcolumn = NumberToColumn(CountColumns(Range("1:1")))
    Range("A:" & lastcolumn).EntireColumn.AutoFit
    
    Range("A1").Select
    
    'rename sheet
    Call PrepWorksheet("Finale Products")
    
    FinaleProducts.Hide

End Sub