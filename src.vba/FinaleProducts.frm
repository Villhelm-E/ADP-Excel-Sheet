Option Explicit

Public UBFields As Integer
Public UBCats As Integer

Private Sub UserForm_Initialize()

    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'Redim FinaleFields Array
    RedimFinaleFields
    UBFields = UBound(FinaleFields())
    
    'Redim CategoriesArr Array
    RedimCategoriesArr
    UBCats = UBound(CategoriesArr())
    
    'Add Categories to Listbox
    Dim i As Integer
    For i = 0 To UBCats
        Me.CategoryListBox.AddItem CategoriesArr(i, 0)
    Next i

End Sub

'FinaleFields is global array
Private Sub RedimFinaleFields()

    'Query Finale Fields
    Set rst = MstrDb.Execute("SELECT [Field], [Category] FROM [FinaleProductFields] WHERE ((Not ([Category])=" & Chr(34) & "Default" & Chr(34) & ")) ORDER BY [ID];")
    
    'convert query to array
    Dim count As Integer
    With rst
        rst.MoveFirst

        'count the number of records
        count = rst.RecordCount
        
        'resize the array
        ReDim FinaleFields(count - 1, 1)
        
        'loop through rst
        While Not .EOF
            'first value in array is the field name
            FinaleFields(rst.AbsolutePosition - 1, 0) = rst.Fields("Field").Value
            
            'the second value in array is the category
            FinaleFields(rst.AbsolutePosition - 1, 1) = rst.Fields("Category").Value
            
            rst.MoveNext
        Wend
    End With
    
    rst.Close

End Sub

'CategoriesArr is global array
Private Sub RedimCategoriesArr()

    'Query unique Categories
    Set rst = MstrDb.Execute("SELECT DISTINCT FinaleProductFields.Category, Active FROM FinaleProductFields WHERE (Not FinaleProductFields.Category=""Default"") ORDER BY FinaleProductFields.Category")
    
    Dim count As Integer
    With rst
        rst.MoveFirst
        
        count = rst.RecordCount
        
        ReDim CategoriesArr(count - 1, 1)
        
        While Not .EOF
            
            CategoriesArr(rst.AbsolutePosition - 1, 0) = rst.Fields("Category").Value
            
            CategoriesArr(rst.AbsolutePosition - 1, 1) = "False"
            
            rst.MoveNext
        Wend
    End With
    
    rst.Close

End Sub

Private Sub PopulateCategoriesArray()
    
    Set rst = MstrDb.Execute("SELECT FinaleProductFields.Field, FinaleProductFields.Category FROM FinaleProductFields WHERE (Not FinaleProductFields.Category=""Default"") ORDER BY FinaleProductFields.ID")
    
    'Add Categories and Fields to global Array
    Dim i As Integer
    
    With rst
        Do Until .EOF
            i = i + 1
            rst.MoveNext
        Loop
    End With
    
    ReDim FinaleFields(i - 1, 2)
    
    rst.MoveFirst
    
    For i = 0 To UBound(FinaleFields)
        FinaleFields(i, 0) = rst.Fields("Field").Value
        rst.MoveNext
    Next i
    
    rst.MoveFirst
    
    For i = 0 To UBound(FinaleFields)
        FinaleFields(i, 1) = rst.Fields("Category").Value
        rst.MoveNext
    Next i
    
    rst.Close
    
    For i = 0 To UBound(FinaleFields)
        FinaleFields(i, 2) = 0
    Next i

End Sub

Private Sub CategoryListBox_AfterUpdate()

    Dim i As Integer
    Dim Cat As String
    Dim x As Integer
    
    'loop through CategoriesListBox to find the selected option
    For i = 0 To Me.CategoryListBox.ListCount - 1
        If Me.CategoryListBox.Selected(i) = True Then
            'when loop finds selected option, save to variable
            Cat = Me.CategoryListBox.List(i)
            'and end the loop
            i = Me.CategoryListBox.ListCount - 1
        End If
    Next i
    
    'Query the Fields in the Category the user chose above
    Set rst = MstrDb.Execute("SELECT FinaleProductFields.Field FROM FinaleProductFields WHERE (FinaleProductFields.Category = " & Chr(34) & Cat & Chr(34) & ") ORDER BY FinaleProductFields.ID")
    rst.MoveFirst
    
    If Me.ListBox2.ListCount = 0 Then
        'add Fields to Listbox2
        With rst
            Do While Not .EOF
                Me.ListBox2.AddItem rst.Fields("Field").Value
                rst.MoveNext
            Loop
        End With
    Else
        'remove items from Listbox2
        i = 0
        
        Do While Me.ListBox2.ListCount > 0
            Me.ListBox2.RemoveItem i
        Loop
        
        rst.MoveFirst
        
        'add Fields to Listbox2
        With rst
            Do While Not .EOF
                Me.ListBox2.AddItem rst.Fields("Field").Value
                rst.MoveNext
            Loop
        End With
    End If
    
    rst.Close

End Sub

Private Sub ListBox2_Change()

    Dim CurrCat As Integer
    Dim UBFields As Integer
    Dim i As Long
    
    'test
    
    
    'loop through listbox to find the selected option
    For i = 0 To Me.CategoryListBox.ListCount - 1
        If Me.CategoryListBox.Selected(i) = True Then
            'when loop finds selected option, save to variable
            CurrCat = Me.CategoryListBox.ListIndex
            'and end the loop
            i = Me.CategoryListBox.ListCount - 1
        End If
    Next i
    
    UBFields = Me.ListBox2.ListCount - 1
    
    ReDim Preserve CategoriesArr(Me.CategoryListBox.ListCount - 1, UBFields)
    
    'loop through listbox to find the selected option
    For i = 0 To UBFields
        CategoriesArr(CurrCat, i) = Me.ListBox2.Selected(i)
    Next i

End Sub

Private Sub SubmitBtn_Click()

'    FinaleFields
    
    Dim i As Integer
    Dim R As Integer
    
    ReDim FinaleFields(0)
    FinaleFields(0) = "Product ID"
    
    For i = 0 To Me.ListBox2.ListCount - 1
        If Me.ListBox2.Selected(i) = True Then
            ReDim Preserve FinaleFields(UBound(FinaleFields) + 1)
            FinaleFields(UBound(FinaleFields)) = Me.ListBox2.List(i)
        End If
    Next i
    
    Unload Me

End Sub