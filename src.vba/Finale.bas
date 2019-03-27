Option Explicit

Public Sub FinaleProductsMain()
    
    'open FinaleProducts userform
    FinaleProducts.Show
    
    'if FinaleFields array is empty, user canceled the operation. exit sub
    If IsArrayAllocated(FinaleFields) = False Then Exit Sub
    
    'turn off screen updating
    Application.ScreenUpdating = False

    'open or create Finale Products worksheet
    Call PrepWorksheet("Finale Products")               'WorksheetConnections Module
        
    'Fill in the headers
    FinaleProductsAddHeaders

    'Autofit
    columns("A:AA").AutoFit
    
    'clear FinaleFields array
    Erase FinaleFields
    
    'turn on screen updating
    Application.ScreenUpdating = True
    
    'Refresh Ribbon
    RibbonCategories

End Sub

Public Sub FinaleStockTakeMain()

    Dim SheetName As String
    SheetName = "Finale Stock Take"
    
    'open or create Stock Take worksheet
    Call PrepWorksheet(SheetName)           'WorksheetConnections Module
    
    'Fill in the headers
    FinaleStockTakeAddHeaders
    
    'Autofit
    columns("A:B").AutoFit

End Sub

Public Sub FinaleBoMMain()

    Dim SheetName As String
    SheetName = "Finale Bill of Materials"
    
    'open or create Bill of Materials worksheet
    Call PrepWorksheet(SheetName)           'WorksheetConnections Module
    
    'Fill in the headers
    FinaleBoMAddHeaders
    
    'Autofit
    columns("A:C").AutoFit

End Sub

Public Sub FinaleLookupsMain()

    Dim SheetName As String
    SheetName = "Finale Lookups"
    
    'open or create Lookups worksheet
    Call PrepWorksheet(SheetName)           'WorksheetConnections Module
    
    'Fill in the headers
    FinaleLookupsAddHeaders
    
    'Autofit
    columns("A:C").AutoFit

End Sub

Private Sub FinaleProductsAddHeaders()

    Dim i As Integer
    
    For i = 0 To UBound(FinaleFields)
        Cells(1, i + 1).Value = FinaleFields(i)
    Next i
    
End Sub

Private Sub FinaleStockTakeAddHeaders()

    Range("A1").Value = "Product ID"
    Range("B1").Value = "Quantity"

End Sub

Private Sub FinaleBoMAddHeaders()

    Range("A1").Value = "Product ID"
    Range("B1").Value = "Quantity"
    Range("C1").Value = "Item product ID"

End Sub

Private Sub FinaleLookupsAddHeaders()

    Range("A1").Value = "Product ID"
    Range("B1").Value = "Product lookup"
    Range("C1").Value = "Stores to add"

End Sub
