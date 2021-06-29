Option Explicit

Public Sub FinaleProductsMain()
    
    'open FinaleProducts userform
    FinaleProducts.Show
    
'''''for some reason none of this code runs after clicking the button in the FinaleProducts userform
    
    'if FinaleFields array is empty, user canceled the operation. exit sub
    If IsArrayAllocated(FinaleFields) = False Then Exit Sub
    
    'turn off screen updating
    Application.ScreenUpdating = False

    'open or create Finale Products worksheet
    Call PrepWorksheet("Finale Products")               'TableConnections Module
        
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
    Call PrepWorksheet(SheetName)           'TableConnections Module
    
    'Fill in the headers
    FinaleStockTakeAddHeaders
    
    'Autofit
    columns("A:B").AutoFit

End Sub

Public Sub FinaleBoMMain()

    Dim SheetName As String
    SheetName = "Finale Bill of Materials"
    
    'open or create Bill of Materials worksheet
    Call PrepWorksheet(SheetName)           'TableConnections Module
    
    'Fill in the headers
    FinaleBoMAddHeaders
    
    'Autofit
    columns("A:D").AutoFit

End Sub

Public Sub FinaleLookupsMain()

    Dim SheetName As String
    SheetName = "Finale Lookups"
    
    'open or create Lookups worksheet
    Call PrepWorksheet(SheetName)           'TableConnections Module
    
    'Fill in the headers
    FinaleLookupsAddHeaders
    
    'Autofit
    columns("A:C").AutoFit

End Sub

Private Sub FinaleProductsAddHeaders()

    Dim i As Integer
    
    For i = 0 To UBound(FinaleFields)
        Cells(1, i + 1).value = FinaleFields(i)
    Next i
    
End Sub

Private Sub FinaleStockTakeAddHeaders()

    range("A1").value = "Product ID"
    range("B1").value = "Quantity"

End Sub

Private Sub FinaleBoMAddHeaders()

    range("A1").value = "Product ID"
    range("B1").value = "Quantity"
    range("C1").value = "Component product ID"
    range("D1").value = "Component note"

End Sub

Private Sub FinaleLookupsAddHeaders()

    range("A1").value = "Product ID"
    range("B1").value = "Product lookup"
    range("C1").value = "Stores to add"

End Sub

Public Sub ShippingMain()

    Dim SheetName As String
    SheetName = "Shipping Methods"
    
    'open or create Stock Take worksheet
    Call PrepWorksheet(SheetName)           'TableConnections Module
    
    'Fill in the headers
    ShippingMethodHeaders
    
    'Autofit
    columns("A:E").AutoFit

End Sub

Private Sub ShippingMethodHeaders()

    range("A1").value = "Product ID"
    range("B1").value = "Ebay Shipping Method"
    range("C1").value = "Ebay Shipping Cost"
    range("D1").value = "Amazon Shipping Method"
    range("E1").value = "Amazon Shipping Cost"

End Sub

'Ribbon Module triggers this Module
Public Sub ShippingValidation()

    Dim methods
    Dim i As Integer
    Dim str As String
    Dim SKUs As range
    Dim endOfRange As Integer
    Dim r As range
    Dim lookup As String
    
    Application.ScreenUpdating = False
    
    'remove any current data validation
    range("B:B,D:D").Validation.Delete
    
    'conect to shipping methods table
    OpenShippingMethods
    
    'count SKUs
    endOfRange = CountRows("A") - 1
    
    'if there are SKUs
    If endOfRange > 0 Then
        'define range for validation based on SKUs
        Set SKUs = range("B2:B" & endOfRange + 1 & ",D2:D" & endOfRange + 1)
        
        'create list of table values from table
        With rst
            For i = 1 To .RecordCount - 1
                If i = .RecordCount - 1 Then
                    str = str & .fields("ShippingMethod").value
                Else
                    str = str & .fields("ShippingMethod").value & ", "
                End If
                rst.MoveNext
            Next i
        End With
    
        'use list generated above as array
        methods = Array(str)
        
        'set data validation
        Call ValidateRange(SKUs, methods)
        
        'Add the shipping prices from lookup
        For Each r In SKUs
            
            'lookup the price based on shipping method
            lookup = ShippingLookup(r.value)
            
            If r <> "" Then
                'only add lookup if the price doesn't match
                If r.Offset(0, 1).value <> lookup Then
                    r.Offset(0, 1).value = lookup
                End If
            End If
        Next r
        
    End If
    
    Application.ScreenUpdating = True

End Sub

Private Sub ValidateRange(range As range, list)

    With range.Validation
        'create list from array
        .Add Type:=xlValidateList, _
        AlertStyle:=xlValidAlertInformation, _
        Formula1:=Join(list, ", ")
    End With

End Sub

Private Function ShippingLookup(method As String)
    
    If method = "" Then
        ShippingLookup = ""
    Else
        rst.MoveFirst
        Do While Not rst.EOF
            If rst.fields("ShippingMethod") = method Then
                ShippingLookup = rst.fields("ShippingCost")
                'end loop
            End If
            rst.MoveNext
        Loop
    End If

End Function

Public Sub POMain()

    Dim SheetName As String
    SheetName = "Purchase Order Items"
    
    'open or create Lookups worksheet
    Call PrepWorksheet(SheetName)           'TableConnections Module
    
    'Fill in the headers
    POHeaders
    
    'Autofit
    columns("A:D").AutoFit

End Sub

Private Sub POHeaders()

    range("A1").value = "Product ID"
    range("B1").value = "Quantity"
    range("C1").value = "Unit price"
    range("D1").value = "Item note"

End Sub
