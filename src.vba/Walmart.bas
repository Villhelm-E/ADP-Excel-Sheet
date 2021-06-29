Option Explicit

Public Sub WalmartMain()

    Dim SheetName As String
    SheetName = "Walmart Template"
    
    'open or create Stock Take worksheet
    Call PrepWorksheet(SheetName)           'WorksheetConnections Module
    
    'turn on wrap text
    range("A1:N3").WrapText = True
    
    'resize columns
    Rows("1:1").RowHeight = 63
    
    'Fill in the headers
    WalmartHeaders
    
    'hide rows 2 and 3
    Rows("2:3").Hidden = True
    
    'fill in default fields
    range("B4:B1000").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="GTIN,UPC"
    range("B4").value = "GTIN"
    range("D4").value = "2038710"
    range("I4").value = "USD"
    range("L4").value = "Vehicle"
    range("N4").value = "VehiclePartsAndAccessories"
    
    'Set width to 20.14
    range("A:N").ColumnWidth = 20.14
    
    'hide column A
    columns(1).Hidden = True
    
End Sub

Public Sub WalmartHeaders()

    'run through WalmartFields table in Master Database and add headers
    Set rst = MstrDb.Execute("SELECT * FROM WalmartFields")
    rst.MoveFirst
    Dim c As Integer
    
    With rst
        For c = 1 To rst.RecordCount
            Cells(1, c).value = rst.fields("FieldName").value
            Cells(2, c).value = rst.fields("FieldDesc").value
            Cells(3, c).value = rst.fields("ThirdField").value
            
            If rst.fields("Required").value = True Then
                Cells(1, c).Font.Bold = True
                Cells(1, c).Interior.Color = RGB(204, 255, 204)
            Else
                Cells(1, c).Interior.Color = RGB(204, 204, 255)
            End If
            
            rst.MoveNext
        Next c
    End With

End Sub
