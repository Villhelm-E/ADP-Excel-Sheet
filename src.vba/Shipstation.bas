Option Explicit

Public Sub ShipstationFieldsMain()

    Dim SheetName As String
    SheetName = "Shipstation"
    
    'open or create Stock Take worksheet
    Call PrepWorksheet(SheetName)           'WorksheetConnections Module
    
    'Fill in the headers
    ShipstationAddHeaders
    
    'Autofit
    columns("A:Y").AutoFit

End Sub

Private Sub ShipstationAddHeaders()

    range("A1").value = "SKU"
    range("B1").value = "Name"
    range("C1").value = "WarehouseLocation"
    range("D1").value = "WeightOZ"
    range("E1").value = "Weight"
    range("F1").value = "Category"
    range("G1").value = "Tag1"
    range("H1").value = "Tag2"
    range("I1").value = "Tag3"
    range("J1").value = "Tag4"
    range("K1").value = "Tag5"
    range("L1").value = "CustomsDescription"
    range("M1").value = "CustomsValue"
    range("N1").value = "CustomsTariffNo"
    range("O1").value = "CustomsCountry"
    range("P1").value = "ThumbnailUrl"
    range("Q1").value = "UPC"
    range("R1").value = "FillSku"
    range("S1").value = "Length"
    range("T1").value = "Width"
    range("U1").value = "Height"
    range("V1").value = "UseProductName"
    range("W1").value = "Active"
    range("X1").value = "SKUAlias"
    range("Y1").value = "IsReturnable"

End Sub
