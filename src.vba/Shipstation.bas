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

    Range("A1").Value = "SKU"
    Range("B1").Value = "Name"
    Range("C1").Value = "WarehouseLocation"
    Range("D1").Value = "WeightOZ"
    Range("E1").Value = "Weight"
    Range("F1").Value = "Category"
    Range("G1").Value = "Tag1"
    Range("H1").Value = "Tag2"
    Range("I1").Value = "Tag3"
    Range("J1").Value = "Tag4"
    Range("K1").Value = "Tag5"
    Range("L1").Value = "CustomsDescription"
    Range("M1").Value = "CustomsValue"
    Range("N1").Value = "CustomsTariffNo"
    Range("O1").Value = "CustomsCountry"
    Range("P1").Value = "ThumbnailUrl"
    Range("Q1").Value = "UPC"
    Range("R1").Value = "FillSku"
    Range("S1").Value = "Length"
    Range("T1").Value = "Width"
    Range("U1").Value = "Height"
    Range("V1").Value = "UseProductName"
    Range("W1").Value = "Active"
    Range("X1").Value = "SKUAlias"
    Range("Y1").Value = "IsReturnable"

End Sub
