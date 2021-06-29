Option Explicit

Public Sub UPC()

    Dim wsSheet As Worksheet

    On Error Resume Next
    
    'GS1 Template sheet name
    Dim SheetName As String
    SheetName = "GS1 Template"
    
    Set wsSheet = Sheets(SheetName)
    On Error GoTo 0
    
    'Determine what to do based on whether "GS1 Template" exists or not
    If Not wsSheet Is Nothing Then
        'if sheet exists
        If CheckUPC = True Then
            'if current sheet is "GS1 Template"
            
        Else
            'if current sheet is not "GS1 Template"
            Worksheets(SheetName).Activate
        End If
    Else
        'if "GS1 Template" doesn't exist
        Call PrepWorksheet(SheetName)
        
        'Add headers to create Amazon Template worksheet
        GS1Headers
        
        'Add default values
        DefaultValues
    End If
    
    'refresh ribbon
    RibbonCategories
    
    range("A2").Select

End Sub

Private Sub GS1Headers()

    range("A1").value = "Action"
    range("B1").value = "GS1CompanyPrefix"
    range("C1").value = "GTIN"
    range("D1").value = "PackagingLevel"
    range("E1").value = "Description"
    range("F1").value = "SKU"
    range("G1").value = "BrandName"
    range("H1").value = "Status"
    range("I1").value = "IsVariable"
    range("J1").value = "IsPurchasable"
    range("K1").value = "Certified"
    range("L1").value = "Height"
    range("M1").value = "Width"
    range("N1").value = "Depth"
    range("O1").value = "DimensionMeasure"
    range("P1").value = "GrossWeight"
    range("Q1").value = "NetWeight"
    range("R1").value = "WeightMeasure"
    range("S1").value = "Comments"
    range("T1").value = "CountryOfOrigin"
    range("U1").value = "ChildGTINs"
    range("V1").value = "Quantity"
    range("W1").value = "SubBrandName"
    range("X1").value = "ProductDescriptionShort"
    range("Y1").value = "LabelDescription"
    range("Z1").value = "NetContent1Count"
    range("AA1").value = "NetContent1UnitOfMeasure"
    range("AB1").value = "NetContent2Count"
    range("AC1").value = "NetContent2UnitOfMeasure"
    range("AD1").value = "NetContent3Count"
    range("AE1").value = "NetContent3UnitOfMeasure"
    range("AF1").value = "GlobalProductClassification"
    range("AG1").value = "ImageURL"
    range("AH1").value = "TargetMarket"

End Sub

Private Sub DefaultValues()

    range("A2").value = "Create"
    range("D2").value = "Each"
    range("G2").value = "AD Auto Parts"
    range("H2").value = "In Use"
    range("I2").value = "N"
    range("J2").value = "Y"

End Sub
