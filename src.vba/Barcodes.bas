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
    
    Range("A2").Select

End Sub

Private Sub GS1Headers()

    Range("A1").Value = "Action"
    Range("B1").Value = "GS1CompanyPrefix"
    Range("C1").Value = "GTIN"
    Range("D1").Value = "PackagingLevel"
    Range("E1").Value = "Description"
    Range("F1").Value = "SKU"
    Range("G1").Value = "BrandName"
    Range("H1").Value = "Status"
    Range("I1").Value = "IsVariable"
    Range("J1").Value = "IsPurchasable"
    Range("K1").Value = "Certified"
    Range("L1").Value = "Height"
    Range("M1").Value = "Width"
    Range("N1").Value = "Depth"
    Range("O1").Value = "DimensionMeasure"
    Range("P1").Value = "GrossWeight"
    Range("Q1").Value = "NetWeight"
    Range("R1").Value = "WeightMeasure"
    Range("S1").Value = "Comments"
    Range("T1").Value = "CountryOfOrigin"
    Range("U1").Value = "ChildGTINs"
    Range("V1").Value = "Quantity"
    Range("W1").Value = "SubBrandName"
    Range("X1").Value = "ProductDescriptionShort"
    Range("Y1").Value = "LabelDescription"
    Range("Z1").Value = "NetContent1Count"
    Range("AA1").Value = "NetContent1UnitOfMeasure"
    Range("AB1").Value = "NetContent2Count"
    Range("AC1").Value = "NetContent2UnitOfMeasure"
    Range("AD1").Value = "NetContent3Count"
    Range("AE1").Value = "NetContent3UnitOfMeasure"
    Range("AF1").Value = "GlobalProductClassification"
    Range("AG1").Value = "ImageURL"
    Range("AH1").Value = "TargetMarket"

End Sub

Private Sub DefaultValues()

    Range("A2").Value = "Create"
    Range("D2").Value = "Each"
    Range("G2").Value = "AD Auto Parts"
    Range("H2").Value = "In Use"
    Range("I2").Value = "N"
    Range("J2").Value = "Y"

End Sub

Public Sub ComputerName()

    MsgBox CreateObject("WScript.Network").ComputerName

End Sub
