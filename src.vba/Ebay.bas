Option Explicit

Public Sub VolumePricing()

    Dim wsSheet As Worksheet

    On Error Resume Next
    
    'Amazon Listings sheet name
    Dim SheetName As String
    SheetName = "Volume Pricing"
    
    Set wsSheet = Sheets(SheetName)
    On Error GoTo 0
    
    'Determine what to do based on whether "Upload Template" exists or not
    If Not wsSheet Is Nothing Then
        'if current sheet is not "Volume Pricing"
        Worksheets(SheetName).Activate
    Else
        'if "Upload Template" doesn't exist
        Call PrepWorksheet(SheetName)
        
        'Add headers to create Amazon Template worksheet
        VolumePricingHeaders
    End If
    
    'refresh ribbon
    RibbonCategories

End Sub

Private Sub VolumePricingHeaders()

    Range("A1").Value = "SKU"
    Range("B1").Value = "Offset Type(Amount or Percentage)"
    Range("B2").Value = "Percentage"
    Range("C1").Value = "T1 Min. Qty"
    Range("D1").Value = "T1 Max. Qty"
    Range("E1").Value = "T1 Offset Value"
    Range("F1").Value = "T2 Min. Qty"
    Range("G1").Value = "T2 Max. Qty"
    Range("H1").Value = "T2 Offset Value"
    Range("I1").Value = "T3 Min. Qty"
    Range("J1").Value = "T3 Max. Qty"
    Range("K1").Value = "T3 Offset Value"
    Range("L1").Value = "T4 Min. Qty"
    Range("M1").Value = "T4 Max. Qty"
    Range("N1").Value = "T4 Offset Value"

End Sub