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

    range("A1").value = "SKU"
    range("B1").value = "Offset Type(Amount or Percentage)"
    range("B2").value = "Percentage"
    range("C1").value = "T1 Min. Qty"
    range("D1").value = "T1 Max. Qty"
    range("E1").value = "T1 Offset Value"
    range("F1").value = "T2 Min. Qty"
    range("G1").value = "T2 Max. Qty"
    range("H1").value = "T2 Offset Value"
    range("I1").value = "T3 Min. Qty"
    range("J1").value = "T3 Max. Qty"
    range("K1").value = "T3 Offset Value"
    range("L1").value = "T4 Min. Qty"
    range("M1").value = "T4 Max. Qty"
    range("N1").value = "T4 Offset Value"

End Sub
