Option Explicit

Public Sub InheritanceMain()

    Dim SheetName As String
    SheetName = "MyFitment Inheritance"
    
    'open or create Stock Take worksheet
    Call PrepWorksheet(SheetName)           'WorksheetConnections Module
    
    'Fill in the headers
    AddHeaders
    
    'Format Headers
    FormatHeaders
    
    'Autofit
    columns("A:J").AutoFit

End Sub

Private Sub AddHeaders()

    range("A1").value = "SKU"
    range("B1").value = "Your Part #"
    range("C1").value = "Inherits Fitment From Part #"
    range("D1").value = "ASIN"
    range("E1").value = "UPC"
    range("F1").value = "Description"
    range("G1").value = "Label"
    range("H1").value = "Landing Page URL"
    range("I1").value = "AAIA Part Type"
    range("J1").value = "AAIA Brand Code"

End Sub

Private Sub FormatHeaders()

    'color
    range("A1:B1").Interior.Color = RGB(0, 176, 240)    'part number fields
    range("C1:H1").Interior.Color = RGB(204, 255, 204)  'MyFitment fields
    range("I1:J1").Interior.Color = RGB(255, 255, 0)    'AAIA fields
    
    'Borders
    range("A1:J1").Borders(xlEdgeBottom).LineStyle = xlContinuous
    With range("A1:J1")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With

End Sub
