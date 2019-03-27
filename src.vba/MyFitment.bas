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

    Range("A1").Value = "SKU"
    Range("B1").Value = "Your Part #"
    Range("C1").Value = "Inherits Fitment From Part #"
    Range("D1").Value = "ASIN"
    Range("E1").Value = "UPC"
    Range("F1").Value = "Description"
    Range("G1").Value = "Label"
    Range("H1").Value = "Landing Page URL"
    Range("I1").Value = "AAIA Part Type"
    Range("J1").Value = "AAIA Brand Code"

End Sub

Private Sub FormatHeaders()

    'color
    Range("A1:B1").Interior.Color = RGB(0, 176, 240)    'part number fields
    Range("C1:H1").Interior.Color = RGB(204, 255, 204)  'MyFitment fields
    Range("I1:J1").Interior.Color = RGB(255, 255, 0)    'AAIA fields
    
    'Borders
    Range("A1:J1").Borders(xlEdgeBottom).LineStyle = xlContinuous
    With Range("A1:J1")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With

End Sub
