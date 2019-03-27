Option Explicit

Public Sub BoMMain()

    'turn screen updating off
    Application.ScreenUpdating = False
    
    'move columns C-F up one row
    Range("C2:F2").Delete xlShiftUp
    
    'Delete empty rows
    DeleteEmpty
    
    'Repeat Product ID
    RepeatProductID
    
    'turn screen updating on
    Application.ScreenUpdating = True
    
    'Refresh Ribbon
    RibbonCategories

End Sub

Private Sub DeleteEmpty()
    
    'Find the last non-blank cell in column D(4)
    Dim lRow As Long
    lRow = Cells(Rows.count, 4).End(xlUp).Row
    
    'save range to variable
    Dim R As Range, i As Long
    Set R = ActiveSheet.Range("A1:F" & 5891)

    'loop through rows from bottom to top
    For i = lRow To 1 Step (-1)
        'if row is empty, delete it
        If WorksheetFunction.CountA(R.Rows(i)) = 0 Then R.Rows(i).Delete
    Next

End Sub

Private Sub RepeatProductID()

    'find last row in column D(4)
    Dim lRow As Long
    lRow = Cells(Rows.count, 4).End(xlUp).Row
    
    'loop
    Dim i As Integer
    Dim ProdID As String
    For i = 2 To lRow
        If Cells(i, 1).Value <> "" Then
            ProdID = Cells(i, 1).Value
        Else
            Cells(i, 1).Value = ProdID
        End If
    Next i

End Sub
