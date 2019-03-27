Option Explicit

Public Sub OOSMain()

    'Turn screen updating off
    Application.ScreenUpdating = False
    
    'BypassRibbon = True
    
    'Delete useless columns
    columns("B:D").Delete
    
    'Move My Note to column E
    MoveMyNote
    
    'Deletes rows of any cell in A that's not a SKU
    DeleteNonSKUs
    
    'Delete empy Rows
    DeleteRows
    
    'Get rid of products with "My note"
    PruneProducts
    
    'Format sheet
    Headers
    Range("A:A").EntireColumn.AutoFit
    Range("A2").Select
    
    'Turn screen updating on
    Application.ScreenUpdating = True
    
    BypassRibbon = False
    
    RibbonCategories

End Sub

Private Sub MoveMyNote()

    Dim R As Range
    
    For Each R In Intersect(Range("A:A"), ActiveSheet.UsedRange)
        If R.Value Like "My note:*" Then R.Cut Destination:=R.Offset(-2, 1)  'move My Note to column B and up 2 rows
    Next R

End Sub

Private Sub DeleteNonSKUs()

    Dim R As Range
    
    For Each R In Intersect(Range("A:A"), ActiveSheet.UsedRange)
        If R.Value = "Select this item for performing bulk action" Or R.Value Like "eBay note:*" Then R.EntireRow.Delete    'anything that's not a sku in column A gets deleted
    Next R

End Sub

Private Sub DeleteRows()

    Dim EndRange As Integer
    Dim i As Integer
    
    'find the last row with value
    EndRange = LastRow("A:A")
    
    'go through every row
    For i = 2 To EndRange
        If IsEmpty(Range("A" & i)) Then
            'if cell A is blank, delete the whole row
            Range(i & ":" & i).EntireRow.Delete
            
            'go back one row to account for the rows shifting up after deleting
            i = i - 1
            
            'subtract one from EndRange so the loop doesn't go on longer than it has to and gets stuck in an infinite loop
            If i + 1 < EndRange Then
                EndRange = EndRange - 1
            Else
                'if the EndRange catches up to current row in loop, stop so it doesn't get stuck in an infinite loop
                Exit Sub
            End If
        End If
    Next i

End Sub

Private Sub PruneProducts()

    Dim R As Integer
    Dim i As Integer
    
    'find last row
    i = LastRow("A:A")
    
    'loop through all rows except the first one to reserve it for headers
    For R = 2 To i
        'if Column B starts with "My note" then delete the row
        If Range("B" & R).Value Like "My note:*" Then
            Range("B" & R).EntireRow.Delete
            
            'Go back up one row to account for the rows shifting up after deleting
            R = R - 1
        End If
    Next R

End Sub

Private Sub Headers()

    Range("A1").Value = "Ebay SKU"

End Sub
