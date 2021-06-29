Option Explicit

Public Sub OOSMain()

    'Turn screen updating off
    Application.ScreenUpdating = False
    
    'Delete useless columns
    columns("B:F").Delete
    
    'Move My Note to column E
    MoveMyNote
    
    'Deletes rows with Ebay note:
    DeleteEbayNote
    
    'Deletes anything left that's not SKUs
    DeleteNonSKUs
    
    'Delete empy Rows
    DeleteRows
    
    'Get rid of products with "My note"
    PruneProducts
    
    'Format sheet
    Headers
    range("A:A").EntireColumn.AutoFit
    range("A2").Select
    
    'Turn screen updating on
    Application.ScreenUpdating = True
    
    RibbonCategories

End Sub

Private Sub MoveMyNote()

    Dim r As range
    
    For Each r In Intersect(range("A:A"), ActiveSheet.UsedRange)
        If r.value Like "My note:*" Then r.Cut Destination:=r.Offset(-2, 1)  'move My Note to column B and up 2 rows
    Next r

End Sub

Private Sub DeleteEbayNote()

    Dim r As range
    
    For Each r In Intersect(range("A:A"), ActiveSheet.UsedRange)
        If r.value Like "eBay note:*" Then r.EntireRow.Delete    'anything that's not a sku in column A gets deleted
    Next r

End Sub

Private Sub DeleteNonSKUs()

    Dim r As range
    
    For Each r In Intersect(range("A:A"), ActiveSheet.UsedRange)
        If r.value Like "Select this item for performing bulk action*" Then r.EntireRow.Delete    'anything that's not a sku in column A gets deleted
    Next r

End Sub

Private Sub DeleteRows()

    Dim EndRange As Integer
    Dim i As Integer
    
    'find the last row with value
    EndRange = LastRow("A:A")
    
    'go through every row
    For i = 2 To EndRange
        If IsEmpty(range("A" & i)) Then
            'if cell A is blank, delete the whole row
            range(i & ":" & i).EntireRow.Delete
            
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

    Dim r As Integer
    Dim i As Integer
    
    'find last row
    i = LastRow("A:A")
    
    'loop through all rows except the first one to reserve it for headers
    For r = 2 To i
        'if Column B starts with "My note" then delete the row
        If range("B" & r).value Like "My note:*" Then
            range("B" & r).EntireRow.Delete
            
            'Go back up one row to account for the rows shifting up after deleting
            r = r - 1
        End If
    Next r

End Sub

Private Sub Headers()

    range("A1").value = "Ebay SKU"

End Sub
