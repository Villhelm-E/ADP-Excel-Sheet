Option Explicit

Public Sub ManageInventoryMain()
    
    On Error GoTo Format_Error
    
    Application.ScreenUpdating = False
    
    'Remove Certain cells
    AdjustCells
    
    'Remove parent Listings
    RemoveParents
    
    'Move cells around
    MoveCells
    
    'Add Headers
    Headers
    
'    rows(2).EntireRow.Delete
    
    'Delete blank rows
    DeleteBlank
    
    'Final Formatting
    FormatCells
    
    Range("A:N").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    'message user formatting is done
    MsgBox "Formatting successful."
    
    Exit Sub
    
Format_Error:
    MsgBox "There was an error formatting the Amazon Listings."
    Application.ScreenUpdating = True

End Sub

Private Sub AdjustCells()

Column_Loop:
    If Application.WorksheetFunction.CountA(columns(1)) = 0 Then
        columns(1).Delete
        GoTo Column_Loop
    End If
    
Row_Loop:
    If Application.WorksheetFunction.CountA(Rows(1)) = 0 Then
        Rows(1).Delete
        GoTo Row_Loop
    End If
    
    'delete top 2 rows
    Range("1:2").EntireRow.Delete
    
Column_Loop_2:
    If Application.WorksheetFunction.CountA(columns(1)) = 0 Then
        columns(1).Delete
        GoTo Column_Loop_2
    End If
    
    'insert 3 columns to the left
    columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

End Sub

Private Sub RemoveParents()

    Dim Row As Integer
    Dim numrows As Integer
    
    'count number of rows in A. Multiplying by 4 accounts for blank rows
    numrows = CountRows("D:D") * 4
    
    'loop through rows to find cells in column D that start with Variation. These are parents
    For Row = 1 To numrows
        If Cells(Row, 1).Value Like "Variations*" Then
            'delete 3 rows. these are the rows of the parent Listing
            Rows(Row).EntireRow.Delete
            Rows(Row).EntireRow.Delete
            Rows(Row).EntireRow.Delete
            Row = Row - 1   'go back 1 row in case there are two parent Listings in a row so the loop doesn't skip any
        End If
    Next Row

End Sub

Private Sub MoveCells()

    Dim Row As Integer
    Dim numrows As Integer
    
    numrows = CountRows("D:D") * 4
    
    For Row = 1 To numrows
        If Not Cells(Row, 4).Value = "" Then
            Cells(Row, 4).Cut Cells(Row, 1)
            Cells(Row + 1, 4).Cut Cells(Row, 2)
            Cells(Row + 1, 6).Cut Cells(Row, 3)
            Cells(Row, 6).Cut Cells(Row, 4)
            Cells(Row + 1, 7).Cut Cells(Row, 5)
            Cells(Row, 7).Cut Cells(Row, 6)
            Cells(Row, 8).Cut Cells(Row, 7)
            Cells(Row + 1, 8).Cut Cells(Row, 8)
            Cells(Row, 10).Cut Cells(Row, 9)
            Cells(Row + 1, 11).Cut Cells(Row, 10)
            Cells(Row + 2, 11).Cut Cells(Row, 11)
            Cells(Row + 1, 12).Cut Cells(Row, 13)
            Cells(Row + 2, 12).Cut Cells(Row, 14)
        End If
    Next Row

End Sub

Private Sub Headers()

    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("A1").Value = "Status"
    Range("B1").Value = "Alert"
    Range("C1").Value = "Condition"
    Range("D1").Value = "SKU"
    Range("E1").Value = "ASIN"
    Range("F1").Value = "Title"
    Range("G1").Value = "Date Created"
    Range("H1").Value = "Status Changed Date"
    Range("I1").Value = "Fee Preview"
    Range("J1").Value = "Shipping"
    Range("K1").Value = "Shipping Template"
    Range("L1").Value = "Lowest Price"
    Range("M1").Value = "Lowest Price Shipping"
    Range("N1").Value = "Price Option"

End Sub

Private Sub DeleteBlank()

    columns("A:N").Sort Key1:=columns(7), Order1:=xlDescending, Header:=xlYes

End Sub

Public Sub FormatCells()

    Range("G:H").NumberFormat = "mm/dd/yyyy hh:mm:ss"
    Range("I:J").NumberFormat = "##0.00"
    Range("L:M").NumberFormat = "##0.00"
    
    'count rows
    Dim numrows As Integer
    numrows = CountRows("A:A")
    
    'delete any rows with headers besides the first row
    Dim i As Integer
    For i = 2 To numrows
        If Cells(i, 3).Value = "Condition" Then Rows(i).Delete
    Next i
    
    'remove final spaces
    Dim k As String
    For i = 1 To numrows
        If Cells(i, 4).Value Like "* " Then
            k = Cells(i, 4).Value
            Cells(i, 4).Value = Left(k, Len(k) - 1)
        End If
    Next i

End Sub
