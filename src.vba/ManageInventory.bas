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
    
    range("A:N").EntireColumn.AutoFit
    
    Application.ScreenUpdating = True
    
    'message user formatting is done
    MsgBox ("Formatting successful.")
    
    Exit Sub
    
Format_Error:
    MsgBox ("There was an error formatting the Amazon Listings.")
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
    range("1:2").EntireRow.Delete
    
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

    Dim row As Integer
    Dim numrows As Integer
    
    'count number of rows in A. Multiplying by 4 accounts for blank rows
    numrows = CountRows("D:D") * 4
    
    'loop through rows to find cells in column D that start with Variation. These are parents
    For row = 1 To numrows
        If Cells(row, 1).value Like "Variations*" Then
            'delete 3 rows. these are the rows of the parent Listing
            Rows(row).EntireRow.Delete
            Rows(row).EntireRow.Delete
            Rows(row).EntireRow.Delete
            row = row - 1   'go back 1 row in case there are two parent Listings in a row so the loop doesn't skip any
        End If
    Next row

End Sub

Private Sub MoveCells()

    Dim row As Integer
    Dim numrows As Integer
    
    numrows = CountRows("D:D") * 4
    
    For row = 1 To numrows
        If Not Cells(row, 4).value = "" Then
            Cells(row, 4).Cut Cells(row, 1)
            Cells(row + 1, 4).Cut Cells(row, 2)
            Cells(row + 1, 6).Cut Cells(row, 3)
            Cells(row, 6).Cut Cells(row, 4)
            Cells(row + 1, 7).Cut Cells(row, 5)
            Cells(row, 7).Cut Cells(row, 6)
            Cells(row, 8).Cut Cells(row, 7)
            Cells(row + 1, 8).Cut Cells(row, 8)
            Cells(row, 10).Cut Cells(row, 9)
            Cells(row + 1, 11).Cut Cells(row, 10)
            Cells(row + 2, 11).Cut Cells(row, 11)
            Cells(row + 1, 12).Cut Cells(row, 13)
            Cells(row + 2, 12).Cut Cells(row, 14)
        End If
    Next row

End Sub

Private Sub Headers()

    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    range("A1").value = "Status"
    range("B1").value = "Alert"
    range("C1").value = "Condition"
    range("D1").value = "SKU"
    range("E1").value = "ASIN"
    range("F1").value = "Title"
    range("G1").value = "Date Created"
    range("H1").value = "Status Changed Date"
    range("I1").value = "Fee Preview"
    range("J1").value = "Shipping"
    range("K1").value = "Shipping Template"
    range("L1").value = "Lowest Price"
    range("M1").value = "Lowest Price Shipping"
    range("N1").value = "Price Option"

End Sub

Private Sub DeleteBlank()

    columns("A:N").Sort Key1:=columns(7), Order1:=xlDescending, Header:=xlYes

End Sub

Public Sub FormatCells()

    range("G:H").NumberFormat = "mm/dd/yyyy hh:mm:ss"
    range("I:J").NumberFormat = "##0.00"
    range("L:M").NumberFormat = "##0.00"
    
    'count rows
    Dim numrows As Integer
    numrows = CountRows("A:A")
    
    'delete any rows with headers besides the first row
    Dim i As Integer
    For i = 2 To numrows
        If Cells(i, 3).value = "Condition" Then Rows(i).Delete
    Next i
    
    'remove final spaces
    Dim k As String
    For i = 1 To numrows
        If Cells(i, 4).value Like "* " Then
            k = Cells(i, 4).value
            Cells(i, 4).value = left(k, Len(k) - 1)
        End If
    Next i

End Sub
