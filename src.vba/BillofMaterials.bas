Option Explicit

Public Sub BoMMain()
    
    Select Case CheckBoM
        'raw report downloaded from Finale
        Case "Raw"
            RawBoM
        'Formatted with the BoM presented in columns
        Case "Expanded"
            ExpandedBoM
        'Formatted with the BoM presented in rows, ready to upload to Finale
        Case "Compact"
            CompactBoM
        'something else
        Case Else
            MsgBox ("No BoM found")
            
            RibbonCategories
            
            Exit Sub
    End Select

End Sub

Private Sub RawBoM()

    'turn screen updating off
    Application.ScreenUpdating = False
    
    'move columns C-F up one row
    range("C2:F2").Delete xlShiftUp
    
    'Delete empty rows
    DeleteEmpty
    
    'Repeat Product ID
    RepeatProductID
    
    'replace the second Product ID with Component Product ID so that it's ready to be uploaded
    'idk why Finale doesn't name it this
    range("D1").value = "Component Product ID"
    
    'delete description fields
    range("B:B").EntireColumn.Delete
    range("D:D").EntireColumn.Delete
    
    Dim rowcount As Integer
    Dim rng As range
    Dim cell As range
    
    rowcount = CountRows("A:A")
    Set rng = range("B2:B" & rowcount)
    For Each cell In rng
        If cell.Text = "" Then cell.value = "0"
    Next
    
    'Autofit columns
    range("A:" & NumberToColumn(CountColumns("1:1"))).EntireColumn.AutoFit
    
    'turn screen updating on
    Application.ScreenUpdating = True
    
    'Refresh Ribbon
    RibbonCategories

End Sub

Private Sub ExpandedBoM()

    Dim BoM As Integer
    Dim row As Integer
    Dim numrows As Integer
    Dim NotePresent As Boolean
    
    If range("D1").value Like "note*" Then NotePresent = True
    
    Application.ScreenUpdating = False
    
    numrows = CountRows("A:A")
    row = 2 'start at row 2
    
    'insert rows between product ID and BoM for moving BoM
    If NotePresent = True Then
        BoM = (CountColumns("1:1") - 1) / 3
        range("B:D").EntireColumn.Insert
    Else
        BoM = (CountColumns("1:1") - 1) / 2
        range("B:C").EntireColumn.Insert
    End If
    
    '
    Dim i As Integer
    Dim BoMCount As Integer
    
    Do Until row > numrows
        BoMCount = 0
        If NotePresent = True Then
            For i = 1 To BoM
                If Cells(row, (i * 3) + 2) <> "" Then BoMCount = BoMCount + 1
            Next i
        Else
            For i = 1 To BoM
                If Cells(row, (i * 2) + 2) <> "" Then BoMCount = BoMCount + 1
            Next i
        End If
        
        If BoMCount > 1 Then
            Rows(row + 1 & ":" & row - 1 + BoMCount).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'after counting bill of materials, add one less row below current row
        End If
        
        Call MoveBoM(row, BoMCount, NotePresent)
        
        'adjust number of rows
        numrows = numrows + BoMCount - 1
        'skip some rows to move onto the next product id
        row = row + BoMCount
    Loop
    
    'delete extra columns after moving them
    If NotePresent = True Then
        range("E:" & NumberToColumn(CountColumns(range("1:1")) + 3)).Delete Shift:=xlLeft
    Else
        range("D:" & NumberToColumn(CountColumns(range("1:1")) + 2)).Delete Shift:=xlLeft
    End If
    
    range("B1").value = "Quantity"
    range("C1").value = "Component Product ID"
    If NotePresent = True Then range("D1").value = "Component note"
    
    Application.ScreenUpdating = True
    
    'Autofit columns
    range("A:" & NumberToColumn(CountColumns("1:1"))).EntireColumn.AutoFit
    
    RibbonCategories
    
    MsgBox ("Done")

End Sub

Private Sub CompactBoM()

    Dim prodid As String
    Dim row As Integer
    Dim count As Integer
    Dim BoM As Integer
    Dim maxBoM As Integer
    Dim NotePresent As Boolean
    
    Application.ScreenUpdating = False
    
    If (range("D1").value = "Component Note" Or range("D1").value = "Component note") Then NotePresent = True
    
    row = 2
    
    'go down row by row and move the BoM and its quantity to a new column for each BoM
    Do While range("A" & row).Text <> ""
        prodid = range("A" & row).Text
        BoM = 1
        Do While range("A" & row).Text = prodid
            If NotePresent = True Then
                'move the Component Product ID to a column based on the number of BoM
                range("C" & row).Cut range(NumberToColumn((BoM * 3) + 2) & row - BoM + 1)
                
                'move the Quantity to a column based on the number of BoM
                range("B" & row).Cut range(NumberToColumn((BoM * 3) + 3) & row - BoM + 1)
                
                'move the Component Note to a column based on the number of BoM
                range("D" & row).Cut range(NumberToColumn((BoM * 3) + 4) & row - BoM + 1)
            Else
                'move the Component Product ID to a column based on the number of BoM
                range("C" & row).Cut range(NumberToColumn((BoM * 2) + 2) & row - BoM + 1)
                
                'move the Quantity to a column based on the number of BoM
                range("B" & row).Cut range(NumberToColumn((BoM * 2) + 3) & row - BoM + 1)
            End If
            
            BoM = BoM + 1
            
            'while looping through all the BoM, find the max number of components
            If BoM > maxBoM Then
                maxBoM = BoM
            End If
            
            row = row + 1
        Loop
    Loop
    
    'delete 2nd and 3rd columns
    If NotePresent = True Then
        range("B:D").EntireColumn.Delete
    Else
        range("B:C").EntireColumn.Delete
    End If
    
    'delete empty rows
    'find the last row with value
    count = LastRow("A:A")
    
    'go through every row
    For row = 2 To count
        If IsEmpty(range("B" & row)) Then
            'if cell B is blank, delete the whole row
            range(row & ":" & row).EntireRow.Delete
            
            'go back one row to account for the rows shifting up after deleting
            row = row - 1
            
            'subtract one from EndRange so the loop doesn't go on longer than it has to and gets stuck in an infinite loop
            If row + 1 < count Then
                count = count - 1
            Else
                'if the EndRange catches up to current row in loop, stop so it doesn't get stuck in an infinite loop
                GoTo fields
            End If
        End If
    Next row
    
fields:
    'Name Fields
    For BoM = 1 To maxBoM - 1
        If NotePresent = True Then
            range(NumberToColumn((BoM * 3) - 1) & "1").value = "BoM " & BoM
            range(NumberToColumn(BoM * 3) & "1").value = "Qty " & BoM
            range(NumberToColumn((BoM * 3) + 1) & "1").value = "note " & BoM
        Else
            range(NumberToColumn(BoM * 2) & "1").value = "BoM " & BoM
            range(NumberToColumn((BoM * 2) + 1) & "1").value = "Qty " & BoM
        End If
    Next BoM
    
    Application.ScreenUpdating = True
    
    'Autofit columns
    range("A:" & NumberToColumn(CountColumns("1:1"))).EntireColumn.AutoFit
    
    RibbonCategories
    
    MsgBox ("Done")

End Sub

Private Sub DeleteEmpty()
    
    'Find the last non-blank cell in column D(4)
    Dim lRow As Long
    lRow = Cells(Rows.count, 4).End(xlUp).row
    
    'save range to variable
    Dim r As range, i As Long
    Set r = ActiveSheet.range("A1:F" & 5891)

    'loop through rows from bottom to top
    For i = lRow To 1 Step (-1)
        'if row is empty, delete it
        If WorksheetFunction.CountA(r.Rows(i)) = 0 Then r.Rows(i).Delete
    Next

End Sub

Private Sub RepeatProductID()

    'find last row in column D(4)
    Dim lRow As Long
    lRow = Cells(Rows.count, 4).End(xlUp).row
    
    'loop
    Dim i As Integer
    Dim prodid As String
    For i = 2 To lRow
        If Cells(i, 1).value <> "" Then
            prodid = Cells(i, 1).value
        Else
            Cells(i, 1).value = prodid
        End If
    Next i

End Sub

Private Sub MoveBoM(row As Integer, BoM As Integer, NotePresent As Boolean)
    
    Dim i As Integer
    Dim component As String
    Dim qty As String
    Dim note As String
    Dim BoMNum As Integer
    
    'repeat Product ID on a new row for each BoM
    If BoM > 1 Then range("A" & row & ":A" & row - 1 + BoM).value = Cells(row, 1).value
    
    If NotePresent = True Then
        For i = 1 To BoM
            Cells(row, (i * 3) + 2).Cut Cells(row + i - 1, 3)
            Cells(row, (i * 3) + 3).Cut Cells(row + i - 1, 2)
            Cells(row, (i * 3) + 4).Cut Cells(row + i - 1, 4)
        Next i
    Else
        For i = 1 To BoM
            Cells(row, (i * 2) + 2).Cut Cells(row + i - 1, 3)
            Cells(row, (i * 2) + 3).Cut Cells(row + i - 1, 2)
        Next i
    End If

End Sub

Public Sub ReplaceComp()

    ReplaceComponent.Show
    

End Sub
