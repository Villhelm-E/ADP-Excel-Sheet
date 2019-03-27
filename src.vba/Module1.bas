Option Explicit

Sub billofmaterialstest()
    
    Dim BoM As Integer
    Dim Row As Integer
    Dim numrows As Integer
    
    Application.ScreenUpdating = False
    
    numrows = CountRows("A:A")
    Row = 2 'start at row 2
    
    Do Until Row > numrows
        BoM = (CountColumns(Row) - 1) / 2 'count number of materials in bill of materials
        If BoM = 1 Then
            
        Else
            Rows(Row + 1 & ":" & Row - 1 + BoM).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'after counting bill of materials, add one less row below current row
        End If
        
        Call MoveBoM(Row, BoM)
        
        numrows = numrows + BoM - 1
        Row = Row + BoM
    Loop
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub MoveBoM(Row As Integer, BoM As Integer)
    
    Dim i As Integer
    
    'repeat Product ID
    If BoM = 1 Then
        
    Else
        Range("A" & Row & ":A" & Row - 1 + BoM).Value = Cells(Row, 1).Value
    End If
    
    If BoM = 1 Then
        
    Else
        For i = 1 To BoM - 1
            Cells(Row + i, 2).Value = Cells(Row, (2 * i) + 2).Value
            Cells(Row + i, 3).Value = Cells(Row, (2 * i) + 3).Value
        Next i
    End If

End Sub