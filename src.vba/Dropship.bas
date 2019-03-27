Option Explicit

'variable for user's chosen sheet
Public ChosenSheet As Worksheet

Public Sub DropshipMain()

    Application.ScreenUpdating = False
    
    'initial cleanup
    InitialCleanup
    
    'determine what Dropship report the user is dealing with
    Select Case CheckDropship
        Case "Herko"
            HerkoDropshipMain
        Case "Shipstation"
            ShipstationDropshipMain
    End Select
    
    Application.ScreenUpdating = True

End Sub

Private Sub InitialCleanup()

    Dim numrows As Integer
    Dim numEntries As Integer
    
    numrows = CountRows("A:H")
    numEntries = CountRows("A:A")
    
    If numrows <> numEntries Then
        Do Until Application.CountA(Range("A" & numEntries + 1).EntireRow) = 0
            Range("A" & numEntries + 1).EntireRow.Delete
        Loop
    End If

End Sub

Private Sub HerkoDropshipMain()

    Dim numrows As Integer
    
    Application.ScreenUpdating = False
    
    numrows = CountRows("A:A")
    
    'rename Headers
    Range("H1").Value = "Shipping"
    Range("I1").Value = "Herko Total Price"
    Range("J1").Value = "Selling Price"
    Range("K1").Value = "Net Selling Price"
    Range("L1").Value = "Profit"
    
    'Fill in formulas
    Range("F2:F" & numrows).Formula = "=D2*E2"
    Range("I2:I" & numrows).Formula = "=F2+H2"
    Range("K2:K" & numrows).Formula = "=J2*88%"
    Range("L2:L" & numrows).Formula = "=K2-I2"
    
    'pretty it up
    Call HerkoFinalFormat(numrows)
    
    Application.ScreenUpdating = True
    
    'Refresh Ribbon
    RibbonCategories

End Sub

Private Sub HerkoFinalFormat(numrows As Integer)

    'freeze top row
    FreezeTopRow
    
    'format date and currency columns
    Range("A:A").NumberFormat = "m/d/yy"
    Range("E:F,H:L").NumberFormat = "$#,##0.00"
    
    'Uppercase Names
    Dim i As Integer
    
    For i = 2 To numrows
        Cells(i, 2).Value = UCase(Cells(i, 2))
    Next i
    
    'Conditional Formatting Profit column
    Conditionals
    
    'autofit/filter columns
    Range("A:L").AutoFilter
    Range("A:L").EntireColumn.AutoFit
    
    'Sort by date
    Call Sort(numrows)
    
    'rename sheet
    Call RenameSheet("Herko", numrows)
    
    Range("A1").Select

End Sub

Private Sub FreezeTopRow()

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

End Sub

Private Sub Conditionals()

    'variables
    Dim profitcol As Range
    Dim cond As FormatCondition
    
    'define range
    Set profitcol = Range("L2", Range("L2").End(xlDown))
    
    'delete any existing conditional formatting
    profitcol.FormatConditions.Delete
    
    'define conditions
    Set cond = profitcol.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    
    'apply conditions and formatting
    With cond
        .Interior.Color = 13551615
        .Font.Color = -16383844
    End With

End Sub

Private Sub ShipstationDropshipMain()

    Application.ScreenUpdating = False
    
    Dim numrows As Integer
    
    Range("B:B,D:X,Z:AA,AC:AW,BA:BC,BE:BE").EntireColumn.Delete
    
    ShipstationReorder
    
    numrows = CountRows("A:A")
    
    Dim i As Integer
    
    'capitalize customer names
    For i = 2 To numrows
        Cells(i, 2).Value = UCase(Cells(i, 2))
    Next i
    
    'add columns
    Range("I1").Value = "Herko Total Price"
    Range("J1").Value = "Selling Price"
    Range("K1").Value = "Net Selling Price"
    Range("L1").Value = "Profit"
    
    'fill formulas
    Call JKLFormulas(numrows)
    
    Call ShipstationFinalFormat(numrows)
    
    Application.ScreenUpdating = True
    
    RibbonCategories

End Sub

Private Sub JKLFormulas(numrows As Integer)

    Range("J2:J" & numrows).Formula = "=C2-E2"
    Range("K2:K" & numrows).Formula = "=J2*88%"
    Range("L2:L" & numrows).Formula = "=IF(I2=" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & ",K2-I2)"

End Sub

Private Sub ShipstationReorder()

    'move Date column H to column A
    Application.CutCopyMode = False             'don't want an existing operation to interfere
    columns("A").Insert XlDirection.xlToLeft
    columns("A").Value = columns("I").Value     'Inserting column left of A shifts column H to I
    columns("I").Delete
    
    'move bill - to Customer column from C to B
    columns("B").Insert XlDirection.xlToLeft
    columns("B").Value = columns("D").Value     'Inserting column left of B shifts column C to D
    columns("D").Delete
    
    Application.CutCopyMode = True

End Sub

Private Sub ShipstationFinalFormat(numrows As Integer)

    FreezeTopRow
    
    Range("A2:A" & numrows).NumberFormat = "m/d/yyyy"
    Range("C:F,H:L").NumberFormat = "$#,##0.00"
    
    'Conditional Formatting Profit column
    Conditionals
    
    Range("A:L").EntireColumn.AutoFit
    
    Call Sort(numrows)
    
    Call RenameSheet("Shipstation", numrows)
    
    Range("A1").Select

End Sub

Private Sub Sort(numrows As Integer)

    Range("A1:M" & numrows).Sort Key1:=Range("C1"), Order1:=xlAscending, Header:=xlYes

End Sub

Private Sub RenameSheet(Source As String, numrows As Integer)

    Dim yearStart As String
    Dim yearEnd As String
    
    yearStart = Left(Range("A2"), Len(Range("A2")) - 4) & Right(Range("A2"), 2)
    yearEnd = Left(Range("A" & numrows), Len(Range("A" & numrows)) - 4) & Right(Range("A" & numrows), 2)
    
    
    ActiveSheet.Name = Replace(Source & " " & yearStart & "" & yearEnd, "/", "-")
    
End Sub

Public Sub FindHerkoReport()

    Dim wsSheet As Worksheet
    Dim FoundSheet As Boolean
    'assume foundsheet is false
    FoundSheet = False
    
    For Each wsSheet In Worksheets
        If wsSheet.Name Like "Herko **" Then
            FoundSheet = True
        End If
    Next

    If FoundSheet Then
        HerkoDropshipReports.Show
        Call ImportHerkoReport(ChosenSheet)
    Else
        MsgBox "Couldn't find a Herko dropship report to import."
    End If

End Sub

Public Sub ImportHerkoReport(ChosenSheet As Worksheet)

    Dim HerkoReportName As String
    HerkoReportName = ChosenSheet.Name
    
    Dim ShipstationNumRows As Integer
    ShipstationNumRows = CountRows("A:A")
    
    Call Formula(HerkoReportName, ShipstationNumRows)
    
    Call ExtraOrders(ChosenSheet, ShipstationNumRows)
    
    If Cells(ShipstationNumRows + 1, 2).Value <> "" Then
        ShipstationNumRows = CountRows("A:A")
        Call Formula(HerkoReportName, ShipstationNumRows)
        Call JKLFormulas(ShipstationNumRows)
        Conditionals
    End If
    
    CompareShipping
    
    'resize the profit column
    Range("L:L").EntireColumn.AutoFit

End Sub

Public Sub FindShipstationReport()

    

End Sub

Private Sub CompareShipping()

    'compare the shipping cost from Herko to shipping cost in Shipstation report
    'if shipstation and herko reports both have a shipping cost, something is wrong. highlight it
    'if one report has a shipping cost, the other report should show 0 shipping cost for the same order

End Sub

Private Sub Formula(HerkoReportName As String, LastRow As Integer)

    Dim nestform As String
    'first query to be nested
    nestform = "INDEX('" & HerkoReportName & "'!I:I,MATCH(B2,'" & HerkoReportName & "'!B:B,0))"
    
    'final query
    Range("I2:I" & LastRow).Formula = "=IF(ISERROR(" & nestform & "),IF(H2=0," & Chr(34) & Chr(34) & ",H2)," & nestform & ")"

End Sub

Private Sub ExtraOrders(ChosenSheet As Worksheet, ShipstationNumRows As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim found As Boolean
    Dim HerkoNumRows As Integer
    
    k = ShipstationNumRows + 1
    HerkoNumRows = CountRows("'" & ChosenSheet.Name & "'!A:A")
    
    For i = 2 To HerkoNumRows - 1
        For j = 2 To ShipstationNumRows - 1
            If InStr(1, Cells(j, 2).Value, ChosenSheet.Cells(i, 2).Value) > 0 Or InStr(1, ChosenSheet.Cells(i, 2).Value, Cells(j, 2).Value) > 0 Then
                GoTo exit_i
            End If
        Next j
        Cells(k, 1).Value = ChosenSheet.Cells(i, 1).Value
        Cells(k, 2).Value = ChosenSheet.Cells(i, 2).Value
        Cells(k, 8).Value = ChosenSheet.Cells(i, 8).Value
        Cells(k, 9).Value = ChosenSheet.Cells(i, 9).Value
        k = k + 1
exit_i:
    Next i

End Sub