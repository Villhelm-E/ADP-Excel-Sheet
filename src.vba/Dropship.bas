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
        Do Until Application.CountA(range("A" & numEntries + 1).EntireRow) = 0
            range("A" & numEntries + 1).EntireRow.Delete
        Loop
    End If

End Sub

Private Sub HerkoDropshipMain()

    Dim numrows As Integer
    
    Application.ScreenUpdating = False
    
    numrows = CountRows("A:A")
    
    'clear the tax column, will replace with shipping cost from Shipstation
    range("H:H").Clear
    
    'rename Headers
    range("H1").value = "Shipping Cost"
    range("I1").value = "AD Total Price"
    range("J1").value = "Selling Price"
    range("K1").value = "Profit/Loss"
    
    'Fill in formulas
    range("I2:I" & numrows).Formula = "=G2+H2"
    range("K2:K" & numrows).Formula = "=(J2*88%)-I2"
    
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
    range("A:A").NumberFormat = "m/d/yy"
    range("F:K").NumberFormat = "$#,##0.00"
    
    'Conditional Formatting Profit column
    Conditionals
    
    'autofit/filter columns
    range("A:L").AutoFilter
    range("A:M").EntireColumn.AutoFit
    
    range("A1").Select

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
    Dim profitcol As range
    Dim custcol As range
    Dim uv As UniqueValues
    Dim cond As FormatCondition
    
    'define range
    Set profitcol = range("K2", range("K2").End(xlDown))
    
    'delete any existing conditional formatting
    profitcol.FormatConditions.Delete
    
    'define conditions
    Set cond = profitcol.FormatConditions.Add(xlCellValue, xlLessEqual, "=0")
    
    'apply conditions and formatting
    With cond
        .Interior.Color = 13551615
        .Font.Color = -16383844
    End With
    
    'apply conditional formatting to find duplicate customers
    Set custcol = range("B2", range("B2").End(xlDown))
    
    Set uv = custcol.FormatConditions.AddUniqueValues
    uv.DupeUnique = xlDuplicate
    With uv
        .Interior.Color = 13551615
        .Font.Color = -16383844
    End With

End Sub

Private Sub ShipstationDropshipMain()

    Application.ScreenUpdating = False
    
    Dim numrows As Integer
    
    numrows = CountRows("A:A")
    
    Call ShipstationFinalFormat(numrows)
    
    Application.ScreenUpdating = True
    
    RibbonCategories

End Sub

Private Sub ShipstationFinalFormat(numrows As Integer)

    FreezeTopRow
    
    range("A:A").NumberFormat = "mm/dd/yyyy"
    range("D:E").NumberFormat = "$#,##0.00"
    
    range("A:M").EntireColumn.AutoFit
    
    Call RenameSheet("Shipstation", numrows)
    
    range("A1").value = "Shipped Date"
    range("C1").value = "Ship To"
    range("D1").value = "Order Total"
    range("E1").value = "Shipping Cost"
    
    range("A1").Select

End Sub

Private Sub RenameSheet(Source As String, numrows As Integer)
    
    On Error GoTo dateError
    
    Dim yearStart As String
    Dim yearEnd As String
    
    yearStart = left(range("A2"), Len(range("A2")) - 4) & Right(range("A2"), 2)
    yearEnd = left(range("A" & numrows), Len(range("A" & numrows)) - 4) & Right(range("A" & numrows), 2)
    
    If yearStart = yearEnd Then
        ActiveSheet.name = Replace(Source & " " & yearStart, "/", "-")
    Else
        ActiveSheet.name = Replace(Source & " " & yearStart & "—" & yearEnd, "/", "-")
    End If
    
    Exit Sub
    
dateError:
    ActiveSheet.name = "Herko"
    
End Sub

Public Sub FindShipstationReport()

    Dim wsSheet As Worksheet
    Dim FoundSheet As Boolean
    'assume foundsheet is false
    FoundSheet = False
    
    For Each wsSheet In Worksheets
        If wsSheet.name Like "Shipstation*—*" Then
            FoundSheet = True
            GoTo Exit_Loop
        End If
    Next
Exit_Loop:
    If FoundSheet Then
        HerkoDropshipReports.Show
        Call ImportShipstationReport(ChosenSheet)
    Else
        MsgBox ("Couldn't find a Shipstation report to import.")
    End If

End Sub

Public Sub ImportShipstationReport(ChosenSheet As Worksheet)

    On Error Resume Next
    
    Dim HerkoReportName As String
    HerkoReportName = ChosenSheet.name
    
    Dim ShipstationNumRows As Integer
    ShipstationNumRows = CountRows("A:A")
    
    InsertDate
    
    'format date column
    FormatDate
    
    FormatPercent
    
    'Fill in Ship Date, Shipping Cost, and Selling Price
    Call Formula(HerkoReportName, ShipstationNumRows)
    
    'Add Percent Column
    range("M1").value = "Profit/Loss %"
    range("M2:M" & ShipstationNumRows).Formula = "=L2/J2"
    
    'Totals
    range("H" & ShipstationNumRows + 1).Formula = "=SUM(H2:H" & ShipstationNumRows & ")"
    range("L" & ShipstationNumRows + 1).Formula = "=SUM(L2:L" & ShipstationNumRows & ")"
    range("M" & ShipstationNumRows + 1).Formula = "=AVERAGE(M2:M" & ShipstationNumRows & ")"
    
    'resize the profit column
    range("A:M").EntireColumn.AutoFit
    
    Call RenameSheet("Herko", ShipstationNumRows)
    
    range("A1").Select

End Sub

Private Sub InsertDate()

    range("A1").EntireColumn.Insert
      
    range("A1").value = "Ship Date"

End Sub

Private Sub FormatDate()

    range("A:A").NumberFormat = "mm/dd/yyyy"

End Sub

Private Sub FormatPercent()

    range("M:M").NumberFormat = "0.00%"

End Sub

Private Sub Formula(HerkoReportName As String, LastRow As Integer)
    
    'final query
    range("A2:A" & LastRow).Formula = "=INDEX('" & HerkoReportName & "'!A:A,MATCH(C2,'" & HerkoReportName & "'!C:C,0))"
    range("I2:I" & LastRow).Formula = "=INDEX('" & HerkoReportName & "'!E:E,MATCH(C2,'" & HerkoReportName & "'!C:C,0))"
    range("K2:K" & LastRow).Formula = "=INDEX('" & HerkoReportName & "'!D:D,MATCH(C2,'" & HerkoReportName & "'!C:C,0))"
    
    'Replace formulas in Ship Date, Shipping Cost, and Selling Price with values
    With range("A2:A" & LastRow)
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False
    
    With range("I2:I" & LastRow)
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False
    
    With range("K2:K" & LastRow)
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False

End Sub
