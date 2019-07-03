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
    
    'clear the tax column, will replace with shipping cost from Shipstation
    Range("H:H").Clear
    
    'rename Headers
    Range("H1").Value = "Shipping Cost"
    Range("I1").Value = "AD Total Price"
    Range("J1").Value = "Selling Price"
    Range("K1").Value = "Profit/Loss"
    
    'Fill in formulas
    Range("I2:I" & numrows).Formula = "=G2+H2"
    Range("K2:K" & numrows).Formula = "=(J2*88%)-I2"
    
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
    Range("F:K").NumberFormat = "$#,##0.00"
    
    'Conditional Formatting Profit column
    Conditionals
    
    'autofit/filter columns
    Range("A:K").AutoFilter
    Range("A:K").EntireColumn.AutoFit
    
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
    Dim custcol As Range
    Dim uv As UniqueValues
    Dim cond As FormatCondition
    
    'define range
    Set profitcol = Range("K2", Range("K2").End(xlDown))
    
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
    Set custcol = Range("B2", Range("B2").End(xlDown))
    
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
    
    Range("A:A").NumberFormat = "mm/dd/yyyy"
    Range("D:E").NumberFormat = "$#,##0.00"
    
    Range("A:E").EntireColumn.AutoFit
    
    Call RenameSheet("Shipstation", numrows)
    
    Range("A1").Select

End Sub

Private Sub RenameSheet(Source As String, numrows As Integer)
    
    On Error GoTo dateError
    
    Dim yearStart As String
    Dim yearEnd As String
    
    yearStart = Left(Range("A2"), Len(Range("A2")) - 4) & Right(Range("A2"), 2)
    yearEnd = Left(Range("A" & numrows), Len(Range("A" & numrows)) - 4) & Right(Range("A" & numrows), 2)
    
    If yearStart = yearEnd Then
        ActiveSheet.Name = Replace(Source & " " & yearStart, "/", "-")
    Else
        ActiveSheet.Name = Replace(Source & " " & yearStart & "" & yearEnd, "/", "-")
    End If
    
    Exit Sub
    
dateError:
    ActiveSheet.Name = "Herko"
    
End Sub

Public Sub FindShipstationReport()

    Dim wsSheet As Worksheet
    Dim FoundSheet As Boolean
    'assume foundsheet is false
    FoundSheet = False
    
    For Each wsSheet In Worksheets
        If wsSheet.Name Like "Shipstation**" Then
            FoundSheet = True
            GoTo Exit_Loop
        End If
    Next
Exit_Loop:
    If FoundSheet Then
        HerkoDropshipReports.Show
        Call ImportShipstationReport(ChosenSheet)
    Else
        MsgBox "Couldn't find a Shipstation report to import."
    End If

End Sub

Public Sub ImportShipstationReport(ChosenSheet As Worksheet)

    On Error Resume Next
    
    Dim HerkoReportName As String
    HerkoReportName = ChosenSheet.Name
    
    Dim ShipstationNumRows As Integer
    ShipstationNumRows = CountRows("A:A")
    
    InsertDate
    
    'format date column
    FormatDate
    
    'Fill in Ship Date, Shipping Cost, and Selling Price
    Call Formula(HerkoReportName, ShipstationNumRows)
    
    'resize the profit column
    Range("A:L").EntireColumn.AutoFit
    
    Call RenameSheet("Herko", ShipstationNumRows)
    
    Range("A1").Select

End Sub

Private Sub InsertDate()

    Range("A1").EntireColumn.Insert
      
    Range("A1").Value = "Ship Date"

End Sub

Private Sub FormatDate()

    Range("A:A").NumberFormat = "mm/dd/yyyy"

End Sub

Private Sub Formula(HerkoReportName As String, LastRow As Integer)
    
    'final query
    Range("A2:A" & LastRow).Formula = "=INDEX('" & HerkoReportName & "'!A:A,MATCH(C2,'" & HerkoReportName & "'!C:C,0))"
    Range("I2:I" & LastRow).Formula = "=INDEX('" & HerkoReportName & "'!E:E,MATCH(C2,'" & HerkoReportName & "'!C:C,0))"
    Range("K2:K" & LastRow).Formula = "=INDEX('" & HerkoReportName & "'!D:D,MATCH(C2,'" & HerkoReportName & "'!C:C,0))"
    
    'Replace formulas in Ship Date, Shipping Cost, and Selling Price with values
    With Range("A2:A" & LastRow)
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False
    
    With Range("I2:I" & LastRow)
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False
    
    With Range("K2:K" & LastRow)
        .Cells.Copy
        .Cells.PasteSpecial xlPasteValues
        .Cells(1).Select
    End With
    Application.CutCopyMode = False

End Sub
