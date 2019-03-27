Option Explicit

Public Sub FormatReportMain()

    'Turn screen updating off
    Application.ScreenUpdating = False
    
    'shift rows up to align with Product ID
    Range("B2:F2").Delete Shift:=xlUp
    
    'Remove blank rows
    CleanStock
    
    'Count number of Rows
    Dim numrows As Integer
    numrows = CountRows("A:A")
    
    'Delete Reorder point
    columns("E:E").Delete
    
    'create column to the left
    columns("A:A").Insert Shift:=xlToRight
    
    'Create Count Column
    Range("A1").Value = "Count"
    
    'recount rows
    numrows = CountRows("A:A")
    
    'Add borders
    Call Borders(numrows)
    
    'Add Filters
    Range("A:F").AutoFilter
    
    'AutoFit
    columns("A:G").EntireColumn.AutoFit
    
    'Select Cell A2
    Range("A2").Select
    
    'Turn screen updating on
    Application.ScreenUpdating = True
    
    'Refresh Ribbon
    RibbonCategories

End Sub

Private Sub CleanStock()

    'First sort by Quantity on Hand ascending
    columns("A:F").Sort Key1:=columns(1), Order1:=xlAscending, Header:=xlYes

End Sub

Private Sub Borders(numrows As Integer)

    Range("A1:F" & numrows).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub RemoveKeepMain()

    RemoveKeep.Show

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub ConfirmInvMain()

    'turn off screen updating
    Application.ScreenUpdating = False
    
    'count the number of rows
    Dim numrows As Integer
    numrows = CountRows("A:A")
    
    'Save range in ADP Excel Sheet to variable
    Dim ADPRange As Range
    Set ADPRange = Range("A2:A" & numrows)
    
    'Open the Confirmed Inventory workbook on the Server
    'Screenupdating is off so this opening of the workbook will not be visible
    Dim Confirmed As Workbook
    Dim Sht As Worksheet
    Set Confirmed = Workbooks.Open("\\ADP-SERVER\AD AutoParts Server (Temp)\Inventory\Confirmed Inventory.xlsx")
    Set Sht = Confirmed.Worksheets("Sheet1")
    
    'count the number of rows in the Confirmed Inventory workbook
    Dim k As Integer
    k = Application.WorksheetFunction.CountA(Sht.Range("A:A")) + 1
    
    'Save the first empty cell in the Confirmed Inventory workbook to variable
    Dim ConfirmedRange As Range
    Set ConfirmedRange = Sht.Range("A" & k)
    
    'Copy the values from ADP Excel Sheet and paste in the Confirmed Inventory workbook
    ADPRange.Copy Destination:=ConfirmedRange
    
    'Close Confirmed Inventory workbook and save
    Confirmed.Close SaveChanges:=True
    
    'Select cell A1
    Range("A1").Select
    
    'Turn screen updating on
    Application.ScreenUpdating = True

End Sub

Private Sub ConfirmedInventory(numrows)

    'Create Confirmed Inventory Column
    Range("G1").Value = "Done"
    
    'Do Vlookup in G1
    Range("G2").FormulaR1C1 = _
        "=VLOOKUP(RC[-6],'\\ADP-SERVER\AD AutoParts Server (Temp)\Inventory\[Confirmed Inventory.xlsx]Sheet1'!Confirmed,1,FALSE)"
    
    'autofill formula
    Range("G2").AutoFill Destination:=Range("G2:G" & numrows), Type:=xlFillDefault
    
    'convert formulas to values
    Range("G1:G" & numrows).Cells.Copy
    Range("G1").Cells.PasteSpecial xlPasteValues
    
    'Delete confirmed inventory rows
    Dim i As Integer
    For i = 2 To numrows
        
        If Not IsError(Range("G" & i).Value) Then
            Rows(i).Delete
            If Not IsError(Range("G" & i)) Then
                If Range("G" & i).Value = "" Then
                    i = numrows
                End If
            Else
                i = i - 1
            End If
        End If
    Next i
    
    'Delete Confirmed column
    columns("G:G").Delete

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub RemoveInactiveMain()

    'Start row variable
    Dim i As Integer
    i = 2
    
    'turn off screen updating
    Application.ScreenUpdating = False
    
    'run through every row until empty
    Do While Cells(i, 2).Value <> ""
        'delete row if value in B is "Inactive"
        If Cells(i, 2).Value = "Inactive" Then
            Cells(i, 2).EntireRow.Delete
        Else
            'Move to next row
            i = i + 1
        End If
    Loop
    
    'turn on screen updating
    Application.ScreenUpdating = True
    
    'Refresh Ribbon
    RibbonCategories

End Sub

