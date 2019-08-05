Option Explicit

'Function takes a column as input and returns number of rows in range
Public Function CountRows(col) As Long

    'user can enter a range, an integer, or a string as a row range input
    Select Case VarType(col)
    Case 8204 'range
        CountRows = Application.CountA(col)
        
    Case 2  'integer
        CountRows = Application.CountA(columns(col))
        
    Case 8  'string
        If InStr(1, col, ":") > 0 Then
            CountRows = Application.CountA(Range(col))
        Else
            CountRows = Application.CountA(Range(col & ":" & col))
        End If
        
    Case Else
        CountRows = -1
    
    End Select

End Function

'Function takes a range as input and returns the last row in range
Public Function LastRow(col As String) As Long
    
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, col).End(xlUp).Row

End Function

'Function takes rows as input and returns the number of columns in range
Public Function CountColumns(RowRange) As Integer

    'user can enter a range, an integer, or a string as a row range input
    Select Case VarType(RowRange)
    Case 8204 'range
        CountColumns = Application.CountA(RowRange)
        
    Case 2  'integer
        CountColumns = Application.CountA(Rows(RowRange))
        
    Case 8  'string
        If InStr(1, RowRange, ":") > 0 Then
            CountColumns = Application.CountA(Range(RowRange))
        Else
            CountColumns = Application.CountA(Range(RowRange & ":" & RowRange))
        End If
        
    Case Else
        CountColumns = -1
        
    End Select

End Function

'Function takes column number as input and returns the column letter
Public Function NumberToColumn(Column As Integer) As String

    Dim vArr
    vArr = Split(Cells(1, Column).Address(True, False), "$")
    NumberToColumn = vArr(0)

End Function

'Function takes a double value and rounds up instead of at .5
Public Function RoundUp(Number As Double) As Double

    'if the double is a whole number/integer then don't do anything to it, otherwise round up
    If Int(Number) = Number Then
        RoundUp = Number
    Else
        RoundUp = Round(Number + 0.5)
    End If

End Function

'Function is meant to be used in Amazon Template
Public Function AmazonColumn(lastcolumnletter As String, SearchItem As String) As Integer

    Dim rFind As Range
    Dim R As Range
    
    NameRow = 3
    
    'define range to search
    Set R = Range("A" & NameRow & ":" & lastcolumnletter & NameRow)
    
    'search through range to find column of where field name is found
    Set rFind = R.Find(SearchItem, , , xlWhole, , , False, , False)
    If Not rFind Is Nothing Then
        AmazonColumn = rFind.Column
    End If

End Function
