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
            CountRows = Application.CountA(range(col))
        Else
            CountRows = Application.CountA(range(col & ":" & col))
        End If
        
    Case Else
        CountRows = -1
    
    End Select

End Function

'Used for centering the referenced userform to the user's screen
Public Sub CenterForm(ByRef frm)

    'position the userform
    frm.startupposition = 0
    frm.left = Application.left + (0.5 * Application.width) - (0.5 * frm.width)
    frm.top = Application.top + (0.5 * Application.height) - (0.5 * frm.height)

End Sub

'Function takes a range as input and returns the last row in range
Public Function LastRow(col As String) As Long
    
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.count, col).End(xlUp).row

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
            CountColumns = Application.CountA(range(RowRange))
        Else
            CountColumns = Application.CountA(range(RowRange & ":" & RowRange))
        End If
        
    Case Else
        CountColumns = -1
        
    End Select

End Function

'Function takes column number as input and returns the column letter
Public Function NumberToColumn(column As Integer) As String

    Dim vArr
    vArr = Split(Cells(1, column).Address(True, False), "$")
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

    Dim rFind As range
    Dim r As range
    
    NameRow = 3
    
    'define range to search
    Set r = range("A" & NameRow & ":" & lastcolumnletter & NameRow)
    
    'search through range to find column of where field name is found
    Set rFind = r.Find(SearchItem, , , xlWhole, , , False, , False)
    If Not rFind Is Nothing Then
        AmazonColumn = rFind.column
    End If

End Function

Public Function CHECKSUM(Num As String) As String

    Dim NumLength As Integer
    NumLength = Len(Num)
    Dim sum As Integer
    Dim i As Integer
    
    For i = NumLength To 1 Step -2
        sum = sum + Mid(Num, i, 1) * 3
    Next i
    
    For i = (NumLength - 1) To 1 Step -2
        sum = sum + Mid(Num, i, 1)
    Next i
    
    CHECKSUM = Num & WorksheetFunction.Ceiling(sum, 10) - sum

End Function

Public Sub ComputerName()

    MsgBox (CreateObject("WScript.Network").ComputerName)

End Sub
