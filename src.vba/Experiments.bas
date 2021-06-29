Option Explicit

Sub billofmaterialstest()
    
    Dim BoM As Integer
    Dim row As Integer
    Dim numrows As Integer
    
    Application.ScreenUpdating = False
    
    numrows = CountRows("A:A")
    row = 2 'start at row 2
    
    Do Until row > numrows
        BoM = (CountColumns(row) - 1) / 2 'count number of materials in bill of materials
        If BoM = 1 Then
            
        Else
            Rows(row + 1 & ":" & row - 1 + BoM).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove 'after counting bill of materials, add one less row below current row
        End If
        
        Call MoveBoM(row, BoM)
        
        numrows = numrows + BoM - 1
        row = row + BoM
    Loop
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub MoveBoM(row As Integer, BoM As Integer)
    
    Dim i As Integer
    
    'repeat Product ID
    If BoM = 1 Then
        
    Else
        range("A" & row & ":A" & row - 1 + BoM).value = Cells(row, 1).value
    End If
    
    If BoM = 1 Then
        
    Else
        For i = 1 To BoM - 1
            Cells(row + i, 2).value = Cells(row, (2 * i) + 2).value
            Cells(row + i, 3).value = Cells(row, (2 * i) + 3).value
        Next i
    End If

End Sub

'DOMAIN function limits the domain of a formula
Function DOMAIN(value As Variant, domain_value As Variant, low_limit As Double, high_limit As Double, Optional outside_value As Double, Optional invert As Boolean)

    If domain_value < low_limit Or domain_value > high_limit Then
        If IsNull(outside_value) Then
            DOMAIN = 0
        Else
            DOMAIN = outside_value
        End If
    Else
        DOMAIN = value
    End If

End Function

Function SKUCODE(target As String) As String

    'constants
    Const capitalLetter = 4
    Const wordInitial = 2
    Const syllableInitial = 2
    
    'initialize the array
    Dim charArray()
    ReDim charArray(Len(target) - 1, 1)
    
    'populate the array with the characters in the target
    Dim c As Integer
    For c = 0 To UBound(charArray)
        charArray(c, 0) = Mid(target, c + 1, 1)
        charArray(c, 1) = 0
    Next c
    
    'Prioritize capital letters
    For c = 0 To UBound(charArray)
        If charArray(c, 0) Like "*[A-Z]*" Then charArray(c, 1) = charArray(c, 1) + capitalLetter
    Next c
    
    'prioritize word-initial characters
    charArray(0, 1) = charArray(0, 0) + 1
    For c = 1 To UBound(charArray)
        If charArray(c, 0) <> " " And charArray(c - 1, 0) = " " Then charArray(c, 1) = charArray(0, 1) + wordInitial
    Next c
    
    'prioritize beginning of syllables
    
    
    'prioritize rare characters
    

End Function
