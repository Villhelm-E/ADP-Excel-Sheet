Option Explicit

Private Sub UserForm_Initialize()

    'position the userform
    Call CenterForm(RemoveKeep)
    
    'Set default values
    Me.BelowNum.value = 1
    Me.AboveNum.value = 49
    Me.MinRow.value = 2
    Me.MaxRow.value = 101
    
    'Format Text
    Me.RemoveBelowButton.Font.Size = 11
    Me.RemoveAboveButton.Font.Size = 11
    Me.BelowNum.Font.Size = 11
    Me.AboveNum.Font.Size = 11
    Me.BelowNum.TextAlign = fmTextAlignCenter
    Me.AboveNum.TextAlign = fmTextAlignCenter
    
    Me.KeepButton.Font.Size = 11
    Me.MinRow.Font.Size = 11
    Me.MaxRow.Font.Size = 11
    Me.MinRow.TextAlign = fmTextAlignCenter
    Me.MaxRow.TextAlign = fmTextAlignCenter

End Sub

Private Sub BelowNum_Change()

    If IsNumeric(Me.BelowNum) = True And Me.BelowNum <> "" Then
        Me.RemoveBelowButton.Enabled = True
    Else
        Me.RemoveBelowButton.Enabled = False
    End If

End Sub

Private Sub AboveNum_Change()

    If IsNumeric(Me.AboveNum) = True And Me.AboveNum <> "" Then
        Me.RemoveAboveButton.Enabled = True
    Else
        Me.RemoveAboveButton.Enabled = False
    End If

End Sub

Private Sub RemoveBelowButton_Click()

    'Setup
    Dim r As Integer
    Dim low As Integer
    r = 2
    low = Me.BelowNum

    'Turn screen updating off
    Application.ScreenUpdating = False
    
    'check user entered a value
    If Me.BelowNum <> "" Then
        'check user entered number
        If IsNumeric(Me.BelowNum) Then
            While Not Cells(r, 5).value = ""    'column 5 is QoH column
                If Cells(r, 5).value < low Then
                    Cells(r, 5).EntireRow.Delete
                Else
                    r = r + 1
                End If
            Wend
        End If
    End If
    
    'turn screen updating on
    Application.ScreenUpdating = True
    
    'close userform
    Unload Me

End Sub

Private Sub RemoveAboveButton_Click()

    'Setup
    Dim r As Integer
    Dim high As Integer
    r = 2
    high = Me.AboveNum

    'Turn screen updating off
    Application.ScreenUpdating = False
    
    'check user entered a value
    If Me.AboveNum <> "" Then
        'check user entered number
        If IsNumeric(Me.AboveNum) Then
            While Not Cells(r, 5).value = ""    'column 5 is QoH column
            'NEED TO ADD ERROR HANDLING FOR #VALUE IF I DO THE SUBLOCATION THING BEFORE THIS FUNCTION
                If Cells(r, 5).value > high Then
                    Cells(r, 5).EntireRow.Delete
                Else
                    r = r + 1
                End If
            Wend
        End If
    End If
    
    'turn screen updating on
    Application.ScreenUpdating = True
    
    'close userform
    Unload Me

End Sub

Private Sub MinRow_Change()

    If (IsNumeric(Me.MinRow) = True And Me.MinRow <> "") Or (IsNumeric(Me.MaxRow) = True And Me.MaxRow <> "") Then
        Me.KeepButton.Enabled = True
    Else
        Me.KeepButton.Enabled = False
    End If

End Sub

Private Sub MaxRow_Change()

    If (IsNumeric(Me.MinRow) = True And Me.MinRow <> "") Or (IsNumeric(Me.MaxRow) = True And Me.MaxRow <> "") Then
        Me.KeepButton.Enabled = True
    Else
        Me.KeepButton.Enabled = False
    End If

End Sub

Private Sub KeepButton_Click()

    Dim a As Integer
    Dim b As Integer
    Dim numrows As Integer
    
    'turn screen updating off
    Application.ScreenUpdating = False
    
    a = WorksheetFunction.Min(Me.MinRow.value, Me.MaxRow.value) 'set a to min
    b = WorksheetFunction.Max(Me.MinRow.value, Me.MaxRow.value) 'set b to max
    numrows = CountRows("B:B")  'Column B is Product ID column
    
    If a > 2 Then
        Rows("2:" & a - 1).EntireRow.Delete       'delete rows from 1 to a if user wants to remove some top rows
        Rows(b - a + 3 & ":" & CountRows("B:B")).EntireRow.Delete
    Else
        Rows(b + 1 & ":" & CountRows("B:B")).EntireRow.Delete   'delete from below b to end of sheet
    End If
    
    'turn screen updating on
    Application.ScreenUpdating = True
    
    'close userform
    Unload Me

End Sub

Private Sub UserForm_Deactivate()

    'close userform
    Unload Me

End Sub
