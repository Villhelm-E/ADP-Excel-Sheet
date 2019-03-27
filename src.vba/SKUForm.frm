Option Explicit

Private Sub UserForm_Deactivate()

    Unload Me

End Sub

Private Sub UserForm_Initialize()

    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    Me.SKUBox.SetFocus

End Sub

Private Sub FormatBtn_Click()

    If Me.SKUBox.Value = "" Then
        MsgBox "Please choose a manufacturer."
    Else
        SKU = Me.SKUBox.Value   'SKU is global variable
        Unload Me
    End If

End Sub

Private Sub CancelBtn_Click()

    SKU = ""
    
    Unload Me

End Sub

