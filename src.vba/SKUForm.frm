Option Explicit

Private Sub UserForm_Deactivate()

    Unload Me

End Sub

Private Sub UserForm_Initialize()

    'position the userform
    Call CenterForm(SKUForm)
    
    Me.SKUBox.SetFocus

End Sub

Private Sub FormatBtn_Click()

    If Me.SKUBox.value = "" Then
        MsgBox ("Please choose a manufacturer.")
    Else
        SKU = Me.SKUBox.value   'SKU is global variable
        Unload Me
    End If

End Sub

Private Sub CancelBtn_Click()

    SKU = ""
    
    Unload Me

End Sub
