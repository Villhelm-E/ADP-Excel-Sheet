Option Explicit

Private Sub UserForm_Initialize()

    'position the userform
    Call CenterForm(VariablesForm)
    
    'Load Variables
    LoadVariables

End Sub

Private Sub UserForm_Deactivate()

    rst.Close
    
    Unload Me

End Sub

Private Sub UpdateBtn_Click()
    
    'update the variables in the AmazonTemplateVariables table
    Set rst = MstrDb.Execute("UPDATE AmazonTemplateVariables SET AmazonTemplateVersion = " & Chr(39) & Me.TemplateVersionBox.value & Chr(39) & ", AmazonTemplateSig = " & _
        Chr(39) & Me.TemplateSigBox.value & Chr(39) & ", NameRow = " & Me.NameRowBox.value & ", LabelRow = " & Me.LabelRowBox.value & " WHERE ID = 1;")
        
    MsgBox ("Updated")
    
    Exit Sub
        
Error_Msg:
    MsgBox ("Failed to update variables")

End Sub

Private Sub LoadVariables()

    'load variables from AmazonTemplateVariables table
    Set rst = MstrDb.Execute("SELECT * FROM AmazonTemplateVariables")
    
    Me.TemplateVersionBox.value = rst.fields("AmazonTemplateVersion").value
    Me.TemplateSigBox.value = rst.fields("AmazonTemplateSig").value
    Me.NameRowBox.value = rst.fields("NameRow").value
    Me.LabelRowBox.value = rst.fields("LabelRow").value
    
    rst.Close

End Sub
