Option Explicit

Private Sub UserForm_Initialize()

    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    'Load Variables
    LoadVariables

End Sub

Private Sub UserForm_Deactivate()

    rst.Close
    
    Unload Me

End Sub

Private Sub UpdateBtn_Click()
    
    'update the variables in the AmazonTemplateVariables table
    Set rst = MstrDb.Execute("UPDATE AmazonTemplateVariables SET AmazonTemplateVersion = " & Chr(39) & Me.TemplateVersionBox.Value & Chr(39) & ", AmazonTemplateSig = " & _
        Chr(39) & Me.TemplateSigBox.Value & Chr(39) & ", NameRow = " & Me.NameRowBox.Value & ", LabelRow = " & Me.LabelRowBox.Value & " WHERE ID = 1;")
        
    MsgBox "Updated"
    
    Exit Sub
        
Error_Msg:
    MsgBox "Failed to update variables"

End Sub

Private Sub LoadVariables()

    'load variables from AmazonTemplateVariables table
    Set rst = MstrDb.Execute("SELECT * FROM AmazonTemplateVariables")
    
    Me.TemplateVersionBox.Value = rst.Fields("AmazonTemplateVersion").Value
    Me.TemplateSigBox.Value = rst.Fields("AmazonTemplateSig").Value
    Me.NameRowBox.Value = rst.Fields("NameRow").Value
    Me.LabelRowBox.Value = rst.Fields("LabelRow").Value
    
    rst.Close

End Sub