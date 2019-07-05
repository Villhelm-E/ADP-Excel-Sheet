Option Explicit

'This Form is used for the "Export to find Sets" procedure
'Need to keep the part number and part type to export to Access

Private Sub UserForm_Initialize()

On Error GoTo UserForm_Initialize_Err
    
    'load parttypes into combobox
    LoadPartTypes
    
    'load manufacturers into combobox
    LoadBrands
    
    'run this to see if user needs to reenter info or not
    ReopenForm
    
    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
UserForm_Initialize_Exit:
    On Error Resume Next
    
    Exit Sub

UserForm_Initialize_Err:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error!"
    Resume UserForm_Initialize_Exit

End Sub

Private Sub LoadPartTypes()

    'Open ACESPartTypes Table in Master Database
    OpenACESPartTypes
    
    'list Part types into combobox
    With Me.PartTypeCombo
        .Clear
        Do
            .AddItem rst![ACESPartType]
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    'close table
    rst.Close

End Sub

Private Sub LoadBrands()
    
    'Open ACESPartTypes Table in Master Database
    OpenManufacturers
    
    'list Part types into combobox
    With Me.BrandCombo
        .Clear
        Do
            .AddItem rst![ManufacturerFull]
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    'close table
    rst.Close

End Sub

Private Sub ReopenForm()

    'Reopen is a global variable
    'if reopen = false, the user is entering the part number and interchange source for the first time
    'true means user has to reenter the info
    If Reopen = False Then
        Me.BrandLabel.Top = 66
        Me.BrandCombo.Top = 84
        Me.PartTypeLabel.Top = 120
        Me.PartTypeCombo.Top = 138
        Me.CancelBtn.Top = 174
        Me.FormatBtn.Top = 174
        Me.Height = 231
        Me.FitmentSourceCombo.Enabled = False
        Me.FitmentSourceLabel.Enabled = False
        Me.FitmentSourceCombo.Visible = False
        Me.FitmentSourceLabel.Visible = False
        Me.FormatBtn.Caption = "Format"
    Else
        'if user is reopening the form, set FitmentSource to null so that if user renames the part and then cancels the SourceForm, the cancel works properly
        FitmentSource = ""
        LoadFitmentSources
    End If

End Sub

Private Sub LoadFitmentSources()

    'Open FitmentSources Table in Master Database
    OpenFitmentSources
    
    'list sources into combobox
    With Me.FitmentSourceCombo
        .Clear
        Do
            .AddItem rst![Source]
            rst.MoveNext
        Loop Until rst.EOF
    End With
    
    'close table
    rst.Close

End Sub

Private Sub FormatBtn_Click()

    'Make sure user entered part number and part type
    If Me.PartNumBox.Value = "" Or Me.PartTypeCombo.Value = "" Then
        MsgBox "Please enter required fields"
    Else
        PartName = Me.PartNumBox.Value                          'PartName is global variable
        PartTypeVar = Me.PartTypeCombo.Value                    'PartTypeVar is global variable
        Brand = Me.BrandCombo.Value                             'Brand is global variable
        
        'save user entries to global variables
        If Me.InterchangeBox.Value = "" Then
            InterchangeSource = PartName                        'InterchangeSource is global variable
        Else
            InterchangeSource = Me.InterchangeBox.Value
        End If
        
        If Reopen = True And Me.FitmentSourceCombo.Enabled = True Then
            FitmentSource = Me.FitmentSourceCombo.Value         'FitmentSource is global variable
        End If
        
        'generate SKU
        Dim prefix As String
        Dim suffix As String
        
        'grab suffix code
        Set rst = MstrDb.Execute("SELECT SuffixCode FROM Manufacturers WHERE [ManufacturerFull] = " & Chr(34) & Me.BrandCombo.Value & Chr(34))
        suffix = rst.Fields("SuffixCode").Value
        rst.Close
        
        'grab prefix code
        Set rst = MstrDb.Execute("SELECT PrefixCode FROM AAIAPartTypes WHERE [AAIAPartType] = " & Chr(34) & Me.PartTypeCombo.Value & Chr(34))
        prefix = rst.Fields("PrefixCode").Value
        rst.Close
        
        gendSKU = prefix & "-" & Me.PartNumBox.Value & "-" & suffix
        
        'close
        'only unload userform if user has entered in all required fields
        Unload Me
    End If

End Sub

Private Sub CancelBtn_Click()

    'wipe variables
    PartName = ""
    InterchangeSource = ""
    Me.PartTypeCombo.Value = ""
    
    'Close form
    SourceForm.Hide

End Sub
