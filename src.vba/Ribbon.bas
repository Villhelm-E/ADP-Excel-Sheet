Option Explicit

Public Rib As IRibbonUI
Public MyTag As String

'Callback for customUI.onLoad
Private Sub RibbonOnLoad(Ribbon As IRibbonUI)
    Set Rib = Ribbon
    
    'run on start
    RibbonCategories
    
End Sub

Private Sub GetEnabledMacro(control As IRibbonControl, ByRef Enabled)
    If MyTag = "Enable" Then
        Enabled = True
    ElseIf control.Tag Like MyTag Then
        Enabled = True
    Else
        Enabled = False
    End If
End Sub

'Refresh the Ribbon
Private Sub RefreshRibbon(Tag As String)
    MyTag = Tag
    If Rib Is Nothing Then
        MsgBox "Error, Save and Reopen the workbook"
    Else
        Rib.Invalidate  'refreshes the ribbon
    End If
End Sub

'Check sheet values to determine which buttons to activate/deactivate
Public Sub RibbonCategories()
    Dim Tag As String

    If CheckBlank = True Then
        'check if blank page
        'this has to be the first check in the module
        'if you add additional checks, add them below this one
        Tag = "*f*"                                     'disable all but permanent buttons
    ElseIf CheckManageInv = True Then
        'check if Manage Inventory sheet
        Tag = "*i*"
    ElseIf RawFitments = True Then
        'check if fitments
        Tag = "*b*"                                     'enable format fitments button
    ElseIf CheckBoM = True Then
        'check if Bill of Materials report
        Tag = "*k*"                                     'enable Format BoM button
    ElseIf CheckOOS = True Then
        'check if out of stock sheet
        Tag = "*c*"                                     'enable OOS button
    ElseIf CheckWeekInvEnd = True Then
        'check if Weekly Inventory already formatted
        Tag = "*h*"
    ElseIf CheckWeekInv = True Then
        'check if weekly inventory sheet
        Tag = "*g*"
    ElseIf CheckAllFinaleProducts = True Then
        'check if All Finale Products report
        Tag = "*j*"
    ElseIf CheckACES = True Then
        'check if formatted fitments
        Tag = "*a*"                                     'enable To Database and To Sixbit buttons
    ElseIf CheckUPC = True Then
        'check if new UPC list
        Tag = "*l*"
    ElseIf CheckDropship <> "" Then
        'check if Dropship report
        Tag = "*m*"
    ElseIf CheckFormattedDropship = True Then
        'check if formatted dropship report
        Tag = "*n*"
    ElseIf CheckVolumePricing = True Then
        'check if volume pricing template
        Tag = "*o*"
    Else
        'if not empty and not any of the above formats
        'this has to be the last check in this module
        'if you add additional checks, add them above this one
        Tag = "*e*"                                 'enable export buttons
    End If
    
    'run the ribbon refresh with the appropriate tag to enable/disable appropriate buttons in ribbon
    Call RefreshRibbon(Tag)
    
End Sub
