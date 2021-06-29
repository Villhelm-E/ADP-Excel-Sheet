Option Explicit

Public Sub FixFitments()

    'RawFitments function determines if there are raw fitments
    If RawFitments = False Then GoTo Exit_Format     'Checks Module
    
    'default Reopen global variable to false
    'false means this is the first time the user is openening the SourceForm userform
    Reopen = False                                                  'Reopen is a global variable
    
    'chooses subprocedure to run according to fitment source
    'FitmentSource is determined in RawFitments function above
    Select Case FitmentSource                                       'FitmentSource is Global variable
        
        Case "Metro"
            FormatMetro
        
        Case "Sixbit"
            FormatSixbit
            
        Case "Amazon"
           FormatAmazon
            
Exit_Format:
        Case Else
            MsgBox ("No fitments found.")
    End Select
    
    Formatted = True
    
    'Update ribbon
    RibbonCategories

End Sub

Private Sub FormatMetro()
    
    'Ask for part info
    SourceForm.Show
    
    'call the Metro Formatting Code if user entered info and did not cancel
    If Not PartName = "" And Not PartTypeVar = "" Then                          'PartName and PartTypeVar as Global variables
        Call MetroMain                                                          'Metro Module
    End If

End Sub

Private Sub FormatSixbit()
    
    'Ask for part info
    SourceForm.Show
    
    'Call the Sixbit Fromatting code if user entered info and did not cancel
    If Not PartName = "" And Not PartTypeVar = "" Then
        SixbitMain
    End If

End Sub

Private Sub FormatAmazon()
    
    'Ask for part info
    SourceForm.Show
    
    'Call the Sixbit Fromatting code if user entered info and did not cancel
    If Not PartName = "" And Not PartTypeVar = "" Then
        AmazonMain
    End If

End Sub
