Option Explicit

Private Sub UserForm_Initialize()

    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

    'Add Worksheets to Listbox
    Dim sheetsArr()
    ReDim sheetsArr(0)
    Call SearchSheets(sheetsArr, "Herko *-*-##*-*-##") 'add every sheet that matches that pattern to an array

    Dim i As Integer
    For i = LBound(sheetsArr) To UBound(sheetsArr)
        Me.ListBox1.AddItem sheetsArr(i)
    Next i

End Sub

Public Sub SearchSheets(sheetsArr, MatchString As String)

    Dim wsSheet As Worksheet
    Dim i As Integer
    
    For Each wsSheet In Worksheets
        If wsSheet.Name Like MatchString Then
            'if array already has values, then add a new entry to array
            If UBound(sheetsArr) > 0 Then ReDim Preserve sheetsArr(UBound(sheetsArr) + 1)
            'save the current worksheet's name to the last entry in the array
            sheetsArr(UBound(sheetsArr)) = wsSheet.Name
        End If
    Next

End Sub

Private Sub ImportButton_Click()
    
    Dim temp As String
    temp = ListBox1.Value
    
    Set ChosenSheet = Application.Worksheets(ListBox1.Value)
    
    Unload Me

End Sub