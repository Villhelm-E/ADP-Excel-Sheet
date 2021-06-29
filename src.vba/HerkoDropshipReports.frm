Option Explicit

Private Sub UserForm_Initialize()

    'position the userform
    Call CenterForm(HerkoDropshipReports)

    'Add Worksheets to Listbox
    Dim sheetsArr()
    ReDim sheetsArr(0)
    Call SearchSheets(sheetsArr, "Shipstation*##*—*##") 'add every sheet that matches that pattern to an array

    Dim i As Integer
    For i = LBound(sheetsArr) To UBound(sheetsArr)
        Me.ListBox1.AddItem sheetsArr(i)
    Next i

End Sub

Public Sub SearchSheets(sheetsArr, MatchString As String)

    Dim wsSheet As Worksheet
    Dim i As Integer
    
    For Each wsSheet In Worksheets
        If wsSheet.name Like MatchString Then
            'if array already has values, then add a new entry to array
            If IsEmpty(sheetsArr(0)) = False Then ReDim Preserve sheetsArr(UBound(sheetsArr) + 1)
            'save the current worksheet's name to the last entry in the array
            sheetsArr(UBound(sheetsArr)) = wsSheet.name
        End If
    Next

End Sub

Private Sub ImportButton_Click()
    
    Dim temp As String
    temp = ListBox1.value
    
    Set ChosenSheet = Application.Worksheets(ListBox1.value)
    
    Unload Me

End Sub
