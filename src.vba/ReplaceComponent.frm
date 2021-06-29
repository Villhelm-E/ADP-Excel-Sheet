Option Explicit

'used to save the state of the form
Private defaultComboWidth As Double
Private defaultComboHeight As Double
Private defaultFormHeight As Double
Private defaultTop As Double
Private defaultBtnTop As Double
Private defaultCombo1Left As Double
Private defaultCombo2Left As Double
Private gapBetweenCombos As Double
Private componentCount As Integer
Private replaceCount As Integer
Private defaultTextWidth As Double

Private Sub UserForm_Initialize()

    'Position userform
    Call CenterForm(ReplaceComponent)
    
    'define range of component IDs
    Dim comps As range
    Set comps = ActiveSheet.range("C2:C" & (CountRows("C:C") - 1)) 'select all rows in column C except first row
    
    'populate the "Component to Replace" Combobox
    Dim r As range
    For Each r In comps
        Call PopulateComboBox(r.value)
    Next r
    
    'enter the values into the Replace Combobox
    Call CopyComboBoxValues(Me.ReplaceCombo1)
    
    'Save Default UserForm values
    defaultComboWidth = CDbl(Me.ComponentCombo1.width)
    defaultComboHeight = CDbl(Me.ComponentCombo1.height)
    defaultFormHeight = CDbl(ReplaceComponent.height)
    defaultTop = CDbl(Me.ComponentCombo1.top)
    defaultBtnTop = CDbl(Me.ReplaceBtn.top)
    defaultCombo1Left = CDbl(Me.ComponentCombo1.left)
    defaultCombo2Left = CDbl(Me.ReplaceCombo1.left)
    defaultTextWidth = CDbl(Me.ComponentQty1.width)
    
    'define gap between comboboxes to determine how much to resize everything when adding comboboxes
    Const gapBetweenCombos = 4.5
    
    'Initialize Buttons
    Me.RemoveComponent1Btn.Enabled = False
    Me.RemoveComponent1Btn.Visible = False
    Me.RemoveComponent2Btn.Enabled = False
    Me.RemoveComponent2Btn.Visible = False
    
    Me.AddComponent1Btn.top = CStr(defaultTop)
    Me.AddComponent2Btn.top = CStr(defaultTop)
    
    'set default count values
    componentCount = 1
    replaceCount = 1

End Sub

Private Sub AddComponent1Btn_Click()
    
    'increment component Count
    componentCount = componentCount + 1
    
    'Add the combobox for the new component combobox
    AddComponent

End Sub

Private Sub AddComponent2Btn_Click()
    
    'increment component Count
    replaceCount = replaceCount + 1
    
    'Add the combobox for the new replace combobox
    AddReplace

End Sub

Private Sub RemoveComponent1Btn_Click()

    RemoveComponent

End Sub

Private Sub RemoveComponent2Btn_Click()
    
    RemoveReplace

End Sub

Private Sub PopulateComboBox(cellValue As Variant)

    Dim i As Integer
    Dim inList As Boolean
    inList = False
    
    With Me.ComponentCombo1
        For i = 0 To Me.ComponentCombo1.ListCount - 1
            If Me.ComponentCombo1.list(i) = cellValue Then
                'If the cellValue in the cell is in the combobox
                inList = True
                Exit For
            End If
        Next i
        
        'only add value if not already in combobox
        If Not inList Then
            .AddItem cellValue
        End If
    End With

End Sub

Private Sub CopyComboBoxValues(copyCombo As control)
     
     Dim i As Integer
     
     For i = 0 To Me.ComponentCombo1.ListCount - 1
        copyCombo.AddItem ReplaceComponent.ComponentCombo1.list(i)
        'adding items of combobox1 to another combobox
     Next
End Sub

Private Sub ResizeForm()

    Dim multiplier As Integer
    Dim deltaHeight As Double
    
    'calculate the highest number of comboboxes
    multiplier = MaxComboBoxes()
    
    deltaHeight = (defaultComboHeight + gapBetweenCombos) * multiplier
    
    'resize the userForm according the maximum umber of comboboxes
    ReplaceComponent.height = CStr(defaultFormHeight + deltaHeight)
    
    'Reposition Replace Button
    Me.ReplaceBtn.top = CStr(defaultBtnTop + (deltaHeight / 2))
    
    'Add Remove Component button
    If componentCount > 1 Then
        Me.RemoveComponent1Btn.Enabled = True
        Me.RemoveComponent1Btn.Visible = True
        
        Me.RemoveComponent1Btn.top = CStr(defaultTop + (defaultComboHeight + gapBetweenCombos) * (componentCount - 1))
    Else
        Me.RemoveComponent1Btn.Enabled = False
        Me.RemoveComponent1Btn.Visible = False
    End If
    
    'Add Remove Replace button
    If replaceCount > 1 Then
        Me.RemoveComponent2Btn.Enabled = True
        Me.RemoveComponent2Btn.Visible = True
        
        Me.RemoveComponent2Btn.top = CStr(defaultTop + (defaultComboHeight + gapBetweenCombos) * (replaceCount - 1))
    Else
        Me.RemoveComponent2Btn.Enabled = False
        Me.RemoveComponent2Btn.Visible = False
    End If

End Sub

Private Function MaxComboBoxes() As Integer
    
    'calculate the highest number of comboboxes
    MaxComboBoxes = Application.WorksheetFunction.Max(componentCount, replaceCount) - 1    'subtract 1 so code doesn't double count the starting comboboxes

End Function

Private Sub AddComponent()

    Dim NewCombo As control
    Set NewCombo = Me.Controls.Add("Forms.ComboBox.1", "ComponentCombo" & CStr(componentCount))
    
    'Size and position the new Component Combobox
    NewCombo.left = CStr(defaultCombo1Left)
    NewCombo.height = CStr(defaultComboHeight)
    NewCombo.width = CStr(defaultComboWidth)
    NewCombo.top = CStr(defaultTop + (defaultComboHeight + gapBetweenCombos) * (componentCount - 1))
    
    Dim NewText As control
    Set NewText = Me.Controls.Add("Forms.TextBox.1", "ComponentQty" & CStr(componentCount))
    
    'Size and position the new Component Text box
    NewText.left = CStr(defaultCombo1Left - 30)
    NewText.height = CStr(defaultComboHeight)
    NewText.width = CStr(defaultTextWidth)
    NewText.top = CStr(defaultTop + (defaultComboHeight + gapBetweenCombos) * (componentCount - 1))
    
    'populate the new combobox
    Call CopyComboBoxValues(NewCombo)
    
    'Resize Form
    Call ResizeForm

End Sub

Private Sub AddReplace()

    Dim NewCombo As control
    Set NewCombo = Me.Controls.Add("Forms.ComboBox.1", "ReplaceCombo" & CStr(replaceCount))
    
    'Size and position the new Replace Combobox
    NewCombo.left = CStr(defaultCombo2Left)
    NewCombo.height = CStr(defaultComboHeight)
    NewCombo.width = CStr(defaultComboWidth)
    NewCombo.top = CStr(defaultTop + (defaultComboHeight + gapBetweenCombos) * (replaceCount - 1))
    
    Dim NewText As control
    Set NewText = Me.Controls.Add("Forms.TextBox.1", "ReplaceQty" & CStr(replaceCount))
    
    'Size and position the new Component Text box
    NewText.left = CStr(defaultCombo2Left - 30)
    NewText.height = CStr(defaultComboHeight)
    NewText.width = CStr(defaultTextWidth)
    NewText.top = CStr(defaultTop + (defaultComboHeight + gapBetweenCombos) * (replaceCount - 1))
    
    'populate the new combobox
    Call CopyComboBoxValues(NewCombo)
    
    'Resize Form
    Call ResizeForm

End Sub

Private Sub RemoveComponent()

    'remove the controls
    Me.Controls.Remove ("ComponentCombo" & CStr(componentCount))
    Me.Controls.Remove ("ComponentQty" & CStr(componentCount))
    
    'count down
    componentCount = componentCount - 1
    
    'reposition and resize everything
    Call ResizeForm

End Sub

Private Sub RemoveReplace()

    'remove the controls
    Me.Controls.Remove ("ReplaceCombo" & CStr(replaceCount))
    Me.Controls.Remove ("ReplaceQty" & CStr(replaceCount))
    
    'count down
    replaceCount = replaceCount - 1
    
    'repositiona and resize everything
    Call ResizeForm

End Sub

Private Sub ReplaceBtn_Click()

    Dim components() As Variant
    Dim replacements() As Variant
    Dim productBoM() As Variant
    Dim ProductID As String
    Dim nextProductID As range
    Dim i As Integer
    Dim b
    Dim r As range
    Dim nr As range
    Dim inProductID As Boolean
    
''''build aray of original components and replacement components
    'Redim arrays to accomodate number of components determined by user
    ReDim components(componentCount - 1, 1) 'componentCount is public variable
    ReDim replacements(replaceCount - 1, 1) 'replaceCount is public variable

    'save the values determined by user into arrays
    Call ComponentsToArray("Component", components)
    Call ComponentsToArray("Replace", replacements)
    
    'build array of Bill of Materials for one Product ID
    ProductID = range("A2").value
    i = 2
    b = 2
    
    'define range for Product ID rows
    While Cells(i, 1).value = ProductID
        i = i + 1
    Wend
    
    Set nr = range("A" & b & ":A" & i - 1)
    
    For Each r In nr
        Set nextProductID = r.Offset(1, 0)
    Next r
    
    'reuse i
    i = 0
    
    'make array of BoM
    For Each r In nr
        If r.value = ProductID Then
            ReDim Preserve productBoM(1, nr.count - 1)
            'CompoenentID 1  | Component ID 2  | Component ID 3
            'Qty 1           | Qty 2           | Qty 3
            productBoM(0, i) = r.Offset(0, 2).value
            productBoM(1, i) = r.Offset(0, 1).value
            i = i + 1
        End If
    Next r
    
    'check to see if the replacement is in the product ID
    For Each r In nr
        If r = ReplaceCombo1.value Then
            inProductID = True
            Exit For
        End If
    Next r
    
    If inProductID = False Then
        nextProductID.EntireRow.Insert
        nextProductID.Offset(-1, 0).value = ProductID
        nextProductID.Offset(-1, 1).value = 0
        nextProductID.Offset(-1, 2).value = ReplaceCombo1.value
        nextProductID.Offset(-1, 3).value = "Replacement"
    End If
    
''''loop through every instance of ProductID to make replacements
    For Each r In nr
        
    Next r
    
''''move to the next ProductID and repeat steps above until end of sheet
    

End Sub

'Private Sub ReplaceBtn_Click()
'
'    Dim components() As Variant
'    Dim replacements() As Variant
'    Dim ProductID As String
'    Dim i As Integer
'
'    'Redim arrays to accomodate number of components determined by user
'    ReDim components(componentCount - 1, 1)
'    ReDim replacements(replaceCount - 1, 1)
'
'    'save the values determined by user into arrays
'    Call ComponentsToArray("Component", components)
'    Call ComponentsToArray("Replace", replacements)
'
'    'determine the gcd
'     Dim componentGCD As Integer
'    Dim replaceGCD As Integer
'    Dim i As Integer
'    Dim compqtys()
'    Dim repqtys()
'
'    'figure out set sizes
'    'save component qtys to array
'    For i = 0 To UBound(components)
'        ReDim Preserve compqtys(i)
'        compqtys(i) = components(i, 0)
'    Next i
'
'    'save replacement qtys to array
'    For i = 0 To UBound(replacements)
'        ReDim Preserve repqtys(i)
'        repqtys(i) = replacements(i, 0)
'    Next i
'
'    ProductID = range("A2").value
'    i = 2
'
'    'find the Greatest Common Denominator
'    componentGCD = Application.WorksheetFunction.Gcd(compqtys())
'    replaceGCD = Application.WorksheetFunction.Gcd(repqtys())
'
'    'edit array by dividing each item by the gcd
'    For i = 0 To UBound(components)
'        compqtys(i) = compqtys(i) / componentGCD
'    Next i
'
'    For i = 0 To UBound(replacements)
'        repqtys(i) = repqtys(i) / replaceGCD
'    Next i
'
'    'go through the Excel sheet and replace values
'    For i = 2 To CountRows("A")
'        If InArray(Cells(i, "C").value, components) Then
'            Cells(i, 2).value = "0"
'        End If
'    Next i
'
'    For i = 2 To CountRows("A")
'        If InArray(Cells(i, "C").value, replacements) Then
'            Cells(i, 2).value = replaceGCD
'        End If
'    Next i
'
'    'end
'    ReplaceComponent.Hide
'
'End Sub

Private Function InArray(val As String, arr() As Variant) As Boolean

    InArray = False
    
    Dim i As Integer
    For i = 0 To UBound(arr)
        If arr(i, 1) = val Then
            InArray = True
            Exit Function
        End If
    Next i

End Function

Private Sub ComponentsToArray(control As String, targetArray() As Variant)
    
    Call ArrayRoutine("Qty", control, targetArray, 0)
    Call ArrayRoutine("Combo", control, targetArray, 1)
    
End Sub

Private Sub ArrayRoutine(typeStr As String, control As String, targetArray() As Variant, field As Integer)

    Dim row As Integer
    Dim targetControl As control
    
    row = 0
    For Each targetControl In Me.Controls
        If targetControl.name = control & typeStr & CStr(row + 1) Then
            targetArray(row, field) = targetControl.value
            row = row + 1
        End If
    Next

End Sub
