Option Explicit

Public Sub PrepWorksheet(SheetName As String)

    Dim wsSheet As Worksheet

    On Error Resume Next
    'save current sheet name to worksheet variable
    Set wsSheet = Sheets(SheetName)
    On Error GoTo 0
    
    'check to see if the worksheet name already exists
    If Not wsSheet Is Nothing Then
        'if sheet exists, open worksheet
        Worksheets(SheetName).Activate
    Else
        'Otherwise create the worksheet
        If CheckBlank = True Then
            'if current worksheet is blank, just rename the worksheet to sheetName
            ActiveSheet.name = SheetName
        Else
            'if current worksheet is not blank, add worksheet to the end and name it
            With Sheets.Add(, Sheets(Sheets.count))
                .name = SheetName
            End With
        End If
    End If

End Sub

Public Sub RenameSheet()

    Dim WS_Count As Integer
    Dim i As Integer
    Dim FoundSheet As Integer
    Dim FoundCopy As Integer
    Dim ws As Worksheet
    
    'start count of copies and duplicate copies at 0
    FoundSheet = 0
    FoundCopy = 0
    Set ws = ActiveWorkbook.ActiveSheet
    
    'set WS_Count equal to the number of worksheets in the active workbook
    WS_Count = ActiveWorkbook.Worksheets.count
    
    'begin the loop
    For i = 1 To WS_Count
        If ActiveWorkbook.Worksheets(i).name = PartName & " " & FitmentSource Then
            'count the number of existing sheets with the same name
            FoundSheet = FoundSheet + 1
        End If
        
        If ActiveWorkbook.Worksheets(i).name = PartName & " " & FitmentSource & " (Copy)" Then
            'count the number of existing copies ot existing sheets witht he same name
            FoundCopy = FoundCopy + 1
        End If
    Next i
    
    Select Case FoundSheet
    
        'if there are no existing sheets with the target name
        Case 0
            'if there are copies, append a number at the end
            If FoundCopy > 0 Then
                ws.name = PartName & " " & FitmentSource & " (Copy) " & FoundCopy
            Else
                'name the worksheet
                ws.name = PartName & " " & FitmentSource
            End If
        
        'if there are copies
        Case Else
            'if there are copies, append a number at the end
            If FoundCopy > 0 Then
                ws.name = PartName & " " & FitmentSource & " (Copy) " & FoundCopy
            Else
                'add (Copy) to the copy
                ws.name = PartName & " " & FitmentSource & " (Copy)"
            End If
    End Select
End Sub

Public Sub OpenExcelVersion()

    'Open Excel Sheet Version table from Master Database
    Set rst = MstrDb.Execute("SELECT [Version] FROM [Excel Sheet Version]") 'rst is global variable
    
    'Move to first record in table
    rst.MoveFirst

End Sub

Public Sub OpenACESPartTypes()

    'run query to return part types from Master Database
    Set rst = MstrDb.Execute("SELECT DISTINCT * FROM PartTypes ORDER BY [ACESPartType]")    'rst is global variable
    
    'Move to first record in table
    rst.MoveFirst

End Sub

Public Sub OpenManufacturers()
    
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT DISTINCT * FROM Manufacturers ORDER BY [ManufacturerFull]")    'rst is global variable
    
    'Move to first record in table
    rst.MoveFirst

End Sub

Public Sub FillFitmentSources()

    'run query to return part types
    Set rst = MstrDb.Execute("SELECT DISTINCT [Source] FROM FitmentSources ORDER BY [Source]")  'rst is global variable
    
    'Move to first record in table
    rst.MoveFirst

End Sub

Public Sub OpenFitmentSources()

    'run query to return part types
    rst.Open "SELECT DISTINCT [Source] FROM FitmentSources ORDER BY [Source];", _
             FndStsDb, adOpenStatic
    rst.MoveFirst
    
End Sub

Public Sub OpenCompatibilityListTable()

    'run query to return Compatibilities
    rst.Open "Compatibilitieslist", FndStsDb, adOpenKeyset, adLockOptimistic, adCmdTable
    
    rst.MoveFirst

End Sub

Public Sub OpenPrimaryPartTable()

    'opens the PrimaryPart Table
    'Table in Find Sets database
    rst.Open "Part1", FndStsDb, adOpenKeyset, adLockOptimistic, adCmdTable
    
    rst.MoveFirst

End Sub

Public Sub OpenSecondaryPartTable()

    'opens the SecondaryPart Table
    'Table in Find Sets database
    rst.Open "Part2", FndStsDb, adOpenKeyset, adLockOptimistic, adCmdTable
    
    rst.MoveFirst

End Sub

Public Sub OpenOrdinalID()
    
    Set rst = SxbtDb.Execute("SELECT * FROM dbo.CompatibilitySets ORDER BY OrdinalID DESC")
    
    rst.MoveFirst

End Sub

Public Sub OpenFinaleProductFields()

    Set rst = MstrDb.Execute("SELECT * FROM FinaleProductFields WHERE ((Not (FinaleProductFields.Field)=""Product ID"") AND ((FinaleProductFields.Active)=True)) ORDER BY FinaleProductFields.ID")
    
    rst.MoveFirst

End Sub

Public Sub OpenShippingMethods()

    Set rst = MstrDb.Execute("SELECT * FROM ShippingMethods")
    
    rst.MoveFirst

End Sub
