Option Explicit

Public Sub FindSetsMain()

    'Check to see if user-entered part number matches part number in sheet
    If PartNumMatch(range("A2")) = False Then   'Checks module
        Reopen = True   'Reopen makes SourceForm ask the user for fitment Source
        SourceForm.Show
    End If
    
    'Check to see if fitments are already in the database
    If DuplicateFound = False Then
        'if FitmentSource is not blank, everything is good to export fitments to the database (necessary for when user cancels SourceForm)
        If FitmentSource <> "" Then
            ACCExport
        End If
    Else
        If PartName = InterchangeSource Then
            MsgBox ("Fitments for this part are already in the database. See " & PartName)
        Else
            MsgBox ("Fitments for this part are already in database. See " & PartName & " or " & InterchangeSource & ".")
        End If
    End If

End Sub

Private Function DuplicateFound() As Boolean
    
    'Open the PrimaryPart table in FindSets database
    rst.Open "Part1", FndStsDb, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.MoveFirst
    
    'Default to False
    DuplicateFound = False
    
    'checks to see if user-entered part number is already in the database
    While rst.EOF = False
        If (rst.fields("PartNum").value = PartName Or rst.fields("InterchangeSource").value = InterchangeSource) And _
        rst.fields("Source").value = FitmentSource And rst.fields("BrandName").value = Brand Then
            'if part is found, close the sub, don't have to search the interchangesource
            DuplicateFound = True
            'if duplicate is found, stop searching
            GoTo Close_rst
        End If
        rst.MoveNext
    Wend
    
Close_rst:
    rst.Close

End Function

Private Sub ACCExport()

    'Replace Spaces with non-breaking space and ampersands with &amp;
    MatchSixbitFormatting
    
    'Add each row as new record in CompatibilityList table
    AddFields
    
    'Add First Part's info into PrimaryPart Table
    FirstPart
    
    'Add Second Part's info into SecondaryPart Table
    SecondPart
    
    'Revert excel sheet back to regular spaces and ampersands
    RevertExcel
    
    MsgBox ("Fitments exported successfully")

End Sub

Private Sub MatchSixbitFormatting()

    'replace ampersands with html code for ampersand
    range("A:AX").Replace "&", "&amp;"
    
    'replace spaces with non-breaking spaces
    range("A:J").Replace " ", " "              'first string in regular space, second string is non-breaking space
    range("L:O").Replace " ", " "
    range("Q:AX").Replace " ", " "
    
    'Column K (Aspiration) and column P (Body Type) need to have regular spaces

End Sub

Private Sub AddFields()

    'Open CompatibilityList Table to enter values
    OpenCompatibilityListTable
    
    'all records in a table
    Dim r As Long
    r = 2
    
    'loop until first empty cell in column A
    Do While Len(range("A" & r).Formula) > 0
        With rst
            'create a new record
            .AddNew
            'add values to each field in the appropriate Fitments table
            .fields("part") = range("A" & r).value
            .fields("brand_code") = range("B" & r).value
            .fields("make") = range("C" & r).value
            .fields("model") = range("D" & r).value
            .fields("year") = range("E" & r).value
            .fields("partterminologyname") = range("F" & r).value
            .fields("notes") = range("G" & r).value
            .fields("qty") = range("H" & r).value
            .fields("mfrlabel") = range("I" & r).value
            .fields("position") = range("J" & r).value
            .fields("aspiration") = range("K" & r).value
            .fields("bedlength") = range("L" & r).value
            .fields("bedtype") = range("M" & r).value
            .fields("block") = range("N" & r).value
            .fields("bodynumdoors") = range("O" & r).value
            .fields("bodytype") = range("P" & r).value
            .fields("brakeabs") = range("Q" & r).value
            .fields("brakesystem") = range("R" & r).value
            .fields("cc") = range("S" & r).value
            .fields("cid") = range("T" & r).value
            .fields("cylinderheadtype") = range("U" & r).value
            .fields("cylinders") = range("V" & r).value
            .fields("drivetype") = range("W" & r).value
            .fields("enginedesignation") = range("X" & r).value
            .fields("enginemfr") = range("Y" & r).value
            .fields("engineversion") = range("Z" & r).value
            .fields("enginevin") = range("AA" & r).value
            .fields("frontbraketype") = range("AB" & r).value
            .fields("frontspringtype") = range("AC" & r).value
            .fields("fueldeliverysubtype") = range("AD" & r).value
            .fields("fueldeliverytype") = range("AE" & r).value
            .fields("fuelsystemcontroltype") = range("AF" & r).value
            .fields("fuelsystemdesign") = range("AG" & r).value
            .fields("fueltype") = range("AH" & r).value
            .fields("ignitionsystemtype") = range("AI" & r).value
            .fields("liters") = range("AJ" & r).value
            .fields("mfrbodycode") = range("AK" & r).value
            .fields("rearbraketype") = range("AL" & r).value
            .fields("rearspringtype") = range("AM" & r).value
            .fields("region") = range("AN" & r).value
            .fields("steeringsystem") = range("AO" & r).value
            .fields("steeringtype") = range("AP" & r).value
            .fields("submodel") = range("AQ" & r).value
            .fields("transmissioncontroltype") = range("AR" & r).value
            .fields("transmissionmfr") = range("AS" & r).value
            .fields("transmissionmfrcode") = range("AT" & r).value
            .fields("transmissionnumspeeds") = range("AU" & r).value
            .fields("transmissiontype") = range("AV" & r).value
            .fields("valvesperengine") = range("AW" & r).value
            .fields("wheelbase") = range("AX" & r).value
            .fields("Source") = FitmentSource
            .fields("InterchangeSource") = InterchangeSource
            .fields("BrandName") = Brand
    
            'stores the new record
            .Update
        End With
    'next row
    r = r + 1
    Loop
    
    'Close connection and wipe data
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub FirstPart()
    
    'Sends each row to a row in the Second Part Table
    OpenPrimaryPartTable

    Dim r As Long
    r = 2
    
    With rst
        'create a new record
        .AddNew
        'add values to each field in the record
        .fields("PartNum") = PartName
        .fields("PartType") = PartTypeVar
        .fields("Source") = FitmentSource
        .fields("InterchangeSource") = InterchangeSource
        .fields("BrandName") = Brand
        .Update
    End With
    
    'Close connection and wipe data
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub SecondPart()
    
    'Sends each row to a row in the Second Part Table
    OpenSecondaryPartTable

    Dim r As Long
    r = 2
    
    With rst
        'create a new record
        .AddNew
        'add values to each field in the record
        .fields("PartNum") = PartName
        .fields("PartType") = PartTypeVar
        .fields("Source") = FitmentSource
        .fields("InterchangeSource") = InterchangeSource
        .fields("BrandName") = Brand
        .Update
    End With

    'Close connection and wipe data
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub RevertExcel()

    'replace html code for ampersand with regular ampersands
    range("A:AX").Replace "&amp;", "&"
    
    'replace spaces with non-breaking spaces
    range("A:AX").Replace " ", " "              'first string is non-brekaing space, second string is regular space

End Sub
