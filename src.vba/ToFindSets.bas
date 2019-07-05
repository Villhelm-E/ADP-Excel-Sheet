Option Explicit

Public Sub FindSetsMain()

    'Check to see if user-entered part number matches part number in sheet
    If PartNumMatch(Range("A2")) = False Then   'Checks module
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
            MsgBox "Fitments for this part are already in the database. See " & PartName
        Else
            MsgBox "Fitments for this part are already in database. See " & PartName & " or " & InterchangeSource & "."
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
        If (rst.Fields("PartNum").Value = PartName Or rst.Fields("InterchangeSource").Value = InterchangeSource) And _
        rst.Fields("Source").Value = FitmentSource And rst.Fields("BrandName").Value = Brand Then
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
    
    MsgBox "Fitments exported successfully"

End Sub

Private Sub MatchSixbitFormatting()

    'replace ampersands with html code for ampersand
    Range("A:AX").Replace "&", "&amp;"
    
    'replace spaces with non-breaking spaces
    Range("A:J").Replace " ", " "              'first string in regular space, second string is non-breaking space
    Range("L:O").Replace " ", " "
    Range("Q:AX").Replace " ", " "
    
    'Column K (Aspiration) and column P (Body Type) need to have regular spaces

End Sub

Private Sub AddFields()

    'Open CompatibilityList Table to enter values
    OpenCompatibilityListTable
    
    'all records in a table
    Dim R As Long
    R = 2
    
    'loop until first empty cell in column A
    Do While Len(Range("A" & R).Formula) > 0
        With rst
            'create a new record
            .AddNew
            'add values to each field in the appropriate Fitments table
            .Fields("part") = Range("A" & R).Value
            .Fields("brand_code") = Range("B" & R).Value
            .Fields("make") = Range("C" & R).Value
            .Fields("model") = Range("D" & R).Value
            .Fields("year") = Range("E" & R).Value
            .Fields("partterminologyname") = Range("F" & R).Value
            .Fields("notes") = Range("G" & R).Value
            .Fields("qty") = Range("H" & R).Value
            .Fields("mfrlabel") = Range("I" & R).Value
            .Fields("position") = Range("J" & R).Value
            .Fields("aspiration") = Range("K" & R).Value
            .Fields("bedlength") = Range("L" & R).Value
            .Fields("bedtype") = Range("M" & R).Value
            .Fields("block") = Range("N" & R).Value
            .Fields("bodynumdoors") = Range("O" & R).Value
            .Fields("bodytype") = Range("P" & R).Value
            .Fields("brakeabs") = Range("Q" & R).Value
            .Fields("brakesystem") = Range("R" & R).Value
            .Fields("cc") = Range("S" & R).Value
            .Fields("cid") = Range("T" & R).Value
            .Fields("cylinderheadtype") = Range("U" & R).Value
            .Fields("cylinders") = Range("V" & R).Value
            .Fields("drivetype") = Range("W" & R).Value
            .Fields("enginedesignation") = Range("X" & R).Value
            .Fields("enginemfr") = Range("Y" & R).Value
            .Fields("engineversion") = Range("Z" & R).Value
            .Fields("enginevin") = Range("AA" & R).Value
            .Fields("frontbraketype") = Range("AB" & R).Value
            .Fields("frontspringtype") = Range("AC" & R).Value
            .Fields("fueldeliverysubtype") = Range("AD" & R).Value
            .Fields("fueldeliverytype") = Range("AE" & R).Value
            .Fields("fuelsystemcontroltype") = Range("AF" & R).Value
            .Fields("fuelsystemdesign") = Range("AG" & R).Value
            .Fields("fueltype") = Range("AH" & R).Value
            .Fields("ignitionsystemtype") = Range("AI" & R).Value
            .Fields("liters") = Range("AJ" & R).Value
            .Fields("mfrbodycode") = Range("AK" & R).Value
            .Fields("rearbraketype") = Range("AL" & R).Value
            .Fields("rearspringtype") = Range("AM" & R).Value
            .Fields("region") = Range("AN" & R).Value
            .Fields("steeringsystem") = Range("AO" & R).Value
            .Fields("steeringtype") = Range("AP" & R).Value
            .Fields("submodel") = Range("AQ" & R).Value
            .Fields("transmissioncontroltype") = Range("AR" & R).Value
            .Fields("transmissionmfr") = Range("AS" & R).Value
            .Fields("transmissionmfrcode") = Range("AT" & R).Value
            .Fields("transmissionnumspeeds") = Range("AU" & R).Value
            .Fields("transmissiontype") = Range("AV" & R).Value
            .Fields("valvesperengine") = Range("AW" & R).Value
            .Fields("wheelbase") = Range("AX" & R).Value
            .Fields("Source") = FitmentSource
            .Fields("InterchangeSource") = InterchangeSource
            .Fields("BrandName") = Brand
    
            'stores the new record
            .Update
        End With
    'next row
    R = R + 1
    Loop
    
    'Close connection and wipe data
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub FirstPart()
    
    'Sends each row to a row in the Second Part Table
    OpenPrimaryPartTable

    Dim R As Long
    R = 2
    
    With rst
        'create a new record
        .AddNew
        'add values to each field in the record
        .Fields("PartNum") = PartName
        .Fields("PartType") = PartTypeVar
        .Fields("Source") = FitmentSource
        .Fields("InterchangeSource") = InterchangeSource
        .Fields("BrandName") = Brand
        .Update
    End With
    
    'Close connection and wipe data
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub SecondPart()
    
    'Sends each row to a row in the Second Part Table
    OpenSecondaryPartTable

    Dim R As Long
    R = 2
    
    With rst
        'create a new record
        .AddNew
        'add values to each field in the record
        .Fields("PartNum") = PartName
        .Fields("PartType") = PartTypeVar
        .Fields("Source") = FitmentSource
        .Fields("InterchangeSource") = InterchangeSource
        .Fields("BrandName") = Brand
        .Update
    End With

    'Close connection and wipe data
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub RevertExcel()

    'replace html code for ampersand with regular ampersands
    Range("A:AX").Replace "&amp;", "&"
    
    'replace spaces with non-breaking spaces
    Range("A:AX").Replace " ", " "              'first string is non-brekaing space, second string is regular space

End Sub

