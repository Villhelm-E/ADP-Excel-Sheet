Option Explicit

Public Sub SixbitMain()

    'turn screenupdating off
    Application.ScreenUpdating = False
    
    'if downloaded from Herko app
    If range("A1").value = "Engine" And range("B1").value = "Make" And _
    range("C1").value = "Model" And range("D1").value = "Trim" And _
    range("E1").value = "Year" And range("F1").value = "Notes" And _
    range("G1").value = "Part Type" And range("H1").value = "Quantity Required" And _
    range("I1").value = "Position" And range("J1").value = "MFRLabel" Then
        ReorderEbayColumns
    End If
    
    'Replaces every non-breaking space with regular spaces
    ReplaceNonBreakingSpace
    
    'Count how many rows there are
    Dim numrows As Integer
    numrows = CountRows("C:C")
    
    'Format Model column so Excel doesn't assume Saab 9-3 is a date
    Call DateToModel(numrows)
    
    'move columns to their location
    columns("E:E").Cut Destination:=columns("AQ:AQ")
    columns("A:A").Cut Destination:=columns("G:G")
    columns("B:D").Cut Destination:=columns("C:E")
    
    'Format columns
    FormatColumns
    
    'Loop that moves things around
    Call BigLoop(numrows)
    
    'Count Rows
    numrows = CountRows("A:A")
    
    'For Engine/Transmission/Torque Strut Mounts
    Call CheckMount(numrows)
    
    'Fill in Headers
    Headers
    
    'Autofit Columns
    columns("A:AY").AutoFit
    
    'Rename the sheet and select A1
    'Found in FixFitmentsModule
    FitmentSource = "Sixbit"
    RenameSheet
    range("A1").Select
    
    'turn screenupdating on
    Application.ScreenUpdating = True
    
    'Done
    MsgBox ("Finished formatting.")

End Sub

Private Sub ReorderEbayColumns()

    range("G:J").EntireColumn.Delete
    
    range("B:C").Cut range("G1")
    
    range("E:E").Cut range("I1")
    
    range("D:D").Cut range("J1")
    
    range("A:A").Cut range("K1")
    
    range("A:E").EntireColumn.Delete

End Sub

Public Sub ReplaceNonBreakingSpace()

    'Sixbit fitments are saved with non-breaking spaces
    'non-breaking space is a character that looks exactly like a regular space
    'used for databases to avoid being delimited (probably)
    
    'replace every non-breaking space with a regular space
    With range("A:F")  'probably change it to A:F
        .Replace " ", " ", xlPart 'the first string is a non-breaking space, not a regular space; the second string is a regular space
    End With

End Sub

Private Sub FormatColumns()

    range("A:AX").NumberFormat = "@"

End Sub

Private Sub DateToModel(numrows)

    Dim row As Integer
    Dim daynum
    Dim monthnum
    
    'start loop
    For row = 1 To numrows
        'if Excel formatted the cell as a date
        'd is day number without leading 0
        'mmm is short month name, eg. "Sep"
        If range("C" & row).NumberFormat = "d-mmm" Then
            'save month number and day number as variables
            monthnum = Month(range("C" & row))
            daynum = Day(range("C" & row))
            
            'format cell to text instead of date
            range("C" & row).NumberFormat = "@"
            
            'replace cell value with month number, hyphen, and day nummber
            'For example the Saab 9-3 shows up as 3-Sep
            'this will force Excel to show 9-3
            range("C" & row).value = monthnum & "-" & daynum
        End If
    Next row

End Sub

Private Sub BigLoop(numrows As Integer)

    Dim i As Integer
    
    'Start loop that formats the fitments row by row
    For i = 2 To numrows
        
        'cut out liters
        Call CutLiters(i)
        
        'Cut CC
        Call CutCC(i)
        
        'Cut Cu In.
        Call CutCuIn(i)
        
        'Cut cylinders
        Call CutCylinders(i)
        
        'Cutsomething
        Call CutFuelType(i)
        
        'Cut Cylinder Head Type
        Call CutCylHeadType(i)
        
        'Clean up notes of duplicate Cylinder head types
        Call DuplicateCylHeadTyp(i)
        
        'Cut Door Count
        Call CutDoorCount(i)
        
        'Cut Body Type
        Call CutBodyType(i)
        
        'Cut Aspiration
        Call CutAspiration(i)
        
        'Cut quantity
        Call CutQuantity(i)
        
        'Cut Part type
        Call CutPartType(i)
        
        'Cut MfrLabel
        Call CutMfrLabel(i)
        
        'Cut VIN
        Call CutVIN(i)
        
        'Check to see if Fuel Type is duplicated in the notes
        Call DuplicateFuelType(i)
        
        'Cut Drive Type
        Call CutDriveType(i)
        
        'Remove "All" from submodels
        Call RemoveAll(i)
        
        'Cut fuel delivery subtype
        Call CutFuelDelivSubtype(i)
        
        'Cut fuel delivery subtype
        Call CutFuelDeliveryType(i)
        
        'Check to see if PartType is duplicated in the notes
        Call DuplicatePartType(i)
        
        'Cut out position
        Call CutPosition(i)
        
        'Cut valves per engine
        Call CutValves(i)
        
        'Cut engine designation
        Call CutEngineDesignation(i)
        
        'Cut transmission control type
        Call CutTransmissionControlType(i)
        
        'Cut out speeds
        Call CutSpeeds(i)
        
        'Cut out Transmission type
        Call CutTransType(i)
        
        'Cut out Wheelbase
        Call CutWheelBase(i)
        
        'Cut out transmission manufacturer code
        Call CutTransmissionMfrCode(i)
        
        'Cut out transmission type
        Call CutTransmissionType(i)
        
        'Cut out valves per engine
        Call CutValvesPerEngine(i)
        
        'remove duplicated submodel from notes
        Call DuplicateBodyType(i)
        
        'remove duplicated submodel from notes
        Call DuplicateSubmodel(i)
        
        'Fix up Gap for Spark Plugs
        If PartTypeVar = "Spark Plug" Then Call CleanGap(i)
        
        'Trim Notes field
        Call Cleanup(i)
        
        'Fill in the part number and part type
        Call ReplicatePart(i)
        
        'Highlight any possible errors
        Call ErrorChecker(i)
        
    Next i

End Sub

Private Sub CutLiters(row As Integer)

    'Cuts out Engine liters from Notes
    Dim Volume As String

    'if liters is two digits
    If Cells(row, 6).value Like "##.#L*" Then
        Volume = left(Cells(row, 6), 4)
        Cells(row, 6).value = Right(Cells(row, 6), Len(Cells(row, 6)) - 6)
        Cells(row, 6).value = Volume
    'the only other option is if liters is 1 digit
    ElseIf Cells(row, 6).value Like "#.#L*" Then
        Volume = left(Cells(row, 6), 3)
        Cells(row, 6).value = Replace(Cells(row, 6), Volume & "L ", "")
        Cells(row, 36).value = Volume
    End If

End Sub

Private Sub CutCC(row As Integer)

    'Cuts out CC
    Dim cc As String

    If Cells(row, 6).value Like "###CC *" Then
        cc = left(Cells(row, 6), 3)
        Cells(row, 6).value = Right(Cells(row, 6).value, Len(Cells(row, 6)) - 6)
        Cells(row, 19).value = cc
    ElseIf Cells(row, 6).value Like "####CC *" Then
        cc = left(Cells(row, 6), 4)
        Cells(row, 6).value = Right(Cells(row, 6).value, Len(Cells(row, 6)) - 7)
        Cells(row, 19).value = cc
    End If

End Sub

Private Sub CutCuIn(row As Integer)

    'Cuts out Cu. In.
    Dim CUIN As String

    If Cells(row, 6).value Like "###Cu*" Then
        CUIN = left(Cells(row, 6), 3)
        Cells(row, 6).value = Right(Cells(row, 6), Len(Cells(row, 6)) - 11)
        Cells(row, 20).value = CUIN
    ElseIf Cells(row, 6).value Like "##Cu*" Then
        CUIN = left(Cells(row, 6), 2)
        Cells(row, 6).value = Right(Cells(row, 6), Len(Cells(row, 6)) - 10)
        Cells(row, 20).value = CUIN
    End If

End Sub

Private Sub CutCylinders(row As Integer)

    'Cuts out Cylinders
    Dim Cyl
    Dim CYL2
    Dim i As Integer
    Dim InArray As Long
    
    Cyl = Array("l## ", "V## ", "H## ")
    CYL2 = Array("l# ", "V# ", "H# ")
    
    'capitalizes the L in block field
    Dim checkl As Integer
    
    'figures out at which position the string appears
    With range("F" & row)
        For InArray = LBound(Cyl) To UBound(Cyl)
            For i = 1 To Len(Cells(row, 6))
                If Mid(Cells(row, 6), i, 4) Like "*" & Cyl(InArray) & "*" Then
                    'This will give the string position
                    Cells(row, 22).value = Mid(Cells(row, 6), i + 1, 2)
                    Cells(row, 14).value = left(Cyl(InArray), 1)
                    Cells(row, 6).value = Right(Cells(row, 6), Len(Cells(row, 6)) - 4)
                    'Exit
                    Exit For
                End If
            Next i
        Next InArray
        For InArray = LBound(CYL2) To UBound(CYL2)
            For i = 1 To Len(Cells(row, 6))
                If Mid(Cells(row, 6), i, 3) Like "*" & CYL2(InArray) & "*" Then
                    'This will give the string position
                    Cells(row, 22).value = Mid(Cells(row, 6), i + 1, 1)
                    Cells(row, 14).value = left(CYL2(InArray), 1)
                    Cells(row, 6).value = Replace(Cells(row, 6), Cyl(InArray) & " ", "")
                    'Exit
                    Exit For
                End If
            Next i
        Next InArray
    End With
    
    'Place the block type in column N
    If range("N" & row).value = "l" Then
        range("N" & row).value = "L"
    End If

End Sub

Private Sub CutFuelType(row As Integer)
    
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [FuelType] FROM FuelTypes ORDER BY [ID]")
    
    'go through each row in column F
    With range("F" & row)
        With rst
            'start at beginning of BodyType field
            rst.MoveFirst
            'loop through the Body Types in the master database
            While (Not .EOF)
                'if any of the values in the BodyType field is found in the notes, cut it out and put it in column AH
                If range("F" & row).value Like "*" & .fields("FuelType").value & "*" Then
                    Cells(row, 34).value = .fields("FuelType").value
                    Cells(row, 6).value = trim(Replace(Cells(row, 6), .fields("FuelType").value, ""))
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutCylHeadType(row As Integer)
    
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [CylinderHeadType] FROM CylinderHeadTypes ORDER BY [ID]")
    
    'go through each row in column F (engine)
    With range("F" & row)
        With rst
            'start at beginning of Cylinder Head Type field
            rst.MoveFirst
            'loop through the Cylinder HEad Types in the master database
            While (Not .EOF)
                'if any values in CylinderHEadTypefield is found in the notes, cut it out and put it in column U
                If range("F" & row).value Like "*" & .fields("CylinderHeadType").value & "*" Then
                    Cells(row, 6).value = trim(Replace(Cells(row, 6), .fields("CylinderHeadType").value, ""))
                    Cells(row, 21).value = .fields("CylinderHeadType").value
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub DuplicateCylHeadTyp(row As Integer)

    Dim chrarray() As String
    Dim i As Integer
    Dim dupecylheadtyp As String
    Dim cylheadtyp As String
    
    'grab cylinder head type from cylinder head type column U
    cylheadtyp = range("U" & row).value
    
    'split up each character of cylinder head type into an array if there is a cylinder head type
    If cylheadtyp <> "" Then
        ReDim chrarray(Len(cylheadtyp) - 1)
        For i = 1 To Len(cylheadtyp)
            chrarray(i - 1) = Mid$(cylheadtyp, i, 1)
        Next
        
        'combine array with periods in between to make an acronym
        For i = 0 To Len(cylheadtyp) - 1
            dupecylheadtyp = dupecylheadtyp & chrarray(i) & "."
        Next
        
        
        'look in the description for the acronym version of cylinder head type
        If range("G" & row).value Like "*" & dupecylheadtyp & "*" Then
            'remove the cylinder head type acronym from column G if found
            range("G" & row).value = Replace(range("G" & row), dupecylheadtyp, "")
        End If
    End If

End Sub

Private Sub CutDoorCount(row As Integer)

    'removes door count from trim
    Dim DOOR As String
    
    If Cells(row, 43).value Like "*#-Door*" Then
        DOOR = Mid(Cells(row, 43), InStr(1, Cells(row, 43), "-Door") - 1, 1)
        Cells(row, 43).value = Replace(Cells(row, 43), " " & DOOR & "-Door", "")
        Cells(row, 15).value = DOOR
    End If

End Sub

Private Sub CutBodyType(row As Integer)
    
    With rst
        .ActiveConnection = MstrDb
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Source = "SELECT [BodyType] FROM BodyTypes ORDER BY [ID]"
        .Open
    End With
    
    'go through each row in column AQ (submodel)
    With range("AQ" & row)
        With rst
            'start at beginning of CylinderHeadType field
            rst.MoveFirst
            'loop through the Cylinder Head Types in the master database
            While (Not .EOF)
                'if any of the values in the CylinderHeadType field is found in the notes, cut it out and put it in column U
                If range("AQ" & row).value Like "*" & .fields("BodyType").value & "*" Then
                    Cells(row, 16).value = .fields("BodyType").value
                    Cells(row, 43).value = Replace(Cells(row, 43), " " & .fields("BodyType").value, "")
                    rst.MoveLast
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutAspiration(row As Integer)
    
    With rst
        .ActiveConnection = MstrDb
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Source = "SELECT [Aspiration] FROM Aspirations ORDER By [ID]"
        .Open
    End With
    
    'Check to see if aspiration is in submodel
    If range("AQ" & row).value Like "*Turbo*" Then
        If range("G" & row).value Like "*" & range("AQ" & row).value & "*" Then
            range("G" & row).value = Replace(range("G" & row), range("AQ" & row), "")
        End If
    End If
    
    'go through each row in column F (notes)
    With range("F" & row)
        With rst
            'start at beginning of Aspiration field
            rst.MoveFirst
            'loop through the Aspiration Types in the master database
            While (Not .EOF)
                'if any of the values in the Aspiration field is found in the notes, cut it out and put it in column K
                If range("F" & row).value Like "*" & .fields("Aspiration").value & "*" Then
                    'if notes says turbo, change to turbocharged
                    If .fields("Aspiration").value = "Turbo" Then
                        Cells(row, 11).value = "Turbocharged"
                        Cells(row, 6).value = Replace(Cells(row, 6), "Turbo", "")
                        rst.MoveLast
                    'otherwise continue as normal
                    Else
                        Cells(row, 11).value = .fields("Aspiration").value
                        Cells(row, 6).value = Replace(Cells(row, 6), .fields("Aspiration").value, "")
                        rst.MoveLast
                    End If
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'Look for "Turbo" in column G and remove it
    With range("G" & row)
        With rst
            rst.MoveFirst
            While (Not .EOF)
                If range("G" & row).value Like "*" & "Turbo" & "*" Then
                    Cells(row, 7).value = Replace(Cells(row, 7), "Turbo", "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'Look for "Supercharged in column G and remove it
    With range("G" & row)
        With rst
            rst.MoveFirst
            While (Not .EOF)
                If range("G" & row).value Like "*Supercharged*" Then
                    Cells(row, 7).value = Replace(Cells(row, 7), "Supercharged", "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutQuantity(row As Integer)

    'cut out quantity of part from notes
    Dim qty As String
    
    If Cells(row, 7).value Like "*Quantity Required ##*" Then
        qty = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "Quantity Required") + 18, 2)
        Cells(row, 7).value = Replace(Cells(row, 7), "Quantity Required " & qty, "")
        Cells(row, 8).value = qty
    Else
        If Cells(row, 7).value Like "*Quantity Required #*" Then
            qty = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "Quantity Required") + 18, 1)
            Cells(row, 7).value = Replace(Cells(row, 7), "Quantity Required " & qty, "")
            Cells(row, 8).value = qty
        Else
            If Cells(row, 7).value Like "* ## USED*" Then
                qty = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "USED") - 3, 2)
                Cells(row, 7).value = Replace(Cells(row, 7), qty & " USED", "")
                Cells(row, 8).value = qty
            Else
                If Cells(row, 7).value Like "* # USED*" Then
                    qty = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "USED") - 2, 1)
                    Cells(row, 7).value = Replace(Cells(row, 7), qty & " USED", "")
                    Cells(row, 8).value = qty
                Else
                    If Cells(row, 7).value Like "*## Per Veh;*" Then
                        qty = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "Per Veh;") - 3, 2)
                        Cells(row, 7).value = Replace(Cells(row, 7), qty & " Per Veh; ", "")
                        Cells(row, 8).value = qty
                    Else
                        If Cells(row, 7).value Like "*# Per Veh;*" Then
                            qty = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "Per Veh;") - 2, 1)
                            Cells(row, 7).value = Replace(Cells(row, 7), qty & " Per Veh;", "")
                            Cells(row, 8).value = qty
                        Else
                            'MsgBox ("check")
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub CutPartType(row As Integer)

    'cuts out parttype
    Dim PART As String
    
    If Right(PartTypeVar, 5) = "Mount" Then
        If range("G" & row).value Like "*PartType Automatic Transmission Mount*" Then
            range("G" & row).value = Replace(range("G" & row), "PartType Automatic Transmission Mount", "")
        Else
            If range("G" & row).value Like "*PartType Manual Transmission Mount*" Then
                range("G" & row).value = Replace(range("G" & row), "PartType Manual Transmission Mount", "")
            End If
        End If
    Else
        If InStr(1, Cells(row, 7), PartTypeVar) > 0 Then
            Cells(row, 7).value = Replace(Cells(row, 7), "PartType " & PartTypeVar, "")
        End If
    End If

End Sub

Private Sub CutMfrLabel(row As Integer)

    On Error GoTo MfrError

    'cuts out mfrlabel
    Dim LABEL As String, mfrstart, mfrend
    
    If Cells(row, 7).value Like "*Mfrlabel*" Then
        mfrstart = InStr(1, Cells(row, 7), "Mfrlabel")
        mfrend = InStr(mfrstart, Cells(row, 7), "  ") - 3 'double space indicates end of mfrlabel
        LABEL = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "Mfrlabel") + 9, mfrend - mfrstart - 6) 'not sure why 6, it should be 0 or 3
        If Cells(row, 7).value Like "*Mfrlabel*  *" Then
            If Cells(row, 7).value Like "*;  Mfrlabel*" Then
                Cells(row, 7).value = Replace(Cells(row, 7), ";  Mfrlabel " & LABEL, "")
                Cells(row, 9).value = LABEL
            Else
                Cells(row, 7).value = Replace(Cells(row, 7), "Mfrlabel " & LABEL, "")
                Cells(row, 9).value = LABEL
            End If
        Else
            'if there is no triple-space after the Mfrlabel in the notes
            If Cells(row, 7).value Like "*" & LABEL Then
                Cells(row, 7).value = Replace(Cells(row, 7), "Mfrlabel " & LABEL, "")
                Cells(row, 9).value = LABEL
            End If
        End If
    End If
    
MfrError:
    Exit Sub

End Sub

Private Sub CutVIN(row As Integer)

    'Cuts out VIN
    Dim VIN As String
    
    If Cells(row, 7).value Like "*VIN:*" Then
        If Cells(row, 7).value Like "*VIN: ?, ? *" Then
            'not sure what to do when it shows two VINs
        Else
            If Cells(row, 7).value Like "*VIN: ?, *" Then
                VIN = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "VIN: ") + 5, 1)
                Cells(row, 7).value = Replace(Cells(row, 7), "VIN: " & VIN & ", ", "")
                Cells(row, 27).value = VIN
            Else
                VIN = Mid(Cells(row, 7), InStr(1, Cells(row, 7), "VIN: ") + 5, 1)
                Cells(row, 7).value = Replace(Cells(row, 7), "VIN: " & VIN, "")
                Cells(row, 27).value = VIN
            End If
        End If
    End If

End Sub

Private Sub DuplicateFuelType(row As Integer)
    
    'rst is the Database established with Global variables in FixFitmentsModule
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [FuelType] FROM FuelTypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With range("G" & row)
        With rst
            'start at beginning of BodyType field
            rst.MoveFirst
            'loop through the Body Types in the master database
            While (Not .EOF)
                'if any of the values in the BodyType field is found in the notes, cut it out and put it in column AH
                If range("G" & row).value Like "*" & .fields("FuelType").value & "*" Then
                    If Cells(row, 34).value = .fields("FuelType").value Then
                        Cells(row, 7).value = Replace(Cells(row, 7), .fields("FuelType").value, "")
                    End If
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub CutDriveType(row As Integer)
    
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [DriveType] FROM DriveTypes ORDER BY [ID]")
    
    'got through each row in column G (notes)
    With range("G" & row)
        With rst
            'start at the beginning of the Drive Type field
            rst.MoveFirst
            'loop through the Drive Types in the master database
            While (Not .EOF)
                'if any of the values in the DriveType field is found in the notes, cut it out and put it in column W
                If range("G" & row).value Like "*" & .fields("DriveType").value & "*" Then
                    Cells(row, 7).value = Replace(Cells(row, 7), .fields("DriveType").value, "")
                    Cells(row, 23).value = .fields("DriveType").value
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close the connectiona to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub RemoveAll(row As Integer)

    'replaces "All" submodels with null
    If Cells(row, 43).value = "All" Then
        Cells(row, 43).value = ""
    End If

End Sub

Private Sub CutFuelDelivSubtype(row As Integer)
    
    'rst is the Database established with Global vriables in FixFitmentsModule
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT [FuelDeliverySubtype] FROM FuelDeliverySubtypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With range("G" & row)
        With rst
            'start at beginning of FuelDeliveryType field
            rst.MoveFirst
            'loop through the Fuel Delivery Types in the master database
            While (Not .EOF)
                'if any of the values in the Fuel Delivery Type field is found in the notes, cut it out and put it in column AE
                If range("G" & row).value Like "*" & .fields("FuelDeliverySubtype").value & "*" Then
                    Cells(row, 30).value = .fields("FuelDeliverySubtype").value
                    Cells(row, 7).value = Replace(Cells(row, 7), .fields("FuelDeliverySubtype").value, "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutFuelDeliveryType(row As Integer)
    
    'rst is the Database established with Global vriables in FixFitmentsModule
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT [FuelDeliveryType] FROM FuelDeliveryTypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With range("G" & row)
        With rst
            'start at beginning of FuelDeliveryType field
            rst.MoveFirst
            'loop through the Fuel Delivery Types in the master database
            While (Not .EOF)
                'if any of the values in the Fuel Delivery Type field is found in the notes, cut it out and put it in column AE
                If range("G" & row).value Like "*" & .fields("FuelDeliveryType").value & "*" Then
                    Cells(row, 31).value = .fields("FuelDeliveryType").value
                    Cells(row, 7).value = Replace(Cells(row, 7), .fields("FuelDeliveryType").value, "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub DuplicatePartType(row As Integer)

    'Checks to see if part type is duplicated in notes field
    If PartTypeVar = "Electric Fuel Pump Repair Kit" Then
        Cells(row, 7).value = Replace(Cells(row, 7), "PartType Electric Fuel Pump", "")
    Else
        If Right(PartTypeVar, 5) = "Mount" Then
            If Cells(row, 7).value Like "*PartType " & Cells(row, 9).value & "*" Then
                Cells(row, 7).value = Replace(Cells(row, 7), "PartType " & Cells(row, 9).value, "")
            End If
        Else
            If Cells(row, 7).value Like "*" & PartTypeVar & "*" Then
                Cells(row, 7).value = Replace(Cells(row, 7), PartTypeVar, "")
            End If
        End If
    End If

End Sub

Private Sub CutPosition(row As Integer)
    
    'rst is the Database established with Global vriables in FixFitmentsModule
    'run query to return oxygen sensor positions
    Set rst = MstrDb.Execute("SELECT [Position] FROM OxygenSensorPositions ORDER BY [ID]")                             'ORDER BY [ID] is important, organized in Master Database so it searches Downstream Left before Downstream
    
    'go through each row in column G (notes)
    With range("G" & row)
            With rst
                'start at beginning of position field
                rst.MoveFirst
                'loop through the Oxygen Sensor Positions in the master database
                While (Not .EOF)
                    'if any of the values in the Position fiel is found in the notes
                    If range("G" & row).value Like "*Position " & .fields("Position").value & "*" Then
                        'cut out position and put it in column J
                        Cells(row, 10).value = .fields("Position").value
                        Cells(row, 7).value = Replace(Cells(row, 7), "Position " & .fields("Position").value, "")
                        rst.MoveLast
                    End If
                    rst.MoveNext
                Wend
            End With
        End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutValves(row As Integer)

    Dim valves As String

    'Test for string "VALVE"
    If range("G" & row).value Like "*VALVE*" Then
        'if G contains 2-digit valves
        If range("G" & row).value Like "*## VALVE*" Then
            valves = Mid(range("G" & row).value, InStr(1, range("G" & row), " VALVE") - 2, 2)
            range("AW" & row).value = valves
            range("G" & row).value = Replace(range("G" & row), range("AW" & row) & " VALVES", "")
        Else
            'if G contains single-digit valves
            If range("G" & row).value Like "*# VALVE*" Then
                valves = Mid(range("G" & row), InStr(1, range("G" & row), " VALVE") - 1, 1)
                range("AW" & row).value = valves
                range("G" & row).value = Replace(range("G" & row), range("AW" & row) & " VALVES", "")
            Else
                'if G ends in 2-digit valves
                If range("G" & row).value Like "*## VALVE" Then
                    valves = Mid(range("G" & row).value, InStr(1, range("G" & row), " VALVES") - 2, 2)
                    range("AW" & row).value = valves
                    range("G" & row).value = left(range("G" & row), Len(range("G" & row)) - 9)
                Else
                    'if G ends in single-digit valves
                    If range("G" & row).value Like "*# VALVE" Then
                        valves = Mid(range("G" & row), InStr(1, range("G" & row), " VALVES") - 1, 1)
                        range("AW" & row).value = valves
                        range("G" & row).value = left(range("G" & row), Len(range("G" & row)) - 8)
                    Else
                        'if G starts with 2-digit valves
                        If range("G" & row).value Like "## VALVE*" Then
                            valves = left(range("G" & row), 2)
                            range("AW" & row).value = valves
                            range("G" & row).value = Right(range("G" & row), Len(range("G" & row)) - 9)
                        Else
                            'if G starts with single-digit valves
                            If range("G" & row).value Like "# VALVE*" Then
                                valves = left(range("G" & row), 1)
                                range("AW" & row).value = valves
                                range("G" & row).value = Right(range("G" & row), Len(range("G" & row)) - 8)
                            Else
                                'if G is only 2-digit valves
                                If range("G" & row).value Like "## Valve" Then
                                    valves = left(range("G" & row), 2)
                                    range("G" & row).value = ""
                                    range("AW" & row).value = valves
                                Else
                                    'if G is only single-digit valves
                                    If range("G" & row).value Like "# Valve" Then
                                        valves = left(range("G" & row), 1)
                                        range("G" & row).value = ""
                                        range("AW" & row).value = valves
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        'Test for "Valves: " string
        If range("G" & row).value Like "*Valves: ##*" Then
            valves = Mid(range("G" & row), InStr(1, range("G" & row), "Valves: ") + 8, 2)
            range("AW" & row).value = valves
            range("G" & row).value = Replace(range("G" & row), "Valves: " & range("AW" & row).value, "")
        Else
            If range("G" & row).value Like "*Valves: #*" Then
                valves = Mid(range("G" & row), InStr(1, range("G" & row), "Valves: ") + 8, 1)
                range("AW" & row).value = valves
                range("G" & row).value = Replace(range("G" & row), "Valves: " & range("AW" & row).value, "")
            End If
        End If
    End If

End Sub

Private Sub CutEngineDesignation(row As Integer)

    'run query to return oxygen sensor positions
    Set rst = MstrDb.Execute("SELECT [EngineDesignation] FROM EngineDesignations ORDER BY [ID]")                           'ORDER BY [ID] is important, organized in Master Database
        
    'go through each row in column G (notes)
    'check if notes contains the string "Engine: "
    If range("G" & row).value Like "*Engine: *" Then
        'if yes, search through Engien Designations in Master Database"
        With range("G" & row)
            With rst
                'start at beginning of position field
                rst.MoveFirst
                'loop through the Oxygen Sensor Positions in the master database
                While (Not .EOF)
                    'if any of the values in the Position fiel is found in the notes
                    If range("G" & row).value Like "*" & "Engine: " & .fields("EngineDesignation").value & "*" Then
                        'cut out position and put it in column J
                        Cells(row, 24).value = .fields("EngineDesignation").value
                        Cells(row, 7).value = Replace(Cells(row, 7), "Engine: " & .fields("EngineDesignation").value, "")
                    End If
                    rst.MoveNext
                Wend
            End With
        End With
    End If
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutTransmissionControlType(row As Integer)

    'run query to return transmission control types
    Set rst = MstrDb.Execute("SELECT [TransControlType] FROM TransControlTypes ORDER BY [ID]")                         'ORDER BY [ID] is important, organized in Master Database so it searches Automatic CVT before Automatic
    
    'go through each row in column G (notes)
    With range("G" & row)
            With rst
                'start at beginning of position field
                rst.MoveFirst
                'loop through the Oxygen Sensor Positions in the master database
                While (Not .EOF)
                    'if any of the values in the Position fiel is found in the notes
                    If range("G" & row).value Like "*" & .fields("TransControlType").value & " Trans*" Then
                        'cut out position and put it in column AR
                        Cells(row, 44).value = .fields("TransControlType").value
                        Cells(row, 7).value = Replace(Cells(row, 7), .fields("TransControlType").value & " Trans", "")
                        GoTo End_Loop
                    End If
                    rst.MoveNext
                Wend
            End With
        End With
    
End_Loop:
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutSpeeds(row As Integer)
    
    'Check to see if Notes contains a two-digit number of speeds
    If range("G" & row).value Like "*## Speed Trans*" Then
        'Cuts out the two-digit number and places in AU
        Cells(row, 47).value = Mid(range("G" & row).value, InStr(1, range("G" & row), " Speed Trans") - 2, 2)
        'erases the "## Speed Trans" string from G
        Cells(row, 7).value = Replace(Cells(row, 7), range("AU" & row).value & " Speed Trans", "")
    Else
        'Check to see if Notes contains a single-digit number of speeds
        If range("G" & row).value Like "*# Speed Trans*" Then
            'Cuts out the single-digit number and places in AU
            Cells(row, 47).value = Mid(Cells(row, 7).value, InStr(1, range("G" & row), " Speed Trans") - 1, 1)
            'erases the "# Speed Trans" string from G
            Cells(row, 7).value = Replace(Cells(row, 7), range("AU" & row).value & " Speed Trans", "")
        End If
    End If

End Sub

Private Sub CutTransType(row As Integer)

    If range("G" & row).value Like "* Transaxle*" Then
        range("AV" & row).value = "Transaxle"
        Cells(row, 7).value = Replace(Cells(row, 7), "Transaxle", "")
    End If

End Sub

Private Sub CutWheelBase(row As Integer)

    Dim wb As Integer
    Dim WheelBase As String
    
    If range("G" & row).value Like "*###.#"" WB*" Then
        wb = InStr(1, Cells(row, 7), "WB")
        WheelBase = Mid(Cells(row, 7), wb - 7, 5)   'take the 5 characters that start 7 characters before the "WB" in the notes
        range("AX" & row).value = WheelBase
        range("G" & row).value = Replace(range("G" & row), WheelBase & """ WB", "")
    End If

End Sub

Private Sub CutTransmissionMfrCode(row As Integer)

    'open Transmission Mfr Codes table in Master Database
    Set rst = MstrDb.Execute("SELECT [TransmissionMfrCode] FROM TransmissionMfrCodes ORDER BY [ID]")
    
    'first check to see if there is something in the G column left
    If range("G" & row).value <> "" Then
        With rst
        .MoveFirst
            While Not .EOF
                'if the non-blank cell contains a TransmissionMfrCode
                If range("G" & row).value Like "* " & .fields("TransmissionMfrCode").value & ", *" Then
                    range("AT" & row).value = .fields("TransmissionMfrCode").value
                    range("G" & row).value = Replace(range("G" & row), .fields("TransmissionMfrCode").value & ", ", "")
                    GoTo Exit_Loop
                Else
                    .MoveNext
                End If
            Wend
        End With
    End If
    
    'close table
Exit_Loop:
    rst.Close

End Sub

Private Sub CutTransmissionType(row As Integer)

    'open Transmission types in Master Database
    Set rst = MstrDb.Execute("SELECT [TransmissionType] FROM TransmissionTypes ORDER BY [ID]")
    
    rst.MoveFirst
    With rst
        If range("G" & row).value Like "*" & .fields("TransmissionType").value & "*" Then
            range("AV" & row).value = .fields("TransmissionType").value
            range("G" & row).value = Replace(range("G" & row), .fields("TransmissionType").value, "")
            GoTo Exit_Loop
        Else
            .MoveNext
        End If
    End With

Exit_Loop:
    rst.Close

End Sub

Private Sub CutValvesPerEngine(row As Integer)

    Dim valves As String
    Dim percyl As String
    
    If range("G" & row).value Like "*valves per cylinder*" Then
        percyl = Mid(range("G" & row), InStr(1, range("G" & row).value, "valves per cylinder") - 2, 1)
        valves = percyl * range("V" & row).value
        range("AW" & row).value = valves
        range("G" & row).value = Replace(range("G" & row), percyl & " valves per cylinder", "")
    End If
    
    If range("G" & row).value Like "*## Valve*" Then
        valves = Mid(range("G" & row), InStr(1, range("G" & row), "Valve") - 3, 2)
        range("AW" & row).value = valves
        range("G" & row).value = Replace(range("G" & row), valves & " Valve", "")
    ElseIf range("G" & row).value Like "*# Valve*" Then
        valves = Mid(range("G" & row), InStr(1, range("G" & row), "Valve") - 3, 1)
        range("AW" & row).value = valves
        range("G" & row).value = Replace(range("G" & row), valves & " Valve", "")
    End If

End Sub

Private Sub DuplicateSubmodel(row As Integer)

    If range("G" & row).value Like "*" & range("AQ" & row).value & "*" Then
        range("G" & row).value = Replace(range("G" & row), range("AQ" & row).value, "")
    End If

End Sub

Private Sub DuplicateBodyType(row As Integer)

    If range("G" & row).value Like "*" & range("P" & row).value & "*" Then
        range("G" & row).value = Replace(range("G" & row), range("P" & row).value, "")
    End If

End Sub

Private Sub CleanGap(row As Integer)

    'Dim GapPre As String
    '
    'If PartTypeVar = "Spark Plug" And (InStr(1, Range("G" & Row), "Gap") > 0 Or InStr(1, Range("G" & Row), "GAP") > 0) Then
    '    If Range("G" & Row).Value Like "GAP=#.###" Then
    '        GapPre = Mid(Range("G" & Row), InStr(1, Range("G" & Row), "GAP="), 9)
    '        Range("G" & Row).Value = Replace(Range("G" & Row), GapPre, "Gap ")
    '    End If
    'End If
    
    'GAP=0.028
    '.028"
    '.028 Gap
    

End Sub

Private Sub Cleanup(row As Integer)

    Dim l As Integer, BigLoop
    
    l = 0
    BigLoop = 0
    
    'If the notes field contains the part type, remove it
    If range("G" & row).value Like "*Part Note: *" Then
        range("G" & row).value = Replace(range("G" & row), "Part Note: ", "")
    End If
    
    'If the notes field contains the number of doors, remove it
    If range("G" & row).value Like "*" & range("O" & row).value & " Door*" Then
        range("G" & row).value = Replace(range("G" & row), range("O" & row).value & " Door", "")
    End If
    
    'If the notes field contains the cylinder head type, remove it
    If range("G" & row).value Like "*" & range("U" & row).value & "*" Then
        range("G" & row).value = Replace(range("G" & row), range("U" & row).value, "")
    End If
    
    'If notes has Natural
    If range("G" & row).value Like "*Natural*" And range("K" & row).value = "Naturally Aspirated" Then
        range("G" & row).value = Replace(range("G" & row), "Natural", "")
    End If
    
    Do While BigLoop < 6
        Do While l = 0
            'remove double semicolons
            If range("G" & row).value Like "*;;*" Then
                range("G" & row).value = Replace(range("G" & row), ";;", ";")
            Else
                l = 1
            End If
        Loop
        
        l = 0
            
        Do While l = 0
            'remove double spaces
            If range("G" & row).value Like "*  *" Then
                range("G" & row).value = Replace(range("G" & row), "  ", " ")
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        Do While l = 0
            'remove double commas
            If range("G" & row).value Like "*,,*" Then
                range("G" & row).value = Replace(range("G" & row), ",,", ",")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove beginning spaces
            If range("G" & row).value Like " *" Then
                range("G" & row).value = Replace(range("G" & row), range("G" & row), Right(range("G" & row), Len(range("G" & row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove beginning semicolons
            If range("G" & row).value Like ";*" Then
                range("G" & row).value = Replace(range("G" & row), range("G" & row), Right(range("G" & row), Len(range("G" & row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove beginning semicolons
            If range("G" & row).value Like ",*" Then
                range("G" & row).value = Replace(range("G" & row), range("G" & row), Right(range("G" & row), Len(range("G" & row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove final semicolons
            If range("G" & row).value Like "*;" Then
                range("G" & row).value = Replace(range("G" & row), Right(range("G" & row), 1), "")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove final spaces
            If range("G" & row).value Like "* " Then
                range("G" & row).value = Replace(range("G" & row), range("G" & row), left(range("G" & row), Len(range("G" & row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove lonely semicolons
            If range("G" & row).value Like "* ; *" Then
                range("G" & row).value = Replace(range("G" & row), " ; ", "; ")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove lonely commas
            If range("G" & row).value Like "* , *" Then
                range("G" & row).value = Replace(range("G" & row), " , ", ", ")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove final hyphens
            If range("G" & row).value Like "*-" Then
                range("G" & row).value = Replace(range("G" & row), range("G" & row), left(range("G" & row), Len(range("G" & row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove comma-semicolons
            If range("G" & row).value Like "* ,; *" Then
                range("G" & row).value = Replace(range("G" & row), " ,; ", " ")
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        Do While l = 0
            'remove semi-colon-commas
            If range("G" & row).value Like "*;, *" Then
                range("G" & row).value = Replace(range("G" & row), ";, ", "; ")
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        Do While l = 0
            'remove final commas
            If range("G" & row).value Like "*," Then
                range("G" & row).value = left(range("G" & row), Len(range("G" & row)) - 1)
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        BigLoop = BigLoop + 1
    Loop

End Sub

Private Sub ReplicatePart(row As Integer)

    'Adds Part number, part type, brand_code, and sku
    Cells(row, 1).value = PartName
    Cells(row, 2).value = "FVKX"
    Cells(row, 6).value = PartTypeVar
    Cells(row, 51).value = gendSKU

End Sub

Private Sub ErrorChecker(row As Integer)

    

End Sub

Private Sub CheckMount(numrows As Integer)

    Dim unmatched As Boolean
    Dim LResponse As Integer
    Dim row As Integer

    unmatched = False
    
    If Right(PartTypeVar, 5) = "Mount" Then
        For row = 2 To numrows
            If range("I" & row).value = "" Then GoTo Next_row
            
            If Not (Replace(range("F" & row).value, "Automatic Transmission", "Auto Trans") = range("I" & row).value Or Replace(range("F" & row).value, "Manual Transmission", "Manual Trans") = range("I" & row).value) Then
                unmatched = True
                row = numrows   'ends the loop
            End If
Next_row:
        Next row
    Else
        Exit Sub
    End If
    
    If unmatched = True Then
        LResponse = MsgBox("Some of the notes don't match the part type entered. Do you want to update to the part type in the notes?", vbYesNo, "Part Type Mismatch")
    
        'Determine Yes/No options
        If LResponse = vbYes Then
            For row = 2 To numrows
                If range("I" & row).value = "Auto Trans Mount" Then
                    range("F" & row).value = "Automatic Transmission Mount"
                Else
                    If range("I" & row).value = "Manual Trans Mount" Then
                        range("F" & row).value = "Manual Transmission Mount"
                    End If
                End If
            Next row
        End If
    End If

End Sub

Private Sub Headers()

    'These are the ACES headers
    range("A1:V1").value = [{"part", "brand_code", "make", "model", "year", "partterminologyname", "notes", "qty", "mfrlabel", "position", "aspiration","bedlength","bedtype","block","bodynumdoors","bodytype","brakeabs","brakesystem","cc","cid","cylinderheadtype","cylinders"}]
    range("W1:AK1").value = [{"drivetype", "enginedesignation","enginemfr","engineversion","enginevin","frontbraketype","frontspringtype","fueldeliverysubtype","fueldeliverytype","fuelsystemcontroltype","fuelsystemdesign","fueltype","ignitionsystemtype", "liters","mfrbodycode"}]
    range("AL1:AX1").value = [{"rearbraketype", "rearspringtype","region","steeringsystem","steeringtype","submodel","transmissioncontroltype","transmissionmfr","transmissionmfrcode","transmissionnumspeeds", "transmissiontype", "valvesperengine", "wheelbase"}]
    range("AY1").value = "sku"

End Sub
