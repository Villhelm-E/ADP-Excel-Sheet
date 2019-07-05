Option Explicit

Public Sub SixbitMain()

    'turn screenupdating off
    Application.ScreenUpdating = False
    
    'Replaces every non-breaking space with regular spaces
    ReplaceNonBreakingSpace
    
    'Count how many rows there are
    Dim numrows As Integer
    numrows = CountRows("C:C")
    
    'Format Model column so Excel doesn't assume Saab 9-3 is a date
    Call DateToModel(numrows)
    
    'Move columns to their location
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
    columns("A:AX").AutoFit
    
    'Rename the sheet and select A1
    'Found in FixFitmentsModule
    RenameSheet
    Range("A1").Select
    
    'turn screenupdating on
    Application.ScreenUpdating = True
    
    'Done
    MsgBox "Finished formatting."

End Sub

Public Sub ReplaceNonBreakingSpace()

    'Sixbit fitments are saved with non-breaking spaces
    'non-breaking space is a character that looks exactly like a regular space
    'used for databases to avoid being delimited (probably)
    
    'replace every non-breaking space with a regular space
    With Range("A:F")  'probably change it to A:F
        .Replace " ", " ", xlPart 'the first string is a non-breaking space, not a regular space; the second string is a regular space
    End With

End Sub

Private Sub FormatColumns()

    Range("A:AX").NumberFormat = "@"

End Sub

Private Sub DateToModel(numrows)

    Dim Row As Integer
    Dim daynum
    Dim monthnum
    
    'start loop
    For Row = 1 To numrows
        'if Excel formatted the cell as a date
        'd is day number without leading 0
        'mmm is short month name, eg. "Sep"
        If Range("C" & Row).NumberFormat = "d-mmm" Then
            'save month number and day number as variables
            monthnum = Month(Range("C" & Row))
            daynum = Day(Range("C" & Row))
            
            'format cell to text instead of date
            Range("C" & Row).NumberFormat = "@"
            
            'replace cell value with month number, hyphen, and day nummber
            'For example the Saab 9-3 shows up as 3-Sep
            'this will force Excel to show 9-3
            Range("C" & Row).Value = monthnum & "-" & daynum
        End If
    Next Row

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

Private Sub CutLiters(Row As Integer)

    'Cuts out Engine liters from Notes
    Dim Volume As String

    'if liters is two digits
    If Cells(Row, 6).Value Like "##.#L*" Then
        Volume = Left(Cells(Row, 6), 4)
        Cells(Row, 6).Value = Right(Cells(Row, 6), Len(Cells(Row, 6)) - 6)
        Cells(Row, 6).Value = Volume
    'the only other option is if liters is 1 digit
    ElseIf Cells(Row, 6).Value Like "#.#L*" Then
        Volume = Left(Cells(Row, 6), 3)
        Cells(Row, 6).Value = Replace(Cells(Row, 6), Volume & "L ", "")
        Cells(Row, 36).Value = Volume
    End If

End Sub

Private Sub CutCC(Row As Integer)

    'Cuts out CC
    Dim cc As String

    If Cells(Row, 6).Value Like "###CC *" Then
        cc = Left(Cells(Row, 6), 3)
        Cells(Row, 6).Value = Right(Cells(Row, 6).Value, Len(Cells(Row, 6)) - 6)
        Cells(Row, 19).Value = cc
    ElseIf Cells(Row, 6).Value Like "####CC *" Then
        cc = Left(Cells(Row, 6), 4)
        Cells(Row, 6).Value = Right(Cells(Row, 6).Value, Len(Cells(Row, 6)) - 7)
        Cells(Row, 19).Value = cc
    End If

End Sub

Private Sub CutCuIn(Row As Integer)

    'Cuts out Cu. In.
    Dim CUIN As String

    If Cells(Row, 6).Value Like "###Cu*" Then
        CUIN = Left(Cells(Row, 6), 3)
        Cells(Row, 6).Value = Right(Cells(Row, 6), Len(Cells(Row, 6)) - 11)
        Cells(Row, 20).Value = CUIN
    ElseIf Cells(Row, 6).Value Like "##Cu*" Then
        CUIN = Left(Cells(Row, 6), 2)
        Cells(Row, 6).Value = Right(Cells(Row, 6), Len(Cells(Row, 6)) - 10)
        Cells(Row, 20).Value = CUIN
    End If

End Sub

Private Sub CutCylinders(Row As Integer)

    'Cuts out Cylinders
    Dim Cyl
    Dim CYL2
    Dim i As Integer
    Dim inarray As Long
    
    Cyl = Array("l## ", "V## ", "H## ")
    CYL2 = Array("l# ", "V# ", "H# ")
    
    'capitalizes the L in block field
    Dim checkl As Integer
    
    'figures out at which position the string appears
    With Range("F" & Row)
        For inarray = LBound(Cyl) To UBound(Cyl)
            For i = 1 To Len(Cells(Row, 6))
                If Mid(Cells(Row, 6), i, 4) Like "*" & Cyl(inarray) & "*" Then
                    'This will give the string position
                    Cells(Row, 22).Value = Mid(Cells(Row, 6), i + 1, 2)
                    Cells(Row, 14).Value = Left(Cyl(inarray), 1)
                    Cells(Row, 6).Value = Right(Cells(Row, 6), Len(Cells(Row, 6)) - 4)
                    'Exit
                    Exit For
                End If
            Next i
        Next inarray
        For inarray = LBound(CYL2) To UBound(CYL2)
            For i = 1 To Len(Cells(Row, 6))
                If Mid(Cells(Row, 6), i, 3) Like "*" & CYL2(inarray) & "*" Then
                    'This will give the string position
                    Cells(Row, 22).Value = Mid(Cells(Row, 6), i + 1, 1)
                    Cells(Row, 14).Value = Left(CYL2(inarray), 1)
                    Cells(Row, 6).Value = Replace(Cells(Row, 6), Cyl(inarray) & " ", "")
                    'Exit
                    Exit For
                End If
            Next i
        Next inarray
    End With
    
    'Place the block type in column N
    If Range("N" & Row).Value = "l" Then
        Range("N" & Row).Value = "L"
    End If

End Sub

Private Sub CutFuelType(Row As Integer)
    
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [FuelType] FROM FuelTypes ORDER BY [ID]")
    
    'go through each row in column F
    With Range("F" & Row)
        With rst
            'start at beginning of BodyType field
            rst.MoveFirst
            'loop through the Body Types in the master database
            While (Not .EOF)
                'if any of the values in the BodyType field is found in the notes, cut it out and put it in column AH
                If Range("F" & Row).Value Like "*" & .Fields("FuelType").Value & "*" Then
                    Cells(Row, 34).Value = .Fields("FuelType").Value
                    Cells(Row, 6).Value = trim(Replace(Cells(Row, 6), .Fields("FuelType").Value, ""))
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutCylHeadType(Row As Integer)
    
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [CylinderHeadType] FROM CylinderHeadTypes ORDER BY [ID]")
    
    'go through each row in column F (engine)
    With Range("F" & Row)
        With rst
            'start at beginning of Cylinder Head Type field
            rst.MoveFirst
            'loop through the Cylinder HEad Types in the master database
            While (Not .EOF)
                'if any values in CylinderHEadTypefield is found in the notes, cut it out and put it in column U
                If Range("F" & Row).Value Like "*" & .Fields("CylinderHeadType").Value & "*" Then
                    Cells(Row, 6).Value = trim(Replace(Cells(Row, 6), .Fields("CylinderHeadType").Value, ""))
                    Cells(Row, 21).Value = .Fields("CylinderHeadType").Value
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub DuplicateCylHeadTyp(Row As Integer)

    Dim chrarray() As String
    Dim i As Integer
    Dim dupecylheadtyp As String
    Dim cylheadtyp As String
    
    'grab cylinder head type from cylinder head type column U
    cylheadtyp = Range("U" & Row).Value
    
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
        If Range("G" & Row).Value Like "*" & dupecylheadtyp & "*" Then
            'remove the cylinder head type acronym from column G if found
            Range("G" & Row).Value = Replace(Range("G" & Row), dupecylheadtyp, "")
        End If
    End If

End Sub

Private Sub CutDoorCount(Row As Integer)

    'removes door count from trim
    Dim DOOR As String
    
    If Cells(Row, 43).Value Like "*#-Door*" Then
        DOOR = Mid(Cells(Row, 43), InStr(1, Cells(Row, 43), "-Door") - 1, 1)
        Cells(Row, 43).Value = Replace(Cells(Row, 43), " " & DOOR & "-Door", "")
        Cells(Row, 15).Value = DOOR
    End If

End Sub

Private Sub CutBodyType(Row As Integer)
    
    With rst
        .ActiveConnection = MstrDb
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Source = "SELECT [BodyType] FROM BodyTypes ORDER BY [ID]"
        .Open
    End With
    
    'go through each row in column AQ (submodel)
    With Range("AQ" & Row)
        With rst
            'start at beginning of CylinderHeadType field
            rst.MoveFirst
            'loop through the Cylinder Head Types in the master database
            While (Not .EOF)
                'if any of the values in the CylinderHeadType field is found in the notes, cut it out and put it in column U
                If Range("AQ" & Row).Value Like "*" & .Fields("BodyType").Value & "*" Then
                    Cells(Row, 16).Value = .Fields("BodyType").Value
                    Cells(Row, 43).Value = Replace(Cells(Row, 43), " " & .Fields("BodyType").Value, "")
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

Private Sub CutAspiration(Row As Integer)
    
    With rst
        .ActiveConnection = MstrDb
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Source = "SELECT [Aspiration] FROM Aspirations ORDER By [ID]"
        .Open
    End With
    
    'Check to see if aspiration is in submodel
    If Range("AQ" & Row).Value Like "*Turbo*" Then
        If Range("G" & Row).Value Like "*" & Range("AQ" & Row).Value & "*" Then
            Range("G" & Row).Value = Replace(Range("G" & Row), Range("AQ" & Row), "")
        End If
    End If
    
    'go through each row in column F (notes)
    With Range("F" & Row)
        With rst
            'start at beginning of Aspiration field
            rst.MoveFirst
            'loop through the Aspiration Types in the master database
            While (Not .EOF)
                'if any of the values in the Aspiration field is found in the notes, cut it out and put it in column K
                If Range("F" & Row).Value Like "*" & .Fields("Aspiration").Value & "*" Then
                    'if notes says turbo, change to turbocharged
                    If .Fields("Aspiration").Value = "Turbo" Then
                        Cells(Row, 11).Value = "Turbocharged"
                        Cells(Row, 6).Value = Replace(Cells(Row, 6), "Turbo", "")
                        rst.MoveLast
                    'otherwise continue as normal
                    Else
                        Cells(Row, 11).Value = .Fields("Aspiration").Value
                        Cells(Row, 6).Value = Replace(Cells(Row, 6), .Fields("Aspiration").Value, "")
                        rst.MoveLast
                    End If
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'Look for "Turbo" in column G and remove it
    With Range("G" & Row)
        With rst
            rst.MoveFirst
            While (Not .EOF)
                If Range("G" & Row).Value Like "*" & "Turbo" & "*" Then
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), "Turbo", "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'Look for "Supercharged in column G and remove it
    With Range("G" & Row)
        With rst
            rst.MoveFirst
            While (Not .EOF)
                If Range("G" & Row).Value Like "*Supercharged*" Then
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), "Supercharged", "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutQuantity(Row As Integer)

    'cut out quantity of part from notes
    Dim QTY As String
    
    If Cells(Row, 7).Value Like "*Quantity Required ##*" Then
        QTY = Right(Cells(Row, 7), 2)
        Cells(Row, 7).Value = Left(Cells(Row, 7), Len(Cells(Row, 7)) - 20)
        Cells(Row, 8).Value = QTY
    Else
        If Cells(Row, 7).Value Like "*Quantity Required #*" Then
            QTY = Right(Cells(Row, 7), 1)
            Cells(Row, 7).Value = Left(Cells(Row, 7), Len(Cells(Row, 7)) - 19)
            Cells(Row, 8).Value = QTY
        Else
            If Cells(Row, 7).Value Like "* ## USED*" Then
                QTY = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "USED") - 3, 2)
                Cells(Row, 7).Value = Replace(Cells(Row, 7), QTY & " USED", "")
                Cells(Row, 8).Value = QTY
            Else
                If Cells(Row, 7).Value Like "* # USED*" Then
                    QTY = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "USED") - 2, 1)
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), QTY & " USED", "")
                    Cells(Row, 8).Value = QTY
                Else
                    If Cells(Row, 7).Value Like "*## Per Veh;*" Then
                        QTY = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "Per Veh;") - 3, 2)
                        Cells(Row, 7).Value = Replace(Cells(Row, 7), QTY & " Per Veh; ", "")
                        Cells(Row, 8).Value = QTY
                    Else
                        If Cells(Row, 7).Value Like "*# Per Veh;*" Then
                            QTY = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "Per Veh;") - 2, 1)
                            Cells(Row, 7).Value = Replace(Cells(Row, 7), QTY & " Per Veh;", "")
                            Cells(Row, 8).Value = QTY
                        Else
                            'MsgBox "check"
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub CutPartType(Row As Integer)

    'cuts out parttype
    Dim PART As String
    
    If Right(PartTypeVar, 5) = "Mount" Then
        If Range("G" & Row).Value Like "*PartType Automatic Transmission Mount*" Then
            Range("G" & Row).Value = Replace(Range("G" & Row), "PartType Automatic Transmission Mount", "")
        Else
            If Range("G" & Row).Value Like "*PartType Manual Transmission Mount*" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), "PartType Manual Transmission Mount", "")
            End If
        End If
    Else
        If InStr(1, Cells(Row, 7), PartTypeVar) > 0 Then
            Cells(Row, 7).Value = Replace(Cells(Row, 7), "PartType " & PartTypeVar, "")
        End If
    End If

End Sub

Private Sub CutMfrLabel(Row As Integer)

    On Error GoTo MfrError

    'cuts out mfrlabel
    Dim LABEL As String, mfrstart, mfrend
    
    If Cells(Row, 7).Value Like "*Mfrlabel*" Then
        mfrstart = InStr(1, Cells(Row, 7), "Mfrlabel")
        mfrend = InStr(mfrstart, Cells(Row, 7), "  ") - 3 'double space indicates end of mfrlabel
        LABEL = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "Mfrlabel") + 9, mfrend - mfrstart - 6) 'not sure why 6, it should be 0 or 3
        If Cells(Row, 7).Value Like "*Mfrlabel*  *" Then
            If Cells(Row, 7).Value Like "*;  Mfrlabel*" Then
                Cells(Row, 7).Value = Replace(Cells(Row, 7), ";  Mfrlabel " & LABEL, "")
                Cells(Row, 9).Value = LABEL
            Else
                Cells(Row, 7).Value = Replace(Cells(Row, 7), "Mfrlabel " & LABEL, "")
                Cells(Row, 9).Value = LABEL
            End If
        Else
            'if there is no triple-space after the Mfrlabel in the notes
            If Cells(Row, 7).Value Like "*" & LABEL Then
                Cells(Row, 7).Value = Replace(Cells(Row, 7), "Mfrlabel " & LABEL, "")
                Cells(Row, 9).Value = LABEL
            End If
        End If
    End If
    
MfrError:
    Exit Sub

End Sub

Private Sub CutVIN(Row As Integer)

    'Cuts out VIN
    Dim VIN As String
    
    If Cells(Row, 7).Value Like "*VIN:*" Then
        If Cells(Row, 7).Value Like "*VIN: ?, ? *" Then
            'not sure what to do when it shows two VINs
        Else
            If Cells(Row, 7).Value Like "*VIN: ?, *" Then
                VIN = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "VIN: ") + 5, 1)
                Cells(Row, 7).Value = Replace(Cells(Row, 7), "VIN: " & VIN & ", ", "")
                Cells(Row, 27).Value = VIN
            Else
                VIN = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "VIN: ") + 5, 1)
                Cells(Row, 7).Value = Replace(Cells(Row, 7), "VIN: " & VIN, "")
                Cells(Row, 27).Value = VIN
            End If
        End If
    End If

End Sub

Private Sub DuplicateFuelType(Row As Integer)
    
    'rst is the Database established with Global variables in FixFitmentsModule
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [FuelType] FROM FuelTypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at beginning of BodyType field
            rst.MoveFirst
            'loop through the Body Types in the master database
            While (Not .EOF)
                'if any of the values in the BodyType field is found in the notes, cut it out and put it in column AH
                If Range("G" & Row).Value Like "*" & .Fields("FuelType").Value & "*" Then
                    If Cells(Row, 34).Value = .Fields("FuelType").Value Then
                        Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelType").Value, "")
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

Private Sub CutDriveType(Row As Integer)
    
    'run query to return body types
    Set rst = MstrDb.Execute("SELECT [DriveType] FROM DriveTypes ORDER BY [ID]")
    
    'got through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at the beginning of the Drive Type field
            rst.MoveFirst
            'loop through the Drive Types in the master database
            While (Not .EOF)
                'if any of the values in the DriveType field is found in the notes, cut it out and put it in column W
                If Range("G" & Row).Value Like "*" & .Fields("DriveType").Value & "*" Then
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("DriveType").Value, "")
                    Cells(Row, 23).Value = .Fields("DriveType").Value
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close the connectiona to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub RemoveAll(Row As Integer)

    'replaces "All" submodels with null
    If Cells(Row, 43).Value = "All" Then
        Cells(Row, 43).Value = ""
    End If

End Sub

Private Sub CutFuelDelivSubtype(Row As Integer)
    
    'rst is the Database established with Global vriables in FixFitmentsModule
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT [FuelDeliverySubtype] FROM FuelDeliverySubtypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at beginning of FuelDeliveryType field
            rst.MoveFirst
            'loop through the Fuel Delivery Types in the master database
            While (Not .EOF)
                'if any of the values in the Fuel Delivery Type field is found in the notes, cut it out and put it in column AE
                If Range("G" & Row).Value Like "*" & .Fields("FuelDeliverySubtype").Value & "*" Then
                    Cells(Row, 30).Value = .Fields("FuelDeliverySubtype").Value
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelDeliverySubtype").Value, "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutFuelDeliveryType(Row As Integer)
    
    'rst is the Database established with Global vriables in FixFitmentsModule
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT [FuelDeliveryType] FROM FuelDeliveryTypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at beginning of FuelDeliveryType field
            rst.MoveFirst
            'loop through the Fuel Delivery Types in the master database
            While (Not .EOF)
                'if any of the values in the Fuel Delivery Type field is found in the notes, cut it out and put it in column AE
                If Range("G" & Row).Value Like "*" & .Fields("FuelDeliveryType").Value & "*" Then
                    Cells(Row, 31).Value = .Fields("FuelDeliveryType").Value
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelDeliveryType").Value, "")
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub DuplicatePartType(Row As Integer)

    'Checks to see if part type is duplicated in notes field
    If PartTypeVar = "Electric Fuel Pump Repair Kit" Then
        Cells(Row, 7).Value = Replace(Cells(Row, 7), "PartType Electric Fuel Pump", "")
    Else
        If Right(PartTypeVar, 5) = "Mount" Then
            If Cells(Row, 7).Value Like "*PartType " & Cells(Row, 9).Value & "*" Then
                Cells(Row, 7).Value = Replace(Cells(Row, 7), "PartType " & Cells(Row, 9).Value, "")
            End If
        Else
            If Cells(Row, 7).Value Like "*" & PartTypeVar & "*" Then
                Cells(Row, 7).Value = Replace(Cells(Row, 7), PartTypeVar, "")
            End If
        End If
    End If

End Sub

Private Sub CutPosition(Row As Integer)
    
    'rst is the Database established with Global vriables in FixFitmentsModule
    'run query to return oxygen sensor positions
    Set rst = MstrDb.Execute("SELECT [Position] FROM OxygenSensorPositions ORDER BY [ID]")                             'ORDER BY [ID] is important, organized in Master Database so it searches Downstream Left before Downstream
    
    'go through each row in column G (notes)
    With Range("G" & Row)
            With rst
                'start at beginning of position field
                rst.MoveFirst
                'loop through the Oxygen Sensor Positions in the master database
                While (Not .EOF)
                    'if any of the values in the Position fiel is found in the notes
                    If Range("G" & Row).Value Like "*Position " & .Fields("Position").Value & "*" Then
                        'cut out position and put it in column J
                        Cells(Row, 10).Value = .Fields("Position").Value
                        Cells(Row, 7).Value = Replace(Cells(Row, 7), "Position " & .Fields("Position").Value, "")
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

Private Sub CutValves(Row As Integer)

    Dim valves As String

    'Test for string "VALVE"
    If Range("G" & Row).Value Like "*VALVE*" Then
        'if G contains 2-digit valves
        If Range("G" & Row).Value Like "*## VALVE*" Then
            valves = Mid(Range("G" & Row).Value, InStr(1, Range("G" & Row), " VALVE") - 2, 2)
            Range("AW" & Row).Value = valves
            Range("G" & Row).Value = Replace(Range("G" & Row), Range("AW" & Row) & " VALVES", "")
        Else
            'if G contains single-digit valves
            If Range("G" & Row).Value Like "*# VALVE*" Then
                valves = Mid(Range("G" & Row), InStr(1, Range("G" & Row), " VALVE") - 1, 1)
                Range("AW" & Row).Value = valves
                Range("G" & Row).Value = Replace(Range("G" & Row), Range("AW" & Row) & " VALVES", "")
            Else
                'if G ends in 2-digit valves
                If Range("G" & Row).Value Like "*## VALVE" Then
                    valves = Mid(Range("G" & Row).Value, InStr(1, Range("G" & Row), " VALVES") - 2, 2)
                    Range("AW" & Row).Value = valves
                    Range("G" & Row).Value = Left(Range("G" & Row), Len(Range("G" & Row)) - 9)
                Else
                    'if G ends in single-digit valves
                    If Range("G" & Row).Value Like "*# VALVE" Then
                        valves = Mid(Range("G" & Row), InStr(1, Range("G" & Row), " VALVES") - 1, 1)
                        Range("AW" & Row).Value = valves
                        Range("G" & Row).Value = Left(Range("G" & Row), Len(Range("G" & Row)) - 8)
                    Else
                        'if G starts with 2-digit valves
                        If Range("G" & Row).Value Like "## VALVE*" Then
                            valves = Left(Range("G" & Row), 2)
                            Range("AW" & Row).Value = valves
                            Range("G" & Row).Value = Right(Range("G" & Row), Len(Range("G" & Row)) - 9)
                        Else
                            'if G starts with single-digit valves
                            If Range("G" & Row).Value Like "# VALVE*" Then
                                valves = Left(Range("G" & Row), 1)
                                Range("AW" & Row).Value = valves
                                Range("G" & Row).Value = Right(Range("G" & Row), Len(Range("G" & Row)) - 8)
                            Else
                                'if G is only 2-digit valves
                                If Range("G" & Row).Value Like "## Valve" Then
                                    valves = Left(Range("G" & Row), 2)
                                    Range("G" & Row).Value = ""
                                    Range("AW" & Row).Value = valves
                                Else
                                    'if G is only single-digit valves
                                    If Range("G" & Row).Value Like "# Valve" Then
                                        valves = Left(Range("G" & Row), 1)
                                        Range("G" & Row).Value = ""
                                        Range("AW" & Row).Value = valves
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
        If Range("G" & Row).Value Like "*Valves: ##*" Then
            valves = Mid(Range("G" & Row), InStr(1, Range("G" & Row), "Valves: ") + 8, 2)
            Range("AW" & Row).Value = valves
            Range("G" & Row).Value = Replace(Range("G" & Row), "Valves: " & Range("AW" & Row).Value, "")
        Else
            If Range("G" & Row).Value Like "*Valves: #*" Then
                valves = Mid(Range("G" & Row), InStr(1, Range("G" & Row), "Valves: ") + 8, 1)
                Range("AW" & Row).Value = valves
                Range("G" & Row).Value = Replace(Range("G" & Row), "Valves: " & Range("AW" & Row).Value, "")
            End If
        End If
    End If

End Sub

Private Sub CutEngineDesignation(Row As Integer)

    'run query to return oxygen sensor positions
    Set rst = MstrDb.Execute("SELECT [EngineDesignation] FROM EngineDesignations ORDER BY [ID]")                           'ORDER BY [ID] is important, organized in Master Database
        
    'go through each row in column G (notes)
    'check if notes contains the string "Engine: "
    If Range("G" & Row).Value Like "*Engine: *" Then
        'if yes, search through Engien Designations in Master Database"
        With Range("G" & Row)
            With rst
                'start at beginning of position field
                rst.MoveFirst
                'loop through the Oxygen Sensor Positions in the master database
                While (Not .EOF)
                    'if any of the values in the Position fiel is found in the notes
                    If Range("G" & Row).Value Like "*" & "Engine: " & .Fields("EngineDesignation").Value & "*" Then
                        'cut out position and put it in column J
                        Cells(Row, 24).Value = .Fields("EngineDesignation").Value
                        Cells(Row, 7).Value = Replace(Cells(Row, 7), "Engine: " & .Fields("EngineDesignation").Value, "")
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

Private Sub CutTransmissionControlType(Row As Integer)

    'run query to return transmission control types
    Set rst = MstrDb.Execute("SELECT [TransControlType] FROM TransControlTypes ORDER BY [ID]")                         'ORDER BY [ID] is important, organized in Master Database so it searches Automatic CVT before Automatic
    
    'go through each row in column G (notes)
    With Range("G" & Row)
            With rst
                'start at beginning of position field
                rst.MoveFirst
                'loop through the Oxygen Sensor Positions in the master database
                While (Not .EOF)
                    'if any of the values in the Position fiel is found in the notes
                    If Range("G" & Row).Value Like "*" & .Fields("TransControlType").Value & " Trans*" Then
                        'cut out position and put it in column AR
                        Cells(Row, 44).Value = .Fields("TransControlType").Value
                        Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("TransControlType").Value & " Trans", "")
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

Private Sub CutSpeeds(Row As Integer)
    
    'Check to see if Notes contains a two-digit number of speeds
    If Range("G" & Row).Value Like "*## Speed Trans*" Then
        'Cuts out the two-digit number and places in AU
        Cells(Row, 47).Value = Mid(Range("G" & Row).Value, InStr(1, Range("G" & Row), " Speed Trans") - 2, 2)
        'erases the "## Speed Trans" string from G
        Cells(Row, 7).Value = Replace(Cells(Row, 7), Range("AU" & Row).Value & " Speed Trans", "")
    Else
        'Check to see if Notes contains a single-digit number of speeds
        If Range("G" & Row).Value Like "*# Speed Trans*" Then
            'Cuts out the single-digit number and places in AU
            Cells(Row, 47).Value = Mid(Cells(Row, 7).Value, InStr(1, Range("G" & Row), " Speed Trans") - 1, 1)
            'erases the "# Speed Trans" string from G
            Cells(Row, 7).Value = Replace(Cells(Row, 7), Range("AU" & Row).Value & " Speed Trans", "")
        End If
    End If

End Sub

Private Sub CutTransType(Row As Integer)

    If Range("G" & Row).Value Like "* Transaxle*" Then
        Range("AV" & Row).Value = "Transaxle"
        Cells(Row, 7).Value = Replace(Cells(Row, 7), "Transaxle", "")
    End If

End Sub

Private Sub CutWheelBase(Row As Integer)

    Dim wb As Integer
    Dim WheelBase As String
    
    If Range("G" & Row).Value Like "*###.#"" WB*" Then
        wb = InStr(1, Cells(Row, 7), "WB")
        WheelBase = Mid(Cells(Row, 7), wb - 7, 5)   'take the 5 characters that start 7 characters before the "WB" in the notes
        Range("AX" & Row).Value = WheelBase
        Range("G" & Row).Value = Replace(Range("G" & Row), WheelBase & """ WB", "")
    End If

End Sub

Private Sub CutTransmissionMfrCode(Row As Integer)

    'open Transmission Mfr Codes table in Master Database
    Set rst = MstrDb.Execute("SELECT [TransmissionMfrCode] FROM TransmissionMfrCodes ORDER BY [ID]")
    
    'first check to see if there is something in the G column left
    If Range("G" & Row).Value <> "" Then
        With rst
        .MoveFirst
            While Not .EOF
                'if the non-blank cell contains a TransmissionMfrCode
                If Range("G" & Row).Value Like "* " & .Fields("TransmissionMfrCode").Value & ", *" Then
                    Range("AT" & Row).Value = .Fields("TransmissionMfrCode").Value
                    Range("G" & Row).Value = Replace(Range("G" & Row), .Fields("TransmissionMfrCode").Value & ", ", "")
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

Private Sub CutTransmissionType(Row As Integer)

    'open Transmission types in Master Database
    Set rst = MstrDb.Execute("SELECT [TransmissionType] FROM TransmissionTypes ORDER BY [ID]")
    
    rst.MoveFirst
    With rst
        If Range("G" & Row).Value Like "*" & .Fields("TransmissionType").Value & "*" Then
            Range("AV" & Row).Value = .Fields("TransmissionType").Value
            Range("G" & Row).Value = Replace(Range("G" & Row), .Fields("TransmissionType").Value, "")
            GoTo Exit_Loop
        Else
            .MoveNext
        End If
    End With

Exit_Loop:
    rst.Close

End Sub

Private Sub CutValvesPerEngine(Row As Integer)

    Dim valves As Integer
    Dim percyl As String
    
    If Range("G" & Row).Value Like "*valves per cylinder*" Then
        percyl = Mid(Range("G" & Row), InStr(1, Range("G" & Row).Value, "valves per cylinder") - 2, 1)
        valves = percyl * Range("V" & Row).Value
        Range("AW" & Row).Value = valves
        Range("G" & Row).Value = Replace(Range("G" & Row), percyl & " valves per cylinder", "")
    End If
    
    If Range("G" & Row).Value Like "*## Valve*" Then
        valves = Mid(Range("G" & Row), InStr(1, Range("G" & Row), "Valve") - 3, 2)
        Range("AW" & Row).Value = valves
        Range("G" & Row).Value = Replace(Range("G" & Row), valves & " Valve", "")
    ElseIf Range("G" & Row).Value Like "*# Valve*" Then
        valves = Mid(Range("G" & Row), InStr(1, Range("G" & Row), "Valve") - 3, 1)
        Range("AW" & Row).Value = valves
        Range("G" & Row).Value = Replace(Range("G" & Row), valves & " Valve", "")
    End If

End Sub

Private Sub DuplicateSubmodel(Row As Integer)

    If Range("G" & Row).Value Like "*" & Range("AQ" & Row).Value & "*" Then
        Range("G" & Row).Value = Replace(Range("G" & Row), Range("AQ" & Row).Value, "")
    End If

End Sub

Private Sub DuplicateBodyType(Row As Integer)

    If Range("G" & Row).Value Like "*" & Range("P" & Row).Value & "*" Then
        Range("G" & Row).Value = Replace(Range("G" & Row), Range("P" & Row).Value, "")
    End If

End Sub

Private Sub CleanGap(Row As Integer)

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

Private Sub Cleanup(Row As Integer)

    Dim l As Integer, BigLoop
    
    l = 0
    BigLoop = 0
    
    'If the notes field contains the part type, remove it
    If Range("G" & Row).Value Like "*Part Note: *" Then
        Range("G" & Row).Value = Replace(Range("G" & Row), "Part Note: ", "")
    End If
    
    'If the notes field contains the number of doors, remove it
    If Range("G" & Row).Value Like "*" & Range("O" & Row).Value & " Door*" Then
        Range("G" & Row).Value = Replace(Range("G" & Row), Range("O" & Row).Value & " Door", "")
    End If
    
    'If the notes field contains the cylinder head type, remove it
    If Range("G" & Row).Value Like "*" & Range("U" & Row).Value & "*" Then
        Range("G" & Row).Value = Replace(Range("G" & Row), Range("U" & Row).Value, "")
    End If
    
    'If notes has Natural
    If Range("G" & Row).Value Like "*Natural*" And Range("K" & Row).Value = "Naturally Aspirated" Then
        Range("G" & Row).Value = Replace(Range("G" & Row), "Natural", "")
    End If
    
    Do While BigLoop < 6
        Do While l = 0
            'remove double semicolons
            If Range("G" & Row).Value Like "*;;*" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), ";;", ";")
            Else
                l = 1
            End If
        Loop
        
        l = 0
            
        Do While l = 0
            'remove double spaces
            If Range("G" & Row).Value Like "*  *" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), "  ", " ")
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        Do While l = 0
            'remove double commas
            If Range("G" & Row).Value Like "*,,*" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), ",,", ",")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove beginning spaces
            If Range("G" & Row).Value Like " *" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), Range("G" & Row), Right(Range("G" & Row), Len(Range("G" & Row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove beginning semicolons
            If Range("G" & Row).Value Like ";*" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), Range("G" & Row), Right(Range("G" & Row), Len(Range("G" & Row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove beginning semicolons
            If Range("G" & Row).Value Like ",*" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), Range("G" & Row), Right(Range("G" & Row), Len(Range("G" & Row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove final semicolons
            If Range("G" & Row).Value Like "*;" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), Right(Range("G" & Row), 1), "")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove final spaces
            If Range("G" & Row).Value Like "* " Then
                Range("G" & Row).Value = Replace(Range("G" & Row), Range("G" & Row), Left(Range("G" & Row), Len(Range("G" & Row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove lonely semicolons
            If Range("G" & Row).Value Like "* ; *" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), " ; ", "; ")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove lonely commas
            If Range("G" & Row).Value Like "* , *" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), " , ", ", ")
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove final hyphens
            If Range("G" & Row).Value Like "*-" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), Range("G" & Row), Left(Range("G" & Row), Len(Range("G" & Row)) - 1))
            Else
                l = 1
            End If
        Loop
        
        l = 0

        Do While l = 0
            'remove comma-semicolons
            If Range("G" & Row).Value Like "* ,; *" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), " ,; ", " ")
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        Do While l = 0
            'remove semi-colon-commas
            If Range("G" & Row).Value Like "*;, *" Then
                Range("G" & Row).Value = Replace(Range("G" & Row), ";, ", "; ")
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        Do While l = 0
            'remove final commas
            If Range("G" & Row).Value Like "*," Then
                Range("G" & Row).Value = Left(Range("G" & Row), Len(Range("G" & Row)) - 1)
            Else
                l = 1
            End If
        Loop
        
        l = 0
        
        BigLoop = BigLoop + 1
    Loop

End Sub

Private Sub ReplicatePart(Row As Integer)

    'Adds Part number, part type, brand_code, and sku
    Cells(Row, 1).Value = PartName
    Cells(Row, 2).Value = "FVKX"
    Cells(Row, 6).Value = PartTypeVar
    Cells(Row, 51).Value = gendSKU

End Sub

Private Sub ErrorChecker(Row As Integer)

    

End Sub

Private Sub CheckMount(numrows As Integer)

    Dim unmatched As Boolean
    Dim LResponse As Integer
    Dim Row As Integer

    unmatched = False
    
    If Right(PartTypeVar, 5) = "Mount" Then
        For Row = 2 To numrows
            If Range("I" & Row).Value = "" Then GoTo Next_row
            
            If Not (Replace(Range("F" & Row).Value, "Automatic Transmission", "Auto Trans") = Range("I" & Row).Value Or Replace(Range("F" & Row).Value, "Manual Transmission", "Manual Trans") = Range("I" & Row).Value) Then
                unmatched = True
                Row = numrows   'ends the loop
            End If
Next_row:
        Next Row
    Else
        Exit Sub
    End If
    
    If unmatched = True Then
        LResponse = MsgBox("Some of the notes don't match the part type entered. Do you want to update to the part type in the notes?", vbYesNo, "Part Type Mismatch")
    
        'Determine Yes/No options
        If LResponse = vbYes Then
            For Row = 2 To numrows
                If Range("I" & Row).Value = "Auto Trans Mount" Then
                    Range("F" & Row).Value = "Automatic Transmission Mount"
                Else
                    If Range("I" & Row).Value = "Manual Trans Mount" Then
                        Range("F" & Row).Value = "Manual Transmission Mount"
                    End If
                End If
            Next Row
        End If
    End If

End Sub

Private Sub Headers()

    'These are the ACES headers
    Range("A1:V1").Value = [{"part", "brand_code", "make", "model", "year", "partterminologyname", "notes", "qty", "mfrlabel", "position", "aspiration","bedlength","bedtype","block","bodynumdoors","bodytype","brakeabs","brakesystem","cc","cid","cylinderheadtype","cylinders"}]
    Range("W1:AK1").Value = [{"drivetype", "enginedesignation","enginemfr","engineversion","enginevin","frontbraketype","frontspringtype","fueldeliverysubtype","fueldeliverytype","fuelsystemcontroltype","fuelsystemdesign","fueltype","ignitionsystemtype", "liters","mfrbodycode"}]
    Range("AL1:AX1").Value = [{"rearbraketype", "rearspringtype","region","steeringsystem","steeringtype","submodel","transmissioncontroltype","transmissionmfr","transmissionmfrcode","transmissionnumspeeds", "transmissiontype", "valvesperengine", "wheelbase"}]
    Range("AY1").Value = "sku"

End Sub
