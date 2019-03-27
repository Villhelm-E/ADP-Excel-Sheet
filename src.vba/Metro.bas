Option Explicit

Public Sub MetroMain()

    'user successfully expanded all fitments in Metro, run code, otherwise warn user and highlight missed expansions
    If Expanded = True Then
    
        'turn screen updating off
        Application.ScreenUpdating = False
        
        'Move the data to the left by deleting first 3 columns
        columns("A:C").Delete Shift:=xlToLeft
        
        'Count how many rows there are in columns ABC
        Dim numrows As Integer
        numrows = CountRows("A:C")
        
        'Move the Engines from column A to column G
        Call MoveEngines(numrows)
        
        'Remove year range from Makes and move the count of Models to column D
        Call CleanMakes(numrows)
        
        'remove the year range from Models
        Call ModelsPass1(numrows)
        
        'Count number of engines per model and finish cleaning Models
        Call ModelsPass2(numrows)
        
''''''''Start removing blank cells between values
        
        'Remove empty cells in Engines in column G
        ConsolidateEngines
        ConsolidateNumPerVeh
        ConsolidateModels
        ConsolidateMakes
        ConsolidateMakeCounts
        ConsolidateModelCounts
        
        'Repeat Makes according to how many models
        RepeatMakes
        
        'Repeat Models according to how many engines
        RepeatModels
        
        'Don't need info in columns A, B, E, F anymore
        Range("A:B").Clear
        Range("E:F").Clear
        
''''''''Loop through each row, moving things around as appropriate
        
        'Format columns
        FormatColumns
        
        'count number of rows in column C
        numrows = CountRows("C:C")
        
        'Loop through every row to move everything in it's place
        Call BigLoop(numrows)
        
        'Add Headers
        Headers
        
        'Autofit fields and select A1
        columns("A:AX").AutoFit
        Range("A1").Select
        
        'Rename sheet
        RenameSheet         'WorksheetConnections Module
        
        'turn screen updating on
        Application.ScreenUpdating = True
        
        'Let user know tha formatting has completed
        MsgBox "Done formatting fitments"
    Else
        MsgBox ("You missed the expansion of some Makes or Models in Metro. They have been highlighted. Go back to Metro and make sure you expanded all Makes and Models.")
    End If

End Sub

'returns true if user successfully expanded all fitments in Metro, false if user missed one
'also highlights missed expansions
Private Function Expanded() As Boolean

    'default function to true
    Expanded = True

    Dim R As Range
    
    'checks every line for two consecutive blanks in column D
    For Each R In Intersect(Range("D:D"), ActiveSheet.UsedRange)
        If R.Value = "" And R.Offset(0, 1).Value = "" And R.Offset(0, 2).Value = "" Then
            'if it finds consecutive blanks, it sets function to false
            Expanded = False
            If R.Offset(-1, 0).Value = "" Then
                'if the cell to the left of the blank cell found above is blank, it highlights the cell above the cell to the left
                'this is the Make that was not expanded in Metro
                With R.Offset(-1, 1).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                End With
            Else
                'if the cell to the left is not blank, this code will highlight it
                'this is a Model that was not expanded in Metro
                With R.Offset(-1, 0).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                End With
            End If
            
        End If
        
        'need special code to determine if the last Make or Model in the fitment list is expanded or not
        'can probably merge this code with the code above in the future
        If IsNumeric(Left(R.Value, 4)) = False And R.Offset(0, 1).Value = "" And R.Offset(0, 2).Value = "" And R.Offset(1, 0).Value = "" And R.Offset(2, 0).Value = "" And _
        R.Offset(1, 2).Value = "" And Not R.Offset(1, 1).Value = "" Then
            Expanded = False
            With R.Offset(1, 1).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
            End With
        Else
            If IsNumeric(Left(R.Value, 4)) = False And R.Offset(0, 1).Value = "" And R.Offset(0, 2).Value = "" And R.Offset(1, 0).Value = "" And R.Offset(2, 0).Value = "" And _
            R.Offset(1, 1).Value = "" And R.Offset(1, 2).Value = "" Then
                Expanded = False
                With R.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                End With
            End If
        End If
        
    Next R

End Function

Private Sub MoveEngines(numrows As Integer)

    Dim k As Integer
    
    For k = 1 To numrows
        'loop through each row and move the engine to column G
        If IsNumeric(Left(Cells(k, 1).Value, 4)) Then
            Cells(k, 1).Cut ActiveSheet.Cells(k, 7)
        End If
    Next k

End Sub

Private Sub CleanMakes(numrows As Integer)

    Dim k As Integer
    Dim ModelCount As Integer
    
    'loop through to remove year range from Make
    For k = 1 To numrows
        If Not Cells(k, 1).Value = "" Then
            'remove the year range
            Cells(k, 1).Value = Left(Cells(k, 1), Len(Cells(k, 1)) - 9)
            
            'remove the count of models and move to column D
            If IsNumeric(Right(Cells(k, 1), 4)) Then
                ModelCount = Right(Cells(k, 1), 4)
                Cells(k, 1).Value = Left(Cells(k, 1), Len(Cells(k, 1)) - 4)
                Cells(k, 4).Value = ModelCount
            Else
                If IsNumeric(Right(Cells(k, 1), 3)) Then
                    ModelCount = Right(Cells(k, 1), 3)
                    Cells(k, 1).Value = Left(Cells(k, 1), Len(Cells(k, 1)) - 3)
                    Cells(k, 4).Value = ModelCount
                Else
                    If IsNumeric(Right(Cells(k, 1), 2)) Then
                        ModelCount = Right(Cells(k, 1), 2)
                        Cells(k, 1).Value = Left(Cells(k, 1), Len(Cells(k, 1)) - 2)
                        Cells(k, 4).Value = ModelCount
                    Else
                        ModelCount = Right(Cells(k, 1), 1)
                        Cells(k, 1).Value = Left(Cells(k, 1), Len(Cells(k, 1)) - 1)
                        Cells(k, 4).Value = ModelCount
                    End If
                End If
            End If
        End If
    Next k

End Sub

Private Sub ModelsPass1(numrows)

    Dim k As Integer
    
    For k = 1 To numrows
        If Not Cells(k, 2).Value = "" Then
            'removes year range from Model
            Cells(k, 2).Value = Left(Cells(k, 2), Len(Cells(k, 2)) - 9)
        End If
    Next k

End Sub

Private Sub ModelsPass2(numrows As Integer)

    Dim ModelPos As Integer
    Dim EnginePos As Integer
    Dim ModelCount As Integer
    
    For ModelPos = 1 To numrows
        If Not Cells(ModelPos, 2).Value = "" Then
            'count the Engines in column G under Model in column B
            For EnginePos = ModelPos + 1 To numrows
                If Not Cells(EnginePos, 7).Value = "" Then
                    'Count the Engine
                    ModelCount = ModelCount + 1
                Else
                    'place modelcount in column E
                    Cells(ModelPos, 5).Value = ModelCount
                    
                    'remove model count from Model
                    Cells(ModelPos, 2).Value = Left(Cells(ModelPos, 2), Len(Cells(ModelPos, 2)) - Len(CStr(ModelCount)))
                    
                    'reset ModelCount
                    ModelCount = 0
                    
                    'exit engine loop to move onto next model
                    GoTo Exit_Loop
                End If
            Next EnginePos
        End If
Exit_Loop:
    Next ModelPos

End Sub

Private Sub ConsolidateEngines()

    Dim k As Integer
    Dim R As Range
    
    k = 1
    
    'loop through all cells in column G to remove blank cells
    For Each R In Intersect(Range("G:G"), ActiveSheet.UsedRange)
        If Not R.Value = "" Then
            R.Cut ActiveSheet.Cells(k, 7)
            k = k + 1
        End If
    Next R

End Sub

Private Sub ConsolidateNumPerVeh()

    Dim k As Integer
    Dim R As Range
    
    k = 1
    
    'loop through all cells in column C and move them to column H without blank cells
    For Each R In Intersect(Range("C:C"), ActiveSheet.UsedRange)
        If Not R.Value = "" Then
            If R.Value Like "## per Vehicle" Then
                Cells(k, 8).Value = Left(R, 2)
                R.Value = ""
            Else
                Cells(k, 8).Value = Left(R, 1)
                R.Value = ""
            End If
            k = k + 1
        End If
    Next R

End Sub

Private Sub ConsolidateModels()

    Dim k As Integer
    Dim R As Range
    
    k = 1
    
    'loop through all cells in column B to remove blank cells
    For Each R In Intersect(Range("B:B"), ActiveSheet.UsedRange)
        If Not R.Value = "" Then
            R.Cut ActiveSheet.Cells(k, 2)
            k = k + 1
        End If
    Next R

End Sub

Private Sub ConsolidateMakes()

    Dim k As Integer
    Dim R As Range
    
    k = 1
    
    'loop through all cells in column A to remove blank cells
    For Each R In Intersect(Range("A:A"), ActiveSheet.UsedRange)
        If Not R.Value = "" Then
            R.Cut ActiveSheet.Cells(k, 1)
            k = k + 1
        End If
    Next R

End Sub
        
Private Sub ConsolidateMakeCounts()

    Dim k As Integer
    Dim R As Range
    
    k = 1
    
    'loop through every cell in column E and move to column F without blank cells
    For Each R In Intersect(Range("E:E"), ActiveSheet.UsedRange)
        If Not R.Value = "" Then
            R.Cut ActiveSheet.Cells(k, 6)
            k = k + 1
        End If
    Next R

End Sub
        
Private Sub ConsolidateModelCounts()

    Dim k As Integer
    Dim R As Range
    
    k = 1
    
    'loop through every cell in column D and move to column E withouth blank cells
    For Each R In Intersect(Range("D:D"), ActiveSheet.UsedRange)
        If Not R.Value = "" Then
            R.Cut ActiveSheet.Cells(k, 5)
            k = k + 1
        End If
    Next R

End Sub

Private Sub RepeatMakes()

    'Repeats Make number of times in column E to D
    Dim lRow As Integer
    Dim LQty As Integer
    Dim LProduct As String
    Dim LColCPosition As Integer
    Dim j As Integer
    Dim lStart As Integer
    Dim LEnd As Integer
            
    'Search for values in column E starting at row 1
    lRow = 1
            
    'Copy values to column B starting at row 1
    LColCPosition = 1
            
    'Search through values in column E until a blank cell is encountered
    While Len(Range("A" & CStr(lRow)).Value) > 0
        'Retrieve quantity and Model
        LQty = Range("E" & CStr(lRow)).Value
        LProduct = Range("A" & CStr(lRow)).Value
        
        'Set start and end position for copy to column B
        lStart = LColCPosition
        LEnd = LColCPosition + LQty
        
        'Copy Model name the number of times that is given by the quantity
        For j = lStart To LEnd - 1
            Range("C" & CStr(j)).Value = LProduct
        Next
      
        'Update column B position
        LColCPosition = LEnd
        
        lRow = lRow + 1
    Wend

End Sub

Private Sub RepeatModels()

    Dim lRow As Integer
    Dim LQty As Integer
    Dim LProduct As String
    Dim LColCPosition As Integer
    Dim j As Integer
    Dim lStart As Integer
    Dim LEnd As Integer
    
    lRow = 1
    LColCPosition = 1

    'Repeats Model number of times in column F
    While Len(Range("B" & CStr(lRow)).Value) > 0
        'Retrieve quantity and Model
        LQty = Range("F" & CStr(lRow)).Value
        LProduct = Range("B" & CStr(lRow)).Value
            
        'Set start and end position for copy to column B
        lStart = LColCPosition
        LEnd = LColCPosition + LQty
      
        'Copy Model name the number of times that is given by the quantity
        For j = lStart To LEnd - 1
            Range("D" & CStr(j)).Value = LProduct
        Next
      
        'Update column B position
        LColCPosition = LEnd
    
        lRow = lRow + 1
    Wend

End Sub

Private Sub FormatColumns()

    'Format column 36 as text so that Excel doesn't remove ".0" from liters like "5.0"
    Range("AJ:AJ").NumberFormat = "@"

End Sub

Private Sub BigLoop(numrows As Integer)

    Dim i As Integer
    
    For i = 1 To numrows
        'Cut Year out of engine info
        Call CutYear(i)
        
        'Cut Liters out of engine info
        Call CutLiters(i)
        
        'Cut CC out of engine info
        Call CutCC(i)
        
        'Cut cid out of engine info
        Call CutCID(i)
        
        'Cut cylinders out of engine info
        Call CutCylinders(i)
        
        'Cut Cylinder Head Type out of engine info
        Call CutCylinderHeadType(i)
        
        'Cut Aspirations out of engine info
        Call CutAspiration(i)
        
        'Cut valves per engine out of engine info
        Call CutValvesPerEngine(i)
        
        'Cut fuel type out of engine info
        Call CutFuelType(i)
        
        'Cut Fuel Delivery Type out of engine info
        Call CutFuelDeliveryType(i)
        
        'Cut VIN out of engine info
        Call CutVIN(i)
        
        'Cut Trim out of engine info
        Call CutTrim(i)
        
        'Cut Mfr Label
        Call CutMfrLabel(i)
        
        'Fill column A with the part number
        Call FillPartNum(i)
    
    Next i

End Sub

Private Sub CutYear(Row As Integer)

    Dim year As String
    
    'save first 4 numbers to variable
    year = Left(Cells(Row, 7), 4)
    
    'cut out the year from the engine info
    Cells(Row, 7).Value = Right(Cells(Row, 7), Len(Cells(Row, 7)) - 5)
    
    'Place year in what will be the ACES year field
    Cells(Row, 5).Value = year

End Sub

Private Sub CutLiters(Row As Integer)

    Dim Volume As String
 
    If Cells(Row, 7).Value Like "##.#L *" Then
        'save the liters to variable
        Volume = Left(Cells(Row, 7), 4)
        
        'Cut liters from column G
        Cells(Row, 7).Value = Right(Cells(Row, 7), Len(Cells(Row, 7)) - 6)
        
        'place liters in column
        Cells(Row, 7).Value = Volume
    Else
        'save the liters to variable
        Volume = Left(Cells(Row, 7), 3)
        
        'cut liters from column G
        Cells(Row, 7).Value = Right(Cells(Row, 7), Len(Cells(Row, 7)) - 5)
        
        'place liters in column
        Cells(Row, 36).Value = Volume
    End If

End Sub

Private Sub CutCC(Row As Integer)

    'Cuts out CC
    Dim cc As String

    If Cells(Row, 7).Value Like "###cc *" Then
        'save cc to variable
        cc = Left(Cells(Row, 7).Value, 3)
        
        'cut cc out from column G
        Cells(Row, 7).Value = Right(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 6)
        
        'place cc in column S
        Cells(Row, 19).Value = cc
    Else
        If Cells(Row, 7).Value Like "####cc *" Then
            'save cc to variable
            cc = Left(Cells(Row, 7).Value, 4)
            
            'cut cc out of column G
            Cells(Row, 7).Value = Right(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 7)
            
            'place cc in column S
            Cells(Row, 19).Value = cc
        End If
    End If

End Sub

Private Sub CutCID(Row As Integer)

    'Cuts out cid
    Dim cid As String
    
    If Cells(Row, 7).Value Like "##cid *" Then
        'save cid to variable
        cid = Left(Cells(Row, 7).Value, 2)
        
        'cut cid out of column G
        Cells(Row, 7).Value = Right(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 6)
        
        'place cid in column T
        Cells(Row, 20).Value = cid
    Else
        If Cells(Row, 7).Value Like "###cid *" Then
            'save cid to variable
            cid = Left(Cells(Row, 7).Value, 3)
            
            'cut cid out of column G
            Cells(Row, 7).Value = Right(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 7)
            
            'place cid in column T
            Cells(Row, 20).Value = cid
        End If
    End If

End Sub

Private Sub CutCylinders(Row As Integer)

    Dim block As String
    Dim Cyl As String

    'Find Block+Cylinders in column G
    If Cells(Row, 7).Value Like "L##*" Or Cells(Row, 7).Value Like "V##*" Or Cells(Row, 7).Value Like "H##*" Then
        'if the number of cylinders is in the double digits
        block = Left(Cells(Row, 7), 1)
        Cyl = Mid(Cells(Row, 7), 2, 2)
        
        Cells(Row, 14).Value = block    'column 14 = N
        Cells(Row, 22).Value = Cyl      'column 22 = V
        
        'if there's nothing after the block+cylinders, then leave cell blank
        If Cells(Row, 7).Value = block & Cyl Then
            Cells(Row, 7).Value = ""
        Else
            Cells(Row, 7).Value = Right(Cells(Row, 7), Len(Cells(Row, 7)) - 4)  'remove block+cylinders from column G
        End If
    Else
        'if the number of cylinders is in the single digits
        If Cells(Row, 7).Value Like "L#*" Or Cells(Row, 7).Value Like "V#*" Or Cells(Row, 7).Value Like "H#*" Then
            block = Left(Cells(Row, 7), 1)
            Cyl = Mid(Cells(Row, 7), 2, 1)
            
            Cells(Row, 14).Value = block    'column 14 = N
            Cells(Row, 22).Value = Cyl      'column 22 = V
            
            'if there's nothing after the block+cylinders, then leave cell blank
            If Cells(Row, 7).Value = block & Cyl Then
                Cells(Row, 7).Value = ""
            Else
                Cells(Row, 7).Value = Right(Cells(Row, 7), Len(Cells(Row, 7)) - 3)
            End If
        End If
    End If

End Sub

Private Sub CutCylinderHeadType(Row As Integer)
    
    'run query to return part types
    'Open Excel Sheet Version table from Master Database
    Set rst = MstrDb.Execute("SELECT [CylinderHeadType] FROM CylinderHeadTypes ORDER BY [ID]") 'rst is global variable
    
    'go through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at beginning of CylinderHeadType field
            rst.MoveFirst
            'loop through the Cylinder Head Types in the master database
            While (Not .EOF)
                'if any of the values in the CylinderHEadType field is found in the notes, cut it out and put it in column U
                If Range("G" & Row).Value Like "*" & .Fields("CylinderHeadType").Value & "*" Then
                    Cells(Row, 21).Value = .Fields("CylinderHeadType").Value
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("CylinderHeadType").Value & " ", "")
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("CylinderHeadType").Value, "")
                    GoTo Exit_Loop
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
Exit_Loop:
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutAspiration(Row As Integer)

    Dim aspiration As String
    
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT [Aspiration] FROM Aspirations ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at beginning of Aspiration field
            rst.MoveFirst
            'loop through the Aspiration Types in the master database
            While (Not .EOF)
                'if any of the values in the Aspiration field is found in the notes, cut it out and put it in column K
                If Range("G" & Row).Value Like "*" & .Fields("Aspiration").Value & "*" Then
                    If .Fields("Aspiration").Value = "Turbo" Then
                        aspiration = "Turbocharged"
                    Else
                        aspiration = .Fields("Aspiration").Value
                    End If
                    Cells(Row, 11).Value = aspiration
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("Aspiration").Value & " ", "")
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("Aspiration").Value, "")
                    GoTo Exit_Loop
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
Exit_Loop:
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutValvesPerEngine(Row As Integer)
    Dim valve As String
    
    If Cells(Row, 7).Value Like "# Valve *" Then
        valve = Left(Cells(Row, 7), 1)
        Cells(Row, 7).Value = Right(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 2)
        Cells(Row, 49).Value = valve
    Else
        If Cells(Row, 7).Value Like "## Valve *" Then
            valve = Left(Cells(Row, 7), 2)
            Cells(Row, 7).Value = Right(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 3)
            Cells(Row, 49).Value = valve
        Else
            If Cells(Row, 7).Value Like "# Valve" Then
                valve = Left(Cells(Row, 7), 1)
                Cells(Row, 7).Value = ""
                Cells(Row, 49).Value = valve
            Else
                If Cells(Row, 7).Value Like "## Valve" Then
                    valve = Left(Cells(Row, 7), 2)
                    Cells(Row, 7).Value = ""
                    Cells(Row, 49).Value = valve
                End If
            End If
        End If
    End If

End Sub

Private Sub CutFuelType(Row As Integer)
    
    'run query to return part types
    Set rst = MstrDb.Execute("SELECT [FuelType] FROM FuelTypes ORDER BY [ID]")
    
    'go through each row in column G (notes)
    With Range("G" & Row)
        With rst
            'start at beginning of FuelType field
            rst.MoveFirst
            'loop through the Fuel Types in the master database
            While (Not .EOF)
                'if any of the values in the Fuel Type field is found in the notes, cut it out and put it in column AH
                If Range("G" & Row).Value Like "*" & .Fields("FuelType").Value & "*" Then
                    Cells(Row, 34).Value = .Fields("FuelType").Value
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelType").Value & " ", "")
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelType").Value, "")
                    GoTo Exit_Loop
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
Exit_Loop:
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
            'loop through the Fuel Delivery Types int he master database
            While (Not .EOF)
                'if any of the values in the Fuel Delivery Type field is found in the notes, cut it out and put it in column AE
                If Range("G" & Row).Value Like "*" & .Fields("FuelDeliveryType").Value & "*" Then
                    Cells(Row, 31).Value = .Fields("FuelDeliveryType").Value
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelDeliveryType").Value & " ", "")
                    Cells(Row, 7).Value = Replace(Cells(Row, 7), .Fields("FuelDeliveryType").Value, "")
                    GoTo Exit_Loop
                End If
                rst.MoveNext
            Wend
        End With
    End With
    
Exit_Loop:
    'close connection to recordset
    rst.Close
    Set rst = Nothing

End Sub

Private Sub CutVIN(Row As Integer)

    Dim VIN As String
    
    If Cells(Row, 7).Value Like "* VIN:?" Or Cells(Row, 7).Value Like "* Vin:?" Then
        VIN = Right(Cells(Row, 7), 1)
        Cells(Row, 7).Value = Left(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 6)
        Cells(Row, 27).Value = VIN
    Else
        If Cells(Row, 7).Value Like "*VIN:?" Or Cells(Row, 7).Value Like "*Vin:?" Then
            VIN = Right(Cells(Row, 7), 1)
            Cells(Row, 7).Value = Left(Cells(Row, 7).Value, Len(Cells(Row, 7)) - 5)
            Cells(Row, 27).Value = VIN
        End If
    End If

End Sub

Private Sub CutTrim(Row As Integer)

    'Cuts out the trim
    Dim trim As String
    
    If Cells(Row, 7).Value Like "Trim:*" Then
        trim = Right(Cells(Row, 7), Len(Cells(Row, 7)) - 5)
        Cells(Row, 7).Value = ""
        Cells(Row, 43).Value = trim
    End If

End Sub

Private Sub CutMfrLabel(Row As Integer)

    Dim MFR As String
    
    If Cells(Row, 7).Value Like "*Eng MFG:*" Then
        MFR = Mid(Cells(Row, 7), InStr(1, Cells(Row, 7), "Eng MFG:") + 8, Len(Cells(Row, 7)) - 8)
        Cells(Row, 7).Value = Replace(Cells(Row, 7), "Eng MFG:" & MFR, "")
        Cells(Row, 25).Value = MFR
    End If

End Sub

Private Sub FillPartNum(Row As Integer)
    
    Cells(Row, 1).Value = PartName    'global variable user entered in SourceForm
    Cells(Row, 2).Value = "FVKX"      'FVKX is the Store code for our Amazon store in MyFitment
    Cells(Row, 6).Value = PartTypeVar 'global variable user entered in SourceForm

End Sub

Private Sub Headers()

    'Insert top row
    Rows("1:1").Insert xlDown
    
    'Add ACES headers
    Range("A1:V1").Value = [{"part", "brand_code", "make", "model", "year", "partterminologyname", "notes", "qty", "mfrlabel", "position", "aspiration","bedlength","bedtype","block","bodynumdoors","bodytype","brakeabs","brakesystem","cc","cid","cylinderheadtype","cylinders"}]
    Range("W1:AK1").Value = [{"drivetype", "enginedesignation","enginemfr","engineversion","enginevin","frontbraketype","frontspringtype","fueldeliverysubtype","fueldeliverytype","fuelsystemcontroltype","fuelsystemdesign","fueltype","ignitionsystemtype", "liters","mfrbodycode"}]
    Range("AL1:AX1").Value = [{"rearbraketype", "rearspringtype","region","steeringsystem","steeringtype","submodel","transmissioncontroltype","transmissionmfr","transmissionmfrcode","transmissionnumspeeds", "transmissiontype", "valvesperengine", "wheelbase"}]

End Sub

