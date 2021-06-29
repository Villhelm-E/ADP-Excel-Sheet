Option Explicit

Public Sub FormattedToSixbit()

    On Error GoTo Err_ToSixbit
    
    'If NewColumn column exists, delete it
    'This column is added when combining tables in PowerQuery
    CheckNewColumn
    
    'replace all spaces with non-breaking spaces
    ReplaceNonBreakingSpace     'sixbit module
    
    'Set Category
    Dim Category As String
    Category = SixbitCategory(range("F2").value)
    
    'Grab Part Type from cell F2
    PartTypeVar = range("F2").value
    
    'Set Ordinal ID
    Dim OrdinalID As String
    OrdinalID = MaxOrdinalID(OrdinalID)
    
    'Show SKUForm to generate SKU
    SKUForm.Show
    
    If SKU <> "" Then
        'Create new record in Compatibilities table in Sixbit database
        Call NewCompatibilityRecord(OrdinalID, Category, SKU)
        
        'Converts fitments in sheet to XML format and exports to Sixbit database
        Call XMLParser(SKU)
    End If
    
    Exit Sub
    
Err_ToSixbit:
    MsgBox ("There was a problem adding fitments to Sixbit.")

End Sub

Private Sub CheckNewColumn()

    If range("AY1").value = "NewColumn" And range("AY2").value = "[Table]" Then
        range("AY:AY").EntireColumn.Delete
    End If

End Sub

Public Function SixbitCategory(PartTyp As String) As String
    
    'Open Part Types table
    OpenACESPartTypes
    
    With rst
        While Not .EOF
            If .fields("ACESPartType").value = PartTyp Then
                SixbitCategory = .fields("EbayCategoryID").value  'set ebay category id
                rst.Close
                Exit Function
            End If
            rst.MoveNext
        Wend
    End With
    
    rst.Close

End Function

Private Function MaxOrdinalID(OrdinalID As String) As String
    
    'Open the Compatibilities table in Sixbit database sorting by descending Ordinal ID to find max ordinal id
    OpenOrdinalID
        
    'Add 1 to max value
    MaxOrdinalID = rst.fields("OrdinalID").value + 1

End Function

Private Sub NewCompatibilityRecord(OrdinalID As String, Category As String, SKU As String)

    On Error GoTo Create_Error
    
    rst.Close
    
    rst.Open "INSERT INTO dbo.CompatibilitySets ([Name], CategoryID, OrdinalID) VALUES (" & Chr(39) & SKU & Chr(39) & ", " & val(Category) & ", " & val(OrdinalID) & ")", _
        SxbtDb, adOpenDynamic, adLockOptimistic
            
    MsgBox ("Added Compatibility")
    
    Exit Sub
    
Create_Error:
    MsgBox Err.Description

End Sub

Private Sub XMLParser(SKU As String)

    Dim make As String
    Dim model As String
    Dim year As String
    Dim cc As String
    Dim cid As String
    Dim block As String
    Dim cylinders As String
    Dim enginevin As String
    Dim fuel As String
    Dim liters As String
    Dim cylheadtype As String
    Dim aspiration As String
    Dim trim As String
    Dim prenote As String
    Dim bodynumdoors As String
    Dim bodytype As String
    
    Dim numrows As Integer
    numrows = CountRows("A:A")
    
    Dim i As Integer
    
    For i = 2 To numrows
        'parse liters
        Call loopliters(liters, i)
        
        'parse Cubic Centimeters
        Call loopCC(cc, i)
        
        'parse Cu In.
        Call loopcid(cid, i)
        
        'parse Block
        Call loopBlock(block, i)
        
        'parse Cylinders
        Call loopCyl(cylinders, i)
        
        'parse fuel
        Call loopfuel(fuel, i)
        
        'parse cylinder head type
        Call loopcylhedtyp(cylheadtype, i)
        
        'parse aspiratiojn
        Call loopaspiration(aspiration, i)
        
        'parse vin
        Call loopvin(enginevin, i)
        
        'parse number of doors
        Call loopbodynumdoors(bodynumdoors, i)
        
        'parse body type
        Call loopbodytype(bodytype, i)
        
        'parse trim
        Call looptrim(trim, bodytype, bodynumdoors, i)
        
        'parse notes
        Call loopnotes(prenote, i)
        
        'parse compatiblity
        Call ParseCompatibility(liters, cc, cid, block, cylinders, fuel, cylheadtype, aspiration, make, model, year, trim, prenote, enginevin, SKU, i)
    Next i
    
    Call EndTags(SKU)
    
    MsgBox ("Fitments added successfully")

End Sub

Private Sub loopliters(liters, i As Integer)

    'liters is column 36
    If IsNull(Cells(i, 36).value) = True Then
        liters = ""
    Else
        liters = Cells(i, 36).value & "L"
    End If

End Sub

Private Sub loopCC(cc, i As Integer)

    'cubic centimeters is column 19
    If IsNull(Cells(i, 19).value) = True Then
        cc = ""
    Else
        cc = " " & Cells(i, 19).value & "CC"
    End If

End Sub

Private Sub loopcid(cid, i As Integer)

    'cubic inches is column 20
        If IsNull(Cells(i, 20).value) = True Then
            cid = ""
        Else
            cid = " " & Cells(i, 20).value & "Cu. In."
        End If
        
End Sub

Private Sub loopBlock(block, i As Integer)

        'block is column 14
        If IsNull(Cells(i, 14).value) = True Then
            block = ""
        Else
            'need to lowercase the L for inline block
            If Cells(i, 14).value = "L" Then
                block = " l"
            Else
                block = " " & Cells(i, 14).value
            End If
        End If
        
End Sub

Private Sub loopCyl(cylinders, i As Integer)

        'cylinders is column 22
        If IsNull(Cells(i, 22).value) = True Then
            cylinders = ""
        Else
            cylinders = Cells(i, 22).value
        End If
        
End Sub
        
Private Sub loopfuel(fuel, i As Integer)

        'fuel type is column 34
        If IsNull(Cells(i, 34).value) = True Then
            fuel = ""
        Else
            fuel = " " & Cells(i, 34).value
        End If
        
End Sub

Private Sub loopcylhedtyp(cylhedtyp, i As Integer)

        'cylinder head type is column 21
        If IsNull(Cells(i, 21).value) = True Then
            cylhedtyp = ""
        Else
            cylhedtyp = " " & Cells(i, 21).value
        End If
        
End Sub

Private Sub loopaspiration(aspiration, i As Integer)

        'aspiration is column 11
        If IsNull(Cells(i, 11).value) = True Then
            aspiration = ""
        Else
            aspiration = " " & Cells(i, 11).value
        End If
        
End Sub

Private Sub loopvin(enginevin, i As Integer)

        'engine vin is column 27
        If IsNull(Cells(i, 27).value) = True Then
            enginevin = ""
        Else
            enginevin = "VIN: " & Cells(i, 27).value
        End If
        
End Sub

Private Sub loopbodynumdoors(bodynumdoors, i As Integer)
        
        'bodynumdoors is column 15
        If IsNull(Cells(i, 15).value) = True Then
            bodynumdoors = ""
        Else
            bodynumdoors = Cells(i, 15).value & "-Door"
        End If
        
End Sub

Private Sub loopbodytype(bodytype, i As Integer)
        
        'body type is column 16
        If IsNull(Cells(i, 16).value) = True Then
            bodytype = ""
        Else
            bodytype = Cells(i, 16).value & " "
        End If
        
End Sub

Private Sub looptrim(trim, bodytype, bodynumdoors, i As Integer)
        
        'trim is column 43
        If IsNull(Cells(i, 43).value) = True Then
            trim = ""
        Else
            trim = Cells(i, 43).value & " " & bodytype & bodynumdoors
        End If
        
End Sub

Private Sub loopnotes(prenote, i As Integer)
        
        'notes is column 7
        If IsNull(Cells(i, 7).value) = True Then
            prenote = ""
        Else
            prenote = Cells(i, 7).value & " "
        End If
        
End Sub

Private Sub ParseCompatibility(liters, cc, cid, block, cylinders, fuel, cylheadtype, aspiration, make, model, year, trim, prenote, enginevin, SKU As String, i As Integer)

    Dim notes As String
    Dim fitment As String
    Dim engine As String
    
    'combines some of the fields with formatting
    'fields are grouped into <NameValue></NameValue>
    'Name Value is split into <Name>Field</Name> and <Value></Value>
    engine = "<NameValue><Name>Engine</Name><Value>" & liters & cc & cid & block & cylinders & fuel & cylheadtype & aspiration & "</Value></NameValue>"
    make = "<NameValue><Name>Make</Name><Value>" & Cells(i, 3).value & "</Value></NameValue>"
    model = "<NameValue><Name>Model</Name><Value>" & Cells(i, 4).value & "</Value></NameValue>"
    year = "<NameValue><Name>Year</Name><Value>" & Cells(i, 5).value & "</Value></NameValue>"
    
    'If fitment doesn't specify submodel, enter ALL, otherwise put the submodel in the trim xml tag
    If IsNull(Cells(i, 43).value) = True Then
        trim = "<NameValue><Name>Trim</Name><Value>All</Value></NameValue>"
    Else
        trim = "<NameValue><Name>Trim</Name><Value>" & trim & "</Value></NameValue>"
    End If
    
    'Notes doesn't use <NameValue><Name></Name><Value></Value></NameValue>, just use <Notes></Notes>
    'add the part type from column 6
    notes = "<Notes>" & prenote & enginevin & " PartType " & Cells(i, 6).value & "</Notes>"
    
    'each compatibility is enclosed by <Compatibility></Compatibility>
    fitment = "<Compatibility>" & engine & make & model & trim & year & notes & "</Compatibility>"
    
    On Error GoTo fitment_error
    
    rst.Open "UPDATE dbo.CompatibilitySets SET CompatibilitySetDefinition= CompatibilitySetDefinition" & " & " & Chr(34) & engine & Chr(34) & " WHERE [Name]=" & Chr(39) & SKU & Chr(39) & ";", _
        SxbtDb, adOpenDynamic, adLockOptimistic
        
    rst.Open "UPDATE dbo.CompatibilitySets SET CompatibilitySetDefinition= CompatibilitySetDefinition" & " & " & Chr(34) & engine & Chr(34) & " WHERE [Name]=" & Chr(39) & SKU & Chr(39) & ";", _
        SxbtDb, adOpenDynamic, adLockOptimistic
    
    'Update Query to dbo_CompatibilitySets
'    Call AppendCompatibility(fitment, SKU)

fitment_error:
    MsgBox Err.Description

End Sub

Private Sub AppendCompatibility(fitment As String, SKU As String)

    On Error GoTo fitment_error
    
    rst.Open "UPDATE dbo.CompatibilitySets SET CompatibilitySetDefinition= CompatibilitySetDefinition" & " & " & Chr(34) & fitment & Chr(34) & " WHERE [Name]=" & Chr(39) & SKU & Chr(39) & ";", _
        SxbtDb, adOpenDynamic, adLockOptimistic
    
    Exit Sub
    
fitment_error:
    MsgBox Err.Description
    MsgBox ("There was an error appending a compatibility set into the table.")

End Sub

Private Sub EndTags(SKU As String)

    On Error GoTo end_tag_error
    
    rst.Open "UPDATE dbo.CompatibilitySets SET CompatibilitySetDefinition= " & Chr(34) & "<Compatibilities>" & Chr(34) & " & CompatibilitySetDefinition & " & Chr(34) & "</Compatibilities>" & _
        Chr(34) & " WHERE [Name]=" & Chr(34) & SKU & Chr(34) & ";", SxbtDb, adOpenDynamic, adLockOptimistic
    
    Exit Sub
    
end_tag_error:
    MsgBox Err.Description
    MsgBox ("There was an error appending the end tags.")

End Sub
