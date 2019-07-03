Option Explicit

Public Function CheckBlank() As Boolean

    'if the entire sheet is blank, return true, otherwise return false
    If WorksheetFunction.CountA(Cells) = 0 Then
        CheckBlank = True
    Else
        CheckBlank = False
    End If

End Function

Public Function RawFitments() As Boolean

    'default to true
    'if page doesn't contain raw fitments function will return false
    RawFitments = True
    
    'looks for field values that indicate whether fitments came from Sixbit or Metro
    'will probably add other sources in the future
    If Range("A1").Value = "" And Range("B1") = "" And Range("C1") = "" And IsNull(Range("D1")) = False And Range("E1").Value = "" And Range("F1").Value = "" Then
        'cells as above indicate fitments came from Metro
        If Formatted = False Then FitmentSource = "Metro"
    Else
        If Range("A1").Value = "Notes" And IsNull(Range("A2")) = False And _
            Range("B1").Value = "Make" And IsNull(Range("B2")) = False And _
            Range("C1").Value = "Model" And IsNull(Range("C2")) = False And _
            Range("D1").Value = "Year" And IsNull(Range("D2")) = False And _
            Range("E1").Value = "Trim" And IsNull(Range("E2")) = False And _
            Range("F1").Value = "Engine" And IsNull(Range("F2")) = False Then
            
            'cells as above indicate fitments came from Sixbit
            If Formatted = False Then FitmentSource = "Sixbit"
        Else
            'if cells don't match either configuration above, then the sheet doesn't contain fitments
            If Range("A1").Value = "Make" And Range("B1").Value = "Model" And Range("C1").Value = "Year" And Range("D1").Value = "Trim" And Range("E1").Value = "Engine" And Range("F1").Value = "Notes" Then
                If Formatted = False Then FitmentSource = "Amazon"
            Else
                If Formatted = False Then FitmentSource = ""
                
                RawFitments = False     'if none of the criteria above is met, sheet doesn't contain raw fitments
            End If
        End If
    End If

End Function

Public Function CheckACES() As Boolean

    'check if all the field headers indicate ACES fields
    If Range("A1").Value = "part" And Range("B1").Value = "brand_code" And Range("C1").Value = "make" And Range("D1").Value = "model" And _
    Range("E1").Value = "year" And Range("F1").Value = "partterminologyname" And Range("G1").Value = "notes" And Range("H1").Value = "qty" And _
    Range("I1").Value = "mfrlabel" And Range("J1").Value = "position" And Range("K1").Value = "aspiration" And Range("L1").Value = "bedlength" And _
    Range("M1").Value = "bedtype" And Range("N1").Value = "block" And Range("O1").Value = "bodynumdoors" And Range("P1").Value = "bodytype" And _
    Range("Q1").Value = "brakeabs" And Range("R1").Value = "brakesystem" And Range("S1").Value = "cc" And Range("T1").Value = "cid" And _
    Range("U1").Value = "cylinderheadtype" And Range("V1").Value = "cylinders" And Range("W1").Value = "drivetype" And Range("X1").Value = "enginedesignation" And _
    Range("Y1").Value = "enginemfr" And Range("Z1").Value = "engineversion" And Range("AA1").Value = "enginevin" And Range("AB1").Value = "frontbraketype" And _
    Range("AC1").Value = "frontspringtype" And Range("AD1").Value = "fueldeliverysubtype" And Range("AE1").Value = "fueldeliverytype" And Range("AF1").Value = "fuelsystemcontroltype" And _
    Range("AG1").Value = "fuelsystemdesign" And Range("AH1").Value = "fueltype" And Range("AI1").Value = "ignitionsystemtype" And Range("AJ1").Value = "liters" And _
    Range("AK1").Value = "mfrbodycode" And Range("AL1").Value = "rearbraketype" And Range("AM1").Value = "rearspringtype" And Range("AN1").Value = "region" And _
    Range("AO1").Value = "steeringsystem" And Range("AP1").Value = "steeringtype" And Range("AQ1").Value = "submodel" And Range("AR1").Value = "transmissioncontroltype" And _
    Range("AS1").Value = "transmissionmfr" And Range("AT1").Value = "transmissionmfrcode" And Range("AU1").Value = "transmissionnumspeeds" And Range("AV1").Value = "transmissiontype" And _
    Range("AW1").Value = "valvesperengine" And Range("AX1").Value = "wheelbase" Then
        
        CheckACES = True
    Else
        CheckACES = False
    End If

End Function

Public Function CheckOOS() As Boolean

    If Range("A1").Value = "Select this item for performing bulk action" And (Range("B1").Value Like "Respond to questions*" Or Range("B1").Value Like "Edit*") Then
        CheckOOS = True
    Else
        CheckOOS = False
    End If
End Function

Public Function CheckWeekInv() As Boolean

    If Range("A1").Text = "Product ID" And Range("B1").Text = "Description" And Range("C1").Text = "Location" And Range("D1").Text = "QoH" And Range("E1").Text = "Reorder Point" _
    And Range("F1").Text = "Quantity Sold" And Not Range("A2").Text = "" And Range("A3").Text = "" And Range("D2").Text = "" And Range("E2").Text = "" And Range("F2").Text = "" Then
        CheckWeekInv = True
    Else
        CheckWeekInv = False
    End If

End Function

Public Function CheckWeekInvEnd() As Boolean

    If Range("A1").Value = "Count" And Range("B1").Value = "Product ID" And Range("C1").Value = "Description" And Range("D1").Value = "Location" And Range("E1").Value = "QoH" And Range("F1").Value = "Quantity Sold" Then
        CheckWeekInvEnd = True
    Else
        CheckWeekInvEnd = False
    End If

End Function

Public Function CheckFinaleProducts() As Boolean

    If Range("A1").Value = "Product ID" And Range("B1").Value = "Description" And Range("C1").Value = "Category" And Range("D1").Value = "Notes" And Range("E1").Value = "Std reorder point" And _
    Range("F1").Value = "Std reorder in qty of" And Range("G1").Value = "Manufacturer" And Range("H1").Value = "Mfg product ID" And Range("I1").Value = "Supplier 1" And _
    Range("J1").Value = "Supplier 1 price" And Range("K1").Value = "Supplier 1 product ID" And Range("L1").Value = "Supplier 1 comments" And Range("M1").Value = "Location" And _
    Range("N1").Value = "Interchange 1" And Range("O1").Value = "Interchange 2" And Range("P1").Value = "Interchange 3" And Range("Q1").Value = "Interchange 4" And _
    Range("R1").Value = "Live Status" And Range("S1").Value = "Sorting" Then
        
        CheckFinaleProducts = True
    Else
        CheckFinaleProducts = False
    End If

End Function

Public Function PartNumMatch(PRange As Range) As Boolean

    If PRange.Value = PartName Then
        PartNumMatch = True
    Else
        PartNumMatch = False
    End If

End Function

Public Function CheckManageInv() As Boolean

    'default to false
    CheckManageInv = False
    
    Dim c As Integer
    Dim R As Integer
    Dim p As Range
    
    'find the top-left-most cell
    For c = 1 To 14
        'loop through columns
        If Application.CountA(columns(c)) > 0 Then
            'loop through rows
            For R = 1 To 5
                If Not Cells(R, c).Text = "" Then
                    'save cell to variable
                    Set p = Cells(R, c)
                    GoTo Exit_Loop

                End If
            Next R
        End If
    Next c
    
Exit_Loop:
    
    'if A1 = "Status" and A2 = "Active" or "Inactive" or "Variations" then the sheet has already been formatted, leave as false
    If p.Value = "Status " And p.Offset(0, 1).Value = "Image" And p.Offset(0, 2).Value = "SKU " Then
        'if top-left-most cell is Status and the other headers are there and it's not been formatted, then set function to true
        If Not Range("A1").Value = "Status " And Not (Range("A2").Value = "Active " Or Range("A2").Value Like "Inactive*" Or Range("A2").Value Like "Variations*" Or Range("A2").Value = "Suppressed ") Then
            CheckManageInv = True
        End If
    End If
    
    'clear variables
    c = 0
    R = 0
    Set p = Nothing

End Function

Public Function CheckAllFinaleProducts() As Boolean

    CheckAllFinaleProducts = False
    
    'if fields for all products report from Finale are found AND column B has "Inactive" values then function returns true
    If Range("A1").Value = "Product ID" And Range("B1").Value = "Status" And Range("C1").Value = "Description" And Range("D1").Value = "Category" And Range("E1").Value = "Notes" And _
    Range("F1").Value = "Std accounting cost" And Range("G1").Value = "Std buy price" And Range("H1").Value = "Item price" And Range("I1").Value = "Case price" And _
    Range("J1").Value = "Std lead days" And Range("K1").Value = "Std packing" And Range("L1").Value = "Std packing units per case" And Range("M1").Value = "Unit of measure" And _
    Not Range("B:B").Find("Inactive") Is Nothing Then CheckAllFinaleProducts = True

End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
    On Error Resume Next
    'checks if array is allocated
    IsArrayAllocated = IsArray(Arr) And Not IsError(LBound(Arr, 1)) And LBound(Arr, 1) <= UBound(Arr, 1)

End Function

Public Function CheckBoM() As Boolean

    CheckBoM = False
    
    'check fields and empty cells to determine if BoM report
    If Range("A1").Text = "Product ID" And Range("B1").Text = "Description" And Range("C1").Text = "Quantity" And Range("D1").Text = "Product ID" And Range("E1").Text = "Description" And _
    Range("F1").Text = "Component note" And Range("A2").Text <> "" And Range("A3").Text = "" Then CheckBoM = True

End Function

Public Function CheckAmazonTemplate() As Boolean

    CheckAmazonTemplate = False
    
    If Range("A1").Value Like "TemplateType=*" And _
    Range("B1").Value Like "Version=*" And _
    Range("C1").Value Like "TemplateSignature=*" And _
    Range("D1").Value = "The top 3 rows are for Amazon.com use only. Do not modify or delete the top 3 rows." Then CheckAmazonTemplate = True

End Function

Public Function CheckUPC() As Boolean

    'check to see if every column except column A is empty
    If CountRows("A:A") > 0 And CountRows("B:D") = 0 Then
        'If column A is only column with values
        CheckUPC = True
    Else
        'if any column besides column A has values, make false
        CheckUPC = False
        Exit Function
    End If
    
    Dim i As Integer
    'loop through every row in column A
    For i = 1 To CountRows("A:A")
        If IsNumeric(Cells(i, 1)) = False Or Len(Cells(i, 1)) <> 12 Then
            'if any row in column A is not a number, make false
            CheckUPC = False
        End If
    Next i

End Function

Public Function CheckDropship() As String

    If Range("A1").Value = "Item" And Range("B1").Value = "Desc" And Range("C1").Value = "customer" And Range("D1").Value = "so" And Range("E1").Value = "Qty Sold" And _
    Range("F1").Value = "Unit Price" And Range("G1").Value = "Total Amount" And Range("H1").Value = "TaxCode" And Range("I1").Value = "" Then
        CheckDropship = "Herko"
    ElseIf Range("A1").Value = "Date - Shipped Date" And Range("B1").Value = "Customer Email" And Range("C1").Value = "Ship To - Name" And Range("D1").Value = "Amount - Order Total" And _
    Range("E1").Value = "Amount - Shipping Cost" And Range("F1").Value = "" Then
        CheckDropship = "Shipstation"
    Else
        CheckDropship = ""
    End If

End Function

Public Function CheckHerkoDropship() As Boolean
    
    If Range("A1").Value = "Date" And Range("B1").Value = "Client Name" And Range("C1").Value = "qb#" And Range("D1").Value = "Qty" And Range("E1").Value = "Unit Price" And Range("F1").Value = "Total" And Range("G1").Value = "SO #" And Range("H1").Value = "SHIPPING" And Range("I1").Value = "" Then
        CheckHerkoDropship = True
    Else
        CheckHerkoDropship = False
    End If

End Function

Public Function CheckShipstationDropship()

    If Range("A1").Value = "Amount - Order Total" And Range("B1").Value = "Ship To - Company" And Range("C1").Value = "Ship To - Name" And Range("D1").Value = "Ship To - Country" And Range("E1").Value = "Bill To - Name" And _
    Range("F1").Value = "Service - Package Type" And Range("G1").Value = "Service - Confirmation Type" And Range("H1").Value = "Dimensions - Height" And Range("I1").Value = "Tags" And Range("J1").Value = "Insurance - Cost" And _
    Range("K1").Value = "Ship To - Zone" And Range("L1").Value = "Count - Number of Items" Then
        CheckShipstationDropship = True
    Else
        CheckShipstationDropship = False
    End If

End Function

Public Function CheckFormattedHerko() As Boolean

    If Range("A1").Value = "Date" And Range("B1").Value = "Client Name" And Range("C1").Value = "qb#" Then
        CheckFormattedHerko = True
    End If

End Function

Public Function CheckFormattedDropship() As Boolean

    If Range("A1").Value = "Item" And Range("B1").Value = "Desc" And Range("C1").Value = "customer" And Range("D1").Value = "so" And Range("E1").Value = "Qty Sold" And _
    Range("F1").Value = "Unit Price" And Range("G1").Value = "Total Amount" And Range("H1").Value = "Shipping Cost" And Range("I1").Value = "AD Total Price" And _
    Range("K1").Value = "Profit/Loss" And Range("L1").Value = "" Then
        CheckFormattedDropship = True
    End If

End Function

Public Function NameValid(FileName As String) As Boolean

    'assume name is valid
    NameValid = True
    
    'check file name length
    If Len(FileName) > 31 Then
        NameValid = False
        Exit Function
    End If
    
    Dim InvalidChars As Variant
    
    'Hard code the invalid characters for now
    InvalidChars = Array("<", ">", "?", "[", "]", ":", "|", "*")
    
    'loop through array to see if any invalid character is used in the sheet title
    Dim i As Integer
    For i = LBound(InvalidChars) To UBound(InvalidChars)
        If InStr(1, FileName, InvalidChars(i)) > 0 Then
            NameValid = False
            Exit Function
        End If
    Next i

End Function

Public Function CheckVolumePricing()

    CheckVolumePricing = False
    
    If Range("A1").Value = "SKU" And Range("B1").Value = "Offset Type(Amount or Percentage)" And Range("C1").Value = "T1 Min. Qty" And Range("D1").Value = "T1 Max. Qty" And _
    Range("E1").Value = "T1 Offset Value" And Range("F1").Value = "T2 Min. Qty" And Range("G1").Value = "T2 Max. Qty" And Range("H1").Value = "T2 Offset Value" And _
    Range("I1").Value = "T3 Min. Qty" And Range("J1").Value = "T3 Max. Qty" And Range("K1").Value = "T3 Offset Value" And Range("L1").Value = "T4 Min. Qty" And _
    Range("M1").Value = "T4 Max. Qty" And Range("N1").Value = "T4 Offset Value" Then CheckVolumePricing = True

End Function

