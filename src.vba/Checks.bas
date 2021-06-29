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
    If range("A1").value = "" And range("B1") = "" And range("C1") = "" And IsNull(range("D1")) = False And range("E1").value = "" And range("F1").value = "" Then
        'cells as above indicate fitments came from Metro
        If Formatted = False Then FitmentSource = "Metro"
    Else
        If (range("A1").value = "Notes" And IsNull(range("A2")) = False And _
            range("B1").value = "Make" And IsNull(range("B2")) = False And _
            range("C1").value = "Model" And IsNull(range("C2")) = False And _
            range("D1").value = "Year" And IsNull(range("D2")) = False And _
            range("E1").value = "Trim" And IsNull(range("E2")) = False And _
            range("F1").value = "Engine" And IsNull(range("F2")) = False) Or _
            (range("A1").value = "Engine" And range("B1").value = "Make" And _
            range("C1").value = "Model" And range("G1").value = "Part Type" And _
            range("H1").value = "Quantity Required" And range("I1").value = "Position" And _
            range("J1").value = "MFRLabel") Then
            
            'cells as above indicate fitments came from Sixbit
            If Formatted = False Then FitmentSource = "Sixbit"
        Else
            'if cells don't match either configuration above, then the sheet doesn't contain fitments
            If range("A1").value = "Make" And range("B1").value = "Model" And range("C1").value = "Year" And range("D1").value = "Trim" And range("E1").value = "Engine" And range("F1").value = "Notes" Then
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
    If range("A1").value = "part" And range("B1").value = "brand_code" And range("C1").value = "make" And range("D1").value = "model" And _
    range("E1").value = "year" And range("F1").value = "partterminologyname" And range("G1").value = "notes" And range("H1").value = "qty" And _
    range("I1").value = "mfrlabel" And range("J1").value = "position" And range("K1").value = "aspiration" And range("L1").value = "bedlength" And _
    range("M1").value = "bedtype" And range("N1").value = "block" And range("O1").value = "bodynumdoors" And range("P1").value = "bodytype" And _
    range("Q1").value = "brakeabs" And range("R1").value = "brakesystem" And range("S1").value = "cc" And range("T1").value = "cid" And _
    range("U1").value = "cylinderheadtype" And range("V1").value = "cylinders" And range("W1").value = "drivetype" And range("X1").value = "enginedesignation" And _
    range("Y1").value = "enginemfr" And range("Z1").value = "engineversion" And range("AA1").value = "enginevin" And range("AB1").value = "frontbraketype" And _
    range("AC1").value = "frontspringtype" And range("AD1").value = "fueldeliverysubtype" And range("AE1").value = "fueldeliverytype" And range("AF1").value = "fuelsystemcontroltype" And _
    range("AG1").value = "fuelsystemdesign" And range("AH1").value = "fueltype" And range("AI1").value = "ignitionsystemtype" And range("AJ1").value = "liters" And _
    range("AK1").value = "mfrbodycode" And range("AL1").value = "rearbraketype" And range("AM1").value = "rearspringtype" And range("AN1").value = "region" And _
    range("AO1").value = "steeringsystem" And range("AP1").value = "steeringtype" And range("AQ1").value = "submodel" And range("AR1").value = "transmissioncontroltype" And _
    range("AS1").value = "transmissionmfr" And range("AT1").value = "transmissionmfrcode" And range("AU1").value = "transmissionnumspeeds" And range("AV1").value = "transmissiontype" And _
    range("AW1").value = "valvesperengine" And range("AX1").value = "wheelbase" Then
        
        CheckACES = True
    Else
        CheckACES = False
    End If

End Function

Public Function CheckOOS() As Boolean

    If range("A1").value Like "Select this item for performing bulk action*" And (range("B1").value Like "EditLink. Edit.*" Or range("B1").value Like "Respond to questions*") Then
        CheckOOS = True
    Else
        CheckOOS = False
    End If
End Function

Public Function CheckWeekInv() As Boolean

    If range("A1").value = "Product ID" And range("B1").value = "Description" And range("C1").value = "Location" And range("D1").value = "QoH" And range("E1").value = "Reorder Point" _
    And range("F1").value = "Quantity Sold" And Not range("A2").value = "" And range("A3").value = "" And range("D2").value = "" And range("E2").value = "" And range("F2").value = "" Then
        CheckWeekInv = True
    Else
        CheckWeekInv = False
    End If

End Function

Public Function CheckWeekInvEnd() As Boolean

    If range("A1").value = "Count" And range("B1").value = "Product ID" And range("C1").value = "Description" And range("D1").value = "Location" And range("E1").value = "QoH" And range("F1").value = "Quantity Sold" Then
        CheckWeekInvEnd = True
    Else
        CheckWeekInvEnd = False
    End If

End Function

Public Function CheckFinaleProducts() As Boolean

    If range("A1").value = "Product ID" And range("B1").value = "Description" And range("C1").value = "Category" And range("D1").value = "Notes" And range("E1").value = "Std reorder point" And _
    range("F1").value = "Std reorder in qty of" And range("G1").value = "Manufacturer" And range("H1").value = "Mfg product ID" And range("I1").value = "Supplier 1" And _
    range("J1").value = "Supplier 1 price" And range("K1").value = "Supplier 1 product ID" And range("L1").value = "Supplier 1 comments" And range("M1").value = "Location" And _
    range("N1").value = "Interchange 1" And range("O1").value = "Interchange 2" And range("P1").value = "Interchange 3" And range("Q1").value = "Interchange 4" And _
    range("R1").value = "Listing Status" And range("S1").value = "Sorting" Then
        
        CheckFinaleProducts = True
    Else
        CheckFinaleProducts = False
    End If

End Function

Public Function PartNumMatch(PRange As range) As Boolean

    If PRange.value = PartName Then
        PartNumMatch = True
    Else
        PartNumMatch = False
    End If

End Function

Public Function CheckManageInv() As Boolean

    'default to false
    CheckManageInv = False
    
    Dim c As Integer
    Dim r As Integer
    Dim p As range
    
    'find the top-left-most cell
    For c = 1 To 14
        'loop through columns
        If Application.CountA(columns(c)) > 0 Then
            'loop through rows
            For r = 1 To 5
                If Not Cells(r, c).value = "" Then
                    'save cell to variable
                    Set p = Cells(r, c)
                    GoTo Exit_Loop

                End If
            Next r
        End If
    Next c
    
Exit_Loop:
    
    'if A1 = "Status" and A2 = "Active" or "Inactive" or "Variations" then the sheet has already been formatted, leave as false
    If Not p Is Nothing Then
        If p.value = "Status " And p.Offset(0, 1).value = "Image" And p.Offset(0, 2).value = "SKU " Then
            'if top-left-most cell is Status and the other headers are there and it's not been formatted, then set function to true
            If Not range("A1").value = "Status " And Not (range("A2").value = "Active " Or range("A2").value Like "Inactive*" Or range("A2").value Like "Variations*" Or range("A2").value = "Suppressed ") Then
                CheckManageInv = True
            End If
        End If
    End If
    
    'clear variables
    c = 0
    r = 0
    Set p = Nothing

End Function

Public Function CheckAllFinaleProducts() As Boolean

    CheckAllFinaleProducts = False
    
    'if fields for all products report from Finale are found AND column B has "Inactive" values then function returns true
    If range("A1").value = "Product ID" And range("B1").value = "Status" And range("C1").value = "Description" And range("D1").value = "Category" And range("E1").value = "Notes" And _
    range("F1").value = "Std accounting cost" And range("G1").value = "Std buy price" And range("H1").value = "Item price" And range("I1").value = "Case price" And _
    range("J1").value = "Std lead days" And range("K1").value = "Std packing" And range("L1").value = "Std packing units per case" And range("M1").value = "Unit of measure" And _
    Not range("B:B").Find("Inactive") Is Nothing Then CheckAllFinaleProducts = True

End Function

Public Function IsArrayAllocated(arr As Variant) As Boolean
    On Error Resume Next
    'checks if array is allocated
    IsArrayAllocated = IsArray(arr) And Not IsError(LBound(arr, 1)) And LBound(arr, 1) <= UBound(arr, 1)

End Function

Public Function CheckBoM() As String
    
    'check fields and empty cells to determine if BoM report
    If range("A1").value = "Product ID" And range("B1").value = "Description" And range("C1").value = "Quantity" And range("D1").value = "Product ID" And range("E1").value = "Description" And _
    range("F1").value = "Component note" And range("A2").value <> "" And range("A3").value = "" Then
        CheckBoM = "Raw"
    Else
        If (range("A1").value = "Product ID" And range("B1").value = "Quantity" And range("C1").value = "Component Product ID") Or _
        (range("A1").value = "Product ID" And range("B1").value = "Quantity" And range("C1").value = "Component Product ID" And range("D1").value = "Component note") Then
            CheckBoM = "Compact"
        Else
            If range("A1").value = "Product ID" And range("B1").value = "BoM 1" And range("C1").value = "Qty 1" Then
                CheckBoM = "Expanded"
            Else
                CheckBoM = ""
            End If
        End If
    End If

End Function

Public Function CheckAmazonTemplate() As Boolean

    CheckAmazonTemplate = False
    
    If range("A1").value Like "TemplateType=*" And _
    range("B1").value Like "Version=*" And _
    range("C1").value Like "TemplateSignature=*" And _
    range("D1").value = "The top 3 rows are for Amazon.com use only. Do not modify or delete the top 3 rows." Then CheckAmazonTemplate = True

End Function

Public Function CheckUPC() As Boolean

    If range("A1").value = "Action" And range("B1").value = "GS1CompanyPrefix" And range("C1").value = "GTIN" And range("D1").value = "PackagingLevel" And _
    range("E1").value = "Description" And range("F1").value = "SKU" And range("G1").value = "BrandName" And range("H1").value = "Status" And _
    range("I1").value = "IsVariable" And range("J1").value = "IsPurchasable" And range("K1").value = "Certified" And range("L1").value = "Height" And _
    range("M1").value = "Width" And range("N1").value = "Depth" And range("O1").value = "DimensionMeasure" And range("P1").value = "GrossWeight" And _
    range("Q1").value = "NetWeight" And range("R1").value = "WeightMeasure" And range("S1").value = "Comments" And range("T1").value = "CountryOfOrigin" And _
    range("U1").value = "ChildGTINs" And range("V1").value = "Quantity" And range("W1").value = "SubBrandName" And range("X1").value = "ProductDescriptionShort" And _
    range("Y1").value = "LabelDescription" And range("Z1").value = "NetContent1Count" And range("AA1").value = "NetContent1UnitOfMeasure" And range("AB1").value = "NetContent2Count" And _
    range("AC1").value = "NetContent2UnitOfMeasure" And range("AD1").value = "NetContent3Count" And range("AE1").value = "NetContent3UnitOfMeasure" And _
    range("AF1").value = "GlobalProductClassification" And range("AG1").value = "ImageURL" And range("AH1").value = "TargetMarket" Then
        CheckUPC = True
    End If

End Function

Public Function CheckDropship() As String

    If range("A1").value = "Item" And range("B1").value = "Desc" And range("C1").value = "customer" And range("D1").value = "so" And range("E1").value = "Qty Sold" And _
    range("F1").value = "Unit Price" And range("G1").value = "Total Amount" And (range("H1").value = "TaxCode" Or range("I1").value = "") Then
        CheckDropship = "Herko"
    ElseIf range("A1").value = "Date - Shipped Date" And range("B1").value = "Customer Email" And range("C1").value = "Ship To - Name" And range("D1").value = "Amount - Order Subtotal" And _
    range("E1").value = "Amount - Shipping Cost" And range("F1").value = "" Then
        CheckDropship = "Shipstation"
    Else
        CheckDropship = ""
    End If

End Function

Public Function CheckHerkoDropship() As Boolean
    
    If range("A1").value = "Date" And range("B1").value = "Client Name" And range("C1").value = "qb#" And range("D1").value = "Qty" And range("E1").value = "Unit Price" And range("F1").value = "Total" And range("G1").value = "SO #" And range("H1").value = "SHIPPING" And range("I1").value = "" Then
        CheckHerkoDropship = True
    Else
        CheckHerkoDropship = False
    End If

End Function

Public Function CheckShipstationDropship()

    If range("A1").value = "Amount - Order Total" And range("B1").value = "Ship To - Company" And range("C1").value = "Ship To - Name" And range("D1").value = "Ship To - Country" And range("E1").value = "Bill To - Name" And _
    range("F1").value = "Service - Package Type" And range("G1").value = "Service - Confirmation Type" And range("H1").value = "Dimensions - Height" And range("I1").value = "Tags" And range("J1").value = "Insurance - Cost" And _
    range("K1").value = "Ship To - Zone" And range("L1").value = "Count - Number of Items" Then
        CheckShipstationDropship = True
    Else
        CheckShipstationDropship = False
    End If

End Function

Public Function CheckFormattedHerko() As Boolean

    If range("A1").value = "Date" And range("B1").value = "Client Name" And range("C1").value = "qb#" Then
        CheckFormattedHerko = True
    End If

End Function

Public Function CheckFormattedDropship() As Boolean

    If range("A1").value = "Item" And range("B1").value = "Desc" And range("C1").value = "customer" And range("D1").value = "so" And range("E1").value = "Qty Sold" And _
    range("F1").value = "Unit Price" And range("G1").value = "Total Amount" And range("H1").value = "Shipping Cost" And range("I1").value = "AD Total Price" And _
    range("K1").value = "Profit/Loss" And range("L1").value = "" Then
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
    
    If range("A1").value = "SKU" And range("B1").value = "Offset Type(Amount or Percentage)" And range("C1").value = "T1 Min. Qty" And range("D1").value = "T1 Max. Qty" And _
    range("E1").value = "T1 Offset Value" And range("F1").value = "T2 Min. Qty" And range("G1").value = "T2 Max. Qty" And range("H1").value = "T2 Offset Value" And _
    range("I1").value = "T3 Min. Qty" And range("J1").value = "T3 Max. Qty" And range("K1").value = "T3 Offset Value" And range("L1").value = "T4 Min. Qty" And _
    range("M1").value = "T4 Max. Qty" And range("N1").value = "T4 Offset Value" Then
        
        CheckVolumePricing = True
        
    End If

End Function

Public Function CheckShippingMethods()

    If range("A1").value = "Product ID" And range("B1").value = "Ebay Shipping Method" And range("C1").value = "Ebay Shipping Cost" And _
    range("D1").value = "Amazon Shipping Method" And range("E1").value = "Amazon Shipping Cost" Then
        
        CheckShippingMethods = True
        
    End If

End Function
