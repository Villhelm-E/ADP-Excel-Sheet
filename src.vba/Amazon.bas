Option Explicit

Public Sub DefineAmazonVariables()

    'They are found on the first row, just copy paste
    'Only version and signature should change
    TemplateType = "TemplateType=fptcustom"
    TemplateAmazonUse = "The top 3 rows are for Amazon.com use only. Do not modify or delete the top 3 rows."   'i modified the top 3 rows. fuck da police
    
    'open AmazonTemplateVariables table
    Set rst = MstrDb.Execute("SELECT * FROM AmazonTemplateVariables")
    
    'every update Amazon makes to the template changes the Version and Signature
    'update these in the Variables button under About in the ADP Tab
    TemplateVersion = rst.fields("AmazonTemplateVersion").value
    TemplateSignature = rst.fields("AmazonTemplateSig").value
    
    'Amazon Template switched the name and label rows from Version 2018.0210 to 2018.0820
    LabelRow = rst.fields("LabelRow").value    'global variable that determines which row on the Amazon template is for the Field Labels
    NameRow = rst.fields("NameRow").value     'global variable that determines which row on the Amazon template is for the Field Names/Codes
    
    rst.Close

End Sub

Public Sub AmazonMain()
    
    Dim wsSheet As Worksheet

    On Error Resume Next
    
    'Amazon Listings sheet name
    Dim SheetName As String
    SheetName = "Amazon Template"
    
    Set wsSheet = Sheets(SheetName)
    On Error GoTo 0
    
    'Determine what to do based on whether "Upload Template" exists or not
    If Not wsSheet Is Nothing Then
        'if sheet exists
        If CheckAmazonTemplate = True Then
            'if current sheet is "Upload Template"
            ListAmazon.Show
        Else
            'if current sheet is not "Upload Template"
            Worksheets(SheetName).Activate
        End If
    Else
        'if "Upload Template" doesn't exist
        Call PrepWorksheet(SheetName)
        
        'Add headers to create Amazon Template worksheet
        AmazonHeaders
    End If
    
    'refresh ribbon
    RibbonCategories
    
    range("A4").Select

End Sub

Public Sub AmazonHeaders()
    
    'Define the global variables that that determine which version of the Amazon template is in use
    DefineAmazonVariables
    
    'do the top row of the header
    FillAmazonVariables
    
    'begin organizing template fields
    Dim rst As Recordset
    Set rst = MstrDb.Execute("SELECT * FROM AmazonTemplateFields WHERE [TemplateOrder] > 0 ORDER BY [TemplateOrder]")
    rst.MoveFirst
    
    Dim yvgdsh As Recordset
    Set yvgdsh = MstrDb.Execute("SELECT MAX(TemplateOrder) as CountCol FROM AmazonTemplateFields")
    
    Dim columns As Integer
    yvgdsh.MoveFirst
    columns = yvgdsh.fields("CountCol").value

    Dim i As Integer

    For i = 1 To columns
        Cells(NameRow, i).value = rst.fields("Field_Name").value  'NameRow is global variable
        Cells(LabelRow, i).value = rst.fields("Label_Name").value  'LabelRow is global variable
        rst.MoveNext
    Next i
    
    
    Dim lastcolumnletter As String
    Dim ColumnCount As Integer
    
    'Save the number and letter of the last column
    lastcolumnletter = NumberToColumn(CountColumns(range(LabelRow & ":" & LabelRow)))   'LabelRow is global variable
    
    'Format the UPC column to font size 14 and text so it doesn't show as scientific notation
    Call InitialFormatHeaders(lastcolumnletter)
    
    'Color the fields
    HeaderColors
    
    FinalFormatHeaders
    
    rst.Close

End Sub

Private Sub FillAmazonVariables()

    range("A1").value = TemplateType
    range("B1").value = TemplateVersion
    range("C1").value = TemplateSignature
    range("D1").value = TemplateAmazonUse

End Sub

Private Sub InitialFormatHeaders(lastcolumnletter As String)

    'format UPC column so it doesn't show scientific notation
    Dim AmznCol As Integer
    Dim col As Integer
    Dim ColLet As String
    
    col = AmazonColumn(lastcolumnletter, "external_product_id")
    ColLet = NumberToColumn(col)
    
    range(ColLet & ":" & ColLet).NumberFormat = "@"
    range(ColLet & ":" & ColLet).Font.Size = "14"

End Sub

Private Sub HeaderColors()

    Dim atf As Recordset
    Dim afg As Recordset
    
    Set atf = MstrDb.Execute("SELECT * FROM [AmazonTemplateFields] WHERE [TemplateOrder] > 0 ORDER BY [TemplateOrder]")
    Set afg = MstrDb.Execute("SELECT * FROM [AmazonFieldGroups]")
    
    Dim GroupVar As String
    
    'count columns
    Dim numcols As Integer
    numcols = CountColumns(Rows(2))
    
    Dim i As Integer
    'loop through columns in Excel sheet
    For i = 1 To numcols
        atf.MoveFirst
        'loop through fields in AmazonTemplateFields
        Do While Not atf.EOF
            'compare the field in the Excel sheet with the Access table and do manual lookup
            If atf.fields("Field_Name").value = Cells(NameRow, i).value Then
                'if Excel and Access table match on a field, save the Organization field value to variable
                GroupVar = atf.fields("Organization").value
                
                'loop through AmazonFieldGroups
                afg.MoveFirst
                Do While Not afg.EOF
                    'do manual lookup of field Group
                    If afg.fields("Group").value = GroupVar Then
                        'paint Rows 1-3 with the color in the AmazonFieldGroups table that matches the variable from previous loop
                        Cells(1, i).Interior.Color = RGB(afg.fields("Red"), afg.fields("Green"), afg.fields("Blue"))
                        Cells(2, i).Interior.Color = RGB(afg.fields("Red"), afg.fields("Green"), afg.fields("Blue"))
                        Cells(3, i).Interior.Color = RGB(afg.fields("Red"), afg.fields("Green"), afg.fields("Blue"))
                    End If
                    afg.MoveNext
                Loop
            End If
            atf.MoveNext
        Loop
    Next i
    
    atf.Close
    afg.Close

End Sub

Private Sub FinalFormatHeaders()

    Dim numcols
    numcols = CountColumns(Rows(2))
    
    'first autofit all the columns
    Dim i As Integer
    For i = 1 To numcols
        columns(i).EntireColumn.AutoFit
    Next i
    
    'next, shorten the fields that contain "TemplateSignature" and "The top 3 rows are for Amazon..."
    For i = 1 To numcols
        If Cells(1, i).value Like "TemplateSignature*" Or Cells(1, i).value Like "The top 3 rows are for Amazon*" Then
            'set the width of the column with "TemplateSignature" or "The top 3 rows are for Amazon..." to the longest of cells in row 2 and 3 (autofit while ignoring row 1)
            columns(i).EntireColumn.ColumnWidth = WorksheetFunction.Max(Len(Cells(2, i)), Len(Cells(3, i)))
        End If
    Next i
    
    'Add orders to headers
    ActiveSheet.UsedRange.Borders.LineStyle = xlContinuous
    
    Rows(1).Font.Bold = True
    Rows("1:3").Font.Size = "11"
    
    'split window into left column and top three rows
    With ActiveWindow
        .SplitColumn = 1    'split left column
        .SplitRow = 3       'split top 3 rows
    End With
    
    'freeze top 3 rows and left column
    ActiveWindow.FreezePanes = True

End Sub
