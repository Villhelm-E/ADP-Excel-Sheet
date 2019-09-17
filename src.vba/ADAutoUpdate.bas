Option Explicit

Declare Function apiCopyFile Lib "KERNEL32" Alias "CopyFileA" _
    (ByVal lpExistingFileName As String, _
        ByVal lpNewFileName As String, _
            ByVal bFailIfExists As Long) As Long
            
Public Function UpdateFeVersion(ThisVersion As String)
    On Error GoTo ProcError
    
    Dim strSourceFile As String
    Dim strDestFile As String
    Dim strExcelExePath As String
    Dim LResponse As Integer
    Dim Updated As String
    Dim currentbook As Workbook
    Dim strCurrentBook As String
    Dim destBook As Workbook
    Dim sourceBook As Workbook
    
'''''Set variables for workbook manipulation
    'current book
    strCurrentBook = Application.ActiveWorkbook.FullName
    Set currentbook = ActiveWorkbook
    
    'Create the source's path and file name.
    strSourceFile = "\\ADP-SERVER\AD AutoParts Server\IT\ADP Systems - Source Code\ADP Excel Sheet\ADP Excel Sheet.xlsm"
    
    Updated = Replace(rst.Fields("Version").Value, ".", "_")
    
    If Right(strCurrentBook, Len(ThisVersion)) = Replace(ThisVersion, ".", "_") Then
        strDestFile = Left(strCurrentBook, Len(strCurrentBook) - (Len(ThisVersion) + 1)) & " " & Updated & ".xlsm"
    Else
        strDestFile = "ADP Excel Sheet " & Updated & ".xlsm"
    End If
    
    'Determine path of current Excel executable.
    Call GetExcelPathFormatVB(strExcelExePath)
    
    'Show message box with Yes/No options
    LResponse = MsgBox("A new version is available. Do you wish to update now?", vbYesNo, "Software Update")
    
    'Determine Yes/No options
    If LResponse = vbYes Then
        If Dir(strSourceFile) = "" Then 'something is wrong and the file is not there.
            
            MsgBox "The file:" & vbCrLf & Chr(34) & strSourceFile & _
                Chr(34) & vbCrLf & vbCrLf & _
                "is not a valid file name. Please see your Administrator.", _
                vbCritical, "Error updating To New Version..."
                GoTo ExitProc
        Else
            'Message box informing update is happening
            MsgBox "Application updated. Please wait while the application" & _
                " restarts.", vbInformation, "Update Successful"
            
            'Copy Excel Sheet from Server to final location
            Set sourceBook = Workbooks.Open(strSourceFile)
            sourceBook.SaveAs strDestFile
            
            'save current workbook to variable
            Set destBook = ActiveWorkbook
            
            'Copy Worksheets from current Sheet to the copy in the previous step
            Call CopySheets(currentbook, destBook)
            destBook.Save
            
            'Close the old workbook
            currentbook.Close savechanges:=True
        End If
    Else
        GoTo ExitProc
    End If
    
ExitProc:
    Exit Function
ProcError:
    MsgBox "Error " & Err.Number & ": " & Err.Description, , _
        "Error in UpdateFEVersion event procedure..."
    Resume ExitProc
End Function

Private Sub GetExcelPathFormatVB(strExcelExePath As String)


    Dim appXL As Object
    Dim s As String
    
    Set appXL = CreateObject("Excel.Application")
    s = appXL.Path
    strExcelExePath = s & "\EXCEL.exe"
    appXL.Quit
    Set appXL = Nothing
    
End Sub

Private Sub CopySheets(currentbook As Workbook, destBook As Workbook)

    Dim currentsheet As Worksheet
    Dim sheetIndex As Integer
    Dim numsheets As Integer
    Dim i As Integer
    Dim exists As Boolean
    Dim oldName As String
    
    'save number of sheets in new workbook
    sheetIndex = destBook.Sheets.Count
    
    'loop through current workbook sheets
    For Each currentsheet In currentbook.Worksheets
        On Error GoTo Exit_Loop
        
        'loop through worksheets in new workbook
        For i = 1 To Worksheets.Count
            'if the worksheet already exists, set exists to true and end the loop
            If currentbook.Sheets(i).name = currentsheet.name Then
                exists = True
                GoTo exit_i_loop 'ends loop
            Else
                'sheet doesn't exist
                exists = False
            End If
        Next i
        
exit_i_loop:
        
        'if sheet doesn't exist, then copy from current sheet to new sheet
        If Not exists Then
            currentsheet.Copy after:=destBook.Sheets(sheetIndex)
            sheetIndex = sheetIndex + 1 'increment sheet index in new sheet
        Else
        '''''if sheet exists, replace the sheet in new workbook with sheet from current workbook
            oldName = currentsheet.name
            
            'copy the sheet
            If oldName = "Amazon Template" Then
                GoTo Amazon_Template
            Else
                currentsheet.Copy after:=destBook.Sheets(i)
            End If
            
            'delete the existing sheet
            Application.DisplayAlerts = False
            destBook.Sheets(currentsheet.name).Delete
            Application.DisplayAlerts = True
            
            'rename the sheet, copying adds " (2)" to the end
            If i < 2 Then
                i = 2
            End If
            
            destBook.Sheets(i).name = oldName
        End If
        
Amazon_Template:
    Next currentsheet
    
Exit_Loop:
    'Ends looping through current workbook sheets
    'sheetindex doesn't have an upper bound, so it errors when it goes beyond scope
End Sub


