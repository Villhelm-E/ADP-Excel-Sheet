Option Explicit

Public Sub XLSX()

    Dim FlSv As Variant
    Dim MyFile As String
    Dim sh As Worksheet
    Dim currentsheetname As Variant
    
    'save current sheet's name to variable
    currentsheetname = Application.ActiveSheet.Name

    'activate current sheet and copy it
    Set sh = Sheets(currentsheetname)
    sh.Copy
    'add extension to sheet name
    MyFile = currentsheetname & ".xlsx"
    
    'open up the Save As Window to verify file name and choose a destination to save
    FlSv = Application.GetSaveAsFilename(MyFile, fileFilter:="Excel Files (*.xlsx), *.xlsx)", Title:="Enter your file name")

    'if user cancels the operation, close the copy and stop code
    If FlSv = False Then
        ActiveWorkbook.Close False
        Exit Sub
    End If

    'overwrite the MyFile variable with the user's input
    MyFile = FlSv

    'saves the copy and closes it
    With ActiveWorkbook
        .SaveAs (MyFile), FileFormat:=51, CreateBackup:=False
        .Close False
    End With

End Sub

Public Sub CSV()

    Dim FlSv As Variant
    Dim MyFile As String
    Dim sh As Worksheet
    Dim currentsheetname As Variant
    
    'save current sheet's name to variable
    currentsheetname = Application.ActiveSheet.Name

    'activate current sheet and copy it
    Set sh = Sheets(currentsheetname)
    sh.Copy
    'add extension to sheet name
    MyFile = currentsheetname & ".csv"
    
    'open up the Save As Window to verify file name and choose a destination to save
    FlSv = Application.GetSaveAsFilename(MyFile, fileFilter:="Excel Files (*.csv), *.csv)", Title:="Enter your file name")

    'if user cancels the operation, close the copy and stop code
    If FlSv = False Then
        ActiveWorkbook.Close False
        Exit Sub
    End If

    'overwrite the MyFile variable with the user's input
    MyFile = FlSv

    'saves the copy and closes it
    With ActiveWorkbook
        .SaveAs (MyFile), FileFormat:=xlCSV, CreateBackup:=False
        .Close False
    End With

End Sub

Public Sub TXT()

    Dim FlSv As Variant
    Dim MyFile As String
    Dim sh As Worksheet
    Dim currentsheetname As Variant
    
    'save current sheet's name to variable
    currentsheetname = Application.ActiveSheet.Name

    'activate current sheet and copy it
    Set sh = Sheets(currentsheetname)
    sh.Copy
    'add extension to sheet name
    MyFile = currentsheetname & ".txt"
    
    'open up the Save As Window to verify file name and choose a destination to save
    FlSv = Application.GetSaveAsFilename(MyFile, fileFilter:="Excel Files (*.txt), *.txt)", Title:="Enter your file name")

    'if user cancels the operation, close the copy and stop code
    If FlSv = False Then
        ActiveWorkbook.Close False
        Exit Sub
    End If

    'overwrite the MyFile variable with the user's input
    MyFile = FlSv

    'saves the copy and closes it
    With ActiveWorkbook
        .SaveAs (MyFile), FileFormat:=xlText, CreateBackup:=False
        .Close False
    End With

End Sub

Public Sub EmailMain()

    Dim oApp As Object
    Dim oMail As Object
    Dim LWorkbook As Workbook
    Dim LFileName As String
    
    'Turn off screen updating
    Application.ScreenUpdating = False
    
    'name email attachments based on how many sheets user has selected
    If ActiveWindow.SelectedSheets.Count > 1 Then
        LFileName = "Workbook Attachment"
    Else
        LFileName = ActiveSheet.Name
    End If
    
    'validate attachment name
    If NameValid(LFileName) = True Then
        
        'copy all selected sheets to new temp workbook
        ActiveWindow.SelectedSheets.Copy
        Set LWorkbook = ActiveWorkbook
        
        On Error Resume Next
        
        'delete the file if it already exists
        Kill LFileName
        
        On Error GoTo 0
        
        'save the temporary file
        LWorkbook.SaveAs FileName:=LFileName
        
        'create an outlook object and new mail message
        Set oApp = CreateObject("Outlook.Application")
        Set oMail = oApp.CreateItem(0)
        
        'set mail attributes
        With oMail
            '.To = "user@yahoo.com"
            .Subject = LFileName
            '.body = "This is the body of the message."
            .Attachments.Add LWorkbook.FullName
            .Display
        End With
        
        'delete the temporary file and close temporary workbook
        LWorkbook.ChangeFileAccess Mode:=xlReadOnly
        Kill LWorkbook.FullName
        LWorkbook.Close savechanges:=False
        
        'turn screen updating back on and clear objects
        Application.ScreenUpdating = True
        Set oMail = Nothing
        Set oApp = Nothing
        
    Else
        MsgBox "The file name is invalid. Please rename attachment."
    End If

End Sub