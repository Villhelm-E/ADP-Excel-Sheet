Option Explicit

Public Sub OfficeVersion()

    MsgBox "Microsoft Office 2013", , "Microsoft Office"

End Sub

Public Sub ADPVersion()

    Dim ThisVersion As String
    Dim CurrentVersion
    
    ThisVersion = "3.0"
    
    'open Excel Sheet Version Table
    If MstrDb.State = adStateOpen Then
        OpenExcelVersion
    Else
        MsgBox "Master Database not connected."
        Exit Sub
    End If
    
    'grab version number from Master Database
    With rst
        CurrentVersion = rst.Fields("Version").Value
    End With
    
    MsgBox "ADP Excel Sheet Version: " & ThisVersion, , "ADP Excel Sheet"
    
    'close table and connection to Master Database
    rst.Close

End Sub
