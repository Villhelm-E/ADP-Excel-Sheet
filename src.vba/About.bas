Option Explicit

Public Sub ADPVersion()

    Dim ThisVersion As String
    Dim CurrentVersion
    
    ThisVersion = "3.11"
    
    'open Excel Sheet Version Table
    If MstrDb.State = adStateOpen Then
        OpenExcelVersion
    Else
        MsgBox ("Master Database not connected.")
        Exit Sub
    End If
    
    'grab version number from Master Database
    With rst
        CurrentVersion = rst.fields("Version").value
    End With
    
    MsgBox ("ADP Excel Sheet Version: " & ThisVersion), , "ADP Excel Sheet"
    
    'close table and connection to Master Database
    rst.Close

End Sub

••••ˇˇˇˇ