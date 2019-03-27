Option Explicit

Private Function Provider() As String

    Provider = "Provider=Microsoft.ACE.OLEDB.15.0;"         '12.0 is Office 2007 version

End Function

Public Sub ConnectMasterDatabase()

    On Error GoTo Connection_Error
    
    Dim DBPath As String
    
    'Database path
    DBPath = "\\ADP-SERVER\AD AutoParts Server\IT\ADP Systems - Source Code\Master Database\Master Database.accdb"
    
    'open connection to Master Database
    MstrDb.Open Provider & "Data Source=" & DBPath      'MstrDb is a global variable
    
    Exit Sub
    
Connection_Error:
    MsgBox "There was a problem connecting to the Master Database"

End Sub

Public Sub ConnectFindSetsDatabase()

    On Error GoTo Connection_Error
    
    Dim DBPath As String
    
    'Database path
    DBPath = "\\ADP-SERVER\AD AutoParts Server\IT\ADP Systems - Source Code\Find Sets\ADP Find Sets.accdb"
    
    'open connection to Find Sets Database
    FndStsDb.Open Provider & "Data Source=" & DBPath    'FndStsDb is global variable
    
    Exit Sub
    
Connection_Error:
    MsgBox "There was a problem connecting to the Find Sets Database"

End Sub

Public Sub ConnectSixbitDatabase()

    Dim Server_Name As String
    Dim Database_Name As String

    On Error GoTo Connection_Error

    'Set Connection Timeouts
    SxbtDb.ConnectionTimeout = 0
    SxbtDb.CommandTimeout = 0

    'open connection to Sixbit database
    SxbtDb.Open "Provider=SQLOLEDB;Server=ADP-SERVER\SIXBITDBSERVER;Database=Sixbit;User Id=sa;Password=S1xb1tR0x;"     'unfortunately have to hard code the usernmae and password

    Exit Sub

Connection_Error:
    MsgBox "There was a problem connecting to the Sixbit Database"

End Sub
