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
    MstrDb.CursorLocation = adUseClient
    
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
    FndStsDb.CursorLocation = adUseClient
    
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
    'open username info
    Dim User As String
    Dim pw As String
    Set rst = MstrDb.Execute("SELECT * FROM Sixbit_DB_Fields")
    rst.MoveLast
    
    User = rst.Fields("Sixbit_UserID").Value
    pw = rst.Fields("Sixbit_PW").Value
    
    SxbtDb.Open "Provider=SQLOLEDB;Server=ADP-SERVER\SIXBITDBSERVER;Database=Sixbit;User Id=" & User & ";Password=" & pw & ";"
    SxbtDb.CursorLocation = adUseClient
    
    'close out connection
    rst.Close
    User = ""
    pw = ""

    Exit Sub

Connection_Error:
    MsgBox "There was a problem connecting to the Sixbit Database"

End Sub
