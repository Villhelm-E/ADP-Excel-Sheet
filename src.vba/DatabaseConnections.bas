Option Explicit

Private Function Provider() As String

    Dim appVersion
    appVersion = Application.Version
    
    Provider = "Provider=Microsoft.ACE.OLEDB." & appVersion & ";"

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
    MsgBox ("There was a problem connecting to the Master Database")

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
    MsgBox ("There was a problem connecting to the Find Sets Database")

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
    
    User = rst.fields("Sixbit_UserID").value
    pw = rst.fields("Sixbit_PW").value
    
    SxbtDb.Open "Provider=SQLOLEDB;Server=ADP-SERVER\SIXBITDBSERVER;Database=Sixbit;User Id=" & User & ";Password=" & pw & ";"
    SxbtDb.CursorLocation = adUseClient
    
    'close out connection
    rst.Close
    User = ""
    pw = ""

    Exit Sub

Connection_Error:
    MsgBox ("There was a problem connecting to the Sixbit Database")

End Sub

Public Sub ConnectADP_SQL_SERVER()

    Dim Server_Name As String
    Dim Database_Name As String
    
    On Error GoTo Connection_Error
    
    'Set Connection Timeouts
    Dim ADP_SQL As New ADODB.Connection
    
    ADP_SQL.ConnectionTimeout = 0
    ADP_SQL.CommandTimeout = 0
    
    'Open connection to ADP_SQ_SERVER\Seller Permits Database
    'Open username info
    Dim User As String
    Dim pw As String
    Set rst = MstrDb.Execute("SELECT * FROM Databases WHERE Database = ""Seller Permits""")
    rst.MoveFirst
    
    User = rst.fields("User").value
    pw = rst.fields("Password").value
    
    ADP_SQL.Open "Provider=SQLOLEDB;Server=192.168.1.101,1450;Database=ADP Seller Permits Database;User Id=AccessUser;Password=H$xZiRvX35^u6ouqoBH64fPUb*BY8%CS"
    ADP_SQL.CursorLocation = adUseClient
    
    rst.Close
    User = ""
    pw = ""
    
    Exit Sub
    
Connection_Error:
    MsgBox ("There was a problem connecting to ADP_SQL_SERVER\Seller Permits Database")

End Sub
