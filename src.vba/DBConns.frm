Option Explicit

Private Sub UserForm_Initialize()

    'set caption of labels to active or inactive based on connection status
    SetCaptions
    
    'format the font color of each connection label
    FormatStatus
    
    'disable the reconnect button if all database connections are active
    RefreshReconnectButton
    
    'position the userform
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)

End Sub

Private Sub SetCaptions()

    'set caption of labels to active or inactive based on connection status
    Me.MstrConn.Caption = MstrDbState
    Me.FndStsConn.Caption = FndStsDbState
    Me.SxbtConn.Caption = SxbtDbState

End Sub

Private Function MstrDbState() As String

    'if connection is open, set to "Active", otherwise "Inactive"
    If MstrDb.State = adStateOpen Then
        MstrDbState = "Active"
    Else
        MstrDbState = "Inactive"
    End If

End Function

Private Function FndStsDbState() As String

    'if connection is open, set to "Active", otherwise "Inactive"
    If FndStsDb.State = adStateOpen Then
        FndStsDbState = "Active"
    Else
        FndStsDbState = "Inactive"
    End If

End Function

Private Function SxbtDbState() As String

    'if connection is open, set to "Active", otherwise "Inactive"
    If SxbtDb.State = adStateOpen Then
        SxbtDbState = "Active"
    Else
        SxbtDbState = "Inactive"
    End If

End Function

Private Sub FormatStatus()

    Dim i As control
    
    For Each i In Me.Controls
        If TypeName(i) = "Label" Then
            'Isolate Labels that end in Conn to format
            If Right(i.name, 4) = "Conn" Then
                'if label caption is "Active", color green
                If i.Caption = "Active" Then
                    i.ForeColor = RGB(0, 192, 0)
                Else
                    'otherwise red
                    i.ForeColor = RGB(255, 0, 0)
                End If
            End If
        End If
    Next i

End Sub

Private Sub RefreshReconnectButton()

    'if any of the databases are inactive, enable the Reconnect button
    If MstrDbState = "Inactive" Or FndStsDbState = "Inactive" Or SxbtDbState = "Inactive" Then
        Me.Reconnect.Enabled = True
    Else
        Me.Reconnect.Enabled = False
    End If
    
    'Refresh color of Statuses
    FormatStatus

End Sub

Private Sub Reconnect_Click()

''''Connect to Master Database if not connected
    If MstrDbState = "Inactive" Then ConnectMasterDatabase
    
''''Connect to ADP Find Sets Database if not connected
    If FndStsDbState = "Inactive" Then ConnectFindSetsDatabase
    
''''Connect to Sixbit Database if not connected
    If SxbtDbState = "Inactive" Then ConnectSixbitDatabase
    
    'check to see if Reconnect button should be deactivated
    RefreshReconnectButton
    
    'Refresh captions
    SetCaptions
    
    'Format color of stauses
    FormatStatus

End Sub

Private Sub CloseBtn_Click()

    Unload DBConns

End Sub
