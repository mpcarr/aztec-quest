
'******************************************************
'*****  Drain                                      ****
'******************************************************

Sub Drain_Hit 
    Drain.kick 57, 20    
    If gameStarted = True Then
        BIP = BIP - 1
        DispatchRelayPinEvent "ball_drain", 1
    End If
End Sub

Sub Drain_UnHit : UpdateTrough : End Sub
