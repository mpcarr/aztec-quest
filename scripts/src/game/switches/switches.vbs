
'******************************************************
'*****  Drain                                      ****
'******************************************************

Sub Drain_Hit 
    BIP = BIP - 1
	Drain.kick 57, 20
    DispatchRelayPinEvent "ball_drain", 1
End Sub

Sub Drain_UnHit : UpdateTrough : End Sub
