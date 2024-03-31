
'******************************************************
'*****  Drain                                      ****
'******************************************************

Sub Drain_Hit 
    BIP = BIP - 1
	Drain.kick 57, 20
    DispatchPinEvent BALL_DRAIN
End Sub

Sub Drain_UnHit : UpdateTrough : End Sub
