
'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener START_GAME,    "ReleaseBall"
AddPinEventListener NEXT_PLAYER,   "ReleaseBall"
'
'*****************************
Sub ReleaseBall()
    swTrough1.kick 90, 10
    BIP = BIP + 1
    RandomSoundBallRelease swTrough1
    PuPlayer.LabelSet   pBackglass, "lblBall",      "Ball " & GetPlayerState(CURRENT_BALL),                        1,  "{}"
End Sub
