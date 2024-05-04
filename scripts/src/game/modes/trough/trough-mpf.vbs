
'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
'
'*****************************
Function ReleaseBall()
    swTrough1.kick 90, 10
    UpdateTrough()
    RandomSoundBallRelease swTrough1
End Function
