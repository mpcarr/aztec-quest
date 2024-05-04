
'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener START_GAME, "start_game_release_ball",   "ReleaseBall", 1000, True
AddPinEventListener NEXT_PLAYER, "next_player_release_ball",   "ReleaseBall", 1000, True
'
'*****************************
Function ReleaseBall(args)
    If Not IsNull(args) Then
        If args(0) = True Then
            DispatchPinEvent "ball_started", Null
        End If
    End If
    debugLog.WriteToLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    swTrough1.kick 90, 10
    ballInReleasePostion = False
    debugLog.WriteToLog "Release Ball", "Just Kicked"
    BIP = BIP + 1
    RandomSoundBallRelease swTrough1
End Function
