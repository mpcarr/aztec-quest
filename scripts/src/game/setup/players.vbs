

'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub AddPlayer()
    Select Case UBound(playerState.Keys())
        Case -1:
            playerState.Add "PLAYER 1", InitNewPlayer()
            currentPlayer = "PLAYER 1"
            PuPlayer.LabelSet   pBackglass, "lblPlayer1",             "Player 1",                        1,  "{}"
            PuPlayer.LabelSet   pBackglass, "lblPlayer1Score",        "00",                        1,  "{}"
        Case 0:     
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 2", InitNewPlayer()
                PuPlayer.LabelSet   pBackglass, "lblPlayer2",         "Player 2",                        1,  "{}"
                PuPlayer.LabelSet   pBackglass, "lblPlayer2Score",    "00",                        1,  "{}"
            End If
        Case 1:
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 3", InitNewPlayer()
                PuPlayer.LabelSet   pBackglass, "lblPlayer3",         "Player 3",                        1,  "{}"
                PuPlayer.LabelSet   pBackglass, "lblPlayer3Score",    "00",                        1,  "{}"
            End If     
        Case 2:   
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 4", InitNewPlayer()
                PuPlayer.LabelSet   pBackglass, "lblPlayer4",         "Player 4",                        1,  "{}"
                PuPlayer.LabelSet   pBackglass, "lblPlayer4Score",    "00",                        1,  "{}"
            End If  
            canAddPlayers = False
    End Select
End Sub

Function InitNewPlayer()

    Dim state: Set state=CreateObject("Scripting.Dictionary")

    state.Add SCORE, 0
    state.Add PLAYER_NAME, ""
    state.Add CURRENT_BALL, 1

    state.Add LANE_1,   0
    state.Add LANE_2,   0
    state.Add LANE_3,   0
    state.Add LANE_4,   0

    state.Add BALLS_LOCKED, 0

    state.Add BALL_SAVE_ENABLED, False
    
    Set InitNewPlayer = state

End Function


'****************************
' Setup Player
' Event Listeners:  
    AddPinEventListener START_GAME,    "SetupPlayer"
    AddPinEventListener NEXT_PLAYER,   "SetupPlayer"
'
'*****************************
Sub SetupPlayer()
    EmitAllPlayerEvents()
End Sub
