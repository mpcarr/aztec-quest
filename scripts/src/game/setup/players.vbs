

'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub AddPlayer()
    Select Case UBound(playerState.Keys())
        Case -1:
            playerState.Add "PLAYER 1", InitNewPlayer()
            currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 2", InitNewPlayer()
            End If
        Case 1:
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 3", InitNewPlayer()
            End If     
        Case 2:   
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 4", InitNewPlayer()
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
    AddPinEventListener START_GAME,  "start_game_setup",  "SetupPlayer", 1000, Null
    AddPinEventListener NEXT_PLAYER, "next_player_setup",  "SetupPlayer", 1000, Null
'
'*****************************
Function SetupPlayer(args)
    EmitAllPlayerEvents()
End Function
