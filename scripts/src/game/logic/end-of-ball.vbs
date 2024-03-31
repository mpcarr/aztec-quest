
'******************************************************
'*****   End of Ball                               ****
'******************************************************

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener BALL_DRAIN, "EndOfBall"
'
'*****************************
Sub EndOfBall()
    debug.print("Ball Saver" & ballSaver)
    If ballSaver = True Then
        DispatchPinEvent BALL_SAVE
    ElseIf BIP - GetPlayerState(BALLS_LOCKED) = 0 Then
        SetPlayerState CURRENT_BALL, GetPlayerState(CURRENT_BALL) + 1

        Select Case currentPlayer
            Case "PLAYER 1":
                If UBound(playerState.Keys()) > 0 Then
                    currentPlayer = "PLAYER 2"
                End If
            Case "PLAYER 2":
                If UBound(playerState.Keys()) > 1 Then
                    currentPlayer = "PLAYER 3"
                Else
                    currentPlayer = "PLAYER 1"
                End If
            Case "PLAYER 3":
                If UBound(playerState.Keys()) > 2 Then
                    currentPlayer = "PLAYER 4"
                Else
                    currentPlayer = "PLAYER 1"
                End If
            Case "PLAYER 4":
                currentPlayer = "PLAYER 1"
        End Select

        If GetPlayerState(CURRENT_BALL) > BALLS_PER_GAME Then
            DispatchPinEvent GAME_OVER
            gameStarted = False
            currentPlayer = Null
            playerState.RemoveAll()
        Else
            SetPlayerState BALL_SAVE_ENABLED, True 
            DispatchPinEvent NEXT_PLAYER
        End If
    End If
End Sub
