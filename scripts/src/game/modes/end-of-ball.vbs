
'******************************************************
'*****   End of Ball                               ****
'******************************************************

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener "ball_drain", "ball_drain", "EndOfBall", 20, Null
'
'*****************************
Function EndOfBall(args)
    
    Dim ballsToSave : ballsToSave = args(1) 
    debugLog.WriteToLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
    debugLog.WriteToLog "end_of_ball, balls in play", CStr(BIP)
    If ballsToSave <= 0 Then
        Exit Function
    End If

    If BIP > 0 Then
        Exit Function
    End If
        
    DispatchPinEvent "ball_ended", Null
    SetPlayerState CURRENT_BALL, GetPlayerState(CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = GetCurrentPlayerNumber()
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
    
    If useBcp Then
        bcpController.SendPlayerVariable "number", GetCurrentPlayerNumber(), previousPlayerNumber
    End If
    If GetPlayerState(CURRENT_BALL) > BALLS_PER_GAME Then
        DispatchPinEvent GAME_OVER, Null
        gameStarted = False
        currentPlayer = Null
        playerState.RemoveAll()
    Else
        DispatchPinEvent NEXT_PLAYER, Null
    End If
    
End Function
