
'******************************************************
'*****   GAME MODE LOGIC START                     ****
'******************************************************

Sub StartGame()
    gameStarted = True
    SetPlayerState BALL_SAVE_ENABLED, True
    DispatchPinEvent START_GAME, Null
End Sub

'****************************
' End Of Game
' Event Listeners:  
    AddPinEventListener GAME_OVER, "end_of_game", "EndOfGame", 1000, Null
'
'*****************************
Function EndOfGame(args)
    
End Function
