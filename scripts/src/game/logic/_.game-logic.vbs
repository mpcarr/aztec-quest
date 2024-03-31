
'******************************************************
'*****   GAME MODE LOGIC START                     ****
'******************************************************

Sub StartGame()
    gameStarted = True
    SetPlayerState BALL_SAVE_ENABLED, True
    DispatchPinEvent START_GAME
End Sub

'****************************
' End Of Game
' Event Listeners:  
    AddPinEventListener GAME_OVER,    "EndOfGame"
'
'*****************************
Sub EndOfGame()
    
End Sub
