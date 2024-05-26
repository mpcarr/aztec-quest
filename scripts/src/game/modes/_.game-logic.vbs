
'******************************************************
'*****   GAME MODE LOGIC START                     ****
'******************************************************

Sub StartGame()
    gameStarted = True
    'SetPlayerState BALL_SAVE_ENABLED, True
    If useBcp Then
        bcpController.Send "player_turn_start?player_num=int:1"
        bcpController.Send "ball_start?player_num=int:1&ball=int:1"
        bcpController.PlaySlide "attract", "base", 1000
        bcpController.SendPlayerVariable "number", 1, 0
    End If

    DispatchPinEvent START_GAME, Null

    
    'mode_start?name=game&priority=int:20
'08:04:31.505 : VERBOSE : BCP : Received BCP command: 
'08:04:31.505 : VERBOSE : BCP : Received BCP command: mode_start?name=base&priority=int:2000
'08:04:31.505 : VERBOSE : BCP : Received BCP command: mode_start?name=beasts&priority=int:2000
'08:04:31.505 : VERBOSE : BCP : Received BCP command: 
End Sub



'****************************
' End Of Game
' Event Listeners:  
    AddPinEventListener GAME_OVER, "end_of_game", "EndOfGame", 1000, Null
'
'*****************************
Function EndOfGame(args)
    
End Function
