


'****************************
' Stat Of Game
' Event Listeners:  
AddPinEventListener START_GAME, "start_game_gi", "GIStartOfGame", 1000, Null
'
'*****************************
Function GIStartOfGame(args)
    Dim x
    For Each x in GI
        lightCtrl.LightOn x
    Next
End Function

'****************************
' End Of Game
' Event Listeners:  
AddPinEventListener GAME_OVER, "game_over_gi", "GIEndOfGame", 1000, Null
'
'*****************************
Function GIEndOfGame(args)
    Dim x
    For Each x in GI
        lightCtrl.LightOff x
    Next
End Function
