


'****************************
' Stat Of Game
' Event Listeners:  
AddPinEventListener START_GAME,    "GIStartOfGame"
'
'*****************************
Sub GIStartOfGame()
    Dim x
    For Each x in GI
        lightCtrl.LightOn x
    Next
End Sub

'****************************
' End Of Game
' Event Listeners:  
AddPinEventListener GAME_OVER,    "GIEndOfGame"
'
'*****************************
Sub GIEndOfGame()
    Dim x
    For Each x in GI
        lightCtrl.LightOff x
    Next
End Sub
