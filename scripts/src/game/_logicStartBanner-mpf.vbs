'***********************************************************************************************************************
'*****     GAME LOGIC START                                                 	                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************

Const AllLightsOnMode = True
Dim canAddPlayers : canAddPlayers = True
Dim currentPlayer : currentPlayer = Null
Dim PlungerDevice
Dim gameStarted : gameStarted = False
Dim pinEvents : Set pinEvents = CreateObject("Scripting.Dictionary")
Dim pinEventsOrder : Set pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerEvents : Set playerEvents = CreateObject("Scripting.Dictionary")
Dim playerEventsOrder : Set playerEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")
Dim bcpController
Dim useBCP : useBCP = False

Dim DT01, DT02, DT03, DT04, DT05, DT06, DT07, DT08, DT09, DT10, DT38, DT40, DT45, DT46, DT47
Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True, Null) 
Set DT04 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False, Array("ball_started"))
Set DT05 = (new DropTarget)(sw05, sw05a, BM_sw05, 5, 0, False, Array("ball_started"))
Set DT06 = (new DropTarget)(sw06, sw06a, BM_sw06, 6, 0, False, Array("ball_started"))
Set DT08 = (new DropTarget)(sw08, sw08a, BM_sw08, 8, 0, False, Array("ball_started"))
Set DT09 = (new DropTarget)(sw09, sw09a, BM_sw09, 9, 0, False, Array("ball_started"))
Set DT10 = (new DropTarget)(sw10, sw10a, BM_sw10, 10, 0, False, Array("ball_started"))
Dim DTArray
DTArray = Array(DT01,DT04, DT05, DT06, DT08, DT09, DT10)

Sub MovePanther(enabled)
    If enabled Then
        DTRaise 1
    Else
        DTDrop 1
    End If
End Sub