'***********************************************************************************************************************
'*****     GAME LOGIC START                                                 	                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************


Dim canAddPlayers : canAddPlayers = True
Dim currentPlayer : currentPlayer = Null
Dim PlungerDevice
Dim gameStarted : gameStarted = False
Dim pinEvents : Set pinEvents = CreateObject("Scripting.Dictionary")
Dim pinEventsOrder : Set pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerEvents : Set playerEvents = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")


'Dim ball_saves_default : Set ball_saves_default = (new BallSave)("default", 10, 3, 2, "ball_started", "balldevice_plunger_ball_eject_success", true, 1, False)
Dim balldevice_plunger : Set balldevice_plunger = (new BallDevice)("plunger", "sw_plunger", Null, 3, True, 0, 50, "y-up", False)
Dim balldevice_cave : Set balldevice_cave = (new BallDevice)("cave", "sw39", Null, 2, False, 0, 60, "z-up", True)

Dim mode_beasts : Set mode_beasts = (new Mode)("beasts", 100, Array("ball_started"), Array("ball_ended"), True)
Dim counter_beasts : Set counter_beasts = (new Counter)("beasts", mode_beasts, Array("mode_beasts_started", "diverter_panther_deactivating"), Array("s_left_ramp_opto_active"), 2, True, True, Array("activate_panther"), True, True)

Dim timer_beasts_panther : Set timer_beasts_panther = (new ModeTimer)("beasts_panther", mode_beasts)
With timer_beasts_panther
    .StartEvents = Array("activate_panther")
    .StopEvents = Null
    .Direction = "down"
    .StartValue = 10
    .EndValue = 0
    .Debug = True
End With

Dim event_player_beasts : Set event_player_beasts = (New EventPlayer)(mode_beasts)
Dim event_player_beasts_events : Set event_player_beasts_events = CreateObject("Scripting.Dictionary")
event_player_beasts_events.Add "timer_beasts_panther_complete", Array("deactivate_panther")
With event_player_beasts
    .Events = event_player_beasts_events
    .Debug = True
End With

Dim diverter_panther : Set diverter_panther = (new Diverter)("panther", Array("ball_started"), Array("ball_ended"), Array("activate_panther"), Array("deactivate_panther"), 0, "MovePanther", True)



Dim DT01, DT02, DT03, DT04, DT05, DT06, DT07, DT08, DT09, DT10, DT38, DT40, DT45, DT46, DT47
Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True, Null) 
Set DT04 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False, "ball_started")
Set DT05 = (new DropTarget)(sw05, sw05a, BM_sw05, 5, 0, False, "ball_started")
Set DT06 = (new DropTarget)(sw06, sw06a, BM_sw06, 6, 0, False, "ball_started")
Set DT08 = (new DropTarget)(sw08, sw08a, BM_sw08, 8, 0, False, "ball_started")
Set DT09 = (new DropTarget)(sw09, sw09a, BM_sw09, 9, 0, False, "ball_started")
Set DT10 = (new DropTarget)(sw10, sw10a, BM_sw10, 10, 0, False, "ball_started")
Dim DTArray
DTArray = Array(DT01,DT04, DT05, DT06, DT08, DT09, DT10)

Sub MovePanther(enabled)
    If enabled Then
        DTRaise 1
    Else
        DTDrop 1
    End If
End Sub