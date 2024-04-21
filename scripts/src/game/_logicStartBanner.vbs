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


Dim ball_saves_default : Set ball_saves_default = (new BallSave)("default", 10, 3, 2, "ball_started", "balldevice_plunger_ball_eject_success", true, 1, True)
Dim ball_saves_mode1 : Set ball_saves_mode1 = (new BallSave)("mode1", 10, 3, 2, "ball_started", "balldevice_plunger_ball_eject_success", false, 1, True)
Dim balldevice_plunger : Set balldevice_plunger = (new BallDevice)("plunger", "sw_plunger", Null, 3, True, 0, 50, "y-up", True)
