

Sub ConfigureDevices
    'Ball Devices
    Dim bd_plunger: Set bd_plunger = (new BallDevice)("bd_plunger")
    With bd_plunger
        .BallSwitches = Array("sw_plunger")
        .EjectTimeout = 1
        .EjectCallback = "PlungerKickBall"
        .MechcanicalEject = True
        .DefaultDevice = True
        .Debug = True
    End With

    Dim bd_cave_scoop: Set bd_cave_scoop = (new BallDevice)("cave_scoop")
    With bd_cave_scoop
        .BallSwitches = Array("sw39")
        .EjectTimeout = 2
        .EjectCallback = "CaveKickBall"
        .Debug = True
    End With

    Dim bd_waterfall_vuk: Set bd_waterfall_vuk = (new BallDevice)("bd_waterfall_vuk")
    With bd_waterfall_vuk
        .BallSwitches = Array("sw46")
        .EjectTimeout = 1
        .EjectCallback = "WaterfallVukKickBall"
        .Debug = True
    End With
    'Diverters
    Dim dv_panther : Set dv_panther = (new Diverter)("dv_panther", Array("ball_started"), Array("ball_ended"))', Array("activate_panther"), Array("deactivate_panther"), 0, False
    dv_panther.ActionCallback = "MovePanther"

    Dim dv_leftorbit : Set dv_leftorbit = (new Diverter)("leftorbit", Array("enable_waterfall"), Array("multiball_waterfall_started"))
    With dv_leftorbit
        .ActivationTime = 2000
        .ActivationSwitches = Array("sw47")
        .ActionCallback = "MoveLeftOrbitDiverter"
        .Debug = True
    End With
    

    Dim dv_waterfall : Set dv_waterfall = (new Diverter)("dv_waterfall", Array("game_started"), Array())
    With dv_waterfall
        .ActivationTime = 2000
        .ActivateEvents = Array("multiball_waterfall_started", "game_ended")
        .ActionCallback = "WaterfallRelease"
        .Debug = True
    End With

    'Drop Targets
    Dim dt_map1 : Set dt_map1 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map2 : Set dt_map2 = (new DropTarget)(sw05, sw05a, BM_sw05, 5, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map3 : Set dt_map3 = (new DropTarget)(sw06, sw06a, BM_sw06, 6, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map4 : Set dt_map4 = (new DropTarget)(sw08, sw08a, BM_sw08, 8, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map5 : Set dt_map5 = (new DropTarget)(sw09, sw09a, BM_sw09, 9, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map6 : Set dt_map6 = (new DropTarget)(sw10, sw10a, BM_sw10, 10, 0, False, Array("ball_started"," machine_reset_phase_3"))

    Dim DT01, DT02
    Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True, Null) 
    Set DT02 = (new DropTarget)(sw02, sw02a, BM_sw02, 2, 0, True, Null)

    
    DTArray = Array(DT01, DT02, dt_map1, dt_map2, dt_map3, dt_map4, dt_map5, dt_map6)

End Sub