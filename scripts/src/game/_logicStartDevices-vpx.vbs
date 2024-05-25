
'Devices
Dim bd_plunger: Set bd_plunger = (new BallDevice)("bd_plunger", "sw_plunger", Null, 1, True, False)
Dim bd_cave_scoop: Set bd_cave_scoop = (new BallDevice)("bd_cave_scoop", "sw39", Null, 2, False, False)

'Diverters
Dim dv_panther : Set dv_panther = (new Diverter)("dv_panther", Array("ball_started"), Array("ball_ended"), Array("activate_panther"), Array("deactivate_panther"), 0, False)
Dim dv_leftorbit : Set dv_leftorbit = (new Diverter)("dv_leftorbit", Array("ball_started"), Array("ball_ended"), Array("activate_panther"), Array(), 0, True)
Dim dv_waterfall : Set dv_waterfall = (new Diverter)("dv_waterfall", Array("ball_started"), Array("ball_ended"), Array("sw44_active"), Array(), 0, True)


'Drop Targets
Dim dt_map1 : Set dt_map1 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False, Array("ball_started"," machine_reset_phase_3"))
Dim dt_map2 : Set dt_map2 = (new DropTarget)(sw05, sw05a, BM_sw05, 5, 0, False, Array("ball_started"," machine_reset_phase_3"))
Dim dt_map3 : Set dt_map3 = (new DropTarget)(sw06, sw06a, BM_sw06, 6, 0, False, Array("ball_started"," machine_reset_phase_3"))
Dim dt_map4 : Set dt_map4 = (new DropTarget)(sw08, sw08a, BM_sw08, 8, 0, False, Array("ball_started"," machine_reset_phase_3"))
Dim dt_map5 : Set dt_map5 = (new DropTarget)(sw09, sw09a, BM_sw09, 9, 0, False, Array("ball_started"," machine_reset_phase_3"))
Dim dt_map6 : Set dt_map6 = (new DropTarget)(sw10, sw10a, BM_sw10, 10, 0, False, Array("ball_started"," machine_reset_phase_3"))
