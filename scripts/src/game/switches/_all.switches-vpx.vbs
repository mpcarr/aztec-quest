
'Switches

Sub sw_plunger_Hit()   : DispatchPinEvent "sw_plunger_active",   ActiveBall : End Sub
Sub sw_plunger_Unhit() : DispatchPinEvent "sw_plunger_inactive", ActiveBall : End Sub

Sub s_start_Hit()   : DispatchPinEvent "s_start_active",   ActiveBall : End Sub
Sub s_start_Unhit() : DispatchPinEvent "s_start_inactive", ActiveBall : End Sub

Sub sw39_Hit()   : DispatchPinEvent "sw39_active",   ActiveBall : End Sub
Sub sw39_Unhit() : DispatchPinEvent "sw39_inactive", ActiveBall : End Sub


