
'***********************************************************************************
'***** Switches                                                         	    ****
'***********************************************************************************

Sub sw11_Hit()
    STHit 11
End Sub

Sub sw11o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw12_Hit()
    STHit 12
End Sub

Sub sw12o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw13_Hit()
    STHit 13
End Sub

Sub sw13o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw15_Hit()
    STHit 15
End Sub

Sub sw15o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw16_Hit()
    STHit 16
End Sub

Sub sw16o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw17_Hit()
    STHit 17
End Sub

Sub sw17o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw41_Hit()
    STHit 41
End Sub

Sub sw01_Hit()
    DTHit 1
End Sub

Sub sw02_Hit()
    DTHit 2
End Sub

Sub sw04_Hit()
    DTHit 4
End Sub

Sub sw05_Hit()
    DTHit 5
End Sub

Sub sw06_Hit()
    DTHit 6
End Sub


Sub sw08_Hit()
    DTHit 8
End Sub

Sub sw09_Hit()
    DTHit 9
End Sub

Sub sw10_Hit()
    DTHit 10
End Sub

Sub sw45_Hit()
    'DTHit 45
End Sub

Sub sw99_Hit()
    DispatchPinEvent "s_left_ramp_opto_active", Null
End Sub
Sub sw99_UnHit()
    DispatchPinEvent "s_left_ramp_opto_inactive", Null
End Sub

Sub sw_plunger_Hit()
    DispatchPinEvent "sw_plunger_active", ActiveBall
End Sub

Sub sw_plunger_Unhit()
    DispatchPinEvent "sw_plunger_inactive", ActiveBall
End Sub

Sub sw39_Hit()
    SoundSaucerLock()
    DispatchPinEvent "sw39_active", ActiveBall
End Sub

Sub sw39_Unhit()
    DispatchPinEvent "sw39_inactive", ActiveBall
End Sub

Sub DTAction(switchid, enabled)
    If enabled = 1 Then
        Select Case switchid
            case 1:
                DispatchPinEvent "sw01_inactive", Null
            case 2:
                DispatchPinEvent "sw02_inactive", Null
            case 4:
                DispatchPinEvent "sw04_active", Null
            case 5:
                DispatchPinEvent "sw05_active", Null
            case 6:
                DispatchPinEvent "sw06_active", Null
            case 8:
                DispatchPinEvent "sw08_active", Null
            case 9:
                DispatchPinEvent "sw09_active", Null
            case 10:
                DispatchPinEvent "sw10_active", Null
        End Select
    ElseIf enabled = 0 Then
        Select Case switchid
            case 1:
                DispatchPinEvent "sw01_active", Null
            case 2:
                DispatchPinEvent "sw02_active", Null
            case 4:
                DispatchPinEvent "sw04_inactive", Null
            case 5:
                DispatchPinEvent "sw05_inactive", Null
            case 6:
                DispatchPinEvent "sw06_inactive", Null
            case 8:
                DispatchPinEvent "sw08_inactive", Null
            case 9:
                DispatchPinEvent "sw09_inactive", Null
            case 10:
                DispatchPinEvent "sw10_inactive", Null            
        End Select
    End If
End Sub


Sub STAction(switchid, enabled)
    If enabled = 1 Then
        Select Case switchid
            case 11:
                DispatchPinEvent "sw11_active", Null
            case 12:
                DispatchPinEvent "sw12_active", Null
            case 13:
                DispatchPinEvent "sw13_active", Null
            case 15:
                DispatchPinEvent "sw15_active", Null
            case 16:
                DispatchPinEvent "sw16_active", Null
            case 17:
                DispatchPinEvent "sw17_active", Null
        End Select
    ElseIf enabled = 0 Then
        Select Case switchid
            case 11:
                DispatchPinEvent "sw11_inactive", Null
            case 12:
                DispatchPinEvent "sw12_inactive", Null
            case 13:
                DispatchPinEvent "sw13_inactive", Null
            case 15:
                DispatchPinEvent "sw15_inactive", Null
            case 16:
                DispatchPinEvent "sw16_inactive", Null
            case 17:
                DispatchPinEvent "sw17_inactive", Null
        End Select
    End If
End Sub