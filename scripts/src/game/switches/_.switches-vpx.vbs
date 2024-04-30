
'***********************************************************************************
'***** Switches                                                         	    ****
'***********************************************************************************

Sub sw11_Hit()
    STHit 11
End Sub

Sub sw12_Hit()
    STHit 12
End Sub

Sub sw13_Hit()
    STHit 13
End Sub

Sub sw15_Hit()
    STHit 15
End Sub

Sub sw16_Hit()
    STHit 16
End Sub

Sub sw17_Hit()
    STHit 17
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
    DTRaise 1
    lightCtrl.pulse l01, 3
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