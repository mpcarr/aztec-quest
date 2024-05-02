
'***********************************************************************************
'***** Switches                                                         	    ****
'***********************************************************************************

Sub sw_plunger_Hit()
    MPFController.Switch("0-0-28") = 1
End Sub

Sub sw_plunger_Unhit()
    MPFController.Switch("0-0-28") = 0
End Sub

Sub sw39_Hit()
    SoundSaucerLock()
    Set KickerBallCave = ActiveBall
    MPFController.Switch("0-0-30") = 1
End Sub

Sub sw39_Unhit()
    KickerBallCave = Null
    MPFController.Switch("0-0-30") = 0
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

Sub sw99_Hit
    MPFController.Switch("0-0-43") = 1
End Sub
Sub sw99_UnHit
    MPFController.Switch("0-0-43") = 0
End Sub

Sub DTAction(switchid, enabled)
	Select Case switchid
		case 4:
			MPFController.Switch("0-0-31") = enabled
        case 5:
			MPFController.Switch("0-0-32") = enabled
        case 6:
			MPFController.Switch("0-0-33") = enabled
        case 8:
			MPFController.Switch("0-0-34") = enabled
        case 9:
			MPFController.Switch("0-0-35") = enabled
        case 10:
			MPFController.Switch("0-0-36") = enabled
	End Select
End Sub

Sub STAction(switchid, enabled)
	Select Case switchid
		case 11:
			MPFController.Switch("0-0-37") = enabled
        case 12:
			MPFController.Switch("0-0-38") = enabled
        case 13:
			MPFController.Switch("0-0-39") = enabled
        case 15:
			MPFController.Switch("0-0-40") = enabled
        case 16:
			MPFController.Switch("0-0-41") = enabled
        case 17:
			MPFController.Switch("0-0-42") = enabled
	End Select
End Sub