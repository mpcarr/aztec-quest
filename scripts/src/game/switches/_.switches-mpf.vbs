
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