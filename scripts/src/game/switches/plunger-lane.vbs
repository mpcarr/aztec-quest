
'******************************************************
'*****  Plunger Lane                               ****
'******************************************************

Sub BIPL_Hit()
    BIPL = True
    If autoPlunge = True Then
        AutoPlungerDelay.Interval = 300
	    AutoPlungerDelay.Enabled = True
    End If
End Sub

Sub BIPL_Top_Hit()
    BIPL = False
    autoPlunge = False
    If GetPlayerState(BALL_SAVE_ENABLED) = True Then
        EnableBallSaver 10
        SetPlayerState BALL_SAVE_ENABLED, False
    End If
End Sub

Sub AutoPlungerDelay_Timer
	plungerIM.Strength = 45
	plungerIM.AutoFire
	AutoPlungerDelay.Enabled = False
End Sub


'****************************
' Auto Plunge Ball
' Event Listeners:  
    AddPinEventListener BALL_SAVE,  "AutoPlungeBall"
    AddPinEventListener ADD_BALL,   "AutoPlungeBall"
'
'*****************************
Sub AutoPlungeBall()
    If BIPL = False And swTrough1.BallCntOver = 1 Then
        ReleaseBall()
        autoPlunge = True
    Else
        ballsInQ = ballsInQ + 1
        BallReleaseTimer.Enabled = True
    End If
End Sub

Dim ballsInQ : ballsInQ = 0
Sub BallReleaseTimer_Timer()
    If BIPL = False And ballsInQ > 0 AND swTrough1.BallCntOver = 1 Then
        ReleaseBall()
        autoPlunge = True
        ballsInQ = ballsInQ - 1
        If ballsInQ = 0 Then
            BallReleaseTimer.Enabled = False
        End If
    End If
End Sub
