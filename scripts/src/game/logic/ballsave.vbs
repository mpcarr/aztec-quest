
'******************************************************
'*****  Ball Saver                                 ****
'******************************************************

dim inGracePeriod : inGracePeriod = False
Sub EnableBallSaver(seconds)
	BallSaverTimerExpired.Interval = (1000 * seconds)
	BallSaverTimerExpired.Enabled = True
    ballSaver = True
    inGracePeriod = False   
End Sub

Sub BallSaverTimerExpired_Timer()
    If inGracePeriod = False Then
        BallSaverTimerExpired.Interval = 3000
        inGracePeriod = True
    Else
        BallSaverTimerExpired.Enabled = False
        ballSaver = False
    End If
End Sub