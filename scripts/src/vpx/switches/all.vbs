
Sub Kicker001_Hit()
    set KickerBall40 = activeball
    SoundSaucerLock()
    Kicker001.TimerEnabled = True
    debug.print("hitsw40")
End Sub
Sub Kicker001_Timer()
    debug.print("kicksw40")
	Kicker001.TimerEnabled = False
    SoundSaucerKick 1, Kicker001
    Kicker001.Kick 0, 80, 1.36
    'KickBall KickerBall40, 0, 0, 100, 50
End Sub

