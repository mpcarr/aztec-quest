
Sub sw39001_Hit()
    set KickerBall39 = activeball
    SoundSaucerLock()
    sw39001.TimerEnabled = True
    debug.print("hitsw39")
End Sub
Sub sw39001_Timer()
    debug.print("kicksw39")
	sw39001.TimerEnabled = False
    SoundSaucerKick 1, sw39001
    KickBall KickerBall39, 0, 0, 60, 30
End Sub


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

