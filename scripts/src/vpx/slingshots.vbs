
'****************************************************************
'  Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShot_Slingshot
	RS.VelocityCorrect(ActiveBall)
	Addscore 10
'	RSling1.Visible = 1
'	Sling1.TransY =  - 20   'Sling Metal Bracket
	RStep = 0
	RightSlingShot.TimerEnabled = 1
	RightSlingShot.TimerInterval = 10
'   vpmTimer.PulseSw 52	'Slingshot Rom Switch
	'RandomSoundSlingshotRight zCol_Rubber_Post043
End Sub

Sub RightSlingShot_Timer
	Select Case RStep
		Case 3
'		RSLing1.Visible = 0
'		RSLing2.Visible = 1
'		Sling1.TransY =  - 10
		Case 4
'		RSLing2.Visible = 0
'		Sling1.TransY = 0
		RightSlingShot.TimerEnabled = 0
	End Select
	RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	LS.VelocityCorrect(ActiveBall)
	Addscore 10
'	LSling1.Visible = 1
'	Sling2.TransY =  - 20   'Sling Metal Bracket
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	LeftSlingShot.TimerInterval = 10
'   vpmTimer.PulseSw 51	'Slingshot Rom Switch
	'RandomSoundSlingshotLeft zCol_Rubber_Post037
End Sub

Sub LeftSlingShot_Timer
	Select Case LStep
		Case 3
'		LSLing1.Visible = 0
'		LSLing2.Visible = 1
'		Sling2.TransY =  - 10
		Case 4
'		LSLing2.Visible = 0
'		Sling2.TransY = 0
'		LeftSlingShot.TimerEnabled = 0
	End Select
	LStep = LStep + 1
End Sub

Sub TestSlingShot_Slingshot
	TS.VelocityCorrect(ActiveBall)
End Sub
