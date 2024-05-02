
'****************************************************************
'  Slingshots
'****************************************************************

Dim RStep, Lstep
LStep = 4
RStep = 4
Sub RightSlingShot_Slingshot
	RS.VelocityCorrect(ActiveBall)
	RStep = 0
	RandomSoundSlingshotRight ActiveBall
	'DOF 104,DOFPulse
	'DOF 202,DOFPulse
	RightSlingShot.TimerInterval = 17
	RightSlingShot.TimerEnabled = 1
End Sub

Sub RightSlingShot_Timer
	Dim x1, x2, y: x1 = True:x2 = False:y = -20
	Select Case RStep
		Case 3:x1 = False:x2 = True:y = -10 :
		Case 4:x1 = False:x2 = False:y = 0:RightSlingShot.TimerEnabled = 0 
	End Select
	Dim x, BL	
	For Each BL in BP_RSling1 : BL.Visible = x1: Next
	For Each BL in BP_RSling2 : BL.Visible = x2: Next
	For Each BL in BP_REMK : BL.transx = -y: Next	
	RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	LS.VelocityCorrect(ActiveBall)
	RandomSoundSlingshotLeft ActiveBall
	'DOF 103,DOFPulse
	'DOF 201,DOFPulse
	LStep = 0
	LeftSlingShot.TimerInterval = 17
	LeftSlingShot.TimerEnabled = 1
End Sub


Sub LeftSlingShot_Timer
	Dim x1, x2, y: x1 = True:x2 = False:y = -20
	Select Case LStep
		Case 3:x1 = False:x2 = True:y = -10 : 
		Case 4:x1 = False:x2 = False:y = 0:LeftSlingShot.TimerEnabled = 0
	End Select

	Dim x, BL	
	For Each BL in BP_LSling1 : BL.Visible = x1: Next
	For Each BL in BP_LSling2 : BL.Visible = x2: Next
	For Each BL in BP_LEMK : BL.transx = -y: Next		
	LStep = LStep + 1
End Sub