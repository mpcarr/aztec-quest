'*******************************************
'  Timers
'*******************************************

Sub GameTimer_Timer() 'The game timer interval; should be 10 ms
	Cor.Update	  'update ball tracking (this sometimes goes in the RDampen_Timer sub)
	RollingUpdate   'update rolling sounds
	DoSTAnim		'handle stand up target animations
	DoDTAnim
	UpdateTargets
	Dim ii
	Dim ChgSol : ChgSol = MPFController.ChangedSolenoids
	if not isempty(ChgSol) Then	
		for ii=0 to UBound(ChgSol)
			debugLog.WriteToLog "coils", "coil: " &  ChgSol(ii,0) & ". State: " & ChgSol(ii,1)
			If ChgSol(ii,0) = "0-0-6" and ChgSol(ii,1) Then
				PlungerDevice.Eject
			End If
		Next
	end If
End Sub

Sub EventTimer_Timer()
	DelayTick
End Sub

Dim FrameTime, InitFrameTime
InitFrameTime = 0
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
	FrameTime = gametime - InitFrameTime
	InitFrameTime = gametime	'Count frametime
	FlipperVisualUpdate		 'update flipper shadows and primitives
	If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
End Sub

Sub LampTimer_Timer()
	lightCtrl.Update
End Sub
