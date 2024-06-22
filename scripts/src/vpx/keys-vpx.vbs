
'*******************************************
'  Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal keycode)
	

	If glf_gameStarted = True Then
		DebugShotTableKeyDownCheck keycode
		
		If keycode = LeftFlipperKey Then
			FlipperActivate LeftFlipper, LFPress
			'FlipperActivate LeftFlipper1, LFPress
			SolLFlipper True	'This would be called by the solenoid callbacks if using a ROM
			If glf_gameStarted = True Then 
				DispatchPinEvent SWITCH_LEFT_FLIPPER_DOWN, Null
			End If
		End If
		
		If keycode = RightFlipperKey Then
			FlipperActivate RightFlipper, RFPress
			SolRFlipper True	'This would be called by the solenoid callbacks if using a ROM
			UpRightFlipper.RotateToEnd
			If glf_gameStarted = True Then 
				DispatchPinEvent SWITCH_RIGHT_FLIPPER_DOWN, Null
			End If
		End If
		
		If keycode = PlungerKey Then
			Plunger.Pullback
			SoundPlungerPull
		End If
		
		If keycode = LeftTiltKey Then
			Nudge 90, 1
			SoundNudgeLeft
		End If
		If keycode = RightTiltKey Then
			Nudge 270, 1
			SoundNudgeRight
		End If
		If keycode = CenterTiltKey Then
			Nudge 0, 1
			SoundNudgeCenter
		End If
		If keycode = MechanicalTilt Then
			SoundNudgeCenter() 'Send the Tilting command to the ROM (usually by pulsing a Switch), or run the tilting code for an orginal table
		End If
	End If

	Glf_KeyDown keycode
End Sub


Sub Table1_KeyUp(ByVal keycode)
	
	If glf_gameStarted = True Then
		DebugShotTableKeyUpCheck keycode
		
		If KeyCode = PlungerKey Then
			Plunger.Fire
			If BIPL = 1 Then
				SoundPlungerReleaseBall()   'Plunger release sound when there is a ball in shooter lane
			Else
				SoundPlungerReleaseNoBall() 'Plunger release sound when there is no ball in shooter lane
			End If
		End If
		
		If keycode = LeftFlipperKey Then
			FlipperDeActivate LeftFlipper, LFPress
			'FlipperDeActivate LeftFlipper1, LFPress
			SolLFlipper False   'This would be called by the solenoid callbacks if using a ROM
		End If
		
		If keycode = RightFlipperKey Then
			UpRightFlipper.RotateToStart
			FlipperDeActivate RightFlipper, RFPress
			SolRFlipper False   'This would be called by the solenoid callbacks if using a ROM
		End If
	End If

	Glf_KeyUp keycode
End Sub
