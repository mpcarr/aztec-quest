
'*******************************************
'  Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this, 
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to 
'	 the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 100 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


Dim ballInReleasePostion : ballInReleasePostion = False
'TROUGH 
Sub swTrough1_Hit
	MPFController.Switch("0-0-3")=1
	ballInReleasePostion = True
	UpdateTrough
End Sub
Sub swTrough1_UnHit
	MPFController.Switch("0-0-3")=0
	ballInReleasePostion = False
	UpdateTrough
End Sub
Sub swTrough2_Hit
	MPFController.Switch("0-0-8")=1
	UpdateTrough
End Sub
Sub swTrough2_UnHit
	MPFController.Switch("0-0-8")=0
	UpdateTrough
End Sub
Sub swTrough3_Hit
	MPFController.Switch("0-0-9")=1
	UpdateTrough
End Sub
Sub swTrough3_UnHit
	MPFController.Switch("0-0-9")=0
	UpdateTrough
End Sub
Sub swTrough4_Hit
	MPFController.Switch("0-0-10")=1
	UpdateTrough
End Sub
Sub swTrough4_UnHit
	MPFController.Switch("0-0-10")=0
	UpdateTrough
End Sub
Sub swTrough5_Hit
	MPFController.Switch("0-0-11")=1
	UpdateTrough
End Sub
Sub swTrough5_UnHit
	MPFController.Switch("0-0-11")=0
	UpdateTrough
End Sub

Sub UpdateTrough
	UpdateTroughTimer.Interval = 100
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
	If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	Me.Enabled = 0
End Sub
