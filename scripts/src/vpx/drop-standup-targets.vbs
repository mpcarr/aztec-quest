'******************************************************
' 	ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
' 	DoDTAnim
'	DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also 
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object. 
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
	Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped
  
	Public Property Get Primary(): Set Primary = m_primary: End Property
	Public Property Let Primary(input): Set m_primary = input: End Property
  
	Public Property Get Secondary(): Set Secondary = m_secondary: End Property
	Public Property Let Secondary(input): Set m_secondary = input: End Property
  
	Public Property Get Prim(): Set Prim = m_prim: End Property
	Public Property Let Prim(input): Set m_prim = input: End Property
  
	Public Property Get Sw(): Sw = m_sw: End Property
	Public Property Let Sw(input): m_sw = input: End Property
  
	Public Property Get Animate(): Animate = m_animate: End Property
	Public Property Let Animate(input): m_animate = input: End Property
  
	Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
	Public Property Let IsDropped(input): m_isDropped = input: End Property
  
	Public default Function init(primary, secondary, prim, sw, animate, isDropped)
	  Set m_primary = primary
	  Set m_secondary = secondary
	  Set m_prim = prim
	  m_sw = sw
	  m_animate = animate
	  m_isDropped = isDropped
  
	  Set Init = Me
	End Function
End Class
  
'Define a variable for each drop target
Dim DT01, DT02, DT03, DT04, DT05, DT06, DT07, DT08, DT09, DT10, DT38, DT40, DT45, DT46, DT47

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:	primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:	   primitive target used for visuals and animation
'				   IMPORTANT!!!
'				   rotz must be used for orientation
'				   rotx to bend the target back
'				   transz to move it up and down
'				   the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:	 ROM switch number
'   animate:	Array slot for handling the animation instrucitons, set to 0
'				   Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
'					Use the function DTDropped(switchid) to check a target's drop status.

'Set DT38 = (new DropTarget)(sw38, sw38a, BM_sw38, 38, 0, False)
'Set DT40 = (new DropTarget)(sw40, sw40a, BM_sw40, 40, 0, False)
'Set DT45 = (new DropTarget)(sw45, sw45a, BM_sw45, 45, 0, False)
Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True) 
'Set DT02 = (new DropTarget)(sw02, sw02a, BM_sw02, 2, 0, False) 
Set DT04 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False)
Set DT05 = (new DropTarget)(sw05, sw05a, BM_sw05, 5, 0, False)
Set DT06 = (new DropTarget)(sw06, sw06a, BM_sw06, 6, 0, False)
'Set DT07 = (new DropTarget)(sw07, sw07a, BM_sw07, 7, 0, False)
Set DT08 = (new DropTarget)(sw08, sw08a, BM_sw08, 8, 0, False)
Set DT09 = (new DropTarget)(sw09, sw09a, BM_sw09, 9, 0, False)
Set DT10 = (new DropTarget)(sw10, sw10a, BM_sw10, 10, 0, False)
'Set DT46 = (new DropTarget)(sw46, sw46a, BM_sw46, 46, 0, False)
'Set DT47 = (new DropTarget)(sw47, sw47a, BM_sw47, 47, 0, False)

Dim DTArray
DTArray = Array(DT01,DT04, DT05, DT06, DT08, DT09, DT10)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 80 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 80 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

  

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
	Dim i
	i = DTArrayID(switch)
	
	PlayTargetSound
	DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
	If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
		DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
	End If
	DoDTAnim
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate =  - 1
	DoDTAnim
End Sub

Sub DTDrop(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate = 1
	DoDTAnim
End Sub

Function DTArrayID(switch)
	Dim i
	For i = 0 To UBound(DTArray)
		If DTArray(i).sw = switch Then
			DTArrayID = i
			Exit Function
		End If
	Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
	Dim rangle,bangle,calc1, calc2, calc3
	rangle = (angle - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	
	calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
	calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)
	
	aBall.velx = calc1 * Cos(rangle) + calc2
	aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
	Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
	rangle = (dtprim.rotz - 90) * 3.1416 / 180
	rangle2 = dtprim.rotz * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)
	
	Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
	Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)
	
	cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)
	
	perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	If perpvel > 0 And  perpvelafter <= 0 Then
		If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
			DTCheckBrick = 3
		Else
			DTCheckBrick = 1
		End If
	ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		DTCheckBrick = 4
	Else
		DTCheckBrick = 0
	End If
End Function

Sub DoDTAnim()
	Dim i
	For i = 0 To UBound(DTArray)
		DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
	Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
	Dim transz, switchid
	Dim animtime, rangle
	
	switchid = switch
	
	Dim ind
	ind = DTArrayID(switchid)
	
	rangle = prim.rotz * PI / 180
	
	DTAnimate = animate
	
	If animate = 0 Then
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	ElseIf primary.uservalue = 0 Then
		primary.uservalue = GameTime
	End If
	
	animtime = GameTime - primary.uservalue
	
	If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
		primary.collidable = 0
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
		DTAnimate = animate
		Exit Function
	ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
		primary.collidable = 0
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
		animate = 2
		SoundDropTargetDrop prim
	End If
	
	If animate = 2 Then
		transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
		If prim.transz >  - DTDropUnits  Then
			prim.transz = transz
		End If
		
		prim.rotx = DTMaxBend * Cos(rangle) / 2
		prim.roty = DTMaxBend * Sin(rangle) / 2
		
		If prim.transz <= - DTDropUnits Then
			prim.transz =  - DTDropUnits
			secondary.collidable = 0
			DTArray(ind).isDropped = True 'Mark target as dropped
			If UsingROM Then
				controller.Switch(Switchid mod 100) = 1
			Else
				DTAction switchid
			End If
			primary.uservalue = 0
			DTAnimate = 0
			Exit Function
		Else
			DTAnimate = 2
			Exit Function
		End If
	End If
	
	If animate = 3 And animtime < DTDropDelay Then
		primary.collidable = 0
		secondary.collidable = 1
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
	ElseIf animate = 3 And animtime > DTDropDelay Then
		primary.collidable = 1
		secondary.collidable = 0
		prim.rotx = 0
		prim.roty = 0
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	End If
	
	If animate =  - 1 Then
		transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1
		
		If prim.transz =  - DTDropUnits Then
			Dim b
			'Dim gBOT
			'gBOT = GetBalls
			
			For b = 0 To UBound(gBOT)
				If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
					gBOT(b).velz = 20
				End If
			Next
		End If
		
		If prim.transz < 0 Then
			prim.transz = transz
		ElseIf transz > 0 Then
			prim.transz = transz
		End If
		
		If prim.transz > DTDropUpUnits Then
			DTAnimate =  - 2
			prim.transz = DTDropUpUnits
			prim.rotx = 0
			prim.roty = 0
			primary.uservalue = GameTime
		End If
		primary.collidable = 0
		secondary.collidable = 1
		DTArray(ind).isDropped = False 'Mark target as not dropped
		If UsingROM Then controller.Switch(Switchid mod 100) = 0
	End If
	
	If animate =  - 2 And animtime > DTRaiseDelay Then
		prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
		If prim.transz < 0 Then
			prim.transz = 0
			primary.uservalue = 0
			DTAnimate = 0
			
			primary.collidable = 1
			secondary.collidable = 0
		End If
	End If
End Function

Function DTDropped(switchid)
	Dim ind
	ind = DTArrayID(switchid)
	
	DTDropped = DTArray(ind).isDropped
End Function

Sub DTAction(switchid)
	Select Case switchid
		
	End Select
End Sub

Sub UpdateTargets

	If DTDropped(1) = True Then
		BM_pantherLid.RotX = -6
	Else
		BM_pantherLid.RotX = 0
	End If
	BM_pantherLid.transz = BM_sw01.transz
	BM_pantherSupport.transz = BM_sw01.transz
End Sub


'******************************************************
'****  END DROP TARGETS
'******************************************************
  


'******************************************************
'	ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Class StandupTarget
	Private m_primary, m_prim, m_sw, m_animate
  
	Public Property Get Primary(): Set Primary = m_primary: End Property
	Public Property Let Primary(input): Set m_primary = input: End Property
  
	Public Property Get Prim(): Set Prim = m_prim: End Property
	Public Property Let Prim(input): Set m_prim = input: End Property
  
	Public Property Get Sw(): Sw = m_sw: End Property
	Public Property Let Sw(input): m_sw = input: End Property
  
	Public Property Get Animate(): Animate = m_animate: End Property
	Public Property Let Animate(input): m_animate = input: End Property
  
	Public default Function init(primary, prim, sw, animate)
	  Set m_primary = primary
	  Set m_prim = prim
	  m_sw = sw
	  m_animate = animate
  
	  Set Init = Me
	End Function
End Class

'Define a variable for each stand-up target
Dim ST11, ST12, ST13

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:	vp target to determine target hit
'   prim:	   primitive target used for visuals and animation
'				   IMPORTANT!!!
'				   transy must be used to offset the target animation
'   switch:	 ROM switch number
'   animate:	Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts


Set ST11 = (new StandupTarget)(sw11, BM_sw11 ,11 , 0)
'Set ST12 = (new StandupTarget)(sw12, psw12,12, 0)
'Set ST13 = (new StandupTarget)(sw13, psw13,13, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array()

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2	  'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'				STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
	Dim i
	i = STArrayID(switch)
	
	PlayTargetSound
	STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)
	
	If STArray(i).animate <> 0 Then
		DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 To UBound(STArray)
		If STArray(i).sw = switch Then
			STArrayID = i
			Exit Function
		End If
	Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
	Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
	rangle = (target.orientation - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)
	
	perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	If perpvel > 0 And  perpvelafter <= 0 Then
		STCheckHit = 1
	ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		STCheckHit = 1
	Else
		STCheckHit = 0
	End If
End Function

Sub DoSTAnim()
	Dim i
	For i = 0 To UBound(STArray)
		STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
	Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
	Dim animtime
	
	STAnimate = animate
	
	If animate = 0  Then
		primary.uservalue = 0
		STAnimate = 0
		Exit Function
	ElseIf primary.uservalue = 0 Then
		primary.uservalue = GameTime
	End If
	
	animtime = GameTime - primary.uservalue
	
	If animate = 1 Then
		primary.collidable = 0
		prim.transy =  - STMaxOffset
		If UsingROM Then
			vpmTimer.PulseSw switch mod 100
		Else
			STAction switch
		End If
		STAnimate = 2
		Exit Function
	ElseIf animate = 2 Then
		prim.transy = prim.transy + STAnimStep
		If prim.transy >= 0 Then
			prim.transy = 0
			primary.collidable = 1
			STAnimate = 0
			Exit Function
		Else
			STAnimate = 2
		End If
	End If
End Function


Sub STAction(Switch)
	
End Sub

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************
  

'******************************************************
'  DROP TARGET
'  SUPPORTING FUNCTIONS 
'******************************************************


' Used for drop targets
'*** Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy)
	Dim AB, BC, CD, DA
	AB = (bx*py) - (by*px) - (ax*py) + (ay*px) + (ax*by) - (ay*bx)
	BC = (cx*py) - (cy*px) - (bx*py) + (by*px) + (bx*cy) - (by*cx)
	CD = (dx*py) - (dy*px) - (cx*py) + (cy*px) + (cx*dy) - (cy*dx)
	DA = (ax*py) - (ay*px) - (dx*py) + (dy*px) + (dx*ay) - (dy*ax)

	If (AB <= 0 AND BC <=0 AND CD <= 0 AND DA <= 0) Or (AB >= 0 AND BC >=0 AND CD >= 0 AND DA >= 0) Then
		InRect = True
	Else
		InRect = False       
	End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
    Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
    Dim rotxy
    rotxy = RotPoint(ax,ay,angle)
    rax = rotxy(0)+px : ray = rotxy(1)+py
    rotxy = RotPoint(bx,by,angle)
    rbx = rotxy(0)+px : rby = rotxy(1)+py
    rotxy = RotPoint(cx,cy,angle)
    rcx = rotxy(0)+px : rcy = rotxy(1)+py
    rotxy = RotPoint(dx,dy,angle)
    rdx = rotxy(0)+px : rdy = rotxy(1)+py

    InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function