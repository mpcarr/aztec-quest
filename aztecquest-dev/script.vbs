'Aztec Quest by Flux


Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs in order to run this table, available in the vp10 package"
On Error GoTo 0

'*******************************************
'  User Options
'*******************************************

'----- DMD Options -----
Const UseFlexDMD = 0			'0 = no FlexDMD, 1 = enable FlexDMD
Const FlexONPlayfield = False	'False = off, True=DMD on playfield ( vrroom overrides this )

'----- Shadow Options -----
Const DynamicBallShadowsOn = 1	'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
Const AmbientBallShadowOn = 1	'0 = Static shadow under ball ("flasher" image, like JP's), 1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that behaves like a true shadow!, 2 = flasher image shadow, but it moves like ninuzzu's

'----- General Sound Options -----
Const VolumeDial = 0.8			'Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Const BallRollVolume = 0.5		'Level of ball rolling volume. Value between 0 and 1
Const RampRollVolume = 0.5		'Level of ramp rolling volume. Value between 0 and 1

'----- VR Room -----
Const VRRoom = 0				'0 - VR Room off, 1 - Minimal Room, 2 - Ultra Minimal Room

Dim haspup : haspup = false

Dim GameTilted : GameTilted = False
Dim gameDebugger : Set gameDebugger = new AdvGameDebugger
'*******************************************
'  Constants and Global Variables
'*******************************************

Const UsingROM = False		'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.

Const BallSize = 50			'Ball diameter in VPX units; must be 50
Const BallMass = 1			'Ball mass must be 1
Const tnob = 7				'Total number of balls the table can hold
Const lob = 0				'Locked balls
Dim gBOT
Const cGameName = "aztecquest"	'The unique alphanumeric name for this table

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP						'Balls in play
BIP = 0
Dim BIPL					'Ball in plunger lane
BIPL = False

'Const IMPowerSetting = 50 			'Plunger Power
'Const IMTime = 1.1        			'Time in seconds for Full Plunge
'Dim plungerIM

Dim gilvl : gilvl = 0  'General Illumination light state tracked for Dynamic Ball Shadows

'*******************************************
'  Table Initialization and Exiting
'*******************************************

Sub Table1_Init
	Glf_Init()

	Dim i
	
	LeftSlingShot_Timer
	RightSlingShot_Timer
	lightCtrl.SyncLightMapColors

	ConfigureDevices
	
	DTDrop 1
	DTDrop 2
	
	Dim xx
	' Add balls to shadow dictionary
	For Each xx In gBOT
		bsDict.Add xx.ID, bsNone
	Next
	
	' Make drop target shadows visible
	For Each xx In ShadowDT
		xx.visible = True
	Next
End Sub


Sub Table1_Exit
	'gameDebugger.Disconnect
	Glf_Exit()
End Sub


'***************************************************************
'****  VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'***************************************************************

'****** INSTRUCTIONS please read ******

'****** Part A:  Table Elements ******
'
' Import the "bsrtx7" and "ballshadow" images
' Import the shadow materials file (3 sets included) (you can also export the 3 sets from this table to create the same file)
' Copy in the BallShadowA flasher set and the sets of primitives named BallShadow#, RtxBallShadow#, and RtxBall2Shadow#
'	* Count from 0 up, with at least as many objects each as there can be balls, including locked balls.  You'll get an "eval" warning if tnob is higher
'	* Warning:  If merging with another system (JP's ballrolling), you may need to check tnob math and add an extra BallShadowA# flasher (out of range error)
' Ensure you have a timer with a -1 interval that is always running
' Set plastic ramps DB to *less* than the ambient shadows (-11000) if you want to see the pf shadow through the ramp
' Place triggers at the start of each ramp *type* (solid, clear, wire) and one at the end if it doesn't return to the base pf
'	* These can share duties as triggers for RampRolling sounds

' Create a collection called DynamicSources that includes all light sources you want to cast ball shadows
' It's recommended that you be selective in which lights go in this collection, as there are limitations:
' 1. The shadows can "pass through" solid objects and other light sources, so be mindful of where the lights would actually able to cast shadows
' 2. If there are more than two equidistant sources, the shadows can suddenly switch on and off, so places like top and bottom lanes need attention
' 3. At this time the shadows get the light on/off from tracking gilvl, so if you have lights you want shadows for that are on at different times you will need to either:
'	a) remove this restriction (shadows think lights are always On)
'	b) come up with a custom solution (see TZ example in script)
' After confirming the shadows work in general, use ball control to move around and look for any weird behavior

'****** End Part A:  Table Elements ******


'****** Part B:  Code and Functions ******

' *** Timer sub
' The "DynamicBSUpdate" sub should be called by a timer with an interval of -1 (framerate)
' Example timer sub:

'Sub FrameTimer_Timer()
'	If DynamicBallShadowsOn Or AmbientBallShadowOn Then DynamicBSUpdate 'update ball shadows
'End Sub

' *** These are usually defined elsewhere (ballrolling), but activate here if necessary
'Const tnob = 10 ' total number of balls
'Const lob = 0	'locked balls on start; might need some fiddling depending on how your locked balls are done
'Dim tablewidth: tablewidth = Table1.width
'Dim tableheight: tableheight = Table1.height

' *** User Options - Uncomment here or move to top for easy access by players
'----- Shadow Options -----
'Const DynamicBallShadowsOn = 1		'0 = no dynamic ball shadow ("triangles" near slings and such), 1 = enable dynamic ball shadow
'Const AmbientBallShadowOn = 1		'0 = Static shadow under ball ("flasher" image, like JP's)
'									'1 = Moving ball shadow ("primitive" object, like ninuzzu's) - This is the only one that shows up on the pf when in ramps and fades when close to lights!
'									'2 = flasher image shadow, but it moves like ninuzzu's

' *** The following segment goes within the RollingUpdate sub, so that if Ambient...=0 and Dynamic...=0 the entire DynamicBSUpdate sub can be skipped for max performance
' ** Change gBOT to BOT if using existing getballs code
' ** Double commented lines commonly found there included for reference:

''	' stop the sound of deleted balls
''	For b = UBound(gBOT) + 1 to tnob
'		If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
''		...rolling(b) = False
''		...StopSound("BallRoll_" & b)
''	Next
''
'' ...rolling and drop sounds...
''
''		If DropCount(b) < 5 Then
''			DropCount(b) = DropCount(b) + 1
''		End If
''
'		' "Static" Ball Shadows
'		If AmbientBallShadowOn = 0 Then
'			BallShadowA(b).visible = 1
'			BallShadowA(b).X = gBOT(b).X + offsetX
'			If gBOT(b).Z > 30 Then
'				BallShadowA(b).height=gBOT(b).z - BallSize/4 + b/1000	'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
'				BallShadowA(b).Y = gBOT(b).Y + offsetY + BallSize/10
'			Else
'				BallShadowA(b).height=gBOT(b).z - BallSize/2 + 1.04 + b/1000
'				BallShadowA(b).Y = gBOT(b).Y + offsetY
'			End If
'		End If

' *** Place this inside the table init, just after trough balls are added to gBOT
' 
' Add balls to shadow dictionary
'	For Each xx in gBOT
'		bsDict.Add xx.ID, bsNone
'	Next

' *** Example RampShadow trigger subs:

'Sub ClearRampStart_hit()
'	bsRampOnClear			'Shadow on ramp and pf below
'End Sub

'Sub SolidRampStart_hit()
'	bsRampOn				'Shadow on ramp only
'End Sub

'Sub WireRampStart_hit()
'	bsRampOnWire			'Shadow only on pf
'End Sub

'Sub RampEnd_hit()
'	bsRampOff ActiveBall.ID	'Back to default shadow behavior
'End Sub

'
'' *** Required Functions, enable these if they are not already present elswhere in your table
'Function max(a,b)
'	If a > b Then
'		max = a
'	Else
'		max = b
'	End If
'End Function

'Function Distance(ax,ay,bx,by)
'	Distance = SQR((ax - bx)^2 + (ay - by)^2)
'End Function

'Dim PI: PI = 4*Atn(1)

'Function Atn2(dy, dx)
'	If dx > 0 Then
'		Atn2 = Atn(dy / dx)
'	ElseIf dx < 0 Then
'		If dy = 0 Then 
'			Atn2 = pi
'		Else
'			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
'		end if
'	ElseIf dx = 0 Then
'		if dy = 0 Then
'			Atn2 = 0
'		else
'			Atn2 = Sgn(dy) * pi / 2
'		end if
'	End If
'End Function

'Function AnglePP(ax,ay,bx,by)
'	AnglePP = Atn2((by - ay),(bx - ax))*180/PI
'End Function

'****** End Part B:  Code and Functions ******


'****** Part C:  The Magic ******

' *** These define the appearance of shadows in your table	***

'Ambient (Room light source)
Const AmbientBSFactor = 0.9	 '0 to 1, higher is darker
Const AmbientMovement = 1	   '1+ higher means more movement as the ball moves left and right
Const offsetX = 0			   'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 5			   'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

'Dynamic (Table light sources)
Const DynamicBSFactor = 0.95	'0 to 1, higher is darker
Const Wideness = 20			 'Sets how wide the dynamic ball shadows can get (20 +5 thinness is technically most accurate for lights at z ~25 hitting a 50 unit ball)
Const Thinness = 5			  'Sets minimum as ball moves away from source

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objrtx1(7), objrtx2(7)
Dim objBallShadow(7)
Dim OnPF(7)
Dim BallShadowA
BallShadowA = Array (BallShadowA0,BallShadowA1,BallShadowA2,BallShadowA3,BallShadowA4,BallShadowA5,BallShadowA6,BallShadowA7)
Dim DSSources(30), numberofsources', DSGISide(30) 'Adapted for TZ with GI left / GI right

' *** The Shadow Dictionary
Dim bsDict
Set bsDict = CreateObject("Scripting.Dictionary")
Const bsNone = "None"
Const bsWire = "Wire"
Const bsRamp = "Ramp"
Const bsRampClear = "Clear"

'Initialization
DynamicBSInit

Sub DynamicBSInit()
	Dim iii, source
	
	'Prepare the shadow objects before play begins
	For iii = 0 To tnob - 1
		Set objrtx1(iii) = Eval("RtxBallShadow" & iii)
		objrtx1(iii).material = "RtxBallShadow" & iii
		objrtx1(iii).z = 1 + iii / 1000 + 0.01  'Separate z for layering without clipping
		objrtx1(iii).visible = 0
		
		Set objrtx2(iii) = Eval("RtxBall2Shadow" & iii)
		objrtx2(iii).material = "RtxBallShadow2_" & iii
		objrtx2(iii).z = 1 + iii / 1000 + 0.02
		objrtx2(iii).visible = 0
		
		Set objBallShadow(iii) = Eval("BallShadow" & iii)
		objBallShadow(iii).material = "BallShadow" & iii
		UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(iii).Z = 1 + iii / 1000 + 0.04
		objBallShadow(iii).visible = 0
		
		BallShadowA(iii).Opacity = 100 * AmbientBSFactor
		BallShadowA(iii).visible = 0
	Next
	
	iii = 0
	
	For Each Source In DynamicSources
		DSSources(iii) = Array(Source.x, Source.y)
		'   If Instr(Source.name , "Left") > 0 Then DSGISide(iii) = 0 Else DSGISide(iii) = 1	'Adapted for TZ with GI left / GI right
		iii = iii + 1
	Next
	numberofsources = iii
End Sub

Sub BallOnPlayfieldNow(onPlayfield, ballNum)	'Whether a ball is currently on the playfield. Only update certain things once, save some cycles
	If onPlayfield Then
		OnPF(ballNum) = True
		bsRampOff gBOT(ballNum).ID
		'   debug.print "Back on PF"
		UpdateMaterial objBallShadow(ballNum).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(ballNum).size_x = 5
		objBallShadow(ballNum).size_y = 4.5
		objBallShadow(ballNum).visible = 1
		BallShadowA(ballNum).visible = 0
		BallShadowA(ballNum).Opacity = 100 * AmbientBSFactor
	Else
		OnPF(ballNum) = False
		'   debug.print "Leaving PF"
	End If
End Sub

Sub DynamicBSUpdate
	Dim falloff 'Max distance to light sources, can be changed dynamically if you have a reason
	falloff = 150
	Dim ShadowOpacity1, ShadowOpacity2
	Dim s, LSd, iii
	Dim dist1, dist2, src1, src2
	Dim bsRampType
	'   Dim gBOT: gBOT=getballs	'Uncomment if you're destroying balls - Not recommended! #SaveTheBalls
	
	'Hide shadow of deleted balls
	For s = UBound(gBOT) + 1 To tnob - 1
		objrtx1(s).visible = 0
		objrtx2(s).visible = 0
		objBallShadow(s).visible = 0
		BallShadowA(s).visible = 0
	Next
	
	If UBound(gBOT) < lob Then Exit Sub 'No balls in play, exit
	
	'The Magic happens now
	For s = lob To UBound(gBOT)
		' *** Normal "ambient light" ball shadow
		'Layered from top to bottom. If you had an upper pf at for example 80 units and ramps even above that, your Elseif segments would be z>110; z<=110 And z>100; z<=100 And z>30; z<=30 And z>20; Else (under 20)
		
		'Primitive shadow on playfield, flasher shadow in ramps
		If AmbientBallShadowOn = 1 Then
			'** Above the playfield
			If gBOT(s).Z > 30 Then
				If OnPF(s) Then BallOnPlayfieldNow False, s		'One-time update
				bsRampType = getBsRampType(gBOT(s).id)
				'   debug.print bsRampType
				
				If Not bsRampType = bsRamp Then 'Primitive visible on PF
					objBallShadow(s).visible = 1
					objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
					objBallShadow(s).Y = gBOT(s).Y + offsetY
					objBallShadow(s).size_x = 5 * ((gBOT(s).Z + BallSize) / 80) 'Shadow gets larger and more diffuse as it moves up
					objBallShadow(s).size_y = 4.5 * ((gBOT(s).Z + BallSize) / 80)
					UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (30 / (gBOT(s).Z)),RGB(0,0,0),0,0,False,True,0,0,0,0
				Else 'Opaque, no primitive below
					objBallShadow(s).visible = 0
				End If
				
				If bsRampType = bsRampClear Or bsRampType = bsRamp Then 'Flasher visible on opaque ramp
					BallShadowA(s).visible = 1
					BallShadowA(s).X = gBOT(s).X + offsetX
					BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
					BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
					If bsRampType = bsRampClear Then BallShadowA(s).Opacity = 50 * AmbientBSFactor
				ElseIf bsRampType = bsWire Or bsRampType = bsNone Then 'Turn it off on wires or falling out of a ramp
					BallShadowA(s).visible = 0
				End If
				
				'** On pf, primitive only
			ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then
				If Not OnPF(s) Then BallOnPlayfieldNow True, s
				objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
				objBallShadow(s).Y = gBOT(s).Y + offsetY
				'   objBallShadow(s).Z = gBOT(s).Z + s/1000 + 0.04		'Uncomment (and adjust If/Elseif height logic) if you want the primitive shadow on an upper/split pf																																						 
				
				'** Under pf, flasher shadow only
			Else
				If OnPF(s) Then BallOnPlayfieldNow False, s
				objBallShadow(s).visible = 0
				BallShadowA(s).visible = 1
				BallShadowA(s).X = gBOT(s).X + offsetX
				BallShadowA(s).Y = gBOT(s).Y + offsetY
				BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
			End If
			
			'Flasher shadow everywhere
		ElseIf AmbientBallShadowOn = 2 Then
			If gBOT(s).Z > 30 Then 'In a ramp
				BallShadowA(s).X = gBOT(s).X + offsetX
				BallShadowA(s).Y = gBOT(s).Y + offsetY + BallSize / 10
				BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000 'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
			ElseIf gBOT(s).Z <= 30 And gBOT(s).Z > 20 Then 'On pf
				BallShadowA(s).visible = 1
				BallShadowA(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
				BallShadowA(s).Y = gBOT(s).Y + offsetY
				BallShadowA(s).height = 1.04 + s / 1000
			Else 'Under pf
				BallShadowA(s).X = gBOT(s).X + offsetX
				BallShadowA(s).Y = gBOT(s).Y + offsetY
				BallShadowA(s).height = gBOT(s).z - BallSize / 4 + s / 1000
			End If
		End If
		
		' *** Dynamic shadows
		If DynamicBallShadowsOn Then
			If gBOT(s).Z < 30 And gBOT(s).X < 850 Then 'Parameters for where the shadows can show, here they are not visible above the table (no upper pf) or in the plunger lane
				dist1 = falloff
				dist2 = falloff
				For iii = 0 To numberofsources - 1 'Search the 2 nearest influencing lights
					LSd = Distance(gBOT(s).x, gBOT(s).y, DSSources(iii)(0), DSSources(iii)(1)) 'Calculating the Linear distance to the Source
					If LSd < falloff And gilvl > 0 Then
						'   If LSd < dist2 And ((DSGISide(iii) = 0 And Lampz.State(100)>0) Or (DSGISide(iii) = 1 And Lampz.State(104)>0)) Then	'Adapted for TZ with GI left / GI right
						dist2 = dist1
						dist1 = LSd
						src2 = src1
						src1 = iii
					End If
				Next
				ShadowOpacity1 = 0
				If dist1 < falloff Then
					objrtx1(s).visible = 1
					objrtx1(s).X = gBOT(s).X
					objrtx1(s).Y = gBOT(s).Y
					'   objrtx1(s).Z = gBOT(s).Z - 25 + s/1000 + 0.01 'Uncomment if you want to add shadows to an upper/lower pf
					objrtx1(s).rotz = AnglePP(DSSources(src1)(0), DSSources(src1)(1), gBOT(s).X, gBOT(s).Y) + 90
					ShadowOpacity1 = 1 - dist1 / falloff
					objrtx1(s).size_y = Wideness * ShadowOpacity1 + Thinness
					UpdateMaterial objrtx1(s).material,1,0,0,0,0,0,ShadowOpacity1 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
				Else
					objrtx1(s).visible = 0
				End If
				ShadowOpacity2 = 0
				If dist2 < falloff Then
					objrtx2(s).visible = 1
					objrtx2(s).X = gBOT(s).X
					objrtx2(s).Y = gBOT(s).Y + offsetY
					'   objrtx2(s).Z = gBOT(s).Z - 25 + s/1000 + 0.02 'Uncomment if you want to add shadows to an upper/lower pf
					objrtx2(s).rotz = AnglePP(DSSources(src2)(0), DSSources(src2)(1), gBOT(s).X, gBOT(s).Y) + 90
					ShadowOpacity2 = 1 - dist2 / falloff
					objrtx2(s).size_y = Wideness * ShadowOpacity2 + Thinness
					UpdateMaterial objrtx2(s).material,1,0,0,0,0,0,ShadowOpacity2 * DynamicBSFactor ^ 3,RGB(0,0,0),0,0,False,True,0,0,0,0
				Else
					objrtx2(s).visible = 0
				End If
				If AmbientBallShadowOn = 1 Then
					'Fades the ambient shadow (primitive only) when it's close to a light
					UpdateMaterial objBallShadow(s).material,1,0,0,0,0,0,AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2)),RGB(0,0,0),0,0,False,True,0,0,0,0
				Else
					BallShadowA(s).Opacity = 100 * AmbientBSFactor * (1 - max(ShadowOpacity1, ShadowOpacity2))
				End If
			Else 'Hide dynamic shadows everywhere else, just in case
				objrtx2(s).visible = 0
				objrtx1(s).visible = 0
			End If
		End If
	Next
End Sub

' *** Ramp type definitions

Sub bsRampOnWire()
	If bsDict.Exists(ActiveBall.ID) Then
		bsDict.Item(ActiveBall.ID) = bsWire
	Else
		bsDict.Add ActiveBall.ID, bsWire
	End If
End Sub

Sub bsRampOn()
	If bsDict.Exists(ActiveBall.ID) Then
		bsDict.Item(ActiveBall.ID) = bsRamp
	Else
		bsDict.Add ActiveBall.ID, bsRamp
	End If
End Sub

Sub bsRampOnClear()
	If bsDict.Exists(ActiveBall.ID) Then
		bsDict.Item(ActiveBall.ID) = bsRampClear
	Else
		bsDict.Add ActiveBall.ID, bsRampClear
	End If
End Sub

Sub bsRampOff(idx)
	If bsDict.Exists(idx) Then
		bsDict.Item(idx) = bsNone
	End If
End Sub

Function getBsRampType(id)
	Dim retValue
	If bsDict.Exists(id) Then
		retValue = bsDict.Item(id)
	Else
		retValue = bsNone
	End If
	getBsRampType = retValue
End Function

'****************************************************************
'****  END VPW DYNAMIC BALL SHADOWS by Iakki, Apophis, and Wylte
'****************************************************************


'*****************************************************************************************************************************************
'  Advance Game Debugger by flux
'*****************************************************************************************************************************************
Class AdvGameDebugger

    Private m_advDebugger, m_connected

    Private Sub Class_Initialize()
        On Error Resume Next
        Set m_advDebugger = CreateObject("vpx_adv_debugger.VPXAdvDebugger")
        m_advDebugger.Connect()
        m_connected = True
        If Err Then Debug.print("Can't start advanced debugger") : m_connected = False
    End Sub

	Public Sub SendPlayerState(key, value)
		If m_connected Then
            m_advDebugger.SendPlayerState key, value
        End If
	End Sub

    Public Sub SendPinEvent(evt)
		If m_connected Then
            m_advDebugger.SendPinEvent evt
        End If
	End Sub

    Public Sub Disconnect()
        If m_connected Then
            m_advDebugger.Disconnect()
        End If
    End Sub
End Class

'*****************************************************************************************************************************************
'  Advance Game Debugger by flux
'*****************************************************************************************************************************************



'****************************************************************
' Section; Debug Shot Tester v3.2
'
' 1.  Raise/Lower outlanes and drain posts by pressing 2 key
' 2.  Capture and Launch ball, Press and hold one of the buttons (W, E, R, Y, U, I, P, A) below to capture ball by flipper.  Release key to shoot ball
' 3.  To change the test shot angles, press and hold a key and use Flipper keys to adjust the shot angle.  Shot angles are saved into the User direction as cgamename.txt
' 4.  Set DebugShotMode = 0 to disable debug shot test code.
'
' HOW TO INSTALL: Copy all debug* objects from Layer 2 to table and adjust. Copy the Debug Shot Tester code section to the script.
'	Add "DebugShotTableKeyDownCheck keycode" to top of Table1_KeyDown sub and add "DebugShotTableKeyUpCheck keycode" to top of Table1_KeyUp sub
'****************************************************************
Const DebugShotMode = 1 'Set to 0 to disable.  1 to enable
Dim DebugKickerForce
DebugKickerForce = 55

' Enable Disable Outlane and Drain Blocker Wall for debug testing
Dim DebugBLState
debug_BLW1.IsDropped = 1
debug_BLP1.Visible = 0
debug_BLR1.Visible = 0
debug_BLW2.IsDropped = 1
debug_BLP2.Visible = 0
debug_BLR2.Visible = 0
debug_BLW3.IsDropped = 1
debug_BLP3.Visible = 0
debug_BLR3.Visible = 0

Sub BlockerWalls
	DebugBLState = (DebugBLState + 1) Mod 4
	'	debug.print "BlockerWalls"
	PlaySound ("Start_Button")
	
	Select Case DebugBLState
		Case 0
		debug_BLW1.IsDropped = 1
		debug_BLP1.Visible = 0
		debug_BLR1.Visible = 0
		debug_BLW2.IsDropped = 1
		debug_BLP2.Visible = 0
		debug_BLR2.Visible = 0
		debug_BLW3.IsDropped = 1
		debug_BLP3.Visible = 0
		debug_BLR3.Visible = 0
		
		Case 1
		debug_BLW1.IsDropped = 0
		debug_BLP1.Visible = 1
		debug_BLR1.Visible = 1
		debug_BLW2.IsDropped = 0
		debug_BLP2.Visible = 1
		debug_BLR2.Visible = 1
		debug_BLW3.IsDropped = 0
		debug_BLP3.Visible = 1
		debug_BLR3.Visible = 1
		
		Case 2
		debug_BLW1.IsDropped = 0
		debug_BLP1.Visible = 1
		debug_BLR1.Visible = 1
		debug_BLW2.IsDropped = 0
		debug_BLP2.Visible = 1
		debug_BLR2.Visible = 1
		debug_BLW3.IsDropped = 1
		debug_BLP3.Visible = 0
		debug_BLR3.Visible = 0
		
		Case 3
		debug_BLW1.IsDropped = 1
		debug_BLP1.Visible = 0
		debug_BLR1.Visible = 0
		debug_BLW2.IsDropped = 1
		debug_BLP2.Visible = 0
		debug_BLR2.Visible = 0
		debug_BLW3.IsDropped = 0
		debug_BLP3.Visible = 1
		debug_BLR3.Visible = 1
	End Select
End Sub

Sub DebugShotTableKeyDownCheck (Keycode)
	'Cycle through Outlane/Centerlane blocking posts
	'-----------------------------------------------
	If Keycode = 3 Then
		BlockerWalls
	End If
	
	If DebugShotMode = 1 Then
		'Capture and launch ball:	
		'	Press and hold one of the buttons (W, E, R, T, Y, U, I, P) below to capture ball by flipper.  Release key to shoot ball
		'	To change the test shot angles, press and hold a key and use Flipper keys to adjust the shot angle.
		'--------------------------------------------------------------------------------------------
		If keycode = 17 Then 'W key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleW
		End If
		If keycode = 18 Then 'E key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleE
		End If
		If keycode = 19 Then 'R key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleR
		End If
		If keycode = 21 Then 'Y key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleY
		End If
		If keycode = 22 Then 'U key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleU
		End If
		If keycode = 23 Then 'I key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleI
		End If
		If keycode = 25 Then 'P key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleP
		End If
		If keycode = 30 Then 'A key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleA
		End If
		If keycode = 31 Then 'S key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleS
		End If
		If keycode = 33 Then 'F key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleF
		End If
		If keycode = 34 Then 'G key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleG
		End If
		
		If debugKicker.enabled = True Then		'Use Flippers to adjust angle while holding key
			If keycode = leftflipperkey Then
				debugKickAim.Visible = True
				TestKickerVar = TestKickerVar - 1
				Debug.print TestKickerVar
			ElseIf keycode = rightflipperkey Then
				debugKickAim.Visible = True
				TestKickerVar = TestKickerVar + 1
				Debug.print TestKickerVar
			End If
			debugKickAim.ObjRotz = TestKickerVar
		End If
	End If
End Sub


Sub DebugShotTableKeyUpCheck (Keycode)
	' Capture and launch ball:
	' Release to shoot ball. Set up angle and force as needed for each shot.  
	'--------------------------------------------------------------------------------------------
	If DebugShotMode = 1 Then
		If keycode = 17 Then 'W key
			TestKickAngleW = TestKickerVar
			debugKicker.kick TestKickAngleW, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 18 Then 'E key
			TestKickAngleE = TestKickerVar
			debugKicker.kick TestKickAngleE, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 19 Then 'R key
			TestKickAngleR = TestKickerVar
			debugKicker.kick TestKickAngleR, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 21 Then 'Y key
			TestKickAngleY = TestKickerVar
			debugKicker.kick TestKickAngleY, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 22 Then 'U key
			TestKickAngleU = TestKickerVar
			debugKicker.kick TestKickAngleU, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 23 Then 'I key
			TestKickAngleI = TestKickerVar
			debugKicker.kick TestKickAngleI, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 25 Then 'P key
			TestKickAngleP = TestKickerVar
			debugKicker.kick TestKickAngleP, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 30 Then 'A key
			TestKickAngleA = TestKickerVar
			debugKicker.kick TestKickAngleA, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 31 Then 'S key
			TestKickAngleS = TestKickerVar
			debugKicker.kick TestKickAngleS, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 33 Then 'F key
			TestKickAngleF = TestKickerVar
			debugKicker.kick TestKickAngleF, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 34 Then 'G key
			TestKickAngleG = TestKickerVar
			debugKicker.kick TestKickAngleG, DebugKickerForce
			debugKicker.enabled = False
		End If
		
		'		EXAMPLE CODE to set up key to cycle through 3 predefined shots
		'		If keycode = 17 Then	 'Cycle through all left target shots
		'			If TestKickerAngle = -28 then 
		'				TestKickerAngle = -24
		'			ElseIf TestKickerAngle = -24 Then
		'				TestKickerAngle = -19
		'			Else
		'				TestKickerAngle = -28
		'			End If
		'			debugKicker.kick TestKickerAngle, DebugKickerForce: debugKicker.enabled = false			 'W key	
		'		End If
		
	End If
	
	If (debugKicker.enabled = False And debugKickAim.Visible = True) Then 'Save Angle changes
		debugKickAim.Visible = False
		SaveTestKickAngles
	End If
End Sub

Dim TestKickerAngle, TestKickerAngle2, TestKickerVar, TeskKickKey, TestKickForce
Dim TestKickAngleWDefault, TestKickAngleEDefault, TestKickAngleRDefault, TestKickAngleYDefault, TestKickAngleUDefault, TestKickAngleIDefault
Dim TestKickAnglePDefault, TestKickAngleADefault, TestKickAngleSDefault, TestKickAngleFDefault, TestKickAngleGDefault
Dim TestKickAngleW, TestKickAngleE, TestKickAngleR, TestKickAngleY, TestKickAngleU, TestKickAngleI
Dim TestKickAngleP, TestKickAngleA, TestKickAngleS, TestKickAngleF, TestKickAngleG
TestKickAngleWDefault =  - 27
TestKickAngleEDefault =  - 20
TestKickAngleRDefault =  - 14
TestKickAngleYDefault =  - 8
TestKickAngleUDefault =  - 3
TestKickAngleIDefault = 1
TestKickAnglePDefault = 5
TestKickAngleADefault = 11
TestKickAngleSDefault = 17
TestKickAngleFDefault = 19
TestKickAngleGDefault = 5
If DebugShotMode = 1 Then LoadTestKickAngles

Sub SaveTestKickAngles
	Dim FileObj, OutFile
	Set FileObj = CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) Then Exit Sub
	Set OutFile = FileObj.CreateTextFile(UserDirectory & cGameName & ".txt", True)
	
	OutFile.WriteLine TestKickAngleW
	OutFile.WriteLine TestKickAngleE
	OutFile.WriteLine TestKickAngleR
	OutFile.WriteLine TestKickAngleY
	OutFile.WriteLine TestKickAngleU
	OutFile.WriteLine TestKickAngleI
	OutFile.WriteLine TestKickAngleP
	OutFile.WriteLine TestKickAngleA
	OutFile.WriteLine TestKickAngleS
	OutFile.WriteLine TestKickAngleF
	OutFile.WriteLine TestKickAngleG
	OutFile.Close
	
	Set OutFile = Nothing
	Set FileObj = Nothing
End Sub

Sub LoadTestKickAngles
	Dim FileObj, OutFile, TextStr
	
	Set FileObj = CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) Then
		MsgBox "User directory missing"
		Exit Sub
	End If
	
	If FileObj.FileExists(UserDirectory & cGameName & ".txt") Then
		Set OutFile = FileObj.GetFile(UserDirectory & cGameName & ".txt")
		Set TextStr = OutFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream = True) Then
			Exit Sub
		End If
		
		TestKickAngleW = TextStr.ReadLine
		TestKickAngleE = TextStr.ReadLine
		TestKickAngleR = TextStr.ReadLine
		TestKickAngleY = TextStr.ReadLine
		TestKickAngleU = TextStr.ReadLine
		TestKickAngleI = TextStr.ReadLine
		TestKickAngleP = TextStr.ReadLine
		TestKickAngleA = TextStr.ReadLine
		TestKickAngleS = TextStr.ReadLine
		TestKickAngleF = TextStr.ReadLine
		TestKickAngleG = TextStr.ReadLine
		TextStr.Close
	Else
		'create file
		TestKickAngleW = TestKickAngleWDefault
		TestKickAngleE = TestKickAngleEDefault
		TestKickAngleR = TestKickAngleRDefault
		TestKickAngleY = TestKickAngleYDefault
		TestKickAngleU = TestKickAngleUDefault
		TestKickAngleI = TestKickAngleIDefault
		TestKickAngleP = TestKickAnglePDefault
		TestKickAngleA = TestKickAngleADefault
		TestKickAngleS = TestKickAngleSDefault
		TestKickAngleF = TestKickAngleFDefault
		TestKickAngleG = TestKickAngleGDefault
		SaveTestKickAngles
	End If
	
	Set OutFile = Nothing
	Set FileObj = Nothing
	
End Sub
'****************************************************************
' End of Section; Debug Shot Tester 3.2
'****************************************************************
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


  
'Define a variable for each drop target

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

Dim DTArray : DTArray = Array()

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
	DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate =  DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw, -1)
End Sub

Sub DTDrop(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw, 1)
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
				DTAction switchid, 1
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

		If DTArray(ind).isDropped = True Then
			DTArray(ind).isDropped = False 'Mark target as not dropped
			If UsingROM Then 
				controller.Switch(Switchid mod 100) = 0
			Else
				DTAction switchid, 0
			End If
		End If
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

Sub UpdateTargets

	If DTDropped(1) = True Then
		BM_pantherLid.RotX = -6
	Else
		BM_pantherLid.RotX = 0
	End If
	BM_pantherLid.transz = BM_sw01.transz
	BM_pantherSupport.transz = BM_sw01.transz

	If DTDropped(2) = True Then
		BM_pantherLid2.RotX = -6
	Else
		BM_pantherLid2.RotX = 0
	End If
	BM_pantherLid2.transz = BM_sw02.transz
	BM_pantherSupport2.transz = BM_sw02.transz
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
Dim ST11, ST12, ST13, ST15, ST16, ST17

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


Set ST11 = (new StandupTarget)(sw11, BM_sw11,11 , 0)
Set ST12 = (new StandupTarget)(sw12, BM_sw11,12, 0)
Set ST13 = (new StandupTarget)(sw13, BM_sw11,13, 0)

Set ST15 = (new StandupTarget)(sw15, BM_sw15,15 , 0)
Set ST16 = (new StandupTarget)(sw16, BM_sw16,16, 0)
Set ST17 = (new StandupTarget)(sw17, BM_sw17,17, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST11, ST12, ST13, ST15, ST16, ST17)

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
			STAction switch, 1
		End If
		STAnimate = 2
		Exit Function
	ElseIf animate = 2 Then
		prim.transy = prim.transy + STAnimStep
		If prim.transy >= 0 Then
			prim.transy = 0
			primary.collidable = 1
			STAnimate = 0
			STAction switch, 0
			Exit Function
		Else
			STAnimate = 2
		End If
	End If
End Function


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

'*******************************************
'  Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
	If Enabled Then
		LF.Fire  'leftflipper.rotatetoend
		'LeftFlipper1.rotatetoend
		If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
			RandomSoundReflipUpLeft LeftFlipper
		Else
			SoundFlipperUpAttackLeft LeftFlipper
			RandomSoundFlipperUpLeft LeftFlipper
		End If
	Else
		LeftFlipper.RotateToStart
		'LeftFlipper1.rotatetostart
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
	If Enabled Then
		RF.Fire 'rightflipper.rotatetoend
		
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
		Else
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
		End If
	Else
		RightFlipper.RotateToStart
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
		End If
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
	CheckLiveCatch Activeball, LeftFlipper, LFCount, parm
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch Activeball, RightFlipper, RFCount, parm
	RightFlipperCollide parm
End Sub

Sub FlipperVisualUpdate 'This subroutine updates the flipper shadows and visual primitives
	FlipperLSh.RotZ = LeftFlipper.CurrentAngle
	FlipperRSh.RotZ = RightFlipper.CurrentAngle
'	LFLogo.RotZ = LeftFlipper.CurrentAngle
'	RFlogo.RotZ = RightFlipper.CurrentAngle
End Sub


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
		End If
		
		If keycode = RightFlipperKey Then
			FlipperActivate RightFlipper, RFPress
			SolRFlipper True	'This would be called by the solenoid callbacks if using a ROM
			UpRightFlipper.RotateToEnd
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


'*******************************************
'  Kickers, Saucers
'*******************************************

'************************* VUKs *****************************
Dim KickerBallCave
	
Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
	dim rangle
	rangle = PI * (kangle - 90) / 180
    
	kball.z = kball.z + kzlift
	kball.velz = kvelz
	kball.velx = cos(rangle)*kvel
	kball.vely = sin(rangle)*kvel
End Sub


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_LEMK: BP_LEMK=Array(BM_LEMK)
Dim BP_LSling1: BP_LSling1=Array(BM_LSling1, LM_inserts_L36_LSling1, LM_Journey_L67_LSling1)
Dim BP_LSling2: BP_LSling2=Array(BM_LSling2, LM_Journey_L67_LSling2)
Dim BP_PF: BP_PF=Array(BM_PF, LM_inserts_L10_PF, LM_inserts_L11_PF, LM_inserts_L12_PF, LM_inserts_L13_PF, LM_inserts_L14_PF, LM_inserts_L15_PF, LM_inserts_L16_PF, LM_inserts_L17_PF, LM_inserts_L18_PF, LM_inserts_L19_PF, LM_inserts_L1_PF, LM_inserts_L20_PF, LM_inserts_L21_PF, LM_inserts_L22_PF, LM_inserts_L23_PF, LM_inserts_L24_PF, LM_inserts_L25_PF, LM_inserts_L26_PF, LM_inserts_L27_PF, LM_inserts_L28_PF, LM_inserts_L29_PF, LM_inserts_L2_PF, LM_inserts_L30_PF, LM_inserts_L31_PF, LM_inserts_L32_PF, LM_inserts_L33_PF, LM_inserts_L34_PF, LM_inserts_L35_PF, LM_inserts_L36_PF, LM_inserts_L37_PF, LM_inserts_L38_PF, LM_inserts_L39_PF, LM_inserts_L3_PF, LM_inserts_L40_PF, LM_inserts_L41_PF, LM_inserts_L42_PF, LM_inserts_L43_PF, LM_inserts_L44_PF, LM_inserts_L4_PF, LM_inserts_L5_PF, LM_Journey_L60_PF, LM_Journey_L61_PF, LM_Journey_L62_PF, LM_Journey_L63_PF, LM_Journey_L64_PF, LM_Journey_L65_PF, LM_Journey_L66_PF, LM_Journey_L67_PF, LM_Journey_L68_PF, LM_Journey_L69_PF, LM_inserts_L6_PF, LM_Journey_L70_PF, _
	LM_Journey_L71_PF, LM_Journey_L72_PF, LM_Journey_L73_PF, LM_Journey_L74_PF, LM_Journey_L75_PF, LM_Journey_L76_PF, LM_Journey_L77_PF, LM_Journey_L78_PF, LM_Journey_L79_PF, LM_inserts_L7_PF, LM_Journey_L80_PF, LM_Journey_L81_PF, LM_Journey_L82_PF, LM_Journey_L83_PF, LM_Journey_L84_PF, LM_Journey_L85_PF, LM_Journey_L86_PF, LM_Journey_L87_PF, LM_Journey_L88_PF, LM_Journey_L89_PF, LM_inserts_L8_PF, LM_Journey_L90_PF, LM_Journey_L91_PF, LM_inserts_L9_PF)
Dim BP_REMK: BP_REMK=Array(BM_REMK)
Dim BP_RSling1: BP_RSling1=Array(BM_RSling1, LM_inserts_L41_RSling1)
Dim BP_RSling2: BP_RSling2=Array(BM_RSling2, LM_inserts_L34_RSling2, LM_inserts_L41_RSling2)
Dim BP_UnderPF: BP_UnderPF=Array(BM_UnderPF, LM_inserts_L10_UnderPF, LM_inserts_L11_UnderPF, LM_inserts_L12_UnderPF, LM_inserts_L13_UnderPF, LM_inserts_L14_UnderPF, LM_inserts_L15_UnderPF, LM_inserts_L16_UnderPF, LM_inserts_L17_UnderPF, LM_inserts_L18_UnderPF, LM_inserts_L19_UnderPF, LM_inserts_L20_UnderPF, LM_inserts_L21_UnderPF, LM_inserts_L22_UnderPF, LM_inserts_L23_UnderPF, LM_inserts_L24_UnderPF, LM_inserts_L25_UnderPF, LM_inserts_L26_UnderPF, LM_inserts_L27_UnderPF, LM_inserts_L28_UnderPF, LM_inserts_L29_UnderPF, LM_inserts_L2_UnderPF, LM_inserts_L30_UnderPF, LM_inserts_L31_UnderPF, LM_inserts_L32_UnderPF, LM_inserts_L33_UnderPF, LM_inserts_L34_UnderPF, LM_inserts_L35_UnderPF, LM_inserts_L36_UnderPF, LM_inserts_L37_UnderPF, LM_inserts_L38_UnderPF, LM_inserts_L39_UnderPF, LM_inserts_L3_UnderPF, LM_inserts_L40_UnderPF, LM_inserts_L41_UnderPF, LM_inserts_L42_UnderPF, LM_inserts_L43_UnderPF, LM_inserts_L44_UnderPF, LM_inserts_L4_UnderPF, LM_inserts_L5_UnderPF, LM_inserts_L6_UnderPF, LM_Journey_L79_UnderPF, _
	LM_inserts_L7_UnderPF, LM_Journey_L80_UnderPF, LM_Journey_L81_UnderPF, LM_Journey_L82_UnderPF, LM_Journey_L83_UnderPF, LM_Journey_L85_UnderPF, LM_Journey_L86_UnderPF, LM_Journey_L87_UnderPF, LM_inserts_L8_UnderPF, LM_inserts_L9_UnderPF)
Dim BP_pantherLid: BP_pantherLid=Array(BM_pantherLid, LM_inserts_L25_pantherLid)
Dim BP_pantherLid2: BP_pantherLid2=Array(BM_pantherLid2)
Dim BP_pantherSupport: BP_pantherSupport=Array(BM_pantherSupport, LM_inserts_L1_pantherSupport, LM_inserts_L20_pantherSupport, LM_inserts_L25_pantherSupport, LM_inserts_L2_pantherSupport, LM_inserts_L34_pantherSupport, LM_inserts_L35_pantherSupport)
Dim BP_pantherSupport2: BP_pantherSupport2=Array(BM_pantherSupport2, LM_inserts_L19_pantherSupport2, LM_inserts_L32_pantherSupport2, LM_inserts_L33_pantherSupport2)
Dim BP_sw01: BP_sw01=Array(BM_sw01, LM_inserts_L1_sw01, LM_inserts_L2_sw01)
Dim BP_sw02: BP_sw02=Array(BM_sw02, LM_inserts_L22_sw02, LM_inserts_L31_sw02)
Dim BP_sw04: BP_sw04=Array(BM_sw04, LM_inserts_L3_sw04, LM_inserts_L4_sw04, LM_inserts_L5_sw04, LM_Journey_L60_sw04, LM_Journey_L63_sw04, LM_Journey_L84_sw04, LM_Journey_L85_sw04, LM_Journey_L86_sw04, LM_Journey_L87_sw04, LM_Journey_L88_sw04, LM_Journey_L89_sw04, LM_Journey_L90_sw04, LM_Journey_L91_sw04)
Dim BP_sw05: BP_sw05=Array(BM_sw05, LM_inserts_L3_sw05, LM_inserts_L4_sw05, LM_inserts_L5_sw05, LM_Journey_L86_sw05, LM_Journey_L87_sw05, LM_Journey_L88_sw05, LM_Journey_L89_sw05)
Dim BP_sw06: BP_sw06=Array(BM_sw06, LM_inserts_L3_sw06, LM_inserts_L4_sw06, LM_inserts_L5_sw06)
Dim BP_sw08: BP_sw08=Array(BM_sw08, LM_inserts_L17_sw08, LM_inserts_L26_sw08, LM_inserts_L6_sw08, LM_inserts_L7_sw08, LM_inserts_L8_sw08)
Dim BP_sw09: BP_sw09=Array(BM_sw09, LM_inserts_L26_sw09, LM_inserts_L6_sw09, LM_inserts_L7_sw09, LM_inserts_L8_sw09)
Dim BP_sw10: BP_sw10=Array(BM_sw10, LM_inserts_L26_sw10, LM_inserts_L7_sw10, LM_inserts_L8_sw10)
Dim BP_sw11: BP_sw11=Array(BM_sw11, LM_inserts_L3_sw11, LM_inserts_L4_sw11, LM_inserts_L5_sw11)
Dim BP_sw12: BP_sw12=Array(BM_sw12, LM_inserts_L3_sw12, LM_inserts_L4_sw12, LM_inserts_L5_sw12)
Dim BP_sw13: BP_sw13=Array(BM_sw13, LM_inserts_L21_sw13, LM_inserts_L3_sw13, LM_inserts_L4_sw13)
Dim BP_sw15: BP_sw15=Array(BM_sw15, LM_inserts_L6_sw15, LM_inserts_L7_sw15, LM_inserts_L8_sw15)
Dim BP_sw16: BP_sw16=Array(BM_sw16, LM_inserts_L6_sw16, LM_inserts_L7_sw16, LM_inserts_L8_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17, LM_inserts_L26_sw17, LM_inserts_L7_sw17, LM_inserts_L8_sw17)
Dim BP_targetbank: BP_targetbank=Array(BM_targetbank)
' Arrays per lighting scenario
Dim BL_Journey_L60: BL_Journey_L60=Array(LM_Journey_L60_PF, LM_Journey_L60_sw04)
Dim BL_Journey_L61: BL_Journey_L61=Array(LM_Journey_L61_PF)
Dim BL_Journey_L62: BL_Journey_L62=Array(LM_Journey_L62_PF)
Dim BL_Journey_L63: BL_Journey_L63=Array(LM_Journey_L63_PF, LM_Journey_L63_sw04)
Dim BL_Journey_L64: BL_Journey_L64=Array(LM_Journey_L64_PF)
Dim BL_Journey_L65: BL_Journey_L65=Array(LM_Journey_L65_PF)
Dim BL_Journey_L66: BL_Journey_L66=Array(LM_Journey_L66_PF)
Dim BL_Journey_L67: BL_Journey_L67=Array(LM_Journey_L67_LSling1, LM_Journey_L67_LSling2, LM_Journey_L67_PF)
Dim BL_Journey_L68: BL_Journey_L68=Array(LM_Journey_L68_PF)
Dim BL_Journey_L69: BL_Journey_L69=Array(LM_Journey_L69_PF)
Dim BL_Journey_L70: BL_Journey_L70=Array(LM_Journey_L70_PF)
Dim BL_Journey_L71: BL_Journey_L71=Array(LM_Journey_L71_PF)
Dim BL_Journey_L72: BL_Journey_L72=Array(LM_Journey_L72_PF)
Dim BL_Journey_L73: BL_Journey_L73=Array(LM_Journey_L73_PF)
Dim BL_Journey_L74: BL_Journey_L74=Array(LM_Journey_L74_PF)
Dim BL_Journey_L75: BL_Journey_L75=Array(LM_Journey_L75_PF)
Dim BL_Journey_L76: BL_Journey_L76=Array(LM_Journey_L76_PF)
Dim BL_Journey_L77: BL_Journey_L77=Array(LM_Journey_L77_PF)
Dim BL_Journey_L78: BL_Journey_L78=Array(LM_Journey_L78_PF)
Dim BL_Journey_L79: BL_Journey_L79=Array(LM_Journey_L79_PF, LM_Journey_L79_UnderPF)
Dim BL_Journey_L80: BL_Journey_L80=Array(LM_Journey_L80_PF, LM_Journey_L80_UnderPF)
Dim BL_Journey_L81: BL_Journey_L81=Array(LM_Journey_L81_PF, LM_Journey_L81_UnderPF)
Dim BL_Journey_L82: BL_Journey_L82=Array(LM_Journey_L82_PF, LM_Journey_L82_UnderPF)
Dim BL_Journey_L83: BL_Journey_L83=Array(LM_Journey_L83_PF, LM_Journey_L83_UnderPF)
Dim BL_Journey_L84: BL_Journey_L84=Array(LM_Journey_L84_PF, LM_Journey_L84_sw04)
Dim BL_Journey_L85: BL_Journey_L85=Array(LM_Journey_L85_PF, LM_Journey_L85_UnderPF, LM_Journey_L85_sw04)
Dim BL_Journey_L86: BL_Journey_L86=Array(LM_Journey_L86_PF, LM_Journey_L86_UnderPF, LM_Journey_L86_sw04, LM_Journey_L86_sw05)
Dim BL_Journey_L87: BL_Journey_L87=Array(LM_Journey_L87_PF, LM_Journey_L87_UnderPF, LM_Journey_L87_sw04, LM_Journey_L87_sw05)
Dim BL_Journey_L88: BL_Journey_L88=Array(LM_Journey_L88_PF, LM_Journey_L88_sw04, LM_Journey_L88_sw05)
Dim BL_Journey_L89: BL_Journey_L89=Array(LM_Journey_L89_PF, LM_Journey_L89_sw04, LM_Journey_L89_sw05)
Dim BL_Journey_L90: BL_Journey_L90=Array(LM_Journey_L90_PF, LM_Journey_L90_sw04)
Dim BL_Journey_L91: BL_Journey_L91=Array(LM_Journey_L91_PF, LM_Journey_L91_sw04)
Dim BL_World: BL_World=Array(BM_LEMK, BM_LSling1, BM_LSling2, BM_PF, BM_REMK, BM_RSling1, BM_RSling2, BM_UnderPF, BM_pantherLid, BM_pantherLid2, BM_pantherSupport, BM_pantherSupport2, BM_sw01, BM_sw02, BM_sw04, BM_sw05, BM_sw06, BM_sw08, BM_sw09, BM_sw10, BM_sw11, BM_sw12, BM_sw13, BM_sw15, BM_sw16, BM_sw17, BM_targetbank)
Dim BL_inserts_L1: BL_inserts_L1=Array(LM_inserts_L1_PF, LM_inserts_L1_pantherSupport, LM_inserts_L1_sw01)
Dim BL_inserts_L10: BL_inserts_L10=Array(LM_inserts_L10_PF, LM_inserts_L10_UnderPF)
Dim BL_inserts_L11: BL_inserts_L11=Array(LM_inserts_L11_PF, LM_inserts_L11_UnderPF)
Dim BL_inserts_L12: BL_inserts_L12=Array(LM_inserts_L12_PF, LM_inserts_L12_UnderPF)
Dim BL_inserts_L13: BL_inserts_L13=Array(LM_inserts_L13_PF, LM_inserts_L13_UnderPF)
Dim BL_inserts_L14: BL_inserts_L14=Array(LM_inserts_L14_PF, LM_inserts_L14_UnderPF)
Dim BL_inserts_L15: BL_inserts_L15=Array(LM_inserts_L15_PF, LM_inserts_L15_UnderPF)
Dim BL_inserts_L16: BL_inserts_L16=Array(LM_inserts_L16_PF, LM_inserts_L16_UnderPF)
Dim BL_inserts_L17: BL_inserts_L17=Array(LM_inserts_L17_PF, LM_inserts_L17_UnderPF, LM_inserts_L17_sw08)
Dim BL_inserts_L18: BL_inserts_L18=Array(LM_inserts_L18_PF, LM_inserts_L18_UnderPF)
Dim BL_inserts_L19: BL_inserts_L19=Array(LM_inserts_L19_PF, LM_inserts_L19_UnderPF, LM_inserts_L19_pantherSupport2)
Dim BL_inserts_L2: BL_inserts_L2=Array(LM_inserts_L2_PF, LM_inserts_L2_UnderPF, LM_inserts_L2_pantherSupport, LM_inserts_L2_sw01)
Dim BL_inserts_L20: BL_inserts_L20=Array(LM_inserts_L20_PF, LM_inserts_L20_UnderPF, LM_inserts_L20_pantherSupport)
Dim BL_inserts_L21: BL_inserts_L21=Array(LM_inserts_L21_PF, LM_inserts_L21_UnderPF, LM_inserts_L21_sw13)
Dim BL_inserts_L22: BL_inserts_L22=Array(LM_inserts_L22_PF, LM_inserts_L22_UnderPF, LM_inserts_L22_sw02)
Dim BL_inserts_L23: BL_inserts_L23=Array(LM_inserts_L23_PF, LM_inserts_L23_UnderPF)
Dim BL_inserts_L24: BL_inserts_L24=Array(LM_inserts_L24_PF, LM_inserts_L24_UnderPF)
Dim BL_inserts_L25: BL_inserts_L25=Array(LM_inserts_L25_PF, LM_inserts_L25_UnderPF, LM_inserts_L25_pantherLid, LM_inserts_L25_pantherSupport)
Dim BL_inserts_L26: BL_inserts_L26=Array(LM_inserts_L26_PF, LM_inserts_L26_UnderPF, LM_inserts_L26_sw08, LM_inserts_L26_sw09, LM_inserts_L26_sw10, LM_inserts_L26_sw17)
Dim BL_inserts_L27: BL_inserts_L27=Array(LM_inserts_L27_PF, LM_inserts_L27_UnderPF)
Dim BL_inserts_L28: BL_inserts_L28=Array(LM_inserts_L28_PF, LM_inserts_L28_UnderPF)
Dim BL_inserts_L29: BL_inserts_L29=Array(LM_inserts_L29_PF, LM_inserts_L29_UnderPF)
Dim BL_inserts_L3: BL_inserts_L3=Array(LM_inserts_L3_PF, LM_inserts_L3_UnderPF, LM_inserts_L3_sw04, LM_inserts_L3_sw05, LM_inserts_L3_sw06, LM_inserts_L3_sw11, LM_inserts_L3_sw12, LM_inserts_L3_sw13)
Dim BL_inserts_L30: BL_inserts_L30=Array(LM_inserts_L30_PF, LM_inserts_L30_UnderPF)
Dim BL_inserts_L31: BL_inserts_L31=Array(LM_inserts_L31_PF, LM_inserts_L31_UnderPF, LM_inserts_L31_sw02)
Dim BL_inserts_L32: BL_inserts_L32=Array(LM_inserts_L32_PF, LM_inserts_L32_UnderPF, LM_inserts_L32_pantherSupport2)
Dim BL_inserts_L33: BL_inserts_L33=Array(LM_inserts_L33_PF, LM_inserts_L33_UnderPF, LM_inserts_L33_pantherSupport2)
Dim BL_inserts_L34: BL_inserts_L34=Array(LM_inserts_L34_PF, LM_inserts_L34_RSling2, LM_inserts_L34_UnderPF, LM_inserts_L34_pantherSupport)
Dim BL_inserts_L35: BL_inserts_L35=Array(LM_inserts_L35_PF, LM_inserts_L35_UnderPF, LM_inserts_L35_pantherSupport)
Dim BL_inserts_L36: BL_inserts_L36=Array(LM_inserts_L36_LSling1, LM_inserts_L36_PF, LM_inserts_L36_UnderPF)
Dim BL_inserts_L37: BL_inserts_L37=Array(LM_inserts_L37_PF, LM_inserts_L37_UnderPF)
Dim BL_inserts_L38: BL_inserts_L38=Array(LM_inserts_L38_PF, LM_inserts_L38_UnderPF)
Dim BL_inserts_L39: BL_inserts_L39=Array(LM_inserts_L39_PF, LM_inserts_L39_UnderPF)
Dim BL_inserts_L4: BL_inserts_L4=Array(LM_inserts_L4_PF, LM_inserts_L4_UnderPF, LM_inserts_L4_sw04, LM_inserts_L4_sw05, LM_inserts_L4_sw06, LM_inserts_L4_sw11, LM_inserts_L4_sw12, LM_inserts_L4_sw13)
Dim BL_inserts_L40: BL_inserts_L40=Array(LM_inserts_L40_PF, LM_inserts_L40_UnderPF)
Dim BL_inserts_L41: BL_inserts_L41=Array(LM_inserts_L41_PF, LM_inserts_L41_RSling1, LM_inserts_L41_RSling2, LM_inserts_L41_UnderPF)
Dim BL_inserts_L42: BL_inserts_L42=Array(LM_inserts_L42_PF, LM_inserts_L42_UnderPF)
Dim BL_inserts_L43: BL_inserts_L43=Array(LM_inserts_L43_PF, LM_inserts_L43_UnderPF)
Dim BL_inserts_L44: BL_inserts_L44=Array(LM_inserts_L44_PF, LM_inserts_L44_UnderPF)
Dim BL_inserts_L5: BL_inserts_L5=Array(LM_inserts_L5_PF, LM_inserts_L5_UnderPF, LM_inserts_L5_sw04, LM_inserts_L5_sw05, LM_inserts_L5_sw06, LM_inserts_L5_sw11, LM_inserts_L5_sw12)
Dim BL_inserts_L6: BL_inserts_L6=Array(LM_inserts_L6_PF, LM_inserts_L6_UnderPF, LM_inserts_L6_sw08, LM_inserts_L6_sw09, LM_inserts_L6_sw15, LM_inserts_L6_sw16)
Dim BL_inserts_L7: BL_inserts_L7=Array(LM_inserts_L7_PF, LM_inserts_L7_UnderPF, LM_inserts_L7_sw08, LM_inserts_L7_sw09, LM_inserts_L7_sw10, LM_inserts_L7_sw15, LM_inserts_L7_sw16, LM_inserts_L7_sw17)
Dim BL_inserts_L8: BL_inserts_L8=Array(LM_inserts_L8_PF, LM_inserts_L8_UnderPF, LM_inserts_L8_sw08, LM_inserts_L8_sw09, LM_inserts_L8_sw10, LM_inserts_L8_sw15, LM_inserts_L8_sw16, LM_inserts_L8_sw17)
Dim BL_inserts_L9: BL_inserts_L9=Array(LM_inserts_L9_PF, LM_inserts_L9_UnderPF)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_LEMK, BM_LSling1, BM_LSling2, BM_PF, BM_REMK, BM_RSling1, BM_RSling2, BM_UnderPF, BM_pantherLid, BM_pantherLid2, BM_pantherSupport, BM_pantherSupport2, BM_sw01, BM_sw02, BM_sw04, BM_sw05, BM_sw06, BM_sw08, BM_sw09, BM_sw10, BM_sw11, BM_sw12, BM_sw13, BM_sw15, BM_sw16, BM_sw17, BM_targetbank)
Dim BG_Lightmap: BG_Lightmap=Array(LM_Journey_L60_PF, LM_Journey_L60_sw04, LM_Journey_L61_PF, LM_Journey_L62_PF, LM_Journey_L63_PF, LM_Journey_L63_sw04, LM_Journey_L64_PF, LM_Journey_L65_PF, LM_Journey_L66_PF, LM_Journey_L67_LSling1, LM_Journey_L67_LSling2, LM_Journey_L67_PF, LM_Journey_L68_PF, LM_Journey_L69_PF, LM_Journey_L70_PF, LM_Journey_L71_PF, LM_Journey_L72_PF, LM_Journey_L73_PF, LM_Journey_L74_PF, LM_Journey_L75_PF, LM_Journey_L76_PF, LM_Journey_L77_PF, LM_Journey_L78_PF, LM_Journey_L79_PF, LM_Journey_L79_UnderPF, LM_Journey_L80_PF, LM_Journey_L80_UnderPF, LM_Journey_L81_PF, LM_Journey_L81_UnderPF, LM_Journey_L82_PF, LM_Journey_L82_UnderPF, LM_Journey_L83_PF, LM_Journey_L83_UnderPF, LM_Journey_L84_PF, LM_Journey_L84_sw04, LM_Journey_L85_PF, LM_Journey_L85_UnderPF, LM_Journey_L85_sw04, LM_Journey_L86_PF, LM_Journey_L86_UnderPF, LM_Journey_L86_sw04, LM_Journey_L86_sw05, LM_Journey_L87_PF, LM_Journey_L87_UnderPF, LM_Journey_L87_sw04, LM_Journey_L87_sw05, LM_Journey_L88_PF, LM_Journey_L88_sw04, _
	LM_Journey_L88_sw05, LM_Journey_L89_PF, LM_Journey_L89_sw04, LM_Journey_L89_sw05, LM_Journey_L90_PF, LM_Journey_L90_sw04, LM_Journey_L91_PF, LM_Journey_L91_sw04, LM_inserts_L1_PF, LM_inserts_L1_pantherSupport, LM_inserts_L1_sw01, LM_inserts_L10_PF, LM_inserts_L10_UnderPF, LM_inserts_L11_PF, LM_inserts_L11_UnderPF, LM_inserts_L12_PF, LM_inserts_L12_UnderPF, LM_inserts_L13_PF, LM_inserts_L13_UnderPF, LM_inserts_L14_PF, LM_inserts_L14_UnderPF, LM_inserts_L15_PF, LM_inserts_L15_UnderPF, LM_inserts_L16_PF, LM_inserts_L16_UnderPF, LM_inserts_L17_PF, LM_inserts_L17_UnderPF, LM_inserts_L17_sw08, LM_inserts_L18_PF, LM_inserts_L18_UnderPF, LM_inserts_L19_PF, LM_inserts_L19_UnderPF, LM_inserts_L19_pantherSupport2, LM_inserts_L2_PF, LM_inserts_L2_UnderPF, LM_inserts_L2_pantherSupport, LM_inserts_L2_sw01, LM_inserts_L20_PF, LM_inserts_L20_UnderPF, LM_inserts_L20_pantherSupport, LM_inserts_L21_PF, LM_inserts_L21_UnderPF, LM_inserts_L21_sw13, LM_inserts_L22_PF, LM_inserts_L22_UnderPF, LM_inserts_L22_sw02, LM_inserts_L23_PF, _
	LM_inserts_L23_UnderPF, LM_inserts_L24_PF, LM_inserts_L24_UnderPF, LM_inserts_L25_PF, LM_inserts_L25_UnderPF, LM_inserts_L25_pantherLid, LM_inserts_L25_pantherSupport, LM_inserts_L26_PF, LM_inserts_L26_UnderPF, LM_inserts_L26_sw08, LM_inserts_L26_sw09, LM_inserts_L26_sw10, LM_inserts_L26_sw17, LM_inserts_L27_PF, LM_inserts_L27_UnderPF, LM_inserts_L28_PF, LM_inserts_L28_UnderPF, LM_inserts_L29_PF, LM_inserts_L29_UnderPF, LM_inserts_L3_PF, LM_inserts_L3_UnderPF, LM_inserts_L3_sw04, LM_inserts_L3_sw05, LM_inserts_L3_sw06, LM_inserts_L3_sw11, LM_inserts_L3_sw12, LM_inserts_L3_sw13, LM_inserts_L30_PF, LM_inserts_L30_UnderPF, LM_inserts_L31_PF, LM_inserts_L31_UnderPF, LM_inserts_L31_sw02, LM_inserts_L32_PF, LM_inserts_L32_UnderPF, LM_inserts_L32_pantherSupport2, LM_inserts_L33_PF, LM_inserts_L33_UnderPF, LM_inserts_L33_pantherSupport2, LM_inserts_L34_PF, LM_inserts_L34_RSling2, LM_inserts_L34_UnderPF, LM_inserts_L34_pantherSupport, LM_inserts_L35_PF, LM_inserts_L35_UnderPF, LM_inserts_L35_pantherSupport, _
	LM_inserts_L36_LSling1, LM_inserts_L36_PF, LM_inserts_L36_UnderPF, LM_inserts_L37_PF, LM_inserts_L37_UnderPF, LM_inserts_L38_PF, LM_inserts_L38_UnderPF, LM_inserts_L39_PF, LM_inserts_L39_UnderPF, LM_inserts_L4_PF, LM_inserts_L4_UnderPF, LM_inserts_L4_sw04, LM_inserts_L4_sw05, LM_inserts_L4_sw06, LM_inserts_L4_sw11, LM_inserts_L4_sw12, LM_inserts_L4_sw13, LM_inserts_L40_PF, LM_inserts_L40_UnderPF, LM_inserts_L41_PF, LM_inserts_L41_RSling1, LM_inserts_L41_RSling2, LM_inserts_L41_UnderPF, LM_inserts_L42_PF, LM_inserts_L42_UnderPF, LM_inserts_L43_PF, LM_inserts_L43_UnderPF, LM_inserts_L44_PF, LM_inserts_L44_UnderPF, LM_inserts_L5_PF, LM_inserts_L5_UnderPF, LM_inserts_L5_sw04, LM_inserts_L5_sw05, LM_inserts_L5_sw06, LM_inserts_L5_sw11, LM_inserts_L5_sw12, LM_inserts_L6_PF, LM_inserts_L6_UnderPF, LM_inserts_L6_sw08, LM_inserts_L6_sw09, LM_inserts_L6_sw15, LM_inserts_L6_sw16, LM_inserts_L7_PF, LM_inserts_L7_UnderPF, LM_inserts_L7_sw08, LM_inserts_L7_sw09, LM_inserts_L7_sw10, LM_inserts_L7_sw15, LM_inserts_L7_sw16, _
	LM_inserts_L7_sw17, LM_inserts_L8_PF, LM_inserts_L8_UnderPF, LM_inserts_L8_sw08, LM_inserts_L8_sw09, LM_inserts_L8_sw10, LM_inserts_L8_sw15, LM_inserts_L8_sw16, LM_inserts_L8_sw17, LM_inserts_L9_PF, LM_inserts_L9_UnderPF)
Dim BG_All: BG_All=Array(BM_LEMK, BM_LSling1, BM_LSling2, BM_PF, BM_REMK, BM_RSling1, BM_RSling2, BM_UnderPF, BM_pantherLid, BM_pantherLid2, BM_pantherSupport, BM_pantherSupport2, BM_sw01, BM_sw02, BM_sw04, BM_sw05, BM_sw06, BM_sw08, BM_sw09, BM_sw10, BM_sw11, BM_sw12, BM_sw13, BM_sw15, BM_sw16, BM_sw17, BM_targetbank, LM_Journey_L60_PF, LM_Journey_L60_sw04, LM_Journey_L61_PF, LM_Journey_L62_PF, LM_Journey_L63_PF, LM_Journey_L63_sw04, LM_Journey_L64_PF, LM_Journey_L65_PF, LM_Journey_L66_PF, LM_Journey_L67_LSling1, LM_Journey_L67_LSling2, LM_Journey_L67_PF, LM_Journey_L68_PF, LM_Journey_L69_PF, LM_Journey_L70_PF, LM_Journey_L71_PF, LM_Journey_L72_PF, LM_Journey_L73_PF, LM_Journey_L74_PF, LM_Journey_L75_PF, LM_Journey_L76_PF, LM_Journey_L77_PF, LM_Journey_L78_PF, LM_Journey_L79_PF, LM_Journey_L79_UnderPF, LM_Journey_L80_PF, LM_Journey_L80_UnderPF, LM_Journey_L81_PF, LM_Journey_L81_UnderPF, LM_Journey_L82_PF, LM_Journey_L82_UnderPF, LM_Journey_L83_PF, LM_Journey_L83_UnderPF, LM_Journey_L84_PF, _
	LM_Journey_L84_sw04, LM_Journey_L85_PF, LM_Journey_L85_UnderPF, LM_Journey_L85_sw04, LM_Journey_L86_PF, LM_Journey_L86_UnderPF, LM_Journey_L86_sw04, LM_Journey_L86_sw05, LM_Journey_L87_PF, LM_Journey_L87_UnderPF, LM_Journey_L87_sw04, LM_Journey_L87_sw05, LM_Journey_L88_PF, LM_Journey_L88_sw04, LM_Journey_L88_sw05, LM_Journey_L89_PF, LM_Journey_L89_sw04, LM_Journey_L89_sw05, LM_Journey_L90_PF, LM_Journey_L90_sw04, LM_Journey_L91_PF, LM_Journey_L91_sw04, LM_inserts_L1_PF, LM_inserts_L1_pantherSupport, LM_inserts_L1_sw01, LM_inserts_L10_PF, LM_inserts_L10_UnderPF, LM_inserts_L11_PF, LM_inserts_L11_UnderPF, LM_inserts_L12_PF, LM_inserts_L12_UnderPF, LM_inserts_L13_PF, LM_inserts_L13_UnderPF, LM_inserts_L14_PF, LM_inserts_L14_UnderPF, LM_inserts_L15_PF, LM_inserts_L15_UnderPF, LM_inserts_L16_PF, LM_inserts_L16_UnderPF, LM_inserts_L17_PF, LM_inserts_L17_UnderPF, LM_inserts_L17_sw08, LM_inserts_L18_PF, LM_inserts_L18_UnderPF, LM_inserts_L19_PF, LM_inserts_L19_UnderPF, LM_inserts_L19_pantherSupport2, _
	LM_inserts_L2_PF, LM_inserts_L2_UnderPF, LM_inserts_L2_pantherSupport, LM_inserts_L2_sw01, LM_inserts_L20_PF, LM_inserts_L20_UnderPF, LM_inserts_L20_pantherSupport, LM_inserts_L21_PF, LM_inserts_L21_UnderPF, LM_inserts_L21_sw13, LM_inserts_L22_PF, LM_inserts_L22_UnderPF, LM_inserts_L22_sw02, LM_inserts_L23_PF, LM_inserts_L23_UnderPF, LM_inserts_L24_PF, LM_inserts_L24_UnderPF, LM_inserts_L25_PF, LM_inserts_L25_UnderPF, LM_inserts_L25_pantherLid, LM_inserts_L25_pantherSupport, LM_inserts_L26_PF, LM_inserts_L26_UnderPF, LM_inserts_L26_sw08, LM_inserts_L26_sw09, LM_inserts_L26_sw10, LM_inserts_L26_sw17, LM_inserts_L27_PF, LM_inserts_L27_UnderPF, LM_inserts_L28_PF, LM_inserts_L28_UnderPF, LM_inserts_L29_PF, LM_inserts_L29_UnderPF, LM_inserts_L3_PF, LM_inserts_L3_UnderPF, LM_inserts_L3_sw04, LM_inserts_L3_sw05, LM_inserts_L3_sw06, LM_inserts_L3_sw11, LM_inserts_L3_sw12, LM_inserts_L3_sw13, LM_inserts_L30_PF, LM_inserts_L30_UnderPF, LM_inserts_L31_PF, LM_inserts_L31_UnderPF, LM_inserts_L31_sw02, LM_inserts_L32_PF, _
	LM_inserts_L32_UnderPF, LM_inserts_L32_pantherSupport2, LM_inserts_L33_PF, LM_inserts_L33_UnderPF, LM_inserts_L33_pantherSupport2, LM_inserts_L34_PF, LM_inserts_L34_RSling2, LM_inserts_L34_UnderPF, LM_inserts_L34_pantherSupport, LM_inserts_L35_PF, LM_inserts_L35_UnderPF, LM_inserts_L35_pantherSupport, LM_inserts_L36_LSling1, LM_inserts_L36_PF, LM_inserts_L36_UnderPF, LM_inserts_L37_PF, LM_inserts_L37_UnderPF, LM_inserts_L38_PF, LM_inserts_L38_UnderPF, LM_inserts_L39_PF, LM_inserts_L39_UnderPF, LM_inserts_L4_PF, LM_inserts_L4_UnderPF, LM_inserts_L4_sw04, LM_inserts_L4_sw05, LM_inserts_L4_sw06, LM_inserts_L4_sw11, LM_inserts_L4_sw12, LM_inserts_L4_sw13, LM_inserts_L40_PF, LM_inserts_L40_UnderPF, LM_inserts_L41_PF, LM_inserts_L41_RSling1, LM_inserts_L41_RSling2, LM_inserts_L41_UnderPF, LM_inserts_L42_PF, LM_inserts_L42_UnderPF, LM_inserts_L43_PF, LM_inserts_L43_UnderPF, LM_inserts_L44_PF, LM_inserts_L44_UnderPF, LM_inserts_L5_PF, LM_inserts_L5_UnderPF, LM_inserts_L5_sw04, LM_inserts_L5_sw05, LM_inserts_L5_sw06, _
	LM_inserts_L5_sw11, LM_inserts_L5_sw12, LM_inserts_L6_PF, LM_inserts_L6_UnderPF, LM_inserts_L6_sw08, LM_inserts_L6_sw09, LM_inserts_L6_sw15, LM_inserts_L6_sw16, LM_inserts_L7_PF, LM_inserts_L7_UnderPF, LM_inserts_L7_sw08, LM_inserts_L7_sw09, LM_inserts_L7_sw10, LM_inserts_L7_sw15, LM_inserts_L7_sw16, LM_inserts_L7_sw17, LM_inserts_L8_PF, LM_inserts_L8_UnderPF, LM_inserts_L8_sw08, LM_inserts_L8_sw09, LM_inserts_L8_sw10, LM_inserts_L8_sw15, LM_inserts_L8_sw16, LM_inserts_L8_sw17, LM_inserts_L9_PF, LM_inserts_L9_UnderPF)
' VLM  Arrays - End


'******************************************************
'****  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these 
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Flippers:	 https://www.youtube.com/watch?v=FWvM9_CdVHw
' Dampeners:	 https://www.youtube.com/watch?v=tqsxx48C6Pg
' Physics:		 https://www.youtube.com/watch?v=UcRMG-2svvE
'
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25) 
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 9.5-10.5 |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
' 
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 4-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |

'******************************************************
'****  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level well need the following:
'	1. flippers with specific physics settings
'	2. custom triggers for each flipper (TriggerLF, TriggerRF)
'	3. an object or point to tell the script where the tip of the flipper is at rest (EndPointLp, EndPointRp)
'	4. and, special scripting
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.  
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end 
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.3            | 0.3                   | 0.275                  | 0.275              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era) 
'******************************************************

Dim LF
Set LF = New FlipperPolarity
Dim RF
Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'   for each x in a
'	   x.AddPoint "Ycoef", 0, RightFlipper.Y - 65, 1
'	   x.AddPoint "Ycoef", 1, RightFlipper.Y - 11, 1
'	   x.enabled = True
'	   x.TimeDelay = 80
'   Next
'
'   AddPt "Polarity", 0, 0, 0
'   AddPt "Polarity", 1, 0.05, - 2.7		
'   AddPt "Polarity", 2, 0.33, - 2.7
'   AddPt "Polarity", 3, 0.37, - 2.7		
'   AddPt "Polarity", 4, 0.41, - 2.7
'   AddPt "Polarity", 5, 0.45, - 2.7
'   AddPt "Polarity", 6, 0.576, - 2.7
'   AddPt "Polarity", 7, 0.66, - 1.8
'   AddPt "Polarity", 8, 0.743, - 0.5
'   AddPt "Polarity", 9, 0.81, - 0.5
'   AddPt "Polarity", 10, 0.88, 0
'
'   addpt "Velocity", 0, 0, 1
'   addpt "Velocity", 1, 0.16, 1.06
'   addpt "Velocity", 2, 0.41, 1.05
'   addpt "Velocity", 3, 0.53, 1 '0.982
'   addpt "Velocity", 4, 0.702, 0.968
'   addpt "Velocity", 5, 0.95,  0.968
'   addpt "Velocity", 6, 1.03, 0.945
'
'   LF.Object = LeftFlipper		
'   LF.EndPoint = EndPointLp
'   RF.Object = RightFlipper
'   RF.EndPoint = EndPointRp
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'   for each x in a
'	   x.AddPoint "Ycoef", 0, RightFlipper.Y - 65, 1
'	   x.AddPoint "Ycoef", 1, RightFlipper.Y - 11, 1
'	   x.enabled = True
'	   x.TimeDelay = 80
'   Next
'
'   AddPt "Polarity", 0, 0, 0
'   AddPt "Polarity", 1, 0.05, - 3.7		
'   AddPt "Polarity", 2, 0.33, - 3.7
'   AddPt "Polarity", 3, 0.37, - 3.7
'   AddPt "Polarity", 4, 0.41, - 3.7
'   AddPt "Polarity", 5, 0.45, - 3.7 
'   AddPt "Polarity", 6, 0.576,- 3.7
'   AddPt "Polarity", 7, 0.66, - 2.3
'   AddPt "Polarity", 8, 0.743, - 1.5
'   AddPt "Polarity", 9, 0.81, - 1
'   AddPt "Polarity", 10, 0.88, 0
'
'   addpt "Velocity", 0, 0, 1
'   addpt "Velocity", 1, 0.16, 1.06
'   addpt "Velocity", 2, 0.41, 1.05
'   addpt "Velocity", 3, 0.53, 1 '0.982
'   addpt "Velocity", 4, 0.702, 0.968
'   addpt "Velocity", 5, 0.95,  0.968
'   addpt "Velocity", 6, 1.03, 0.945
'
'   LF.Object = LeftFlipper		
'   LF.EndPoint = EndPointLp
'   RF.Object = RightFlipper
'   RF.EndPoint = EndPointRp
'End Sub

'*******************************************
'  Late 80's early 90's

Sub InitPolarity()
	Dim x, a
	a = Array(LF, RF)
	For Each x In a
		x.AddPoint "Ycoef", 0, RightFlipper.Y - 65, 1
		x.AddPoint "Ycoef", 1, RightFlipper.Y - 11, 1
		x.enabled = True
		x.TimeDelay = 60
	Next
	
	AddPt "Polarity", 0, 0, 0
	AddPt "Polarity", 1, 0.05, - 5
	AddPt "Polarity", 2, 0.4, - 5
	AddPt "Polarity", 3, 0.6, - 4.5
	AddPt "Polarity", 4, 0.65, - 4.0
	AddPt "Polarity", 5, 0.7, - 3.5
	AddPt "Polarity", 6, 0.75, - 3.0
	AddPt "Polarity", 7, 0.8, - 2.5
	AddPt "Polarity", 8, 0.85, - 2.0
	AddPt "Polarity", 9, 0.9, - 1.5
	AddPt "Polarity", 10, 0.95, - 1.0
	AddPt "Polarity", 11, 1, - 0.5
	AddPt "Polarity", 12, 1.1, 0
	AddPt "Polarity", 13, 1.3, 0
	
	addpt "Velocity", 0, 0, 1
	addpt "Velocity", 1, 0.16, 1.06
	addpt "Velocity", 2, 0.41, 1.05
	addpt "Velocity", 3, 0.53, 1 '0.982
	addpt "Velocity", 4, 0.702, 0.968
	addpt "Velocity", 5, 0.95,  0.968
	addpt "Velocity", 6, 1.03,  0.945
	
	LF.Object = LeftFlipper
	LF.EndPoint = EndPointLp
	RF.Object = RightFlipper
	RF.EndPoint = EndPointRp
End Sub

'
''*******************************************
'' Early 90's and after
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'   for each x in a
'	   x.AddPoint "Ycoef", 0, RightFlipper.Y - 65, 1
'	   x.AddPoint "Ycoef", 1, RightFlipper.Y - 11, 1
'	   x.enabled = True
'	   x.TimeDelay = 60
'   Next
'
'   AddPt "Polarity", 0, 0, 0
'   AddPt "Polarity", 1, 0.05, - 5.5
'   AddPt "Polarity", 2, 0.4, - 5.5
'   AddPt "Polarity", 3, 0.6, - 5.0
'   AddPt "Polarity", 4, 0.65, - 4.5
'   AddPt "Polarity", 5, 0.7, - 4.0
'   AddPt "Polarity", 6, 0.75, - 3.5
'   AddPt "Polarity", 7, 0.8, - 3.0
'   AddPt "Polarity", 8, 0.85, - 2.5
'   AddPt "Polarity", 9, 0.9,- 2.0
'   AddPt "Polarity", 10, 0.95, - 1.5
'   AddPt "Polarity", 11, 1, - 1.0
'   AddPt "Polarity", 12, 1.05, - 0.5
'   AddPt "Polarity", 13, 1.1, 0
'   AddPt "Polarity", 14, 1.3, 0
'
'   addpt "Velocity", 0, 0, 1
'   addpt "Velocity", 1, 0.16, 1.06
'   addpt "Velocity", 2, 0.41, 1.05
'   addpt "Velocity", 3, 0.53, 1 '0.982
'   addpt "Velocity", 4, 0.702, 0.968
'   addpt "Velocity", 5, 0.95,  0.968
'   addpt "Velocity", 6, 1.03, 0.945
'
'   LF.Object = LeftFlipper		
'   LF.EndPoint = EndPointLp
'   RF.Object = RightFlipper
'   RF.EndPoint = EndPointRp
'End Sub

' Flipper trigger hit subs
Sub TriggerLF_Hit()
	LF.Addball activeball
End Sub
Sub TriggerLF_UnHit()
	LF.PolarityCorrect activeball
End Sub
Sub TriggerRF_Hit()
	RF.Addball activeball
End Sub
Sub TriggerRF_UnHit()
	RF.PolarityCorrect activeball
End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt		  'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay		'delay before trigger turns off and polarity is disabled TODO set time!
	Private Flipper, FlipperStart,FlipperEnd, FlipperEndY, LR, PartialFlipCoef
	Private Balls(20), balldata(20)
	
	Dim PolarityIn, PolarityOut
	Dim VelocityIn, VelocityOut
	Dim YcoefIn, YcoefOut
	
	Public Sub Class_Initialize
		ReDim PolarityIn(0)
		ReDim PolarityOut(0)
		ReDim VelocityIn(0)
		ReDim VelocityOut(0)
		ReDim YcoefIn(0)
		ReDim YcoefOut(0)
		Enabled = True
		TimeDelay = 50
		LR = 1
		Dim x
		For x = 0 To UBound(balls)
			balls(x) = Empty
			Set Balldata(x) = New SpoofBall
		Next
	End Sub
	
	Public Property Let Object(aInput)
		Set Flipper = aInput
		StartPoint = Flipper.x
	End Property
	
	Public Property Let StartPoint(aInput)
		If IsObject(aInput) Then
			FlipperStart = aInput.x
		Else
			FlipperStart = aInput
		End If
	End Property
	
	Public Property Get StartPoint
		StartPoint = FlipperStart
	End Property
	
	Public Property Let EndPoint(aInput)
		FlipperEnd = aInput.x
		FlipperEndY = aInput.y
	End Property
	
	Public Property Get EndPoint
		EndPoint = FlipperEnd
	End Property
	
	Public Property Get EndPointY
		EndPointY = FlipperEndY
	End Property
	
	Public Sub AddPoint(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out) 
		Select Case aChooseArray
			Case "Polarity"
			ShuffleArrays PolarityIn, PolarityOut, 1
			PolarityIn(aIDX) = aX
			PolarityOut(aIDX) = aY
			ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity"
			ShuffleArrays VelocityIn, VelocityOut, 1
			VelocityIn(aIDX) = aX
			VelocityOut(aIDX) = aY
			ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef"
			ShuffleArrays YcoefIn, YcoefOut, 1
			YcoefIn(aIDX) = aX
			YcoefOut(aIDX) = aY
			ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
		If gametime > 100 Then Report aChooseArray
	End Sub
	
	Public Sub Report(aChooseArray) 'debug, reports all coords in tbPL.text
		If Not DebugOn Then Exit Sub
		Dim a1, a2
		Select Case aChooseArray
			Case "Polarity"
			a1 = PolarityIn
			a2 = PolarityOut
			Case "Velocity"
			a1 = VelocityIn
			a2 = VelocityOut
			Case "Ycoef"
			a1 = YcoefIn
			a2 = YcoefOut
			Case Else
			tbpl.text = "wrong string"
			Exit Sub
		End Select
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & aChooseArray & " x: " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		tbpl.text = str
	End Sub
	
	Public Sub AddBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If IsEmpty(balls(x)) Then
				Set balls(x) = aBall
				Exit Sub
			End If
		Next
	End Sub
	
	Private Sub RemoveBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If TypeName(balls(x) ) = "IBall" Then
				If aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub
	
	Public Sub Fire()
		Flipper.RotateToEnd
		processballs
	End Sub
	
	Public Property Get Pos 'returns % position a ball. For debug stuff.
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x) ) Then
				balldata(x).Data = balls(x)
			End If
		Next
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = Abs(PartialFlipCoef - 1)
	End Sub
	
	Private Function FlipperOn() 'Timer shutoff for polaritycorrect
		If gameTime < FlipAt + TimeDelay Then FlipperOn = True
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY >  - 8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)	'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)  'find safety coefficient 'ycoef' data
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx * VelCoef
				If Enabled Then aBall.Vely = aBall.Vely * VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				If StartPoint > EndPoint Then LR =  - 1 'Reverse polarity if left flipper
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled Then aBall.VelX = aBall.VelX + 1 * (AddX * ycoef * PartialFlipcoef)
			End If
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  SLINGSHOT CORRECTION FUNCTIONS
'******************************************************
' To add these slingshot corrections:
'	 - On the table, add the endpoint primitives that define the two ends of the Slingshot
'	 - Initialize the SlingshotCorrection objects in InitSlingCorrection
'	 - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
	LS.Object = LeftSlingshot
	LS.EndPoint1 = EndPoint1LS
	LS.EndPoint2 = EndPoint2LS
	
	RS.Object = RightSlingshot
	RS.EndPoint1 = EndPoint1RS
	RS.EndPoint2 = EndPoint2RS
	
	'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
	' These values are best guesses. Retune them if needed based on specific table research.
	AddSlingsPt 0, 0.00, - 4
	AddSlingsPt 1, 0.45, - 7
	AddSlingsPt 2, 0.48,	0
	AddSlingsPt 3, 0.52,	0
	AddSlingsPt 4, 0.55,	7
	AddSlingsPt 5, 1.00,	4
End Sub

Sub AddSlingsPt(idx, aX, aY)		'debugger wrapper for adjusting flipper script in-game
	Dim a
	a = Array(LS, RS)
	Dim x
	For Each x In a
		x.addpoint idx, aX, aY
	Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
'	dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
'	dcos = cos(degrees * Pi/180)
'End Function
'
Function RotPoint(x,y,angle)
	dim rx, ry
	rx = x*dCos(angle) - y*dSin(angle)
	ry = x*dSin(angle) + y*dCos(angle)
	RotPoint = Array(rx,ry)
End Function

Class SlingshotCorrection
	Public DebugOn, Enabled
	Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2
	
	Public ModIn, ModOut
	
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
		Enabled = True
	End Sub
	
	Public Property Let Object(aInput)
		Set Slingshot = aInput
	End Property
	
	Public Property Let EndPoint1(aInput)
		SlingX1 = aInput.x
		SlingY1 = aInput.y
	End Property
	
	Public Property Let EndPoint2(aInput)
		SlingX2 = aInput.x
		SlingY2 = aInput.y
	End Property
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If gametime > 100 Then Report
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
	
	
	Public Sub VelocityCorrect(aBall)
		Dim BallPos, XL, XR, YL, YR
		
		'Assign right and left end points
		If SlingX1 < SlingX2 Then
			XL = SlingX1
			YL = SlingY1
			XR = SlingX2
			YR = SlingY2
		Else
			XL = SlingX2
			YL = SlingY2
			XR = SlingX1
			YR = SlingY1
		End If
		
		'Find BallPos = % on Slingshot
		If Not IsEmpty(aBall.id) Then
			If Abs(XR - XL) > Abs(YR - YL) Then
				BallPos = PSlope(aBall.x, XL, 0, XR, 1)
			Else
				BallPos = PSlope(aBall.y, YL, 0, YR, 1)
			End If
			If BallPos < 0 Then BallPos = 0
			If BallPos > 1 Then BallPos = 1
		End If
		
		'Velocity angle correction
		If Not IsEmpty(ModIn(0) ) Then
			Dim Angle, RotVxVy
			Angle = LinearEnvelope(BallPos, ModIn, ModOut)
			'   debug.print " BallPos=" & BallPos &" Angle=" & Angle 
			'   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely 
			RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
			If Enabled Then aBall.Velx = RotVxVy(0)
			If Enabled Then aBall.Vely = RotVxVy(1)
			'   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely 
			'   debug.print " " 
		End If
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY. RUBBER DAMPENER, AND SLINGSHOT CORRECTION SUPPORTING FUNCTIONS 
'******************************************************

Sub AddPt(aStr, idx, aX, aY)	'debugger wrapper for adjusting flipper script in-game
	Dim a
	a = Array(LF, RF)
	Dim x
	For Each x In a
		x.addpoint aStr, idx, aX, aY
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, ByVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray) 'Shuffle objects in a temp array
		If Not IsEmpty(aArray(x) ) Then
			If IsObject(aArray(x)) Then
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	If offset < 0 Then offset = 0
	ReDim aArray(aCount - 1 + offset)   'Resize original array
	For x = 0 To aCount - 1 'set objects back into original array
		If IsObject(a(x)) Then
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
	BallSpeed = Sqr(ball.VelX ^ 2 + ball.VelY ^ 2 + ball.VelZ ^ 2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)  'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m * X2
	Y = M * x + b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
	Public Property Let Data(aBall)
		With aBall
			x = .x
			y = .y
			z = .z
			velx = .velx
			vely = .vely
			velz = .velz
			id = .ID
			mass = .mass
			radius = .radius
		End With
	End Property
	Public Sub Reset()
		x = Empty
		y = Empty
		z = Empty
		velx = Empty
		vely = Empty
		velz = Empty
		id = Empty
		mass = Empty
		radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	Dim y 'Y output
	Dim L 'Line
	Dim ii
	For ii = 1 To UBound(xKeyFrame) 'find active line
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)	'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L - 1), yLvl(L - 1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )	'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )	'Clamp upper
	
	LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS 
'******************************************************

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b
	'   Dim BOT
	'   BOT = GetBalls
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub

'*****************
' Maths
'*****************

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
	dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
	dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		End If
	ElseIf dx = 0 Then
		If dy = 0 Then
			Atn2 = 0
		Else
			Atn2 = Sgn(dy) * pi / 2
		End If
	End If
End Function


Function max(a,b)
	if a > b then 
		max = a
	Else
		max = b
	end if
end Function




'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point is px,py
	DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	
	If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1 'EM's to late 80's
Const EOSTnew = 0.8 '90's and later
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
	Case 0
	SOSRampup = 2.5
	Case 1
	SOSRampup = 6
	Case 2
	SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity
	
	Flipper.eostorque = EOST
	Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST * EOSReturn / FReturn
	
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b', BOT
		'		BOT = GetBalls
		
		For b = 0 To UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper
	
	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3 * Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
		If FCount = 0 Then FCount = GameTime
		
		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 30  'minimum distance in vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle)	'-1 for Right Flipper
	Dim LiveCatchBounce																														'If live catch is not perfect, it won't freeze ball totally
	Dim CatchTime
	CatchTime = GameTime - FCount
	
	If CatchTime <= LiveCatch And parm > 6 And Abs(Flipper.x - ball.x) > LiveDistanceMin And Abs(Flipper.x - ball.x) < LiveDistanceMax Then
		If CatchTime <= LiveCatch * 0.5 Then												'Perfect catch only when catch time happens in the beginning of the window
			LiveCatchBounce = 0
		Else
			LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)		'Partial catch when catch happens a bit late
		End If
		
		If LiveCatchBounce = 0 And ball.velx * Dir > 0 Then ball.velx = 0
		ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
		ball.angmomx = 0
		ball.angmomy = 0
		ball.angmomz = 0
	Else
		If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf Activeball, parm
	End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************

'******************************************************
'****  PHYSICS DAMPENERS
'******************************************************
' These are data mined bounce curves, 
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
	RubbersD.dampen Activeball
	TargetBouncer Activeball, 1
End Sub

Sub dSleeves_Hit(idx)
	SleevesD.Dampen Activeball
	TargetBouncer Activeball, 0.7
End Sub

Dim RubbersD				'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	  'debug, reports in debugger (in vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1		 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	   'there's clamping so interpolate up to 56 at least

Dim SleevesD	'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	  'debug, reports in debugger (in vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn   'tbpOut.text
	Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
	End Sub
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If gametime > 100 Then Report
	End Sub
	
	Public Sub Dampen(aBall)
		If threshold Then
			If BallSpeed(aBall) < threshold Then Exit Sub
		End If
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " in vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef
		aBall.vely = aBall.vely * coef
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
	Public ballvel, ballvelx, ballvely
	
	Private Sub Class_Initialize
		ReDim ballvel(0)
		ReDim ballvelx(0)
		ReDim ballvely(0)
	End Sub
	
	Public Sub Update()	'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = getballs
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)	'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)	'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)	'set bounds
		
		For Each b In allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
'	Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************


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

'******************************************************
'****  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 To tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate()
	Dim b
	'   Dim BOT
	'   BOT = GetBalls
	
	' stop the sound of deleted balls
	For b = UBound(gBOT) + 1 To tnob - 1
		' Comment the next line if you are not implementing Dyanmic Ball Shadows
		If AmbientBallShadowOn = 0 Then BallShadowA(b).visible = 0
		rolling(b) = False
		StopSound("BallRoll_" & b)
	Next
	
	' exit the sub if no balls on the table
	If UBound(gBOT) =  - 1 Then Exit Sub
	
	' play the rolling sound for each ball
	For b = 0 To UBound(gBOT)
		If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If
		
		' Ball Drop Sounds
		If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If gBOT(b).velz >  - 7 Then
					RandomSoundBallBouncePlayfieldSoft gBOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard gBOT(b)
				End If
			End If
		End If
		
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If
		
		' "Static" Ball Shadows
		' Comment the next If block, if you are not implementing the Dynamic Ball Shadows
		If AmbientBallShadowOn = 0 Then
			If gBOT(b).Z > 30 Then
				BallShadowA(b).height = gBOT(b).z - BallSize / 4		'This is technically 1/4 of the ball "above" the ramp, but it keeps it from clipping the ramp
			Else
				BallShadowA(b).height = 0.1
			End If
			BallShadowA(b).Y = gBOT(b).Y + offsetY
			BallShadowA(b).X = gBOT(b).X + offsetX
			BallShadowA(b).visible = 1
		End If
	Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************


'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.  
' Create the following new collections:
'	 Metals (all metal objects, metal walls, metal posts, metal wire guides)
'	 Apron (the apron walls and plunger wall)
'	 Walls (all wood or plastic walls)
'	 Rollovers (wire rollover triggers, star triggers, or button triggers)
'	 Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'	 Gates (plate gates)
'	 GatesWire (wire gates)
'	 Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.  
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).  
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate 
' how to make these updates. But in summary the following needs to be updated:	
'	- Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
'	- Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
'	- Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
'	- Ball rolling sounds sub
'
' Tutorial vides by Apophis
' Part 1:	 https://youtu.be/PbE2kNiam3g
' Part 2:	 https://youtu.be/B5cm1Y8wQsk
' Part 3:	 https://youtu.be/eLhWyuYOyGg


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1					  'volume level; range [0, 1]
NudgeLeftSoundLevel = 1				 'volume level; range [0, 1]
NudgeRightSoundLevel = 1				'volume level; range [0, 1]
NudgeCenterSoundLevel = 1			   'volume level; range [0, 1]
StartButtonSoundLevel = 0.1			 'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1			   'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010		'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635		'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0					   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45					'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel		'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel	   'sound helper; not configurable
SlingshotSoundLevel = 0.95					  'volume level; range [0, 1]
BumperSoundFactor = 4.25						'volume multiplier; must not be zero
KnockerSoundLevel = 1						   'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2		  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5			 'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5			   'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5			'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025		   'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025		   'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8	  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075				   'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5			'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10	 'volume multiplier; must not be zero
DTSoundLevel = 0.25				 'volume multiplier; must not be zero
RolloverSoundLevel = 0.25		   'volume level; range [0, 1]
SpinnerSoundLevel = 0.5			 'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8				   'volume level; range [0, 1]
BallReleaseSoundLevel = 1			   'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2	'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015	 'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5			 'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, aVol * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, - 1, aVol * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, aVol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	Playsound soundname, 1,aVol * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.y * 2 / tableheight - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 10)
	Else
		AudioFade = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.x * 2 / tablewidth - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 10)
	Else
		AudioPan = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
	BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
	RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
	RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
	PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm / 10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm / 10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
	Dim finalspeed
	finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 5 Then
		RandomSoundRubberStrong 1
	End If
	If finalspeed <= 5 Then
		RandomSoundRubberWeak()
	End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd * 10) + 1
		Case 1
		PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 2
		PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 3
		PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 4
		PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 5
		PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 6
		PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 7
		PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 8
		PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 9
		PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 10
		PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
	RandomSoundWall()
End Sub

Sub RandomSoundWall()
	Dim finalspeed
	finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 5) + 1
			Case 1
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 4) + 1
			Case 1
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 3) + 1
			Case 1
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
			PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
	RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
	RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
	Dim finalspeed
	finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
			PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
			PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
			PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
			PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(activeball.id) < 4) And cor.ballvely(activeball.id) > 7 Then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
	Dim finalspeed
	finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
			PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2
			PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	Dim finalspeed
	finalspeed = Sqr(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
	If finalspeed > 10 Then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft Activeball
	Else
		RandomSoundTargetHitWeak()
	End If
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd * 9) + 1
		Case 1
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9
		PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd * 5) + 1
		Case 1
		PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2
		PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3
		PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4
		PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5
		PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, Activeball
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, Activeball
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
	SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
	If Activeball.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If activeball.velx <  - 8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If Activeball.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If activeball.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, Activeball
End Sub

Sub SoundSaucerLockAtBall(ball)
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ball
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0
		PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1
		PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

'Sub OnBallBallCollision(ball1, ball2, velocity)
'	Dim snd
'	Select Case Int(Rnd * 7) + 1
'		Case 1
'		snd = "Ball_Collide_1"
'		Case 2
'		snd = "Ball_Collide_2"
'		Case 3
'		snd = "Ball_Collide_3"
'		Case 4
'		snd = "Ball_Collide_4"
'		Case 5
'		snd = "Ball_Collide_5"
'		Case 6
'		snd = "Ball_Collide_6"
'		Case 7
'		snd = "Ball_Collide_7"
'	End Select
'	
'	PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
'End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
	PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
	PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05	  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
	Select Case toggle
		Case 1
		PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
		Case 0
		PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
	End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
	Select Case toggle
		Case 1
		PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
		Case 0
		PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
	End Select
End Sub

'/////////////////////////////////////////////////////////////////
'					End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'******************************************************
'**** RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'		  * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'		  * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'		  * Create a Timer called RampRoll, that is enabled, with a interval of 100
'		  * Set RampBAlls and RampType variable to Total Number of Balls
'	Usage:
'		  * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'		  * To stop tracking ball
'				 * call WireRampOff
'				 * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(7,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(7)

Sub WireRampOn(input)
	Waddball ActiveBall, input
	RampRollUpdate
End Sub

Sub WireRampOff()
	WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
	' This will loop through the RampBalls array checking each element of the array x, position 1
	' To see if the the ball was already added to the array.
	' If the ball is found then exit the subroutine
	Dim x
	For x = 1 To UBound(RampBalls)	'Check, don't add balls twice
		If RampBalls(x, 1) = input.id Then
			If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub	'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next
	
	' This will itterate through the RampBalls Array.
	' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
	' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
	' The RampType(BallId) is set to RampInput
	' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
	For x = 1 To UBound(RampBalls)
		If IsEmpty(RampBalls(x, 1)) Then
			Set RampBalls(x, 0) = input
			RampBalls(x, 1) = input.ID
			RampType(x) = RampInput
			RampBalls(x, 2) = 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1	 'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			Exit Sub
		End If
		If x = UBound(RampBalls) Then	 'debug
			Debug.print "WireRampOn error, ball queue is full: " & vbNewLine & _
			RampBalls(0, 0) & vbNewLine & _
			TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
			TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
			TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
			TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
			TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
			" "
		End If
	Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
	'   Debug.Print "In WRemoveBall() + Remove ball from loop array"
	Dim ballcount
	ballcount = 0
	Dim x
	For x = 1 To UBound(RampBalls)
		If ID = RampBalls(x, 1) Then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
	Next
	If BallCount = 0 Then RampBalls(0,0) = False	'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
	RampRollUpdate
End Sub

Sub RampRollUpdate()	'Timer update
	Dim x
	For x = 1 To UBound(RampBalls)
		If Not IsEmpty(RampBalls(x,1) ) Then
			If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
				If RampType(x) Then
					PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
				End If
				RampBalls(x, 2) = RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			End If
			If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then	'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
	Next
	If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()	'debug textbox
	Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
	"1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
	"2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
	"3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
	"4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
	"5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
	"6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
	" "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
	BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
	BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************


'******************************************************
' VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1	  '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.7	 'Level of bounces. Recommmended value of 0.7

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		vel = BallSpeed(aBall)
		If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
		Select Case Int(Rnd * 6) + 1
			Case 1
			zMultiplier = 0.2 * defvalue
			Case 2
			zMultiplier = 0.25 * defvalue
			Case 3
			zMultiplier = 0.3 * defvalue
			Case 4
			zMultiplier = 0.4 * defvalue
			Case 5
			zMultiplier = 0.45 * defvalue
			Case 6
			zMultiplier = 0.5 * defvalue
		End Select
		aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
		aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
		aBall.vely = aBall.velx * vratio
		'   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'   debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
	TargetBouncer activeball, 1
End Sub

'*******************************************
'  Timers
'*******************************************

Sub GameTimer_Timer() 'The game timer interval; should be 10 ms
	Cor.Update	  'update ball tracking (this sometimes goes in the RDampen_Timer sub)
	RollingUpdate   'update rolling sounds
	DoSTAnim		'handle stand up target animations
	DoDTAnim
	UpdateTargets
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




Sub ConfigureDevices
    'Ball Devices
    Dim bd_plunger: Set bd_plunger = (new BallDevice)("bd_plunger")
    With bd_plunger
        .BallSwitches = Array("sw_plunger")
        .EjectTimeout = 1
        .EjectCallback = "PlungerKickBall"
        .MechcanicalEject = True
        .DefaultDevice = True
        .Debug = True
    End With

    Dim bd_cave_scoop: Set bd_cave_scoop = (new BallDevice)("cave_scoop")
    With bd_cave_scoop
        .BallSwitches = Array("sw39")
        .EjectTimeout = 2
        .EjectCallback = "CaveKickBall"
        .Debug = True
    End With

    Dim bd_waterfall_vuk: Set bd_waterfall_vuk = (new BallDevice)("bd_waterfall_vuk")
    With bd_waterfall_vuk
        .BallSwitches = Array("sw46")
        .EjectTimeout = 1
        .EjectCallback = "WaterfallVukKickBall"
        .Debug = True
    End With
    'Diverters
    Dim dv_panther : Set dv_panther = (new Diverter)("dv_panther", Array("ball_started"), Array("ball_ended"))', Array("activate_panther"), Array("deactivate_panther"), 0, False
    dv_panther.ActionCallback = "MovePanther"

    Dim dv_leftorbit : Set dv_leftorbit = (new Diverter)("leftorbit", Array("enable_waterfall"), Array("multiball_waterfall_started"))
    With dv_leftorbit
        .ActivationTime = 2000
        .ActivationSwitches = Array("sw47")
        .ActionCallback = "MoveLeftOrbitDiverter"
        .Debug = True
    End With
    

    Dim dv_waterfall : Set dv_waterfall = (new Diverter)("dv_waterfall", Array("game_started"), Array())
    With dv_waterfall
        .ActivationTime = 2000
        .ActivateEvents = Array("multiball_waterfall_started", "game_ended")
        .ActionCallback = "WaterfallRelease"
        .Debug = True
    End With

    'Drop Targets
    Dim dt_map1 : Set dt_map1 = (new DropTarget)(sw04, sw04a, BM_sw04, 4, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map2 : Set dt_map2 = (new DropTarget)(sw05, sw05a, BM_sw05, 5, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map3 : Set dt_map3 = (new DropTarget)(sw06, sw06a, BM_sw06, 6, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map4 : Set dt_map4 = (new DropTarget)(sw08, sw08a, BM_sw08, 8, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map5 : Set dt_map5 = (new DropTarget)(sw09, sw09a, BM_sw09, 9, 0, False, Array("ball_started"," machine_reset_phase_3"))
    Dim dt_map6 : Set dt_map6 = (new DropTarget)(sw10, sw10a, BM_sw10, 10, 0, False, Array("ball_started"," machine_reset_phase_3"))

    Dim DT01, DT02
    Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True, Null) 
    Set DT02 = (new DropTarget)(sw02, sw02a, BM_sw02, 2, 0, True, Null)

    
    DTArray = Array(DT01, DT02, dt_map1, dt_map2, dt_map3, dt_map4, dt_map5, dt_map6)

End Sub

'Set up ball devices
Function PlungerKickBall(ball)
    dim rangle
    rangle = PI * (0 - 90) / 180
    ball.vely = sin(rangle)*50
    SoundSaucerKick 1, ball
End Function

Function CaveKickBall(ball)
    ball.z = ball.z + 30
    ball.velz = 60        
    SoundSaucerKick 1, ball
End Function

Function WaterfallVukKickBall(ball)
    'ball.z = ball.z + 30
    'ball.velz = 1        
    SoundSaucerKick 1, ball
    sw46.Kick 0, 45, 1.36
End Function


'Set up diverters
Sub MovePanther(enabled)
    If enabled Then
        DTRaise 1
    Else
        DTDrop 1
    End If
End Sub

Sub MoveLeftOrbitDiverter(enabled)
    waterfalldiverter.isdropped=enabled
End Sub

Sub WaterfallRelease(enabled)
    sw44.isdropped=enabled
End Sub



'***********************************************************************************
'***** Switches                                                         	    ****
'***********************************************************************************

Sub sw11_Hit()
    STHit 11
End Sub

Sub sw11o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw12_Hit()
    STHit 12
End Sub

Sub sw12o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw13_Hit()
    STHit 13
End Sub

Sub sw13o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw15_Hit()
    STHit 15
End Sub

Sub sw15o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw16_Hit()
    STHit 16
End Sub

Sub sw16o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw17_Hit()
    STHit 17
End Sub

Sub sw17o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw41_Hit()
    STHit 41
End Sub

Sub sw01_Hit()
    DTHit 1
End Sub

Sub sw02_Hit()
    DTHit 2
End Sub

Sub sw04_Hit()
    DTHit 4
End Sub

Sub sw05_Hit()
    DTHit 5
End Sub

Sub sw06_Hit()
    DTHit 6
End Sub


Sub sw08_Hit()
    DTHit 8
End Sub

Sub sw09_Hit()
    DTHit 9
End Sub

Sub sw10_Hit()
    DTHit 10
End Sub

Sub sw45_Hit()
    'DTHit 45
End Sub

Sub Wall015_Hit()
    ActiveBall.velx = 0
    ActiveBall.vely = 0
End Sub

Sub sw39_Hit()   : DispatchPinEvent "sw39_active",   ActiveBall : End Sub
Sub sw39_Unhit() : DispatchPinEvent "sw39_inactive", ActiveBall : End Sub

Sub sw46_Hit()   : DispatchPinEvent "sw46_active",   ActiveBall : End Sub
Sub sw46_Unhit() : DispatchPinEvent "sw46_inactive", ActiveBall : End Sub

Sub sw99_Hit()   : DispatchPinEvent "sw99_active",   Null : End Sub
Sub sw99_Unhit() : DispatchPinEvent "sw99_inactive", Null : End Sub

Sub sw44_Hit()   : DispatchPinEvent "sw44_active",   Null : End Sub
Sub sw44_Unhit() : DispatchPinEvent "sw44_inactive", Null : End Sub

Sub sw45_Spin()   : DispatchPinEvent "sw45_active",   Null : End Sub

Sub sw47_Hit()   : DispatchPinEvent "sw47_active",   Null : End Sub
Sub sw47_Unhit() : DispatchPinEvent "sw47_inactive", Null : End Sub

Sub DTAction(switchid, enabled)
    If enabled = 1 Then
        Select Case switchid
            case 1:
                DispatchPinEvent "sw01_inactive", Null
            case 2:
                DispatchPinEvent "sw02_inactive", Null
            case 4:
                DispatchPinEvent "sw04_active", Null
            case 5:
                DispatchPinEvent "sw05_active", Null
            case 6:
                DispatchPinEvent "sw06_active", Null
            case 8:
                DispatchPinEvent "sw08_active", Null
            case 9:
                DispatchPinEvent "sw09_active", Null
            case 10:
                DispatchPinEvent "sw10_active", Null
        End Select
    ElseIf enabled = 0 Then
        Select Case switchid
            case 1:
                DispatchPinEvent "sw01_active", Null
            case 2:
                DispatchPinEvent "sw02_active", Null
            case 4:
                DispatchPinEvent "sw04_inactive", Null
            case 5:
                DispatchPinEvent "sw05_inactive", Null
            case 6:
                DispatchPinEvent "sw06_inactive", Null
            case 8:
                DispatchPinEvent "sw08_inactive", Null
            case 9:
                DispatchPinEvent "sw09_inactive", Null
            case 10:
                DispatchPinEvent "sw10_inactive", Null            
        End Select
    End If
End Sub


Sub STAction(switchid, enabled)
    If enabled = 1 Then
        Select Case switchid
            case 11:
                DispatchPinEvent "sw11_active", Null
            case 12:
                DispatchPinEvent "sw12_active", Null
            case 13:
                DispatchPinEvent "sw13_active", Null
            case 15:
                DispatchPinEvent "sw15_active", Null
            case 16:
                DispatchPinEvent "sw16_active", Null
            case 17:
                DispatchPinEvent "sw17_active", Null
        End Select
    ElseIf enabled = 0 Then
        Select Case switchid
            case 11:
                DispatchPinEvent "sw11_inactive", Null
            case 12:
                DispatchPinEvent "sw12_inactive", Null
            case 13:
                DispatchPinEvent "sw13_inactive", Null
            case 15:
                DispatchPinEvent "sw15_inactive", Null
            case 16:
                DispatchPinEvent "sw16_inactive", Null
            case 17:
                DispatchPinEvent "sw17_inactive", Null
        End Select
    End If
End Sub

'Switches

Sub sw_plunger_Hit()   : DispatchPinEvent "sw_plunger_active",   ActiveBall : End Sub
Sub sw_plunger_Unhit() : DispatchPinEvent "sw_plunger_inactive", ActiveBall : End Sub

Sub s_start_Hit()   : DispatchPinEvent "s_start_active",   ActiveBall : End Sub
Sub s_start_Unhit() : DispatchPinEvent "s_start_inactive", ActiveBall : End Sub





'VPX Game Logic Framework (https://mpcarr.github.io/vpx-gle-framework/)

'
Dim glf_currentPlayer : glf_currentPlayer = Null
Dim glf_canAddPlayers : glf_canAddPlayers = True
Dim glf_PI : glf_PI = 4 * Atn(1)
Dim glf_plunger
Dim glf_gameStarted : glf_gameStarted = False
Dim glf_pinEvents : Set glf_pinEvents = CreateObject("Scripting.Dictionary")
Dim glf_pinEventsOrder : Set glf_pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim glf_playerEvents : Set glf_playerEvents = CreateObject("Scripting.Dictionary")
Dim Glf_EventBlocks : Set Glf_EventBlocks = CreateObject("Scripting.Dictionary")
Dim Glf_ShotProfiles : Set Glf_ShotProfiles = CreateObject("Scripting.Dictionary")
Dim Glf_ShowStartQueue : Set Glf_ShowStartQueue = CreateObject("Scripting.Dictionary")
Dim glf_playerEventsOrder : Set glf_playerEventsOrder = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")
Dim glf_Modes : Set glf_Modes = CreateObject("Scripting.Dictionary")
Dim bcpController : bcpController = Null
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = ""
Dim lightCtrl : Set lightCtrl = new LStateController
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0

Dim glf_ballsPerGame : glf_ballsPerGame = 3
Dim glf_troughSize : glf_troughSize = tnob

Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False

lightCtrl.RegisterLights Glf_Lights
lightCtrl.Debug = False
Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController
    Set bcpController = (new GlfVpxBcpController)(bcpPort	, bcpExeName)
End Sub

Public Sub Glf_Init()
	If useBCP = True Then
		Glf_ConnectToBCPMediaController
	End If

	swTrough1.DestroyBall
	swTrough2.DestroyBall
	swTrough3.DestroyBall
	swTrough4.DestroyBall
	swTrough5.DestroyBall
	swTrough6.DestroyBall
	swTrough7.DestroyBall
	swTrough8.DestroyBall
	If glf_troughSize > 0 Then : Set glf_ball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1) : End If
	If glf_troughSize > 1 Then : Set glf_ball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2) : End If
	If glf_troughSize > 2 Then : Set glf_ball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3) : End If
	If glf_troughSize > 3 Then : Set glf_ball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4) : End If
	If glf_troughSize > 4 Then : Set glf_ball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5) : End If
	If glf_troughSize > 5 Then : Set glf_ball6 = swTrough6.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6) : End If
	If glf_troughSize > 6 Then : Set glf_ball7 = swTrough7.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7) : End If
	If glf_troughSize > 7 Then : Set glf_ball8 = swTrough8.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8) : End If
	
	Dim switch, switchHitSubs
	switchHitSubs = ""
	For Each switch in Glf_Switches
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_Hit() : DispatchPinEvent """ & switch.Name & "_active"", ActiveBall : End Sub" & vbCrLf
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_UnHit() : DispatchPinEvent """ & switch.Name & "_inactive"", ActiveBall : End Sub" & vbCrLf
	Next
	ExecuteGlobal switchHitSubs
End Sub

Sub Glf_Options(ByVal eventId)
	Dim ballsPerGame : ballsPerGame = Table1.Option("Balls Per Game", 1, 2, 1, 1, 0, Array("3 Balls", "5 Balls"))
	If ballsPerGame = 1 Then
		glf_ballsPerGame = 3
	Else
		glf_ballsPerGame = 5
	End If

	Dim glfDebug : glfDebug = Table1.Option("Glf Debug Log", 0, 1, 1, 1, 0, Array("Off", "On"))
	If glfDebug = 1 Then
		glf_debugEnabled = True
	Else
		glf_debugEnabled = False
	End If

	Dim glfuseBCP : glfuseBCP = Table1.Option("Glf Backbox Control Protocol", 0, 1, 1, 1, 0, Array("Off", "On"))
	If glfuseBCP = 1 Then
		useBCP = True
		If Not IsNull(bcpController) Then
			bcpController.Disconnect
			bcpController = Null
		End If
		Glf_ConnectToBCPMediaController
	Else
		useBCP = False
		If Not IsNull(bcpController) Then
			bcpController.Disconnect
			bcpController = Null
		End If
	End If
End Sub

Public Sub Glf_Exit()
	If Not IsNull(bcpController) Then
		bcpController.Disconnect
		bcpController = Null
	End If
End Sub

Public Sub Glf_KeyDown(ByVal keycode)
    If glf_gameStarted = True Then
		If keycode = LeftFlipperKey Then
			DispatchPinEvent "s_left_flipper_active", Null
		End If
		
		If keycode = RightFlipperKey Then
			DispatchPinEvent "s_right_flipper_active", Null
		End If
		
		If keycode = StartGameKey Then
			If glf_canAddPlayers = True Then
				Glf_AddPlayer()
			End If
		End If
	Else
		If keycode = StartGameKey Then
			Glf_AddPlayer()
			Glf_StartGame()
		End If
	End If
End Sub

Public Sub Glf_KeyUp(ByVal keycode)
	If glf_gameStarted = True Then
		If KeyCode = PlungerKey Then
			Plunger.Fire
		End If

		If keycode = LeftFlipperKey Then
			DispatchPinEvent "s_left_flipper_inactive", Null
		End If
		
		If keycode = RightFlipperKey Then
			DispatchPinEvent "s_right_flipper_inactive", Null
		End If
	End If
End Sub

Public Sub Glf_GameTimer_Timer()
	
End Sub

Public Sub Glf_EventTimer_Timer()
	DelayTick
End Sub

Public Sub Glf_BCPUpdateTimer_Timer()
	Glf_BcpUpdate
End Sub

Public Sub Glf_LampTimer_Timer()
	lightCtrl.Update
End Sub

Public Function Glf_ParseInput(value, isTime)
	Dim templateCode : templateCode = ""
	Dim tmp: tmp = value
    Select Case VarType(value)
        Case 8 ' vbString
			tmp = Glf_ReplaceCurrentPlayerAttributes(tmp)
			If InStr(tmp, " if ") Then
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & Glf_ConvertIf(tmp, "Glf_" & glf_FuncCount) & vbCrLf
				templateCode = templateCode & "End Function"
			Else
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
				templateCode = templateCode & "End Function"
			End IF
        Case Else
			templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
			If isTime Then
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp * 1000 & vbCrLf
			Else
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
			End If
			templateCode = templateCode & "End Function"
    End Select
	'msgbox templateCode
	ExecuteGlobal templateCode
	Dim funcRef : funcRef = "Glf_" & glf_FuncCount
	glf_FuncCount = glf_FuncCount + 1
	Glf_ParseInput = Array(funcRef, value)
End Function

Public Function Glf_ParseEventInput(value)
	Dim templateCode : templateCode = ""
	Dim condition : condition = Glf_IsCondition(value)
	If IsNull(condition) Then
		Glf_ParseEventInput = Array(value, value, Null)
	Else
		dim conditionReplaced : conditionReplaced = Glf_ReplaceCurrentPlayerAttributes(condition)
		templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
		templateCode = templateCode & vbTab & Glf_ConvertCondition(conditionReplaced, "Glf_" & glf_FuncCount) & vbCrLf
		templateCode = templateCode & "End Function"
		'msgbox templateCode
		ExecuteGlobal templateCode
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		glf_FuncCount = glf_FuncCount + 1
		Glf_ParseEventInput = Array(Replace(value, "{"&condition&"}", funcRef) ,Replace(value, "{"&condition&"}", ""), funcRef)
	End If
End Function

Function Glf_ReplaceCurrentPlayerAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "current_player\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    replacement = "GetPlayerState(""$1"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceCurrentPlayerAttributes = outputString
End Function

Function Glf_ConvertIf(value, retName)
    Dim parts, condition, truePart, falsePart
    parts = Split(value, " if ")
    truePart = Trim(parts(0))
    Dim conditionAndFalsePart
    conditionAndFalsePart = Split(parts(1), " else ")
    condition = Trim(conditionAndFalsePart(0))
    falsePart = Trim(conditionAndFalsePart(1))
    Dim vbscriptIfStatement
    vbscriptIfStatement = "If " & condition & " Then" & vbCrLf & _
                          "    "&retName&" = " & truePart & vbCrLf & _
                          "Else" & vbCrLf & _
                          "    "&retName&" = " & falsePart & vbCrLf & _
                          "End If"
	Glf_ConvertIf = vbscriptIfStatement
End Function

Function Glf_ConvertCondition(value, retName)
	value = Replace(value, "==", "=")
	value = Replace(value, "!=", "<>")
	Glf_ConvertCondition = "    "&retName&" = " & value
End Function



Function GlfShotProfiles(name)
	If Glf_ShotProfiles.Exists(name) Then
		Set GlfShotProfiles = Glf_ShotProfiles(name)
	Else
		Dim new_shotprofile : Set new_shotprofile = (new GlfShotProfile)(name)
		Glf_ShotProfiles.Add name, new_shotprofile
		Set GlfShotProfiles = new_shotprofile
	End If
End Function

Function CreateGlfMode(name, priority)
	If Not glf_Modes.Exists(name) Then 
		Dim mode : Set mode = (new Mode)(name, priority)
		glf_Modes.Add name, mode
		Set CreateGlfMode = mode
	End If
End Function

Function GlfModes(name)
	If glf_Modes.Exists(name) Then 
		Set GlfModes = glf_Modes(name)
	Else
		GlfModes = Null
	End If
End Function

Function GlfKwargs()
	Set GlfKwargs = CreateObject("Scripting.Dictionary")
End Function

Function Glf_ConvertShow(show, tokens)

	Dim showStep, light, lightsCount, x,tagLight, tagLights, lightParts, token, stepIdx
	Dim newShow
	ReDim newShow(UBound(show))
	stepIdx = 0
	For Each showStep in show
		lightsCount = 0 
		For Each light in showStep
			lightParts = Split(light, "|")
			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
					tagLights = lightCtrl.GetLightsForTag(lightParts(0))
					lightsCount = UBound(tagLights)+1
				Else
					If IsNull(token) Then
						lightsCount = lightsCount + 1
					Else
						'resolve token lights
						If IsNull(lightCtrl.GetLightIdx(tokens(token))) Then
							'token is a tag
							tagLights = lightCtrl.GetLightsForTag(tokens(token))
							lightsCount = UBound(tagLights)+1
						Else
							lightsCount = lightsCount + 1
						End If
					End If
				End If
			End If
		Next
	
		Dim seqArray
		ReDim seqArray(lightsCount-1)
		x=0
		For Each light in showStep
			lightParts = Split(light, "|")
			Dim lightColor : lightColor = ""
			If Ubound(lightParts) = 2 Then 
				If IsNull(Glf_IsToken(lightParts(2))) Then
					lightColor = lightParts(2)
				Else
					lightColor = tokens(Glf_IsToken(lightParts(2)))
				End If
			End If

			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
					tagLights = lightCtrl.GetLightsForTag(lightParts(0))
					For Each tagLight in tagLights
						If UBound(lightParts) >=1 Then
							seqArray(x) = tagLight & "|"&lightParts(1)&"|"&lightColor
						Else
							seqArray(x) = tagLight & "|"&lightParts(1)
						End If
						x=x+1
					Next
				Else
					If IsNull(token) Then
						If UBound(lightParts) >= 1 Then
							seqArray(x) = lightParts(0) & "|"&lightParts(1)&"|"&lightColor
						Else
							seqArray(x) = lightParts(0) & "|"&lightParts(1)
						End If
						x=x+1
					Else
						'resolve token lights
						If IsNull(lightCtrl.GetLightIdx(tokens(token))) Then
							'token is a tag
							tagLights = lightCtrl.GetLightsForTag(tokens(token))
							For Each tagLight in tagLights
								If UBound(lightParts) >=1 Then
									seqArray(x) = tagLight & "|"&lightParts(1)&"|"&lightColor
								Else
									seqArray(x) = tagLight & "|"&lightParts(1)
								End If
								x=x+1
							Next
						Else
							If UBound(lightParts) >= 1 Then
								seqArray(x) = tokens(token) & "|"&lightParts(1)&"|"&lightColor
							Else
								seqArray(x) = tokens(token) & "|"&lightParts(1)
							End If
							x=x+1
						End If
					End If
				End If
			End If
		Next
		glf_debugLog.WriteToLog "Convert Show", Join(seqArray)
		newShow(stepIdx) = seqArray
		stepIdx = stepIdx + 1
	Next
	Glf_ConvertShow = newShow
End Function

Private Function Glf_IsToken(mainString)
	' Check if the string contains an opening parenthesis and ends with a closing parenthesis
	If InStr(mainString, "(") > 0 And Right(mainString, 1) = ")" Then
		' Extract the substring within the parentheses
		Dim startPos, subString
		startPos = InStr(mainString, "(")
		subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
		Glf_IsToken = subString
	Else
		Glf_IsToken = Null
	End If
End Function

Private Function Glf_IsCondition(mainString)
	' Check if the string contains an opening { and ends with a closing }
	If InStr(mainString, "{") > 0 And Right(mainString, 1) = "}" Then
		Dim startPos, subString
		startPos = InStr(mainString, "{")
		subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
		Glf_IsCondition = subString
	Else
		Glf_IsCondition = Null
	End If
End Function

Function Glf_RotateArray(arr, direction)
    Dim n, rotatedArray, i
    ReDim rotatedArray(UBound(arr))
 
    If LCase(direction) = "l" Then
        For i = 0 To UBound(arr) - 1
            rotatedArray(i) = arr(i + 1)
        Next
        rotatedArray(UBound(arr)) = arr(0)
    ElseIf LCase(direction) = "r" Then
        For i = UBound(arr) To 1 Step -1
            rotatedArray(i) = arr(i - 1)
        Next
        rotatedArray(0) = arr(UBound(arr))
    Else
        ' Invalid direction
        Glf_RotateArray = arr
        Exit Function
    End If
    
    ' Return the rotated array
    Glf_RotateArray = rotatedArray
End Function

Function Glf_CopyArray(arr)
    Dim newArr, i
    ReDim newArr(UBound(arr))
    For i = 0 To UBound(arr)
        newArr(i) = arr(i)
    Next
    Glf_CopyArray = newArr
End Function

Function Glf_IsInArray(value, arr)
    Dim i
    Glf_IsInArray = False

    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            Glf_IsInArray = True
            Exit Function
        End If
    Next
End Function

'******************************************************
'*****   GLF Shows 		                           ****
'******************************************************

Dim glf_ShowOn : glf_ShowOn = Array(Array("(lights)|100"))
Dim glf_ShowOff : glf_ShowOff = Array(Array("(lights)|0"))
Dim glf_ShowFlash : glf_ShowFlash = Array(Array("(lights)|100"), Array("(lights)|0"))
Dim glf_ShowFlashColor : glf_ShowFlashColor = Array(Array("(lights)|100|(color)"), Array("(lights)|0|(color)"))

With GlfShotProfiles("default")
	With .States("on")
			.Show = glf_ShowFlash
	End With
	With .States("off")
			.Show = glf_ShowOff
	End With
End With

With GlfShotProfiles("flash_color")
	With .States("on")
			.Show = glf_ShowFlashColor
	End With
	With .States("off")
			.Show = glf_ShowOff
	End With
End With


'******************************************************
'*****   GLF Pin Events                            ****
'******************************************************

Const GLF_GAME_STARTED = "game_started"
Const GLF_GAME_OVER = "game_ended"
Const GLF_BALL_ENDED = "ball_ended"
Const GLF_NEXT_PLAYER = "next_player"
Const GLF_BALL_DRAIN = "ball_drain"
Const GLF_BALL_STARTED = "ball_started"

'******************************************************
'*****   GLF Player State                          ****
'******************************************************

Const GLF_SCORE = "score"
Const GLF_CURRENT_BALL = "current_ball"
Const GLF_INITIALS = "initials"




'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************

Class GlfVpxBcpController

    Private m_bcpController, m_connected

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_server.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        Glf_BCPUpdateTimer.Enabled = True
        If Err Then Debug.print("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

	Public Sub Send(commandMessage)
		If m_connected Then
            m_bcpController.Send commandMessage
        End If
	End Sub

    Public Function GetMessages
		If m_connected Then
            GetMessages = m_bcpController.GetMessages
        End If
	End Function

    Public Sub Reset()
		If m_connected Then
            m_bcpController.Send "reset"
        End If
	End Sub
    
    Public Sub PlaySlide(slide, context, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""" & slide & """: {""action"": ""play"", ""expire"": 0}}, ""context"": """ & context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub SendPlayerVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & Getglf_currentPlayerNumber
            '06:34:34.644 : VERBOSE : BCP : Received BCP command: ball_start?player_num=int:1&ball=int:1
        End If
	End Sub

    Private Function EncodeVariable(value)
        Dim retValue
        Select Case VarType(value)
            Case vbInteger, vbLong
                retValue = "int:" & value
            Case vbSingle, vbDouble
                retValue = "float:" & value
            Case vbString
                retValue = "string:" & value
            Case vbBoolean
                retValue = "bool:" & CStr(value)
            Case Else
                retValue = "NoneType:"
        End Select
        EncodeVariable = retValue
    End Function
    
    Private Function VariableVariance(v1, v2)
        Dim retValue
        Select Case VarType(v1)
            Case vbInteger, vbLong, vbSingle, vbDouble
                retValue = Abs(v1 - v2)
            Case Else
                retValue = True 
        End Select
        VariableVariance = retValue
    End Function

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
            Glf_BCPUpdateTimer.Enabled = False
        End If
    End Sub
End Class

Sub Glf_BcpSendPlayerVar(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim player_var : player_var = kwargs(0)
    Dim value : value = kwargs(1)
    Dim prevValue : prevValue = kwargs(2)
    bcpController.SendPlayerVariable player_var, value, prevValue
End Sub

Sub Glf_BcpAddPlayer(playerNum)
    If useBcp Then
        bcpController.Send("player_added?player_num=int:"&playerNum)
    End If
End Sub

Sub Glf_BcpUpdate()
    If IsNull(bcpController) Then
        Exit Sub
    End If
    Dim messages : messages = bcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter
        For Each message in messages
            Select Case message.Command
                case "hello"
                    bcpController.Reset
                case "monitor_start"
                    Dim category : category = message.GetValue("category")
                    If category = "player_vars" Then
                        AddPlayerStateEventListener "score", "bcp_player_var_score", "BcpSendPlayerVar", 1000, Null
                        AddPlayerStateEventListener "current_ball", "bcp_player_var_ball", "BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    Dim eventName : eventName = message.GetValue("event")
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************
Class BallSave

    Private m_name
    Private m_mode
    Private m_active_time
    Private m_grace_period
    Private m_enable_events
    Private m_timer_start_events
    Private m_auto_launch
    Private m_balls_to_save
    Private m_saving_balls
    Private m_enabled
    Private m_timer_started
    Private m_tick
    Private m_in_grace
    Private m_in_hurry_up
    Private m_hurry_up_time
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AutoLaunch(): AutoLaunch = m_auto_launch: End Property
    Public Property Let ActiveTime(value) : m_active_time = Glf_ParseInput(value, True) : End Property
    Public Property Let GracePeriod(value) : m_grace_period = Glf_ParseInput(value, True) : End Property
    Public Property Let HurryUpTime(value) : m_hurry_up_time = Glf_ParseInput(value, True) : End Property
    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let TimerStartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_timer_start_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let AutoLaunch(value) : m_auto_launch = value : End Property
    Public Property Let BallsToSave(value) : m_balls_to_save = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "ball_saves_" & name
        m_mode = mode.Name
        m_active_time = Null
	    m_grace_period = Null
        m_hurry_up_time = Null
        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_timer_start_events = CreateObject("Scripting.Dictionary")
	    m_auto_launch = False
	    m_balls_to_save = 1
        m_enabled = False
        m_timer_started = False
        m_debug = False
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "BallSaveEventHandler", mode.Priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "BallSaveEventHandler", mode.Priority, Array("deactivate", Me)
	  Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events.Keys
            AddPinEventListener m_enable_events(evt).EventName, m_name & "_enable", "BallSaveEventHandler", 1000, Array("enable", Me, evt)
        Next
        For Each evt in m_timer_start_events.Keys
            AddPinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start", "BallSaveEventHandler", 1000, Array("timer_start", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events.Keys
            RemovePinEventListener m_enable_events(evt).EventName, m_name & "_enable"
        Next
        For Each evt in m_timer_start_events.Keys
            RemovePinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start"
        Next
    End Sub

    Public Sub Enable(evt)
        If m_enabled = True Then
            Exit Sub
        End If
        If Not IsNull(m_enable_events(evt).Condition) Then
            If GetRef(m_enable_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        m_enabled = True
        m_saving_balls = m_balls_to_save
        Log "Enabling. Auto launch: "&m_auto_launch&", Balls to save: "&m_balls_to_save
        AddPinEventListener "ball_drain", m_name & "_ball_drain", "BallSaveEventHandler", 1000, Array("drain", Me)
        DispatchPinEvent m_name&"_enabled", Null
        If UBound(m_timer_start_events.Keys) = -1 Then
            Log "Timer Starting as no timer start events are set"
            TimerStart()
        End If
    End Sub

    Public Sub Disable
        'Disable ball save
        If m_enabled = False Then
            Exit Sub
        End If
        m_enabled = False
        m_saving_balls = m_balls_to_save
        m_timer_started = False
        Log "Disabling..."
        RemovePinEventListener "ball_drain", m_name & "_ball_drain"
        RemoveDelay "_ball_saves_"&m_name&"_disable"
        RemoveDelay m_name&"_grace_period"
        RemoveDelay m_name&"_hurry_up_time"
            
    End Sub

    Sub Drain(ballsToSave)
        If m_enabled = True And ballsToSave > 0 Then
            If m_saving_balls > 0 Then
                m_saving_balls = m_saving_balls -1
            End If
            Log "Ball(s) drained while active. Requesting new one(s). Auto launch: "& m_auto_launch
            DispatchPinEvent m_name&"_saving_ball", Null
            SetDelay m_name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", Me),Null), 1000
            If m_saving_balls = 0 Then
                Disable()
            End If
        End If
    End Sub

    Public Sub TimerStart
        'Start the timer.
        'This is usually called after the ball was ejected while the ball save may have been enabled earlier.
        If m_timer_started=True Or m_enabled=False Then
            Exit Sub
        End If
        m_timer_started=True
        DispatchPinEvent m_name&"_timer_start", Null
        If Not IsNull(m_active_time) Then
            Dim active_time : active_time = GetRef(m_active_time(0))()
            Dim grace_period, hurry_up_time
            If Not IsNull(m_grace_period) Then
                grace_period = GetRef(m_grace_period(0))()
            Else
                grace_period = 0
            End If
            If Not IsNull(m_hurry_up_time) Then
                hurry_up_time = GetRef(m_hurry_up_time(0))()
            Else
                hurry_up_time = 0
            End If
            Log "Starting ball save timer: " & active_time
            Log "gametime: "& gametime & ". disabled at: " & gametime+active_time+grace_period
            SetDelay m_name&"_disable", "BallSaveEventHandler" , Array(Array("disable", Me),Null), active_time+grace_period
            SetDelay m_name&"_grace_period", "BallSaveEventHandler", Array(Array("grace_period", Me),Null), active_time
            SetDelay m_name&"_hurry_up_time", "BallSaveEventHandler", Array(Array("hurry_up_time", Me), Null), active_time-hurry_up_time
        End If
    End Sub

    Public Sub EnterGracePeriod
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Sub EnterHurryUpTime
        DispatchPinEvent m_name & "_hurry_up_time", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = Replace(m_name, "ball_saves_", "") & ":" & vbCrLf
        yaml = yaml & "    active_time: " & m_active_time(1) & "s" & vbCrLf
        yaml = yaml & "    grace_period: " & m_grace_period(1) & "s" & vbCrLf
        yaml = yaml & "    hurry_up_time: " & m_hurry_up_time(1) & "s" & vbCrLf
        yaml = yaml & "    enable_events: "
        Dim evt,x : x = 0
        For Each evt in m_enable_events.Keys
            yaml = yaml & m_enable_events(evt).Raw
            If x <> UBound(m_enable_events.Keys) Then
                yaml = yaml & ", "
            End If
            x = x +1
        Next
        yaml = yaml & vbCrLf
        yaml = yaml & "    timer_start_events: "
        x=0
        For Each evt in m_timer_start_events.Keys
            yaml = yaml & m_timer_start_events(evt).Raw
            If x <> UBound(m_timer_start_events.Keys) Then
                yaml = yaml & ", "
            End If
            x = x +1
        Next
        yaml = yaml & vbCrLf
        yaml = yaml & "    auto_launch: " & m_auto_launch & vbCrLf
        yaml = yaml & "    balls_to_save: " & m_balls_to_save & vbCrLf
        yaml = yaml & vbCrLf
        ToYaml = yaml
    End Function
End Class

Function BallSaveEventHandler(args)
    Dim ownProps, ballsToSave : ownProps = args(0) : ballsToSave = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballSave : Set ballSave = ownProps(1)
    Select Case evt
        Case "activate"
            ballSave.Activate
        Case "deactivate"
            ballSave.Deactivate
        Case "enable"
            ballSave.Enable ownProps(2)
        Case "disable"
            ballSave.Disable
        Case "grace_period"
            ballSave.EnterGracePeriod
        Case "hurry_up_time"
            ballSave.EnterHurryUpTime
        Case "drain"
            If ballsToSave > 0 Then
                ballSave.Drain ballsToSave
                ballsToSave = ballsToSave - 1
            End If
        Case "timer_start"
            ballSave.TimerStart
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True Then
                Glf_ReleaseBall(Null)
                If ballSave.AutoLaunch = True Then
                    SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave),Null), 500
                End If
            Else
                SetDelay ballSave.Name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", ballSave), Null), 1000
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave), Null), 500
            End If
    End Select
    BallSaveEventHandler = ballsToSave
End Function

Class GlfCounter

    Private m_name
    Private m_priority
    Private m_mode
    Private m_enable_events
    Private m_count_events
    Private m_count_complete_value
    Private m_disable_on_complete
    Private m_reset_on_complete
    Private m_events_when_complete
    Private m_persist_state
    Private m_debug

    Private m_count

    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CountEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_count_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CountCompleteValue(value) : m_count_complete_value = value : End Property
    Public Property Let DisableOnComplete(value) : m_disable_on_complete = value : End Property
    Public Property Let ResetOnComplete(value) : m_reset_on_complete = value : End Property
    Public Property Let EventsWhenComplete(value) : m_events_when_complete = value : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "counter_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_count = -1
        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_count_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "CounterEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "CounterEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub SetValue(value)
        If value = "" Then
            value = 0
        End If
        m_count = value
        If m_persist_state Then
            SetPlayerState m_name & "_state", m_count
        End If
    End Sub

    Public Sub Activate()
        If m_persist_state And m_count > -1 Then
            If Not IsNull(GetPlayerState(m_name & "_state")) Then
                SetValue GetPlayerState(m_name & "_state")
            Else
                SetValue 0
            End If
        Else
            SetValue 0
        End If
        Dim evt
        For Each evt in m_enable_events.Keys
            AddPinEventListener m_enable_events(evt).EventName, m_name & "_enable", "CounterEventHandler", m_priority, Array("enable", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        If Not m_persist_state Then
            SetValue -1
        End If
        Dim evt
        For Each evt in m_enable_events.Keys
            RemovePinEventListener m_enable_events(evt).EventName, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_count_events.Keys
            AddPinEventListener m_count_events(evt).EventName, m_name & "_count", "CounterEventHandler", m_priority, Array("count", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_count_events.Keys
            RemovePinEventListener m_count_events(evt).EventName, m_name & "_count"
        Next
    End Sub

    Public Sub Count()
        Log "counting: old value: "& m_count & ", new Value: " & m_count+1 & ", target: "& m_count_complete_value
        SetValue m_count + 1
        If m_count = m_count_complete_value Then
            Dim evt
            For Each evt in m_events_when_complete
                DispatchPinEvent evt, Null
            Next
            If m_disable_on_complete Then
                Disable()
            End If
            If m_reset_on_complete Then
                SetValue 0
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function CounterEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim counter : Set counter = ownProps(1)
    Select Case evt
        Case "activate"
            counter.Activate
        Case "deactivate"
            counter.Deactivate
        Case "enable"
            counter.Enable
        Case "count"
            counter.Count
    End Select
    CounterEventHandler = kwargs
End Function


Class GlfEventPlayer

    Private m_priority
    Private m_mode
    private m_base_device
    Private m_events

    Public Property Get Events() : Set Events = m_events : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_event_player_play", "EventPlayerEventHandler", m_priority, Array("play", Me, m_events(evt))
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_event_player_play"
        Next
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml & "    shots: " & Join(m_shots, ",") & vbCrLf
    End Function

End Class

Function EventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            dim evtToFire
            For Each evtToFire in ownProps(2)
                DispatchPinEvent evtToFire, Null
            Next
    End Select
    EventPlayerEventHandler = kwargs
End Function

Class GlfLightPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug
    Private m_name
    Private m_value

    Public Property Let Events(value)
        Set m_events = value
        Dim evt
        For Each evt in m_events
            lightCtrl.CreateSeqRunner m_name & "_" & evt, m_priority
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_name = "light_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "light_player_activate", "LightPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "light_player_deactivate", "LightPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_light_player_play"
            PlayOff evt, m_events(evt)
        Next
    End Sub

    Public Sub Add(evt, lights)
        
        Dim light, lightsCount, x,tagLight, tagLights, lightParts
        lightsCount = 0
        For Each light in lights
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            Else
                If IsNull(lightCtrl.GetLightIdx(lightParts)) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts)
                    Log "Tag Lights: " & Join(tagLights)
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            End If
        Next
        Log "Adding " & lightsCount & " lights for event: " & evt 
        Dim seqArray
        ReDim seqArray(lightsCount-1)
        x=0
        For Each light in lights
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        If UBound(lightParts) >=1 Then
                            seqArray(x) = tagLight & "|100|"&lightParts(1)
                        Else
                            seqArray(x) = tagLight & "|100"
                        End If
                        x=x+1
                    Next
                Else
                    If UBound(lightParts) >= 1 Then
                        seqArray(x) = lightParts(0) & "|100|"&lightParts(1)
                    Else
                        seqArray(x) = lightParts(0) & "|100"
                    End If
                    x=x+1
                End If
            Else
                If IsNull(lightCtrl.GetLightIdx(lightParts)) Then
                    tagLights = lightCtrl.GetLightsForTag(lightParts)
                    For Each tagLight in tagLights
                        seqArray(x) = tagLight & "|100"
                        x=x+1
                    Next
                Else
                    seqArray(x) = lightParts & "|100"
                    x=x+1
                End If
            End If
        Next
        Log "Light List: " & Join(seqArray)
        m_events.Add evt, Array(lights, Array(seqArray))
    End Sub

    Public Sub Play(evt, lights)
        Dim light, lightParts
        lightCtrl.AddLightSeq m_name & "_" & evt, evt, lights(1), -1, 1, Null, m_priority, 0
        For Each light in lights(0)
            lightParts = Split(light, "|")
            If IsArray(lightParts) Then
                If IsNull(lightCtrl.GetLightIdx(lightParts(0))) Then
                    Dim tagLight, tagLights
                    tagLights = lightCtrl.GetLightsForTag(lightParts(0))
                    For Each tagLight in tagLights
                        ProcessLight tagLight, lightParts, evt
                    Next
                Else
                    ProcessLight lightParts(0), lightParts, evt
                End If
            End If
        Next
    End Sub

    Public Sub PlayOff(evt, lights)
        Dim light
        lightCtrl.RemoveLightSeq m_name & "_" & evt, evt
    End Sub

    Private Sub ProcessLight(light_name, light_props, evt)
        If UBound(light_props) = 2 Then
            lightCtrl.FadeLightToColor light_name, light_props(1), light_props(2), m_name & "_" & evt & "_" & light_name, m_priority
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_light_player", message
        End If
    End Sub
End Class

Function LightPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim LightPlayer : Set LightPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            LightPlayer.Activate
        Case "deactivate"
            LightPlayer.Deactivate
        Case "play"
            LightPlayer.Play ownProps(3), ownProps(2)
    End Select
    LightPlayerEventHandler = Null
End Function


Class GlfBaseModeDevice

    Private m_mode
    Private m_priority
    Private m_enable_events
    Private m_disable_events
    Private m_device
    Private m_parent
    Private m_debug

    Public Property Get EnableEvents(): Set EnableEvents = m_enable_events: End Property
    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Get DisableEvents(): Set DisableEvents = m_disable_events: End Property
    Public Property Let DisableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get Mode(): Set Mode = m_mode: End Property

    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode, device, parent)
        Set m_mode = mode
        m_priority = mode.Priority
        m_device = device
        Set m_parent = parent
        m_debug = mode.Debug

        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_disable_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode.Name & "_starting", m_device & "_" & m_parent.Name & "_activate", "BaseModeDeviceEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode.Name & "_stopping", m_device & "_" & m_parent.Name & "_deactivate", "BaseModeDeviceEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_enable_events.Keys()
            AddPinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable", "BaseModeDeviceEventHandler", m_priority, Array("enable", m_parent, evt)
        Next
        For Each evt In m_disable_events.Keys()
            AddPinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable", "BaseModeDeviceEventHandler", m_priority, Array("disable", m_parent, evt)
        Next
        m_parent.Activate
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        m_parent.Disable()
        Dim evt
        For Each evt In m_enable_events.Keys()
            RemovePinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable"
        Next
        For Each evt In m_disable_events.Keys()
            RemovePinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable"
        Next
        m_parent.Deactivate
    End Sub

    Public Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode.Name & m_device & "_" & m_parent.Name & "_play", message
        End If
    End Sub
End Class


Function BaseModeDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Select Case evt
        Case "activate"
            device.Activate
        Case "deactivate"
            device.Deactivate
        Case "enable"
            device.Enable
        Case "disable"
            device.Disable
    End Select
    BaseModeDeviceEventHandler = kwargs
End Function

Class Mode

    Private m_name
    Private m_start_events
    Private m_stop_events
    private m_priority
    Private m_debug

    Private m_ballsaves
    Private m_counters
    Private m_multiball_locks
    Private m_multiballs
    Private m_shots
    Private m_shot_groups
    Private m_timers
    Private m_lightplayer
    Private m_showplayer
    Private m_variableplayer
    Private m_eventplayer

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Get Debug(): Debug = m_debug: End Property
    Public Property Get LightPlayer(): Set LightPlayer = m_lightplayer: End Property
    Public Property Get ShowPlayer(): Set ShowPlayer = m_showplayer: End Property
    Public Property Get EventPlayer(): Set EventPlayer = m_eventplayer.Events: End Property
    Public Property Get VariablePlayer(): Set VariablePlayer = m_variableplayer: End Property
    

    Public Property Get BallSaves(name)
        If m_ballsaves.Exists(name) Then
            Set BallSaves = m_ballsaves(name)
        Else
            Dim new_ballsave : Set new_ballsave = (new BallSave)(name, Me)
            m_ballsaves.Add name, new_ballsave
            Set BallSaves = new_ballsave
        End If
    End Property

    Public Property Get Timers(name)
        If m_timers.Exists(name) Then
            Set Timers = m_timers(name)
        Else
            Dim new_timer : Set new_timer = (new GlfTimer)(name, Me)
            m_timers.Add name, new_timer
            Set Timers = new_timer
        End If
    End Property

    Public Property Get Counters(name)
        If m_counters.Exists(name) Then
            Set Counters = m_counters(name)
        Else
            Dim new_counter : Set new_counter = (new GlfCounter)(name, Me)
            m_counters.Add name, new_counter
            Set Counters = new_counter
        End If
    End Property

    Public Property Get MultiballLocks(name)
        If m_multiball_locks.Exists(name) Then
            Set MultiballLocks = m_multiball_locks(name)
        Else
            Dim new_multiball_lock : Set new_multiball_lock = (new GlfMultiballLocks)(name, Me)
            m_multiball_locks.Add name, new_multiball_lock
            Set MultiballLocks = new_multiball_lock
        End If
    End Property

    Public Property Get Multiballs(name)
        If m_multiballs.Exists(name) Then
            Set Multiballs = m_multiballs(name)
        Else
            Dim new_multiball : Set new_multiball = (new GlfMultiballs)(name, Me)
            m_multiballs.Add name, new_multiball
            Set Multiballs = new_multiball
        End If
    End Property

    Public Property Get Shots(name)
        If m_shots.Exists(name) Then
            Set Shots = m_shots(name)
        Else
            Dim new_shot : Set new_shot = (new GlfShot)(name, Me)
            m_shots.Add name, new_shot
            Set Shots = new_shot
        End If
    End Property

    Public Property Get ShotGroups(name)
        If m_shot_groups.Exists(name) Then
            Set ShotGroups = m_shot_groups(name)
        Else
            Dim new_shot_group : Set new_shot_group = (new GlfShotGroup)(name, Me)
            m_shot_groups.Add name, new_shot_group
            Set ShotGroups = new_shot_group
        End If
    End Property

    Public Property Let StartEvents(value)
        m_start_events = value
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_start", "ModeEventHandler", m_priority, Array("start", Me)
        Next
    End Property
    
    Public Property Let StopEvents(value)
        m_stop_events = value
        Dim evt
        For Each evt in m_stop_events
            AddPinEventListener evt, m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me)
        Next
        AddPinEventListener "ball_ended", m_name & "_stop", "ModeEventHandler", m_priority+1, Array("stop", Me)
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_priority = priority
        Set m_ballsaves = CreateObject("Scripting.Dictionary")
        Set m_counters = CreateObject("Scripting.Dictionary")
        Set m_timers = CreateObject("Scripting.Dictionary")
        Set m_multiball_locks = CreateObject("Scripting.Dictionary")
        Set m_multiballs = CreateObject("Scripting.Dictionary")
        Set m_shots = CreateObject("Scripting.Dictionary")
        Set m_shot_groups = CreateObject("Scripting.Dictionary")
        Set m_lightplayer = (new GlfLightPlayer)(Me)
        Set m_showplayer = (new GlfShowPlayer)(Me)
        Set m_eventplayer = (new GlfEventPlayer)(Me)
        Set m_variableplayer = (new GlfVariablePlayer)(Me)
        Set Init = Me
	End Function

    Public Sub StartMode()
        Log "Starting"
        DispatchPinEvent m_name & "_starting", Null
        DispatchPinEvent m_name & "_started", Null
        Log "Started"
    End Sub

    Public Sub StopMode()
        Log "Stopping"
        DispatchPinEvent m_name & "_stopping", Null
        DispatchPinEvent m_name & "_stopped", Null
        Log "Stopped"
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        dim yaml, child
        yaml = "mode:" & vbCrLf
        yaml = yaml & "  start_events: " & Join(m_start_events, ",") & vbCrLf
        yaml = yaml & "  stop_events: " & Join(m_stop_events, ",") & vbCrLf
        yaml = yaml & "  priority: " & m_priority & vbCrLf
        yaml = yaml & vbCrLf
        If UBound(m_ballsaves.Keys)>-1 Then
            yaml = yaml & "ballsaves: " & vbCrLf
            For Each child in m_ballsaves.Keys
                yaml = yaml & m_ballsaves(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        If UBound(m_shots.Keys)>-1 Then
            yaml = yaml & "shots: " & vbCrLf
            For Each child in m_shots.Keys
                yaml = yaml & m_shots(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        If UBound(m_shot_groups.Keys)>-1 Then
            yaml = yaml & "shot_groups: " & vbCrLf
            For Each child in m_shot_groups.Keys
                yaml = yaml & m_shot_groups(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        ToYaml = yaml
        Log yaml
    End Function
End Class

Function ModeEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim mode : Set mode = ownProps(1)
    Select Case evt
        Case "start"
            mode.StartMode
        Case "stop"
            mode.StopMode
        Case "started"
            DispatchPinEvent mode.Name & "_started", Null
        Case "stopped"
            DispatchPinEvent mode.Name & "_stopped", Null
    End Select
    ModeEventHandler = kwargs
End Function

Class GlfMultiballLocks

    Private m_name
    Private m_priority
    Private m_mode
    Private m_enable_events
    Private m_disable_events
    Private m_balls_to_lock
    Private m_balls_locked
    Private m_lock_events
    Private m_reset_events
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let BallsToLock(value) : m_balls_to_lock = value : End Property
    Public Property Let LockEvents(value) : m_lock_events = value : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "multiball_locks_" & name
        m_mode = mode.Name
        m_lock_events = Array()
        m_reset_events = Array()
        m_balls_to_lock = 0

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "MultiballLocksHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "MultiballLocksHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MultiballLocksHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_lock_events
            AddPinEventListener evt, m_name & "_ball_locked", "MultiballLocksHandler", m_priority, Array("lock", Me)
        Next
        For Each evt in m_reset_events
            AddPinEventListener evt, m_name & "_reset", "MultiballLocksHandler", m_priority, Array("reset", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_lock_events
            RemovePinEventListener evt, m_name & "_ball_locked"
        Next
    End Sub

    Public Sub Lock()
        Dim balls_locked
        If IsNull(GetPlayerState(m_name & "_balls_locked")) Then
            balls_locked = 1
        Else
            balls_locked = GetPlayerState(m_name & "_balls_locked") + 1
        End If
        SetPlayerState m_name & "_balls_locked", balls_locked
        DispatchPinEvent m_name & "_locked_ball", balls_locked
        Log CStr(balls_locked)
        glf_BIP = glf_BIP - 1
        If balls_locked = m_balls_to_lock Then
            DispatchPinEvent m_name & "_full", balls_locked
        Else
            SetDelay m_name & "_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", Me),Null), 1000
        End If
    End Sub

    Public Sub Reset
        SetPlayerState m_name & "_balls_locked", 0
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MultiballLocksHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    Select Case evt
        Case "activate"
            multiball.Activate
        Case "deactivate"
            multiball.Deactivate
        Case "enable"
            multiball.Enable
        Case "disable"
            multiball.Disable
        Case "lock"
            multiball.Lock
        Case "reset"
            multiball.Reset
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True Then
                ReleaseBall(Null)
                SetDelay multiball.Name&"_auto_launch", "MultiballLocksHandler" , Array(Array("auto_launch", multiball),Null), 500
            Else
                SetDelay multiball.Name&"_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", multiball), Null), 1000
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay multiball.Name&"_auto_launch", "MultiballLocksHandler" , Array(Array("auto_launch", multiball), Null), 500
            End If
    End Select
    MultiballLocksHandler = kwargs
End Function

Class GlfMultiball

    Private m_name
    Private m_priority
    Private m_mode
    Private m_multiball_locks
    Private m_enable_events
    Private m_disable_events
    Private m_start_events
    Private m_ball_save
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let StartEvents(value) : m_start_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, multiball_locks, mode)
        m_name = "multiball_" & name
        m_mode = mode.Name
        m_start_events = Array()
        m_multiball_locks = multiball_locks
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "MultiballHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "MultiballHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MultiballHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_starting", "MultiballHandler", m_priority, Array("start", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_start_events
            RemovePinEventListener evt, m_name & "_starting"
        Next
    End Sub

    Public Sub StartMultiball()
        glf_BIP = glf_BIP + GetPlayerState(m_multiball_locks & "_balls_locked")
        DispatchPinEvent m_name & "_started", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MultiballHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    Select Case evt
        Case "activate"
            multiball.Activate
        Case "deactivate"
            multiball.Deactivate
        Case "enable"
            multiball.Enable
        Case "disable"
            multiball.Disable
        Case "start"
            multiball.StartMultiball
    End Select
    MultiballHandler = kwargs
End Function
Class GlfShotGroup
    Private m_name
    Private m_mode
    Private m_priority
    private m_base_device
    private m_shots
    private m_common_state
    Private m_enable_rotation_events
    Private m_disable_rotation_events
    Private m_restart_events
    Private m_reset_events
    Private m_rotate_events
    Private m_rotate_left_events
    Private m_rotate_right_events
    Private rotation_enabled
    Private m_temp_shots
    Private m_rotation_pattern
    Private m_rotation_enabled
 
    Public Property Get Name(): Name = m_name: End Property
    Public Property Get CommonState()
        Dim state : state = m_base_device.Mode.Shots(m_shots(0)).State
        Dim shot
        For Each shot in m_shots
            If state <> m_base_device.Mode.Shots(shot).State Then
                CommonState = Empty
                Exit Property
            End If
        Next
        CommonState = state
    End Property
 
    Public Property Let Shots(value)
        m_shots = value
        Dim shot_name
        For Each shot_name in m_shots
            AddPinEventListener "shot_" & shot_name & "_hit", m_name & "_" & m_mode & "_hit", "ShotGroupEventHandler", m_priority, Array("hit", Me)
        Next
        m_rotation_pattern = Glf_CopyArray(Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).RotationPattern)
    End Property
 
    Public Property Let EnableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_rotation_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let DisableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_rotation_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RotateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RotateLeftEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_left_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RotateRightEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_right_events.Add newEvent.Name, newEvent
        Next
    End Property
 
	Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_common_state = Empty
        m_rotation_enabled = True
        m_rotation_pattern = Empty
 
        Set m_enable_rotation_events = CreateObject("Scripting.Dictionary")
        Set m_disable_rotation_events = CreateObject("Scripting.Dictionary")
        Set m_restart_events = CreateObject("Scripting.Dictionary")
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_left_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_right_events = CreateObject("Scripting.Dictionary")
        Set m_temp_shots = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot_group", Me)
 
        Set Init = Me
	End Function
 
    Public Sub Activate
        For Each evt in m_enable_rotation_events.Keys()
            m_rotation_enabled = False
            AddPinEventListener m_enable_rotation_events(evt).EventName, m_name & "_enable_rotation", "ShotGroupEventHandler", m_priority, Array("enable_rotation", Me, evt)
        Next
        For Each evt in m_disable_rotation_events.Keys()
            AddPinEventListener m_disable_rotation_events(evt).EventName, m_name & "_disable_rotation", "ShotGroupEventHandler", m_priority, Array("disable_rotation", Me, evt)
        Next
        For Each evt in m_restart_events.Keys()
            AddPinEventListener m_restart_events(evt).EventName, m_name & "_restart", "ShotGroupEventHandler", m_priority, Array("restart", Me, evt)
        Next
        For Each evt in m_reset_events.Keys()
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_reset", "ShotGroupEventHandler", m_priority, Array("reset", Me, evt)
        Next
        For Each evt in m_rotate_events.Keys()
            AddPinEventListener m_rotate_events(evt).EventName, m_name & "_rotate", "ShotGroupEventHandler", m_priority, Array("rotate", Me, evt)
        Next
        For Each evt in m_rotate_left_events.Keys()
            AddPinEventListener m_rotate_left_events(evt).EventName, m_name & "_rotate_left", "ShotGroupEventHandler", m_priority, Array("rotate_left", Me, evt)
        Next
        For Each evt in m_rotate_right_events.Keys()
            AddPinEventListener m_rotate_right_events(evt).EventName, m_name & "_rotate_right", "ShotGroupEventHandler", m_priority, Array("rotate_right", Me, evt)
        Next
    End Sub
 
    Public Sub Deactivate
        m_rotation_enabled = True
        For Each evt in m_enable_rotation_events.Keys()
            RemovePinEventListener m_enable_rotation_events(evt).EventName, m_name & "_enable_rotation"
        Next
        For Each evt in m_disable_rotation_events.Keys()
            RemovePinEventListener m_disable_rotation_events(evt).EventName, m_name & "_disable_rotation"
        Next
        For Each evt in m_restart_events.Keys()
            RemovePinEventListener m_restart_events(evt).EventName, m_name & "_restart"
        Next
        For Each evt in m_reset_events.Keys()
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_reset"
        Next
        For Each evt in m_rotate_events.Keys()
            RemovePinEventListener m_rotate_events(evt).EventName, m_name & "_rotate"
        Next
        For Each evt in m_rotate_left_events.Keys()
            RemovePinEventListener m_rotate_left_events(evt).EventName, m_name & "_rotate_left"
        Next
        For Each evt in m_rotate_right_events.Keys()
            RemovePinEventListener m_rotate_right_events(evt).EventName, m_name & "_rotate_right"
        Next
    End Sub
 
    Private Function CheckForComplete()
        Dim state : state = CommonState()
        If state = m_common_state Then
            Exit Function
        End If
 
        m_common_state = state
 
        If state = Empty Then
            Exit Function
        End If
 
        Dim state_name : state_name = Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).StateName(m_state)
 
        Log "Shot group is complete with state: " & state_name
        Dim kwargs : Set kwargs = GlfKwargs()
		With kwargs
            .Add "state", state_name
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        DispatchPinEvent m_name & "_" & state_name & "_complete", Null
 
    End Function
 
    Public Sub Enable()
        Dim shot
        m_base_device.Log "Enabling"
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Enable()
        Next
        Dim evt
    End Sub
 
    Public Sub Disable()
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Disable()
        Next
    End Sub
 
    Public Sub EnableRotation
        m_base_device.Log "Enabling Rotation"
        m_rotation_enabled = True
    End Sub
 
    Public Sub DisableRotation
        m_base_device.Log "Disabling Rotation"
        m_rotation_enabled = False
    End Sub
 
    Public Sub Restart
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Restart()
        Next
    End Sub
 
    Public Sub Reset
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Reset()
        Next
    End Sub
 
    Public Sub Rotate(direction)
 
        If m_rotation_enabled = False Then
            Exit Sub
        End If
        Dim shots_to_rotate : shots_to_rotate = Array()
 
        m_temp_shots.RemoveAll
        Dim shot
        For Each shot in m_shots
            If m_base_device.Mode.Shots(shot).CanRotate Then
                m_temp_shots.Add shot, m_base_device.Mode.Shots(shot)
            End If
        Next
 
        Dim shot_states, x
        x=0
        ReDim shot_states(UBound(m_temp_shots.Keys))
        For Each shot in m_temp_shots.Keys
            shot_states(x) = m_temp_shots(shot).State
            x=x+1
        Next 
 
        If direction = Empty Then
            direction = m_rotation_pattern(0)
            Glf_RotateArray m_rotation_pattern, "l"
        End If
 
        shot_states = Glf_RotateArray(shot_states, direction)
        x=0
        For Each shot in m_temp_shots.Keys
            m_base_device.Log "Rotating Shot:" & shot
            m_temp_shots(shot).Jump shot_states(x), True, False
            x=x+1
        Next 
 
    End Sub
 
    Public Function ToYaml
        Dim yaml
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml & "    shots: " & Join(m_shots, ",") & vbCrLf
 
        If UBound(m_enable_rotation_events.Keys) > -1 Then
            yaml = yaml & "    enable_rotation_events: "
            x=0
            For Each key in m_enable_rotation_events.keys
                yaml = yaml & m_enable_rotation_events(key).Raw
                If x <> UBound(m_enable_rotation_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_disable_rotation_events.Keys) > -1 Then
            yaml = yaml & "    disable_rotation_events: "
            x=0
            For Each key in m_disable_rotation_events.keys
                yaml = yaml & m_disable_rotation_events(key).Raw
                If x <> UBound(m_disable_rotation_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_restart_events.Keys) > -1 Then
            yaml = yaml & "    restart_events: "
            x=0
            For Each key in m_restart_events.keys
                yaml = yaml & m_restart_events(key).Raw
                If x <> UBound(m_restart_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_reset_events.Keys) > -1 Then
            yaml = yaml & "    reset_events: "
            x=0
            For Each key in m_reset_events.keys
                yaml = yaml & m_reset_events(key).Raw
                If x <> UBound(m_reset_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_events.Keys) > -1 Then
            yaml = yaml & "    rotate_events: "
            x=0
            For Each key in m_rotate_events.keys
                yaml = yaml & m_rotate_events(key).Raw
                If x <> UBound(m_rotate_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_left_events.Keys) > -1 Then
            yaml = yaml & "    rotate_left_events: "
            x=0
            For Each key in m_rotate_left_events.keys
                yaml = yaml & m_rotate_left_events(key).Raw
                If x <> UBound(m_rotate_left_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_right_events.Keys) > -1 Then
            yaml = yaml & "    rotate_right_events: "
            x=0
            For Each key in m_rotate_right_events.keys
                yaml = yaml & m_rotate_right_events(key).Raw
                If x <> UBound(m_rotate_right_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_base_device.EnableEvents.Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents.keys
                yaml = yaml & m_base_device.EnableEvents(key).Raw
                If x <> UBound(m_base_device.EnableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_base_device.DisableEvents.Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents.keys
                yaml = yaml & m_base_device.DisableEvents(key).Raw
                If x <> UBound(m_base_device.DisableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        yaml = yaml & "    priority: " & m_priority & vbCrLf
        yaml = yaml & "    rotation_enabled: " & m_rotation_enabled & vbCrLf
 
        ToYaml = yaml
        End Function
End Class
 
Function ShotGroupEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Select Case evt
        Case "hit"
            DispatchPinEvent device.Name & "_hit", Null
            DispatchPinEvent device.Name & "_" & kwargs("state") & "_hit", Null
        Case "enable_rotation"
            device.EnableRotation
        Case "disable_rotation"
            device.DisableRotation
        Case "restart"
            device.Restart
        Case "reset"
            device.Reset
        Case "rotate"
            device.Rotate Empty
        Case "rotate_left"
            device.Rotate "l"
        Case "rotate_right"
            device.Rotate "r"
    End Select
    If IsObject(args(1)) Then
        Set ShotGroupEventHandler = kwargs
    Else
        ShotGroupEventHandler = kwargs
    End If
 
End Function

Class GlfShotProfile

    Private m_name
    Private m_advance_on_hit
    Private m_block
    Private m_loop
    Private m_rotation_pattern
    Private m_states
    Private m_states_not_to_rotate
    Private m_states_to_rotate

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AdvanceOnHit(): AdvanceOnHit = m_advance_on_hit: End Property
    Public Property Get Block(): Block = m_block: End Property
    Public Property Let Block(input): m_block = input: End Property
    Public Property Get ProfileLoop(): ProfileLoop = m_loop: End Property
    Public Property Get RotationPattern(): RotationPattern = m_rotation_pattern: End Property
    Public Property Get States(name)
        If m_states.Exists(name) Then
            Set States = m_states(name)
        Else
            Dim new_state : Set new_state = (new GlfShowPlayerItem)()
            m_states.Add name, new_state
            Set States = new_state
        End If
    End Property
    Public Property Get StateForIndex(index)
        Dim stateItems : stateItems = m_states.Items()
        If UBound(stateItems) >= index Then
            Set StateForIndex = stateItems(index)
        Else
            StateForIndex = Null
        End If
    End Property
    Public Property Get StateName(index)
        Dim stateKeys : stateKeys = m_states.Keys()
        If UBound(stateKeys) >= index Then
            StateName = stateKeys(index)
        Else
            StateName = Empty
        End If
    End Property
    Public Property Get StatesCount()
        StatesCount = UBound(m_states.Keys())
    End Property

    Public Property Get StateNamesToRotate(): StateNamesToRotate = m_states_to_rotate: End Property
    Public Property Let StateNamesToRotate(input): m_states_to_rotate = input: End Property
    Public Property Get StateNamesNotToRotate(): StateNamesNotToRotate = m_states_not_to_rotate: End Property
    Public Property Let StateNamesNotToRotate(input): m_states_not_to_rotate = input: End Property
    
	Public default Function init(name)
        m_name = "shotprofile_" & name
        m_advance_on_hit = True
        m_block = False
        m_loop = False
        m_rotation_pattern = Array("r")
        m_states_to_rotate = Array()
        m_states_not_to_rotate = Array()
        Set m_states = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

End Class

Class GlfShot

    Private m_name
    Private m_mode
    Private m_priority
    Private m_base_device
    Private m_profile
    Private m_control_events
    Private m_advance_events
    Private m_reset_events
    Private m_restart_events
    Private m_switches
    Private m_tokens
    Private m_hit_events
    Private m_start_enabled
    Private m_show_cache
    Private m_state
    Private m_enabled

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get Profile(): Profile = m_profile: End Property
    Public Property Get ShotKey(): ShotKey = m_name & "_" & m_profile: End Property
    Public Property Get State(): State = m_state: End Property
    Public Property Get Tokens() : Set Tokens = m_tokens : End Property
    Public Property Get CanRotate()
        If Glf_IsInArray(Glf_ShotProfiles(m_profile).StateName(m_state), Glf_ShotProfiles(m_profile).StateNamesNotToRotate) Then
            CanRotate = False
        Else
            CanRotate = True
        End If
    End Property
    
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Get ControlEvents(name)
        If m_control_events.Exists(name) Then
            Set ControlEvents = m_control_events(name)
        Else
            Dim newEvent : Set newEvent = (new GlfShotControlEvent)()
            m_control_events.Add name, newEvent
            Set ControlEvents = newEvent
        End If
    End Property
    Public Property Let AdvanceEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_advance_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Name, newEvent
        Next
    End Property   
    Public Property Let Profile(value) : m_profile = value : End Property
    Public Property Let Switch(value) : m_switches = Array(value) : End Property
    Public Property Let Switches(value) : m_switches = value : End Property
    Public Property Let StartEnabled(value) : m_start_enabled = value : End Property
    Public Property Let HitEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_hit_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "shot_" & name
        m_mode = mode.Name
        m_priority = mode.Priority

        m_enabled = False

        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot", Me)

        m_profile = "default"
        m_state = -1
        m_switches = Array()
        m_start_enabled = Empty
        Set m_hit_events = CreateObject("Scripting.Dictionary")
        Set m_tokens = CreateObject("Scripting.Dictionary")
        Set m_show_cache = CreateObject("Scripting.Dictionary")
        Set m_advance_events = CreateObject("Scripting.Dictionary")
        Set m_control_events = CreateObject("Scripting.Dictionary")
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_restart_events = CreateObject("Scripting.Dictionary")

        Set Init = Me
	End Function

    Public Sub Activate()
        m_state = 0
        If m_start_enabled = True Then
            Enable()
        Else
            If IsEmpty(m_start_enabled) And UBound(m_base_device.EnableEvents.Keys) = -1 Then
                Enable()
            End If
        End If
    End Sub

    Public Sub Deactivate()
        m_state = -1
    End Sub

    Public Sub Enable()
        m_base_device.Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_switches
            AddPinEventListener evt & "_active", m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me)
        Next
        For Each evt in m_hit_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me)
        Next
        For Each evt in m_advance_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_advance", "ShotEventHandler", m_priority, Array("advance", Me)
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events
                AddPinEventListener cEvt, m_mode & "_" & m_name & "_control", "ShotEventHandler", m_priority, Array("control", Me, m_control_events(evt))
            Next
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_reset", "ShotEventHandler", m_priority, Array("reset", Me)
        Next
        For Each evt in m_restart_events.Keys
            AddPinEventListener evt, m_mode & "_" & m_name & "_restart", "ShotEventHandler", m_priority, Array("restart", Me)
        Next
        'Play the show for the active state
        PlayShowForState(m_state)
    End Sub

    Public Sub Disable()
        m_base_device.Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_switches
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        For Each evt in m_hit_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        For Each evt in m_advance_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_advance"
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events
                RemovePinEventListener cEvt, m_mode & "_" & m_name & "_control"
            Next
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_reset"
        Next
        For Each evt in m_restart_events.Keys
            RemovePinEventListener evt, m_mode & "_" & m_name & "_restart"
        Next
        Dim x
        For x=0 to Glf_ShotProfiles(m_profile).StatesCount()
            StopShowForState(x)
        Next
    End Sub

    Private Sub StopShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        lightCtrl.RemoveLightSeq m_mode & "_" & m_name, profileState.Key
    End Sub

    Private Sub PlayShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        If IsObject(profileState) Then
            If IsArray(profileState.Show) Then
                If m_show_cache.Exists(CStr(m_state) & "_" & profileState.Key) Then
                    lightCtrl.AddLightSeq m_mode & "_" & m_name, profileState.Key, m_show_cache(CStr(m_state) & "_" & profileState.Key), profileState.Loops, profileState.Speed, Null, m_priority, profileState.SyncMS
                Else
                    Dim key
                    Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
                    Dim profileStateTokens : Set profileStateTokens = profileState.Tokens
                    For Each key In profileStateTokens.Keys
                        mergedTokens.Add key, profileStateTokens(key)
                    Next
                    Dim shotTokens : Set shotTokens = m_tokens
                    For Each key In shotTokens.Keys
                        If mergedTokens.Exists(key) Then
                            mergedTokens(key) = shotTokens(key)
                        Else
                            mergedTokens.Add key, shotTokens(key)
                        End If
                    Next
                    Dim show : show = Glf_ConvertShow(profileState.Show, mergedTokens)
                    m_show_cache.Add CStr(m_state) & "_" & profileState.Key, show
                    lightCtrl.AddLightSeq m_mode & "_" & m_name, profileState.Key, show, profileState.Loops, profileState.Speed, Null, m_priority, profileState.SyncMS
                End If
            End If
        End If
    End Sub

    Public Sub Hit(evt)
        If m_enabled = False Then
            Exit Sub
        End If

        Dim profile : Set profile = Glf_ShotProfiles(m_profile)
        Dim old_state : old_state = m_state
        m_base_device.Log "Hit! Profile: "&m_profile&", State: " & profile.StateName(m_state)

        Dim advancing
        If profile.AdvanceOnHit Then
            m_base_device.Log "Advancing shot because advance_on_hit is True."
            advancing = Advance(False)
        Else
            m_base_device.Log "Not advancing shot"
            advancing = False
        End If

    
        If profile.Block Then
            Glf_EventBlocks(evt).Add Name, True
        Else
            Glf_EventBlocks(evt).Add ShotKey, True
        End If
        Dim kwargs : Set kwargs = GlfKwargs()
		With kwargs
            .Add "profile", m_profile
            .Add "state", profile.StateName(old_state)
            .Add "advancing", advancing
        End With

        DispatchPinEvent m_name & "_hit", kwargs
        DispatchPinEvent m_name & "_" & m_profile & "_hit", kwargs
        DispatchPinEvent m_name & "_" & m_profile & "_" & profile.StateName(old_state) & "_hit", kwargs
        DispatchPinEvent m_name & "_" & profile.StateName(old_state) & "_hit", kwargs
        
    End Sub

    Public Function Advance(force)

        If m_enabled = False And force = False Then
            Advance = False
            Exit Function
        End If
        Dim profile : Set profile = Glf_ShotProfiles(m_profile)

        m_base_device.Log "Advancing 1 step. Profile: "&m_profile&", Current State: " & profile.StateName(m_state)

        If profile.StatesCount() = m_state Then
            If profile.ProfileLoop Then
                StopShowForState(m_state)
                m_state = 0
                PlayShowForState(m_state)
            Else
                Advance = False
                Exit Function
            End If
        Else
            StopShowForState(m_state)
            m_state = m_state + 1
            PlayShowForState(m_state)
        End If

        Advance = True
        
    End Function

    Public Sub Reset()
        Jump 0, True, False
    End Sub

    Public Sub Jump(state, force, force_show)
        m_base_device.Log "Received jump request. State: " & state & ", Force: "& force

        If Not m_enabled And Not force Then
            m_base_device.Log "Profile is disabled and force is False. Not jumping"
            Exit Sub
        End If
        If state = m_state And Not force_show Then
            m_base_device.Log "Shot is already in the jump destination state"
            Exit Sub
        End If
        m_base_device.Log "Jumping to profile state " & state

        StopShowForState(m_state)
        m_state = state
        PlayShowForState(m_state)
    End Sub

    Public Sub Restart()
        Reset()
        Enable()
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & Replace(m_name, "shot_", "") & ":" & vbCrLf
        If UBound(m_switches) = 0 Then
            yaml = yaml & "    switch: " & m_switches(0) & vbCrLf
        Else
            yaml = yaml & "    switches: " & Join(m_switches, ",") & vbCrLf
        End If
        yaml = yaml & "    show_tokens: " & vbCrLf
        dim key
        For Each key in m_tokens.keys
            If IsArray(m_tokens(key)) Then
                yaml = yaml & "      " & key & ": " & Join(m_tokens(key), ",") & vbCrLf
            Else  
                yaml = yaml & "      " & key & ": " & m_tokens(key) & vbCrLf
            End If
        Next

        If UBound(m_base_device.EnableEvents.Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents.keys
                yaml = yaml & m_base_device.EnableEvents(key).Raw
                If x <> UBound(m_base_device.EnableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_base_device.DisableEvents.Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents.keys
                yaml = yaml & m_base_device.DisableEvents(key).Raw
                If x <> UBound(m_base_device.DisableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_advance_events.Keys) > -1 Then
            yaml = yaml & "    advance_events: "
            x=0
            For Each key in m_advance_events.keys
                yaml = yaml & m_advance_events(key).Raw
                If x <> UBound(m_advance_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_hit_events.Keys) > -1 Then
            yaml = yaml & "    hit_events: "
            x=0
            For Each key in m_hit_events.keys
                yaml = yaml & m_hit_events(key).Raw
                If x <> UBound(m_hit_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        yaml = yaml & "    profile: " & m_profile & vbCrLf
        If Not IsEmpty(m_start_enabled) Then
            yaml = yaml & "    start_enabled: " & m_start_enabled & vbCrLf
        End If
        If UBound(m_restart_events.Keys) > -1 Then
            yaml = yaml & "    restart_events: "
            x=0
            For Each key in m_restart_events.keys
                yaml = yaml & m_restart_events(key).Raw
                If x <> UBound(m_restart_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_reset_events.Keys) > -1 Then
            yaml = yaml & "    reset_events: "
            x=0
            For Each key in m_reset_events.keys
                yaml = yaml & m_reset_events(key).Raw
                If x <> UBound(m_reset_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_control_events.Keys) > -1 Then
            yaml = yaml & "    control_events: " & vbCrLf
            For Each key in m_control_events.keys
                yaml = yaml & "      - events: "
                Dim cEvt
                x=0
                For Each cEvt in m_control_events(key).Events
                    yaml = yaml & cEvt
                    If x <> UBound(m_control_events(key).Events) Then
                        yaml = yaml & ", "
                    End If
                    x = x + 1
                Next
                yaml = yaml & vbCrLf
                yaml = yaml & "      state: " & m_control_events(key).State & vbCrLf
            Next
        End If
        
        ToYaml = yaml
    End Function
End Class

Function ShotEventHandler(args)
    Dim ownProps, kwargs, e
    ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    e = args(2)
    Dim evt : evt = ownProps(0)
    Dim shot : Set shot = ownProps(1)
    Select Case evt
        Case "activate"
            shot.Activate
        Case "deactivate"
            shot.Deactivate
        Case "enable"
            shot.Enable
        Case "hit"
            If Not Glf_EventBlocks(e).Exists(shot.Name) And Not Glf_EventBlocks(e).Exists(shot.ShotKey) Then
                shot.Hit e
            End If
        Case "advance"
            shot.Advance False
        Case "control"
            shot.Jump ownProps(2).State, ownProps(2).Force, ownProps(2).ForceShow
        Case "reset"
            shot.Reset
        Case "restart"
            shot.Restart
            
    End Select
    If IsObject(args(1)) Then
        Set ShotEventHandler = kwargs
    Else
        ShotEventHandler = kwargs
    End If
End Function

Class GlfShotControlEvent
	Private m_events, m_state, m_force, m_force_show
  
	Public Property Get Events(): Events = m_events: End Property
    Public Property Let Events(input): m_events = input: End Property

    Public Property Get State(): State = m_state End Property
    Public Property Let State(input): m_state = input End Property

    Public Property Get Force(): Force = m_force: End Property
	Public Property Let Force(input): m_force = input: End Property
  
	Public Property Get ForceShow(): ForceShow = m_force_show: End Property
	Public Property Let ForceShow(input): m_force_show = input: End Property   

	Public default Function init()
        m_events = Array()
        m_state = 0
        m_force = True
        m_force_show = False
	    Set Init = Me
	End Function

End Class

Class GlfShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_show_cache
    Private m_debug
    Private m_name
    Private m_value

    Public Property Get Events(name)
        If m_events.Exists(name) Then
            Set Events = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfShowPlayerItem)()
            m_events.Add name, new_event
            Set Events = new_event
        End If
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = True
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_show_cache = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "show_player_activate", "ShowPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "show_player_deactivate", "ShowPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            If IsObject(m_events(evt)) Then
                AddPinEventListener Replace(evt, "_" & m_events(evt).Key, "") , m_mode & "_" & m_events(evt).Key & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
            Else
                AddPinEventListener Replace(evt, "_" & m_events(evt), "") , m_mode & "_" & m_events(evt) & "_show_player_play", "ShowPlayerEventHandler", -m_priority, Array("play", Me, m_events(evt), evt)
            End If
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            If IsObject(m_events(evt)) Then
                RemovePinEventListener Replace(evt, "_" & m_events(evt).Key, ""), m_mode & "_" & m_events(evt).Key & "_show_player_play"
            Else
                RemovePinEventListener Replace(evt, "_" & m_events(evt), ""), m_mode & "_" & m_events(evt) & "_show_player_play"
            End If
            PlayOff m_events(evt).Key
        Next
    End Sub

    Public Sub Play(evt, show)
        
        If show.Action = "stop" Then
            PlayOff show.Key
        Else
            If m_show_cache.Exists(show.Key) Then
                lightCtrl.AddLightSeq m_name & "_" & show.Key, show.Key, m_show_cache(show.Key), show.Loops, show.Speed, Null, m_priority, 0
            Else
                Dim cachedShow : cachedShow = Glf_ConvertShow(show.Show, show.Tokens)
                m_show_cache.Add show.Key, cachedShow
                lightCtrl.AddLightSeq m_name & "_" & show.Key, show.Key, cachedShow, show.Loops, show.Speed, Null, m_priority, 0
            End If
        End If
    End Sub

    'Public Sub StopShow(evt, key)
    '    m_events.Add evt & "_" & key & ".stop", key & ".stop"
    'End Sub

    Public Sub PlayOff(key)
        lightCtrl.RemoveLightSeq m_name & "_" & key, key
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_show_player", message
        End If
    End Sub
End Class

Function ShowPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ShowPlayer : Set ShowPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            ShowPlayer.Activate
        Case "deactivate"
            ShowPlayer.Deactivate
        Case "play"
            ShowPlayer.Play ownProps(3), ownProps(2)
    End Select
    ShowPlayerEventHandler = Null
End Function

Class GlfShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Show(): Show = m_show: End Property
	Public Property Let Show(input): m_show = input: End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property        

    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property        
  
	Public default Function init()
        m_action = "play"
        m_key = ""
        m_loops = -1
        m_speed = 1
        m_syncms = 0
        Set m_tokens = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class

Class GlfStateMachine
    Private m_name
    Private m_priority
    private m_base_device
    Private m_states
    Private m_transistions
 
    Private m_current_state
 
    Public Property Get Name(): Name = m_name: End Property
 
    Public Property Get States(): States = m_states: End Property
    Public Property Let States(value)
        m_states = value
    End Property
 
    Public Property Get Transistions(): Transistions = m_transistions: End Property
    Public Property Let Transistions(value)
        m_transistions = value
    End Property
 
    Public Property Get CurrentState(): CurrentState = m_current_state: End Property
    Public Property Let CurrentState(value)
        m_current_state = value
    End Property
 
    Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
 
        Set m_states = CreateObject("Scripting.Dictionary")
        Set m_transistions = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "state_machine", Me)
 
        Set Init = Me
    End Function
 
End Class
 
Class GlfStateMachineState
	Private m_name, m_label, m_show_when_active, m_show_tokens, m_events_when_started, m_events_when_stopped
 
	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property
 
    Public Property Get Label(): Label = m_label: End Property
    Public Property Let Label(input): m_label = input: End Property
 
    Public Property Get ShowWhenActive(): ShowWhenActive = m_show_when_active: End Property
    Public Property Let ShowWhenActive(input): m_show_when_active = input: End Property
 
    Public Property Get ShowTokens(): Set ShowTokens = m_show_tokens: End Property
 
    Public Property Get EventsWhenStarted(): EventsWhenStarted = m_events_when_started: End Property
    Public Property Let EventsWhenStarted(input): m_events_when_started = input: End Property
 
    Public Property Get EventsWhenStopped(): EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(input): m_events_when_stopped = input: End Property
 
	Public default Function init(name)
        m_name = name
        m_label = Empty
        m_show_when_active = Empty
        Set m_show_tokens = CreateObject("Scripting.Dictionary")
        Set m_events_when_started = Array()
        Set m_events_when_stopped = Array()
	    Set Init = Me
	End Function
 
End Class
 
Class GlfStateMachineTranistion
	Private m_source, m_target, m_events, m_events_when_transistioning
 
    Public Property Get Source(): Source = m_source: End Property
    Public Property Let Source(input): m_source = input: End Property
 
    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input): m_target = input: End Property
 
    Public Property Get Events(): Events = m_events: End Property
    Public Property Let Events(input): m_events = input: End Property
 
    Public Property Get EventsWhenTransitioning(): EventsWhenTransitioning = m_events_when_transitioning: End Property
    Public Property Let EventsWhenTransitioning(input): m_events_when_transitioning = input: End Property
 
	Public default Function init(name)
        m_source = name
        m_target = Empty
        m_events = Array()
        m_events_when_transistioning = Array()
	    Set Init = Me
	End Function
 
End Class


Class GlfTimer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_start_events
    Private m_stop_events
    Private m_debug

    Private m_value

    Public Property Let StartValue(value) : m_start_value = value : End Property
    Public Property Let EndValue(value) : m_end_value = value : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let StartEvents(value) : m_start_events = value : End Property
    Public Property Let StopEvents(value) : m_stop_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "TimerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "TimerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_start", "TimerEventHandler", m_priority, Array("start", Me)
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events
                AddPinEventListener evt, m_name & "_stop", "TimerEventHandler", m_priority, Array("stop", Me)
            Next
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_start_events
            RemovePinEventListener evt, m_name & "_start"
        Next
        If Not IsNull(m_stop_events) Then
            For Each evt in m_stop_events
                RemovePinEventListener evt, m_name & "_stop"
            Next
        End If
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub StartTimer()
        Log "Started"
        DispatchPinEvent m_name & "_started", Null
        m_value = m_start_value
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
    End Sub

    Public Sub StopTimer()
        Log "Stopped"
        DispatchPinEvent m_name & "_stopped", Null
        RemoveDelay m_name & "_tick"
        m_value = m_start_value
    End Sub

    Public Sub Tick()
        Dim newValue
        If m_direction = "down" Then
            newValue = m_value - 1
        Else
            newValue = m_value + 1
        End If
        Log "ticking: old value: "& m_value & ", new Value: " & newValue & ", target: "& m_end_value
        m_value = newValue
        If m_value = m_end_value Then
            DispatchPinEvent m_name & "_complete", Null
        Else
            DispatchPinEvent m_name & "_tick", Null
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), 1000
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    Select Case evt
        Case "activate"
            timer.Activate
        Case "deactivate"
            timer.Deactivate
        Case "start"
            timer.StartTimer
        Case "stop"
            timer.StopTimer
        Case "tick"
            timer.Tick 
    End Select
    TimerEventHandler = kwargs
End Function
Class GlfVariablePlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Get Events(name)
        Dim newEvent : Set newEvent = (new GlfVariablePlayerEvent)(name)
        m_events.Add newEvent.BaseEvent.Name, newEvent
        Set Events = newEvent
    End Property
   
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        
        AddPinEventListener m_mode & "_starting", "variable_player_activate", "VariablePlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", "variable_player_deactivate", "VariablePlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_variable_player_play", "VariablePlayerEventHandler", m_priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_variable_player_play"
        Next
    End Sub

    Public Sub Play(evt)
        Log "Playing: " & evt
        If Not IsNull(m_events(evt).BaseEvent.Condition) Then
            If GetRef(m_events(evt).BaseEvent.Condition)() = False Then
                Exit Sub
            End If
        End If
        Dim vKey, v
        For Each vKey in m_events(evt).Variables.Keys
            Log "Setting Variable " & vKey
            Set v = m_events(evt).Variable(vKey)
            Dim varValue : varValue = v.VariableValue
            Select Case v.Action
                Case "add"
                    SetPlayerState vKey, GetPlayerState(vKey) + varValue
                Case "set"
                    SetPlayerState vKey, varValue
        End Select
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_variable_player_play", message
        End If
    End Sub
End Class

Class GlfVariablePlayerEvent

    Private m_event
	Private m_variables

    Public Property Get BaseEvent() : Set BaseEvent = m_event : End Property
  
	Public Property Get Variable(name)
        If m_variables.Exists(name) Then
            Set Variable = m_variables(name)
        Else
            Dim new_variable : Set new_variable = (new GlfVariablePlayerItem)()
            m_variables.Add name, new_variable
            Set Variable = new_variable
        End If
    End Property
    
    Public Property Get Variables(): Set Variables = m_variables End Property

	Public default Function init(evt)
        Set m_event = (new GlfEvent)(evt)
        Set m_variables = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class

Class GlfVariablePlayerItem
	Private m_block, m_show, m_flaot, m_int, m_string, m_player, m_action, m_type
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Block(): Block = m_block End Property
    Public Property Let Block(input): m_block = input End Property

	Public Property Let Float(input): m_float = Glf_ParseInput(input, False): m_type = "float" : End Property
  
	Public Property Let Int(input): m_int = Glf_ParseInput(input, False): m_type = "int" : End Property
  
	Public Property Let String(input): m_string = input: m_type = "string" : End Property

    Public Property Get VariableType(): VariableType = m_type: End Property
    Public Property Get VariableValue()
        Select Case m_type
            Case "float"
                VariableValue = GetRef(m_float(0))()
            Case "int"
                VariableValue = GetRef(m_int(0))()
            Case "string"
                VariableValue = m_string
            Case Else
                VariableValue = Empty
        End Select
    End Property

    Public Property Get Player(): Player = m_player: End Property
    Public Property Let Player(input): m_player = input: End Property

	Public default Function init()
        m_action = "add"
        m_type = Empty
        m_block = False
        m_float = Empty
        m_int = Empty
        m_string = Empty
        m_player = Empty
	    Set Init = Me
	End Function

End Class

Function VariablePlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Dim variablePlayer : Set variablePlayer = ownProps(1)
    Select Case evt
        Case "activate"
            variablePlayer.Activate
        Case "deactivate"
            variablePlayer.Deactivate
        Case "play"
            variablePlayer.Play ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set VariablePlayerEventHandler = kwargs
    Else
        VariablePlayerEventHandler = kwargs
    End If
    
End Function


Class DelayObject
	Private m_name, m_callback, m_ttl, m_args
  
	Public Property Get Name(): Name = m_name: End Property
	Public Property Let Name(input): m_name = input: End Property
  
	Public Property Get Callback(): Callback = m_callback: End Property
	Public Property Let Callback(input): m_callback = input: End Property
  
	Public Property Get TTL(): TTL = m_ttl: End Property
	Public Property Let TTL(input): m_ttl = input: End Property
  
	Public Property Get Args(): Args = m_args: End Property
	Public Property Let Args(input): m_args = input: End Property
  
	Public default Function init(name, callback, ttl, args)
	  m_name = name
	  m_callback = callback
	  m_ttl = ttl
	  m_args = args

	  Set Init = Me
	End Function
End Class

Dim delayQueue : Set delayQueue = CreateObject("Scripting.Dictionary")
Dim delayQueueMap : Set delayQueueMap = CreateObject("Scripting.Dictionary")
Dim delayCallbacks : Set delayCallbacks = CreateObject("Scripting.Dictionary")

Sub SetDelay(name, callbackFunc, args, delayInMs)
    Dim executionTime
    executionTime = AlignToNearest5th(gametime + delayInMs)
    
    If delayQueueMap.Exists(name) Then
        delayQueueMap.Remove name
    End If
    

    If delayQueue.Exists(executionTime) Then
        If delayQueue(executionTime).Exists(name) Then
            delayQueue(executionTime).Remove name
        End If
    Else
        delayQueue.Add executionTime, CreateObject("Scripting.Dictionary")
    End If

    glf_debugLog.WriteToLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc & ", ExecutionTime: " & executionTime
    delayQueue(executionTime).Add name, (new DelayObject)(name, callbackFunc, executionTime, args)
    delayQueueMap.Add name, executionTime
    
End Sub

Function AlignToNearest5th(timeMs)
    AlignToNearest5th = Int(timeMs / 200) * 200
End Function

Function RemoveDelay(name)
    If delayQueueMap.Exists(name) Then
        If delayQueue.Exists(delayQueueMap(name)) Then
            If delayQueue(delayQueueMap(name)).Exists(name) Then
                delayQueue(delayQueueMap(name)).Remove name
            End If
            delayQueueMap.Remove name
            RemoveDelay = True
            glf_debugLog.WriteToLog "Delay", "Removing delay for " & name
            Exit Function
        End If
    End If
    RemoveDelay = False
End Function

Sub DelayTick()
    Dim key, delayObject

    Dim executionTime
    executionTime = AlignToNearest5th(gametime)
    If delayQueue.Exists(executionTime) Then
        For Each key In delayQueue(executionTime).Keys()
            Set delayObject = delayQueue(executionTime)(key)
            glf_debugLog.WriteToLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
            GetRef(delayObject.Callback)(delayObject.Args)
        Next
        delayQueue.Remove executionTime
    End If
End Sub
Class BallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeout
    Private m_balls
    Private m_balls_in_device
    Private m_eject_angle
    Private m_eject_pitch
    Private m_eject_strength
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_balls_to_eject
    Private m_ejecting_all
    Private m_ejecting
    Private m_mechcanical_eject
    Private m_eject_targets
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Let DefaultDevice(value)
        m_default_device = value
        If m_default_device = True Then
            Set glf_plunger = Me
        End If
    End Property
	Public Property Get HasBall()
        HasBall = (Not IsNull(m_balls(0)) And m_ejecting = False)
    End Property
    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    
    Public Property Let EjectAngle(value) : m_eject_angle = glf_PI * (0 - 90) / 180 : End Property
    Public Property Let EjectPitch(value) : m_eject_pitch = glf_PI * (0 - 90) / 180 : End Property
    Public Property Let EjectStrength(value) : m_eject_strength = value : End Property
    
    Public Property Let EjectTimeout(value) : m_eject_timeout = value * 1000 : End Property
    Public Property Let EjectAllEvents(value)
        m_eject_all_events = value
        Dim evt
        For Each evt in m_eject_all_events
            AddPinEventListener evt, m_name & "_eject_all", "BallDeviceEventHandler", 1000, Array("ball_eject_all", Me)
        Next
    End Property
    Public Property Let EjectTargets(value)
        m_eject_targets = value
        Dim evt
        For Each evt in m_eject_targets
            AddPinEventListener evt & "_active", m_name & "_eject_target", "BallDeviceEventHandler", 1000, Array("eject_timeout", Me)
        Next
    End Property
    Public Property Let PlayerControlledEjectEvents(value)
        m_player_controlled_eject_event = value
        Dim evt
        For Each evt in m_player_controlled_eject_event
            AddPinEventListener evt, m_name & "_eject_attempt", "BallDeviceEventHandler", 1000, Array("ball_eject", Me)
        Next
    End Property
    Public Property Let BallSwitches(value)
        m_ball_switches = value
        m_balls = Array(Ubound(m_ball_switches))
        Dim x
        For x=0 to UBound(m_ball_switches)
            m_balls(x) = Null
            AddPinEventListener m_ball_switches(x)&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_entering", Me, x)
            AddPinEventListener m_ball_switches(x)&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, x)
        Next
    End Property
    Public Property Let MechcanicalEject(value) : m_mechcanical_eject = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property
        
	Public default Function init(name)
        m_name = "balldevice_" & name
        m_ball_switches = Array()
        m_eject_all_events = Array()
        m_eject_targets = Array()
        m_balls = Array()
        m_debug = False
        m_default_device = False
        m_eject_pitch = 0
        m_eject_angle = 0
        m_eject_strength = 0
        m_ejecting = False
        m_eject_callback = Null
        m_ejecting_all = False
        m_balls_to_eject = 0
        m_balls_in_device = 0
        m_mechcanical_eject = False
        m_eject_timeout = 0
	    Set Init = Me
	End Function

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_eject_timeout"
        'SoundSaucerLockAtBall ball
        Set m_balls(switch) = ball
        m_balls_in_device = m_balls_in_device + 1
        Log "Ball Entered" 
        Dim unclaimed_balls
        unclaimed_balls = DispatchRelayPinEvent(m_name & "_ball_entered", 1)
        Log "Unclaimed Balls: " & unclaimed_balls
        If (m_default_device = False Or m_ejecting = True) And unclaimed_balls > 0 And Not IsNull(m_balls(0)) Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        m_balls(switch) = Null
        m_balls_in_device = m_balls_in_device - 1
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechcanical_eject = True And m_eject_timeout > 0 Then
            SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), ball), m_eject_timeout
        End If
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        m_ejecting = False
        RemoveDelay m_name & "_eject_timeout"
        DispatchPinEvent m_name & "_ball_eject_success", Null
        Log "Ball successfully exited"
        If m_ejecting_all = True Then
            If m_balls_to_eject = 0 Then
                m_ejecting_all = False
                Exit Sub
            End If
            If Not IsNull(m_balls(0)) Then
                m_balls_to_eject = m_balls_to_eject - 1
                Eject()
            Else
                SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 600
            End If
        End If
    End Sub

    Public Sub Eject
        Log "Ejecting."
        SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_balls(0)), m_eject_timeout
        m_ejecting = True
        If m_eject_strength > 0 Then
            If Not IsNull(m_balls(0)) Then
                m_balls(0).VelX = m_eject_strength * Cos(m_eject_pitch) * Sin(m_eject_angle)
                m_balls(0).VelY = m_eject_strength * Cos(m_eject_pitch) * Cos(m_eject_angle)
                m_balls(0).VelZ = m_eject_strength * Sin(m_eject_pitch)
                Log "VelX: " &  m_balls(0).VelX & ", VelY: " &  m_balls(0).VelY & ", VelZ: " &  m_balls(0).VelZ
            End If
        End If

        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(m_balls(0))
        End If
    End Sub

    Public Sub EjectAll
        Log "Ejecting All."
        m_ejecting_all = True
        m_balls_to_eject = m_balls_in_device
        Eject()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallDeviceEventHandler(args)
    Dim ownProps, ball : ownProps = args(0) : Set ball = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballDevice : Set ballDevice = ownProps(1)
    Dim switch
    Select Case evt
        Case "ball_entering"
            switch = ownProps(2)
            SetDelay ballDevice.Name & "_" & switch & "_ball_enter", "BallDeviceEventHandler", Array(Array("ball_enter", ballDevice, switch), ball), 500
        Case "ball_enter"
            switch = ownProps(2)
            ballDevice.BallEnter ball, switch
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_eject_all"
            ballDevice.EjectAll
        Case "ball_exiting"
            switch = ownProps(2)
            If RemoveDelay(ballDevice.Name & "_" & switch & "_ball_enter") = False Then
                ballDevice.BallExiting ball, switch
            End If
        Case "eject_timeout"
            ballDevice.BallExitSuccess ball
    End Select
End Function

Class Diverter

    Private m_name
    Private m_activate_events
    Private m_deactivate_events
    Private m_activation_time
    Private m_enable_events
    Private m_disable_events
    Private m_activation_switches
    Private m_action_cb
    Private m_debug

    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let ActivateEvents(value) : m_activate_events = value : End Property
    Public Property Let DeactivateEvents(value) : m_deactivate_events = value : End Property
    Public Property Let ActivationTime(value) : m_activation_time = value : End Property
    Public Property Let ActivationSwitches(value) : m_activation_switches = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, enable_events, disable_events)
        m_name = "diverter_" & name
        m_enable_events = enable_events
        m_disable_events = disable_events
        m_activate_events = Array()
        m_deactivate_events = Array()
        m_activation_switches = Array()
        m_activation_time = 0
        m_debug = False
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "DiverterEventHandler", 1000, Array("enable", Me)
        Next
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "DiverterEventHandler", 1000, Array("disable", Me)
        Next
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_activate_events
            AddPinEventListener evt, m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
        For Each evt in m_deactivate_events
            AddPinEventListener evt, m_name & "_deactivate", "DiverterEventHandler", 1000, Array("deactivate", Me)
        Next
        For Each evt in m_activation_switches
            AddPinEventListener evt & "_active", m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_activate_events
            RemovePinEventListener evt, m_name & "_activate"
        Next
        For Each evt in m_deactivate_events
            RemovePinEventListener evt, m_name & "_deactivate"
        Next
        For Each evt in m_activation_switches
            RemovePinEventListener evt & "_active", m_name & "_activate"
        Next
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
    End Sub

    Public Sub Activate()
        Log "Activating"
        GetRef(m_action_cb)(1)
        If m_activation_time > 0 Then
            SetDelay m_name & "_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_activation_time
        End If
        DispatchPinEvent m_name & "_activating", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_deactivating", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DiverterEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim diverter : Set diverter = ownProps(1)
    Select Case evt
        Case "enable"
            diverter.Enable
        Case "disable"
            diverter.Disable
        Case "activate"
            diverter.Activate
        Case "deactivate"
            diverter.Deactivate
    End Select
    DiverterEventHandler = kwargs
End Function
Class DropTarget
	Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped
    Private m_reset_events
  
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
  
	Public default Function init(primary, secondary, prim, sw, animate, isDropped, reset_events)
	  Set m_primary = primary
	  Set m_secondary = secondary
	  Set m_prim = prim
	  m_sw = sw
	  m_animate = animate
	  m_isDropped = isDropped
      m_reset_events = reset_events
	  If Not IsNull(reset_events) Then
	  	Dim evt
		For Each evt in reset_events
			AddPinEventListener evt, primary.name & "_droptarget_reset", "DropTargetEventHandler", 1000, Array("droptarget_reset", m_sw)
		Next      	
	  End If
	  Set Init = Me
	End Function
End Class

Function DropTargetEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim kwargs : kwargs = args(1)
    Dim evt : evt = ownProps(0)
    Dim switch : switch = ownProps(1)
    Select Case evt
        Case "droptarget_reset"
            DTRaise switch
    End Select
    DropTargetEventHandler = kwargs
End Function

Class GlfEvent
	Private m_raw, m_name, m_event, m_condition
  
    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventName() : EventName = m_event : End Property
    Public Property Get Condition() : Condition = m_condition : End Property
    Public Property Get Raw() : Raw = m_raw : End Property

	Public default Function init(evt)
        m_raw = evt
        Dim parsedEvent : parsedEvent = Glf_ParseEventInput(evt)
        m_name = parsedEvent(0)
        m_event = parsedEvent(1)
        m_condition = parsedEvent(2)
	    Set Init = Me
	End Function

End Class


'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub Glf_AddPlayer()
    Select Case UBound(playerState.Keys())
        Case -1:
            playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            Glf_BcpAddPlayer 1
            glf_currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 2", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 2
            End If
        Case 1:
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 3", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 3
            End If     
        Case 2:   
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 4", Glf_InitNewPlayer()
                Glf_BcpAddPlayer 4
            End If  
            glf_canAddPlayers = False
    End Select
End Sub

Function Glf_InitNewPlayer()
    Dim state : Set state = CreateObject("Scripting.Dictionary")
    state.Add GLF_SCORE, 0
    state.Add GLF_INITIALS, ""
    state.Add GLF_CURRENT_BALL, 1
    Set Glf_InitNewPlayer = state
End Function


'****************************
' Setup Player
' Event Listeners:  
    AddPinEventListener GLF_GAME_STARTED,   "start_game_setup",   "Glf_SetupPlayer", 1000, Null
    AddPinEventListener GLF_NEXT_PLAYER,    "next_player_setup",  "Glf_SetupPlayer", 1000, Null
'
'*****************************
Function Glf_SetupPlayer(args)
    EmitAllglf_playerEvents()
End Function

'****************************
' StartGame
'
'*****************************
Sub Glf_StartGame()
    glf_gameStarted = True
    If useBcp Then
        bcpController.Send "player_turn_start?player_num=int:1"
        bcpController.Send "ball_start?player_num=int:1&ball=int:1"
        bcpController.PlaySlide "base", "base", 1000
        bcpController.SendPlayerVariable "number", 1, 0
    End If
    DispatchPinEvent GLF_GAME_STARTED, Null
End Sub


'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener GLF_GAME_STARTED, "start_game_release_ball",   "Glf_ReleaseBall", 1000, True
AddPinEventListener GLF_NEXT_PLAYER, "next_player_release_ball",   "Glf_ReleaseBall", 1000, True
'
'*****************************
Function Glf_ReleaseBall(args)
    If Not IsNull(args) Then
        If args(0) = True Then
            DispatchPinEvent GLF_BALL_STARTED, Null
            If useBcp Then
                bcpController.SendPlayerVariable GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL), GetPlayerState(GLF_CURRENT_BALL)-1
                bcpController.SendPlayerVariable GLF_SCORE, GetPlayerState(GLF_SCORE), GetPlayerState(GLF_SCORE)
            End If
        End If
    End If
    glf_debugLog.WriteToLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    swTrough1.kick 90, 10
    glf_debugLog.WriteToLog "Release Ball", "Just Kicked"
    glf_BIP = glf_BIP + 1
End Function


'****************************
' End Of Ball
' Event Listeners:      
    AddPinEventListener GLF_BALL_DRAIN, "ball_drain", "Glf_Drain", 20, Null
'
'*****************************
Function Glf_Drain(args)
    
    Dim ballsToSave : ballsToSave = args(1) 
    glf_debugLog.WriteToLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
    glf_debugLog.WriteToLog "end_of_ball, balls in play", CStr(glf_BIP)
    If ballsToSave <= 0 Then
        Exit Function
    End If

    If glf_BIP > 0 Then
        Exit Function
    End If
        
    DispatchPinEvent GLF_BALL_ENDED, Null
    SetPlayerState GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            If UBound(playerState.Keys()) > 0 Then
                glf_currentPlayer = "PLAYER 2"
            End If
        Case "PLAYER 2":
            If UBound(playerState.Keys()) > 1 Then
                glf_currentPlayer = "PLAYER 3"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 3":
            If UBound(playerState.Keys()) > 2 Then
                glf_currentPlayer = "PLAYER 4"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 4":
            glf_currentPlayer = "PLAYER 1"
    End Select
    
    If useBcp Then
        bcpController.SendPlayerVariable "number", Getglf_currentPlayerNumber(), previousPlayerNumber
    End If
    If GetPlayerState(GLF_CURRENT_BALL) > glf_ballsPerGame Then
        DispatchPinEvent GLF_GAME_OVER, Null
        glf_gameStarted = False
        glf_currentPlayer = Null
        playerState.RemoveAll()
    Else
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000 
    End If
    
End Function

Public Function EndOfBallNextPlayer(args)
    DispatchPinEvent GLF_NEXT_PLAYER, Null
End Function


'***********************************************************************************************************************
' Lights State Controller - 0.9.1
'  
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************


Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqOffset, m_tableSeqSpeed, m_tableSeqDirection, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, m_lightmaps, m_seqOverrideRunners, m_pauseMainLights, m_pausedLights, m_minX, m_minY, m_maxX, m_maxY, m_width, m_height, m_centerX, m_centerY, m_coordsX, m_coordsY, m_angles, m_radii, m_tags, m_debug, m_syncMs, m_syncMsCounter

    Private Lights(260)

    Private Sub Class_Initialize()
        Set m_lights = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_off = CreateObject("Scripting.Dictionary")
        Set m_seqRunners = CreateObject("Scripting.Dictionary")
        Set m_seqOverrideRunners = CreateObject("Scripting.Dictionary")
        Set m_currentFrameState = CreateObject("Scripting.Dictionary")
        Set m_seqs = CreateObject("Scripting.Dictionary")
        Set m_pulse = CreateObject("Scripting.Dictionary")
        Set m_on = CreateObject("Scripting.Dictionary")
        Set m_tags = CreateObject("Scripting.Dictionary")
        m_vpxLightSyncRunning = False
        m_vpxLightSyncCollection = Null
        m_initFrameTime = 0
        m_frameTime = 0
        m_pulseInterval = 26
        m_vpxLightSyncClear = False
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        m_pauseMainLights = False
        Set m_pausedLights = CreateObject("Scripting.Dictionary")
        Set m_lightmaps = CreateObject("Scripting.Dictionary")
        Set m_syncMs = CreateObject("Scripting.Dictionary")
        Set m_syncMsCounter = CreateObject("Scripting.Dictionary")
        m_minX = 1000000
        m_minY = 1000000
        m_maxX = -1000000
        m_maxY = -1000000
        m_centerX = Round(tableWidth/2)
        m_centerY = Round(tableHeight/2)
        m_debug = False
    End Sub

    Public Property Let Debug(value) : m_debug = value : End Property

    Private Sub AssignStateForFrame(key, state)
        If m_currentFrameState.Exists(key) Then
            m_currentFrameState.Remove key
        End If
        m_currentFrameState.Add key, state
    End Sub

    Public Sub LoadLightShows()
        Dim oFile
        Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-out.txt",2,true)
        For Each oFile In oFSO.GetFolder(cGameName & "_LightShows").Files
            If LCase(oFSO.GetExtensionName(oFile.Name)) = "yaml" And Not Left(oFile.Name,6) = "lights" Then
                Dim textStream : Set textStream = oFSO.OpenTextFile(oFile.Path, 1)
                Dim show : show = textStream.ReadAll
                Dim fileName : fileName = "lSeq" & Replace(oFSO.GetFileName(oFile.Name), "."&oFSO.GetExtensionName(oFile.Name), "")
                Dim lcSeq : lcSeq = "Dim " & fileName & " : Set " & fileName & " = New LCSeq"&vbCrLf
                lcSeq = lcSeq + fileName & ".Name = """&fileName&""""&vbCrLf
                Dim seq : seq = ""
                Dim re : Set re = New RegExp
                With re
                    .Pattern    = "- time:.*?\n"
                    .IgnoreCase = False
                    .Global     = True
                End With
                Dim matches : Set matches = re.execute(show)
                Dim steps : steps = matches.Count
                Dim match, nextMatchIndex, uniqueLights
                Set uniqueLights = CreateObject("Scripting.Dictionary")
                nextMatchIndex = 1
                For Each match in matches
                    Dim lightStep
                    If Not nextMatchIndex < steps Then
                        lightStep = Mid(show, match.FirstIndex, Len(show))
                    Else
                        lightStep = Mid(show, match.FirstIndex, matches(nextMatchIndex).FirstIndex - match.FirstIndex)
                        nextMatchIndex = nextMatchIndex + 1
                    End If

                    Dim re1 : Set re1 = New RegExp
                    With re1
                        .Pattern        = ".*:?: '([A-Fa-f0-9]{6})'"
                        .IgnoreCase     = True
                        .Global         = True
                    End With

                    Dim lightMatches : Set lightMatches = re1.execute(lightStep)
                    If lightMatches.Count > 0 Then
                        Dim lightMatch, lightStr, lightSplit
                        lightStr = "Array("
                        lightSplit = 0
                        For Each lightMatch in lightMatches
                            Dim sParts : sParts = Split(lightMatch.Value, ":")
                            Dim lightName : lightName = Trim(sParts(0))
                            Dim color : color = Trim(Replace(sParts(1),"'", ""))
                            If color = "000000" Then
                                lightStr = lightStr + """"&lightName&"|0|000000"","
                            Else
                                lightStr = lightStr + """"&lightName&"|100|"&color&""","
                            End If

                            If Len(lightStr)+20 > 2000 And lightSplit = 0 Then                           
                                lightSplit = Len(lightStr)
                            End If

                            uniqueLights(lightname) = 0
                        Next
                        lightStr = Left(lightStr, Len(lightStr) - 1)
                        lightStr = lightStr & ")"
                        
                        If lightSplit > 0 Then
                            lightStr = Left(lightStr, lightSplit) & " _ " & vbCrLF & Right(lightStr, Len(lightStr)-lightSplit)
                        End If

                        seq = seq + lightStr & ", _"&vbCrLf
                    Else
                        seq = seq + "Array(), _"&vbCrLf
                    End If

                    
                    Set re1 = Nothing
                Next
                
                lcSeq = lcSeq + filename & ".Sequence = Array( " & Left(seq, Len(seq) - 5) & ")"&vbCrLf
                'lcSeq = lcSeq + seq & vbCrLf
                lcSeq = lcSeq + fileName & ".UpdateInterval = 20"&vbCrLf
                lcSeq = lcSeq + fileName & ".Color = Null"&vbCrLf
                lcSeq = lcSeq + fileName & ".Repeat = False"&vbCrLf

                'MsgBox(lcSeq)
                objFileToWrite.WriteLine(lcSeq)
                ExecuteGlobal lcSeq
                Set re = Nothing

                textStream.Close
            End if
        Next
        'Clean up
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Set oFile = Nothing
        Set oFSO = Nothing
    End Sub

    Public Sub CompileLights(collection, name)
        Dim light
        Dim lights : lights = "light:" & vbCrLf
        For Each light in collection
            lights = lights + light.name & ":"&vbCrLf
            lights = lights + "   x: "& light.x/tablewidth & vbCrLf
            lights = lights + "   y: "& light.y/tableheight & vbCrLf
        Next
        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/lights-"&name&".yaml",2,true)
        objFileToWrite.WriteLine(lights)
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Log  "Lights YAML File saved to: " & cGameName & "LightShows/lights-"&name&".yaml"
    End Sub

    Dim leds
    Dim ledGrid()
    Dim lightsToLeds

    Sub PrintLEDs
        Dim light
        Dim lights : lights = ""
    
        Dim row,col,value,i
        For row = LBound(ledGrid, 1) To UBound(ledGrid, 1)
            For col = LBound(ledGrid, 2) To UBound(ledGrid, 2)
                ' Access the array element and do something with it
                value = ledGrid(row, col)
                lights = lights + cstr(value) & vbTab
            Next
            lights = lights + vbCrLf
        Next

        Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/led-grid.txt",2,true)
        objFileToWrite.WriteLine(lights)
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Log "Lights File saved to: " & cGameName & "LightShows/led-grid.txt"

        lights = ""
        For i = 0 To UBound(leds)
            value = leds(i)
            If IsArray(value) Then
                lights = lights + "Index: " & cstr(value(0)) & ". X: " & cstr(value(1)) & ". Y:" & cstr(value(2)) & ". Angle:" & cstr(value(3)) & ". Radius:" & cstr(value(4)) & ". CoordsX:" & cstr(value(5)) & ". CoordsY:" & cstr(value(6)) & ". Angle256:" & cstr(value(7)) &". Radii256:" & cstr(value(8)) &","
            End If
            lights = lights + vbCrLf
            'lights = lights + cstr(value) & ","
            
        Next

        
        Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/coordsX.txt",2,true)
        objFileToWrite.WriteLine(lights)
        objFileToWrite.Close
        Set objFileToWrite = Nothing
        Log "Lights File saved to: " & cGameName & "LightShows/coordsX.txt"


    End Sub

    Public Sub RegisterLights(aLights)


        Dim obj, str, ii
        For Each obj In aLights
            idx = obj.TimerInterval
            If IsArray(Lights(idx)) Then
                str = "Lights(" & idx & ") = Array("
                For Each ii In Lights(idx) : str = str & ii.Name & "," : Next
                ExecuteGlobal str & obj.Name & ")"
            ElseIf IsObject(Lights(idx)) Then
                Lights(idx) = Array(Lights(idx),obj)
            Else
                Set Lights(idx) = obj
            End If
        Next

        Dim idx,tmp,vpxLight,lcItem
    
            Dim colCount : colCount = Round(tablewidth/20)
            Dim rowCount : rowCount = Round(tableheight/20)
            Log "Registering Lights" 
            dim ledIdx : ledIdx = 0
            redim leds(UBound(Lights)-1)
            redim lightsToLeds(UBound(Lights)-1)
            ReDim ledGrid(rowCount,colCount)
            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
                If Not IsNull(vpxLight) Then
                    Log "Registering Light: "& vpxLight.name


                    Dim r : r = Round(vpxLight.y/20)
                    Dim c : c = Round(vpxLight.x/20)
                    If r < rowCount And c < colCount And r >= 0 And c >= 0 Then
                        If Not ledGrid(r,c) = "" Then
                            MsgBox(vpxLight.name & " is too close to another light")
                        End If
                        leds(ledIdx) = Array(ledIdx, c, r, 0,0,0,0,0,0) 'index, row, col, angle, radius, x256, y256, angle256, radius256
                        lightsToLeds(idx) = ledIdx
                        ledGrid(r,c) = ledIdx
                        ledIdx = ledIdx + 1
                        If (c < m_minX) Then : m_minX = c
                        if (c > m_maxX) Then : m_maxX = c
                
                        if (r < m_minY) Then : m_minY = r
                        if (r > m_maxY) Then : m_maxY = r
                    End If
                    Dim e, lmStr: lmStr = "lmArr = Array("    
                    For Each e in GetElements()
                        On Error Resume Next
                        If InStr(LCase(e.Name), LCase("_" & vpxLight.Name & "_")) Or InStr(LCase(e.Name), LCase("_" & vpxLight.UserValue & "_")) Then
                            lmStr = lmStr & e.Name & ","
                        End If
                        If Err Then Log "Error: " & Err
                    Next
                    lmStr = lmStr & "Null)"
                    lmStr = Replace(lmStr, ",Null)", ")")
                    ExecuteGlobal "Dim lmArr : "&lmStr
                    m_lightmaps.Add vpxLight.Name, lmArr
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    CreateSeqRunner "lSeqRunner" & CStr(vpxLight.name), 1000
                End If
            Next
            'ReDim Preserve leds(ledIdx)
            m_width = m_maxX - m_minX + 1
            m_height = m_maxY - m_minY + 1
            m_centerX = m_width / 2
            m_centerY = m_height / 2
            GenerateLedMapCode()
    End Sub

    Private Sub GenerateLedMapCode()

        Dim minX256, minY256, minAngle, minAngle256, minRadius, minRadius256
        Dim maxX256, maxY256, maxAngle, maxAngle256, maxRadius, maxRadius256
        Dim i, led
        minX256 = 1000000
        minY256 = 1000000
        minAngle = 1000000
        minAngle256 = 1000000
        minRadius = 1000000
        minRadius256 = 1000000

        maxX256 = -1000000
        maxY256 = -1000000
        maxAngle = -1000000
        maxAngle256 = -1000000
        maxRadius = -1000000
        maxRadius256 = -1000000

        For i = 0 To UBound(leds)
            led = leds(i)
            If IsArray(led) Then
                
                Dim x : x = led(1)
                Dim y : y = led(2)
            
                Dim radius : radius = Sqr((x - m_centerX) ^ 2 + (y - m_centerY) ^ 2)
                Dim radians: radians = Atn2(m_centerY - y, m_centerX - x)
                Dim angle
                angle = radians * (180 / 3.141592653589793)
                Do While angle < 0
                    angle = angle + 360
                Loop
                Do While angle > 360
                    angle = angle - 360
                Loop
            
                If angle < minAngle Then
                    minAngle = angle
                End If
                If angle > maxAngle Then
                    maxAngle = angle
                End If
            
                If radius < minRadius Then
                    minRadius = radius
                End If
                If radius > maxRadius Then
                    maxRadius = radius
                End If
            
                led(3) = angle
                led(4) = radius
                leds(i) = led
            End If
        Next

        For i = 0 To UBound(leds)
            led = leds(i)
            If IsArray(led) Then
                Dim x256 : x256 = MapNumber(led(1), m_minX, m_maxX, 0, 255)
                Dim y256 : y256 = MapNumber(led(2), m_minY, m_maxY, 0, 255)
                Dim angle256 : angle256 = MapNumber(led(3), 0, 360, 0, 255)
                Dim radius256 : radius256 = MapNumber(led(4), 0, maxRadius, 0, 255)
            
                led(5) = Round(x256)
                led(6) = Round(y256)
                led(7) = Round(angle256)
                led(8) = Round(radius256)
            
                If x256 < minX256 Then minX256 = x256
                If x256 > maxX256 Then maxX256 = x256
            
                If y256 < minY256 Then minY256 = y256
                If y256 > maxY256 Then maxY256 = y256
            
                If angle256 < minAngle256 Then minAngle256 = angle256
                If angle256 > maxAngle256 Then maxAngle256 = angle256
            
                If radius256 < minRadius256 Then minRadius256 = radius256
                If radius256 > maxRadius256 Then maxRadius256 = radius256
                leds(i) = led
            End If
        Next

        reDim m_coordsX(UBound(leds)-1)
        reDim m_coordsY(UBound(leds)-1)
        reDim m_angles(UBound(leds)-1)
        reDim m_radii(UBound(leds)-1)
        
        For i = 0 To UBound(leds)
            led = leds(i)
            If IsArray(led) Then
                m_coordsX(i)    =  leds(i)(5) 'x256
                m_coordsY(i)    =  leds(i)(6) 'y256
                m_angles(i)     =  leds(i)(7) 'angle256
                m_radii(i)      =  leds(i)(8) 'radius256
            End If
        Next

    End Sub

    Private Function MapNumber(l, inMin, inMax, outMin, outMax)
        If (inMax - inMin + outMin) = 0 Then
            MapNumber = 0
        Else
            MapNumber = ((l - inMin) * (outMax - outMin)) / (inMax - inMin) + outMin
        End If
    End Function

    Private Function ReverseArray(arr)
        Dim i, upperBound
        upperBound = UBound(arr)

        ' Create a new array of the same size
        Dim reversedArr()
        ReDim reversedArr(upperBound)

        ' Fill the new array with elements in reverse order
        For i = 0 To upperBound
            reversedArr(i) = arr(upperBound - i)
        Next

        ReverseArray = reversedArr
    End Function

    Private Function Atn2(dy, dx)
        If dx > 0 Then
            Atn2 = Atn(dy / dx)
        ElseIf dx < 0 Then
            If dy = 0 Then 
                Atn2 = pi
            Else
                Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
            end if
        ElseIf dx = 0 Then
            if dy = 0 Then
                Atn2 = 0
            else
                Atn2 = Sgn(dy) * pi / 2
            end if
        End If
    End Function

    Private Function ColtoArray(aDict)	'converts a collection to an indexed array. Indexes will come out random probably.
        redim a(999)
        dim count : count = 0
        dim x  : for each x in aDict : set a(Count) = x : count = count + 1 : Next
        redim preserve a(count-1) : ColtoArray = a
    End Function

    Function IncrementUInt8(x, increment)
        If x + increment > 255 Then
            IncrementUInt8 = x + increment - 256
        Else
            IncrementUInt8 = x + increment
        End If
    End Function

    Public Sub AddLight(light, idx)
        If m_lights.Exists(light.name) Then
            Exit Sub
        End If
        Dim lcItem : Set lcItem = new LCItem
        lcItem.Init idx, light.BlinkInterval, Array(light.color, light.colorFull), light.name, light.x, light.y
        m_lights.Add light.Name, lcItem
        CreateSeqRunner "lSeqRunner" & CStr(light.name), 1000
    End Sub

    Public Sub AddLightTags(light, tags)
        If m_lights.Exists(light.name) Then
            Dim tagArray, tag, lightName
            lightName = light.name
            tagArray = Split(tags, ",")
            
            For Each tag In tagArray
                tag = Trim(tag) ' Remove any leading or trailing spaces
                If Not m_tags.Exists(tag) Then
                    Set m_tags(tag) = CreateObject("Scripting.Dictionary")
                End If
                If Not m_tags(tag).Exists(lightName) Then
                    m_tags(tag).Add lightName, True
                End If
            Next
        End If
    End Sub

    Public Function GetLightsForTag(tag)
        If m_tags.Exists(tag) Then
            GetLightsForTag = m_tags(tag).Keys()
        Else
            GetLightsForTag = Array()
        End If
    End Function

    Public Sub LightState(light, state)
        m_lightOff(light.name)
        If state = 1 Then
            m_lightOn(light.name)
        ElseIF state = 2 Then
            Blink(light)
        End If
    End Sub

    Public Sub LightOn(light)
        If vartype(light) = 8 Then
            m_LightOn(light)
        Else
            m_LightOn(light.name)
        End If
    End Sub

    Public Sub LightOnWithColor(light, color)
        If vartype(light) = 8 Then
            m_LightOnWithColor light, color
        Else
            m_LightOnWithColor light.name, color
        End If
    End Sub

    Public Sub FadeLightToColor(light, color, fadeSpeed, runner, priority)
        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If

        If vartype(color) = 8 Then
            color = RGB( HexToInt(Left(color, 2)), HexToInt(Mid(color, 3, 2)), HexToInt(Right(color, 2)) )
        End If

        If m_lights.Exists(lightName) Then
            dim lightColor, steps
            steps = Round(fadeSpeed/20)
            If steps < 10 Then
                steps = 10
            End If
            lightColor = m_lights(lightName).Color
            Dim seq : seq  = FadeRGB(lightName, lightColor(0), color, steps)
            m_lights(lightName).Color = color
            CreateSeqRunner runner, priority
            m_seqRunners(runner).AddItem lightName & "Fade", seq, 1, 10, Null,0
            If color = RGB(0,0,0) Then
                m_lightOff(lightName)
            End If
        End If
    End Sub

    Public Sub FlickerOn(light)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            m_lightOn(name)

            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70), 0, m_pulseInterval, 1, null)
        End If
    End Sub  
    
    Public Sub LightColor(light, color)

        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If

        If m_lights.Exists(lightName) Then
            m_lights(lightName).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(lightName & "Blink") Then
                m_seqs(lightName & "Blink").Color = color
            End If

        End If
    End Sub

    Public Function GetLightColor(light)
        If m_lights.Exists(light.name) Then
            Dim colors : colors = m_lights(light.name).Color
            GetLightColor = colors(0)
        Else
            GetLightColor = RGB(0,0,0)
        End If
    End Function

    Private Sub m_LightOn(name)
        
        If m_lights.Exists(name) Then
            
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If
            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Private Sub m_LightOnWithColor(name, color)
        If m_lights.Exists(name) Then
            m_lights(name).Color = color
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_on.Exists(name) Then 
                Exit Sub
            End If
            m_on.Add name, m_lights(name)
        End If
    End Sub

    Public Sub LightOff(light)
        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If
        m_lightOff(lightName)
    End Sub

    Private Sub m_lightOff(name)
        If m_lights.Exists(name) Then
            If m_on.Exists(name) Then 
                m_on.Remove(name)
            End If

            If m_seqs.Exists(name & "Blink") Then
                m_seqRunners("lSeqRunner"&CStr(name)).RemoveItem m_seqs(name & "Blink")
            End If

            If m_off.Exists(name) Then 
                Exit Sub
            End If
            LightColor m_lights(name), m_lights(name).BaseColor
            m_off.Add name, m_lights(name)
        End If
    End Sub

    Public Sub UpdateBlinkInterval(light, interval)
        If m_lights.Exists(light.name) Then
            light.BlinkInterval = interval
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs.Item(light.name & "Blink").UpdateInterval = interval
            End If
        End If
    End Sub

    Public Sub Pulse(light, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            'Array(100,94,32,13,6,3,0)
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount, null)
        End If
    End Sub

    Public Sub PulseWithColor(light, color, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            'Array(100,94,32,13,6,3,0)
            m_pulse.Add name, (new PulseState)(m_lights(name), Array(37,100,24,0,70,100,12,0), 0, m_pulseInterval, repeatCount,  Array(color,null))
        End If
    End Sub

    Public Sub PulseWithProfile(light, profile, repeatCount)
        Dim name : name = light.name
        If m_lights.Exists(name) Then
            If m_off.Exists(name) Then 
                m_off.Remove(name)
            End If
            If m_pulse.Exists(name) Then 
                Exit Sub
            End If
            m_pulse.Add name, (new PulseState)(m_lights(name), profile, 0, m_pulseInterval, repeatCount, null)
        End If
    End Sub       

    Public Sub PulseWithState(pulse)
        
        If m_lights.Exists(pulse.Light) Then
            If m_off.Exists(pulse.Light) Then 
                m_off.Remove(pulse.Light)
            End If
            If m_pulse.Exists(pulse.Light) Then 
                Exit Sub
            End If
            m_pulse.Add name, pulse
        End If
    End Sub

    Public Sub LightLevel(light, lvl)
        If m_lights.Exists(light.name) Then
            m_lights(light.name).Level = lvl

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Sequence = m_buildBlinkSeq(light.name, light.BlinkPattern)
            End If
        End If
    End Sub

    Public Sub AddShot(name, light, color)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(name & light.name) Then
                m_seqs(name & light.name).Color = color
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddSequenceItem m_seqs(name & light.name)
            Else
                Dim stateOn : stateOn = light.name&"|100"
                Dim stateOff : stateOff = light.name&"|0"
                Dim seq : Set seq = new LCSeq
                seq.Name = name
                seq.Sequence = Array(stateOn, stateOff,stateOn, stateOff)
                seq.Color = color
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem name, seq, -1, 200/light.BlinkInterval, Null,0
                m_seqs.Add name & light.name, seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub RemoveShot(name, light)
        If m_lights.Exists(light.name) And m_seqs.Exists(name & light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveItem m_seqs(name & light.name)
            If IsNUll(m_seqRunners("lSeqRunner"&CStr(light.name)).CurrentItem) Then
            LightOff(light)
            End If
        End If
    End Sub

    Public Sub RemoveAllShots()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub RemoveShotsFromLight(light)
        If m_lights.Exists(light.name) Then
            m_seqRunners("lSeqRunner"&CStr(light.name)).RemoveAll   
            m_lightOff(light.name)  
        End If
    End Sub

    Public Sub Blink(light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").ResetInterval
                m_seqs(light.name & "Blink").CurrentIdx = 0
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddSequenceItem m_seqs(light.name & "Blink")
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = light.name & "Blink"
                seq.Sequence = m_buildBlinkSeq(light.name, light.BlinkPattern)
                seq.Color = Null
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem light.name & "Blink", seq.Sequence, -1, 200/light.BlinkInterval, Null,0
                m_seqs.Add light.name & "Blink", seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
        End If
    End Sub

    Public Sub AddLightToBlinkGroup(group, light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(group & "BlinkGroup") Then

                Dim i, pattern, buff : buff = Array()
                pattern = m_seqs(group & "BlinkGroup").Pattern
                ReDim buff(Len(pattern)-1)
                For i = 0 To Len(pattern)-1
                    Dim lightInSeq, ii, p, buff2
                    buff2 = Array()
                    If Mid(pattern, i+1, 1) = 1 Then
                        p=1
                    Else
                        p=0
                    End If
                    ReDim buff2(UBound(m_seqs(group & "BlinkGroup").LightsInSeq)+1)
                    ii=0
                    For Each lightInSeq in m_seqs(group & "BlinkGroup").LightsInSeq
                        If p = 1 Then
                            buff2(ii) = lightInSeq & "|100"
                        Else
                            buff2(ii) = lightInSeq & "|0"
                        End If
                        ii = ii + 1
                    Next
                    If p = 1 Then
                        buff2(ii) = light.name & "|100"
                    Else
                        buff2(ii) = light.name & "|0"
                    End If
                    buff(i) = buff2
                Next
                m_seqs(group & "BlinkGroup").Sequence = buff
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = group & "BlinkGroup"
                seq.Sequence = Array(Array(light.name & "|100"), Array(light.name & "|0"))
                seq.Color = Null
                seq.Pattern = "10"
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True
                CreateSeqRunner "lSeqRunner" & group & "BlinkGroup", 1000
                m_seqs.Add group & "BlinkGroup", seq
            End If
        End If
    End Sub

    Public Sub RemoveLightFromBlinkGroup(group, light)
        If m_lights.Exists(light.name) Then

            If m_seqs.Exists(group & "BlinkGroup") Then

                Dim i, pattern, buff : buff = Array()
                pattern = m_seqs(group & "BlinkGroup").Pattern
                ReDim buff(Len(pattern)-1)
                For i = 0 To Len(pattern)-1
                    Dim lightInSeq, ii, p, buff2
                    buff2 = Array()
                    If Mid(pattern, i+1, 1) = 1 Then
                        p=1
                    Else
                        p=0
                    End If
                    ReDim buff2(UBound(m_seqs(group & "BlinkGroup").LightsInSeq)-1)
                    ii=0
                    For Each lightInSeq in m_seqs(group & "BlinkGroup").LightsInSeq
                        If Not lightInSeq = light.name Then
                            If p = 1 Then
                                buff2(ii) = lightInSeq & "|100"
                            Else
                                buff2(ii) = lightInSeq & "|0"
                            End If
                            ii = ii + 1
                        End If
                    Next
                    buff(i) = buff2
                Next
                AssignStateForFrame light.name, (new FrameState)(0, Null, m_lights(light.name).Idx)
                m_seqs(group & "BlinkGroup").Sequence = buff
            End If
        End If
    End Sub

    Public Sub UpdateBlinkGroupPattern(group, pattern)
        If m_seqs.Exists(group & "BlinkGroup") Then

            Dim i, buff : buff = Array()
            m_seqs(group & "BlinkGroup").Pattern = pattern
            ReDim buff(Len(pattern)-1)
            For i = 0 To Len(pattern)-1
                Dim lightInSeq, ii, p, buff2
                buff2 = Array()
                If Mid(pattern, i+1, 1) = 1 Then
                    p=1
                Else
                    p=0
                End If
                ReDim buff2(UBound(m_seqs(group & "BlinkGroup").LightsInSeq))
                ii=0
                For Each lightInSeq in m_seqs(group & "BlinkGroup").LightsInSeq
                    If p = 1 Then
                        buff2(ii) = lightInSeq & "|100"
                    Else
                        buff2(ii) = lightInSeq & "|0"
                    End If
                    ii = ii + 1
                Next
                buff(i) = buff2
            Next
            m_seqs(group & "BlinkGroup").Sequence = buff
        End If
    End Sub

    Public Sub UpdateBlinkGroupInterval(group, interval)
        If m_seqs.Exists(group & "BlinkGroup") Then
            m_seqs(group & "BlinkGroup").UpdateInterval = interval
        End If 
    End Sub
    
    Public Sub StartBlinkGroup(group)
        If m_seqs.Exists(group & "BlinkGroup") Then
            If Not m_seqRunners.Exists("lSeqRunner" & group & "BlinkGroup") Then
                CreateSeqRunner "lSeqRunner" & group & "BlinkGroup", 1000
            End If
            m_seqRunners("lSeqRunner" & group & "BlinkGroup").AddSequenceItem m_seqs(group & "BlinkGroup")
        End If
    End Sub

    Public Sub StopBlinkGroup(group)
        If m_seqs.Exists(group & "BlinkGroup") Then
            RemoveLightSeq "lSeqRunner" & group & "BlinkGroup", m_seqs(group & "BlinkGroup").Name
        End If
    End Sub

    Public Function GetLightState(light)
        GetLightState = 0
        If(m_lights.Exists(light.name)) Then
            If m_on.Exists(light.name) Then
                GetLightState = 1
            Else
                If m_seqs.Exists(light.name & "Blink") Then
                    GetLightState = 2
                End If
            End If
        End If
    End Function

    Public Function IsShotLit(name, light)
        IsShotLit = False
        If(m_lights.Exists(light.name)) Then
            If m_seqRunners("lSeqRunner"&CStr(light.name)).HasSeq(name) Then
                IsShotLit = True
            End If
        End If
    End Function

    Public Sub CreateSeqRunner(name, priority)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        seqRunner.Priority = priority
        m_seqRunners.Add name, seqRunner
        Dim runnerKeys : runnerKeys = m_seqRunners.Keys()
        Dim itemsArray()
        ReDim itemsArray(m_seqRunners.Count - 1)
        Dim i, key
        i=0
        For Each key in runnerKeys
            Set itemsArray(i) = m_seqRunners(key)
            i=i+1
        Next
        
        Dim j, temp
        For i = 0 To UBound(itemsArray) - 1
            For j = i + 1 To UBound(itemsArray)
                If itemsArray(i).Priority > itemsArray(j).Priority Then
                    Set temp = itemsArray(i)
                    Set itemsArray(i) = itemsArray(j)
                    Set itemsArray(j) = temp
                End If
            Next
        Next
        m_seqRunners.RemoveAll
        For i = 0 To UBound(itemsArray)
            m_seqRunners.Add itemsArray(i).Name, itemsArray(i)
        Next
    End Sub

    Private Sub CreateOverrideSeqRunner(name)
        If m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqOverrideRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(runner, key, sequence, loops, speed, tokens, priority, syncMs)
        If Not m_seqRunners.Exists(runner) Then
            CreateSeqRunner runner, priority
        End If

        If syncMs = 0 Then
            If m_seqRunners(runner).Items().Exists(key) Then
                RemoveLightSeq runner, key
            End If
            m_seqRunners(runner).AddItem key, sequence, loops, speed, tokens,0
        Else
            If Not m_seqRunners.Exists(syncMs) Then
                CreateSeqRunner syncMs, 10
                m_seqRunners(syncMs).AddItem syncMs,Array("lXX|100", "lXX|0"), -1, 400/syncMs, Null, syncMs
            End If

            If Not m_syncMs.Exists(syncMs) Then
                m_syncMs.Add syncMs, CreateObject("Scripting.Dictionary")
            End If
            If m_syncMs(syncMs).Exists(runner & "_" & key) Then
                m_syncMs(syncMs).Remove runner & "_" & key
            End If
            m_syncMs(syncMs).Add runner & "_" & key, Array(runner, key, sequence, loops, speed, tokens)
        End If
    End Sub

    Public Sub RemoveLightSeq(runner, key)
        If Not m_seqRunners.Exists(runner) Then
            Exit Sub
        End If

        If m_seqRunners(runner).Items().Exists(key) Then
            Dim light
            For Each light in m_seqRunners(runner).ItemByKey(key).LightsInSeq
                If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
                End If
            Next
            m_seqRunners(runner).RemoveItem key
        End If
    End Sub

    Public Sub RemoveAllLightSeq(lcSeqRunner)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If
        Dim lcSeqKey, light, seqs, lcSeq
        Set seqs = m_seqRunners(lcSeqRunner).Items()
        For Each lcSeqKey in seqs.Keys()
            Set lcSeq = seqs(lcSeqKey)
            For Each light in lcSeq.LightsInSeq
                If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
                End If
            Next
        Next

        m_seqRunners(lcSeqRunner).RemoveAll
    End Sub

    Public Sub AddTableLightSeq(name, lcSeq)
        CreateOverrideSeqRunner(name)

        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
        m_seqOverrideRunners(name).AddSequenceItem lcSeq
    End Sub

    Public Sub RemoveTableLightSeq(name, lcSeq)
        If Not m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        m_seqOverrideRunners(name).RemoveItem lcSeq
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
            Dim light
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
        End If
    End Sub

    Public Sub RemoveAllTableLightSeqs()
        Dim light, runner
        For Each runner in m_seqOverrideRunners.Keys()
            m_seqOverrideRunners(runner).RemoveAll()
        Next
        For Each light in m_lights.Keys()
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
    End Sub

    Public Sub SyncLightMapColors()
        dim light,lm
        For Each light in m_lights.Keys()
            If m_lights(light).IsRGB Then
                dim color : color = m_lights(light).Color
                If m_lightmaps.Exists(light) Then
                    For Each lm in m_lightmaps(light)
                        If not IsNull(lm) Then
                            lm.Color = color(0)
                        End If
                    Next
                End If
            End If
        Next
    End Sub

    Public Sub SyncWithVpxLights(lightSeq)
        m_vpxLightSyncCollection = ColToArray(eval(lightSeq.collection))
        m_vpxLightSyncRunning = True
        m_tableSeqSpeed = Null
        m_tableSeqOffset = 0
        m_tableSeqDirection = Null
    End Sub

    Public Sub StopSyncWithVpxLights()
        m_vpxLightSyncRunning = False
        m_vpxLightSyncClear = True
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        m_tableSeqSpeed = Null
        m_tableSeqOffset = 0
        m_tableSeqDirection = Null
    End Sub

    Public Sub SetVpxSyncLightColor(color)
        m_tableSeqColor = color
    End Sub
    
    Public Sub SetVpxSyncLightsPalette(palette, direction, speed)
        m_tableSeqColor = palette
        Select Case direction:
            Case "BottomToTop": 
                m_tableSeqDirection = m_coordsY
                m_tableSeqColor = ReverseArray(palette)
            Case "TopToBottom": 
                m_tableSeqDirection = m_coordsY
            Case "RightToLeft": 
                m_tableSeqDirection = m_coordsX
            Case "LeftToRight": 
                m_tableSeqDirection = m_coordsX
                m_tableSeqColor = ReverseArray(palette)       
            Case "RadialOut": 
                m_tableSeqDirection = m_radii      
            Case "RadialIn": 
                m_tableSeqDirection = m_radii
                m_tableSeqColor = ReverseArray(palette) 
            Case "Clockwise": 
                m_tableSeqDirection = m_angles
            Case "AntiClockwise": 
                m_tableSeqDirection = m_angles
                m_tableSeqColor = ReverseArray(palette) 
        End Select  

        m_tableSeqSpeed = speed
    End Sub

    Public Sub SetTableSequenceFade(fadeUp, fadeDown)
        m_tableSeqFadeUp = fadeUp
        m_tableSeqFadeDown = fadeDown
    End Sub

    Public Function GetLightIdx(light)
        Dim lightName
        If vartype(light) = 8 Then
            lightName = light
        Else
            lightName = light.name
        End If
        dim syncLight : syncLight = Null
        
        If m_lights.Exists(lightName) Then
            'found a light
            Set syncLight = m_lights(lightName)
        End If
        If Not IsNull(syncLight) Then
            'Found a light to sync.
            GetLightIdx = lightsToLeds(syncLight.Idx)
        Else
            GetLightIdx = Null
        End If
        
    End Function

    Private Function m_buildBlinkSeq(lightName, pattern)
        Dim i, buff : buff = Array()
        ReDim buff(Len(pattern)-1)
        For i = 0 To Len(pattern)-1
            
            If Mid(pattern, i+1, 1) = 1 Then
                buff(i) = lightName & "|100"
            Else
                buff(i) = lightName & "|0"
            End If
        Next
        m_buildBlinkSeq=buff
    End Function

    Private Function GetTmpLight(idx)
        If IsArray(Lights(idx) ) Then	'if array
            Set GetTmpLight = Lights(idx)(0)
        Else
            Set GetTmpLight = Lights(idx)
        End If
    End Function

    Public Sub ResetLights()
        Dim light
        For Each light in m_lights.Keys()
            m_seqRunners("lSeqRunner"&CStr(light)).RemoveAll
            m_lightOff(light) 
            AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
        Next
        RemoveAllTableLightSeqs()
        Dim k
        For Each k in m_seqRunners.Keys()
            Dim lsRunner: Set lsRunner = m_seqRunners(k)
            lsRunner.RemoveAll
        Next

    End Sub

    Public Sub PauseMainLights
        If m_pauseMainLights = False Then
            m_pauseMainLights = True
            m_pausedLights.RemoveAll
            Dim pon
            Set pon = CreateObject("Scripting.Dictionary")
            Dim poff : Set poff = CreateObject("Scripting.Dictionary")
            Dim ppulse : Set ppulse = CreateObject("Scripting.Dictionary")
            Dim pseqs : Set pseqs = CreateObject("Scripting.Dictionary")
            Dim lightProps : Set lightProps = CreateObject("Scripting.Dictionary")
            'Add State in
            Dim light, item
            For Each item in m_on.Keys()
                pon.add item, m_on(Item)
            Next
            For Each item in m_off.Keys()
                poff.add item, m_off(Item)
            Next
            For Each item in m_pulse.Keys()
                ppulse.add item, m_pulse(Item)
            Next
            For Each item in m_seqRunners.Keys()
                dim tmpSeq : Set tmpSeq = new LCSeqRunner
                dim seqItem
                For Each seqItem in m_seqRunners(Item).Items.Items()
                    tmpSeq.AddSequenceItem seqItem
                Next
                tmpSeq.CurrentItemIdx = m_seqRunners(Item).CurrentItemIdx
                pseqs.add item, tmpSeq
            Next
            
            Dim savedProps(1,3)
            
            For Each light in m_lights.Keys()
                    
                savedProps(0,0) = m_lights(light).Color
                savedProps(0,1) = m_lights(light).Level
                If m_seqs.Exists(light & "Blink") Then
                    savedProps(0,2) = m_seqs.Item(light & "Blink").UpdateInterval
                Else
                    savedProps(0,2) = Empty
                End If
                lightProps.add light, savedProps
            Next
            m_pausedLights.Add "on", pon
            m_pausedLights.Add "off", poff
            m_pausedLights.Add "pulse", ppulse
            m_pausedLights.Add "runners", pseqs
            m_pausedLights.Add "lightProps", lightProps
            m_on.RemoveAll
            m_off.RemoveAll
            m_pulse.RemoveAll
            For Each item in m_seqRunners.Items()
                item.removeAll
            Next			
        End If
    End Sub

    Public Sub ResumeMainLights
        If m_pauseMainLights = True Then
            m_pauseMainLights = False
            m_on.RemoveAll
            m_off.RemoveAll
            m_pulse.RemoveAll
            Dim light, item
            For Each light in m_lights.Keys()
                AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            Next
            For Each item in m_seqRunners.Items()
                item.removeAll
            Next
            'Add State in
            For Each item in m_pausedLights("on").Keys()
                m_on.add item, m_pausedLights("on")(Item)
            Next
            For Each item in m_pausedLights("off").Keys()
                m_off.add item, m_pausedLights("off")(Item)
            Next			
            For Each item in m_pausedLights("pulse").Keys()
                m_pulse.add item, m_pausedLights("pulse")(Item)
            Next
            For Each item in m_pausedLights("runners").Keys()
                
                
                Set m_seqRunners(Item) = m_pausedLights("runners")(Item)
            Next
            For Each item in m_pausedLights("lightProps").Keys()
                LightColor Eval(Item), m_pausedLights("lightProps")(Item)(0,0)
                LightLevel Eval(Item), m_pausedLights("lightProps")(Item)(0,1)
                If Not IsEmpty(m_pausedLights("lightProps")(Item)(0,2)) Then
                    UpdateBlinkInterval Eval(Item), m_pausedLights("lightProps")(Item)(0,2)
                End If
            Next			
            m_pausedLights.RemoveAll
        End If
    End Sub

    Public Sub Update()

        'm_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
        m_frameTime = 20
        Dim x
        Dim lk
        dim color
        dim idx
        Dim lightKey
        Dim lcItem
        Dim tmpLight
        Dim seqOverride, hasOverride
        hasOverride = False
        For Each seqOverride In m_seqOverrideRunners.Keys()
            If Not IsNull(m_seqOverrideRunners(seqOverride).CurrentItem) Then
                RunLightSeq m_seqOverrideRunners(seqOverride)
                hasOverride = True
            End If
        Next
        If hasOverride = False Then
        
            If HasKeys(m_on) Then   
                For Each lightKey in m_on.Keys()
                    Set lcItem = m_on(lightKey)
                    If lcItem.IsRGB Then
                        color = m_on(lightKey).Color
                    Else
                        color = Null
                    End If
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_off) Then
                For Each lightKey in m_off.Keys()
                    Set lcItem = m_off(lightKey)
                    AssignStateForFrame lightKey, (new FrameState)(0, Null, lcItem.Idx)
                Next
            End If

        
            If HasKeys(m_seqRunners) Then
                Dim k
                For Each k in m_seqRunners.Keys()
                    Dim lsRunner: Set lsRunner = m_seqRunners(k)
                    If Not IsNull(lsRunner.CurrentItem) Then
                            RunLightSeq lsRunner
                    End If
                Next
            End If

            If HasKeys(m_pulse) Then   
                For Each lightKey in m_pulse.Keys()
                    Dim pulseColor
                    If m_pulse(lightKey).IsRGB Then
                        pulseColor = m_pulse(lightKey).Color
                    Else
                        pulseColor = Null
                    End If
                    If IsNull(pulseColor) Then
                        AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), Null, m_pulse(lightKey).light.Idx)
                    Else
                        AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), pulseColor, m_pulse(lightKey).light.Idx)
                    End If						
                    
                    Dim pulseUpdateInt : pulseUpdateInt = m_pulse(lightKey).interval - m_frameTime
                    Dim pulseIdx : pulseIdx = m_pulse(lightKey).idx
                    If pulseUpdateInt <= 0 Then
                        pulseUpdateInt = m_pulseInterval
                        pulseIdx = pulseIdx + 1
                    End If
                    
                    Dim pulses : pulses = m_pulse(lightKey).pulses
                    Dim pulseCount : pulseCount = m_pulse(lightKey).Cnt
                    
                    
                    If pulseIdx > UBound(m_pulse(lightKey).pulses) Then
                        m_pulse.Remove lightKey    
                        If pulseCount > 0 Then
                            pulseCount = pulseCount - 1
                            pulseIdx = 0
                            m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount, pulseColor)
                        End If
                    Else
                        m_pulse.Remove lightKey
                        m_pulse.Add lightKey, (new PulseState)(m_lights(lightKey),pulses, pulseIdx, pulseUpdateInt, pulseCount, pulseColor)
                    End If
                Next
            End If

            If m_vpxLightSyncRunning = True Then
                Dim lx
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lx in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncLight : syncLight = Null
                        If m_lights.Exists(lx.name) Then
                            'found a light
                            Set syncLight = m_lights(lx.name)
                        End If
                        If Not IsNull(syncLight) Then
                            'Found a light to sync.
                            
                            Dim lightState

                            If syncLight.IsRGB Then
                                If IsNull(m_tableSeqColor) Then
                                    color = syncLight.Color
                                Else
                                    If Not IsArray(m_tableSeqColor) Then
                                        color = Array(m_tableSeqColor, Null)
                                    Else
                                        
                                        If Not IsNull(m_tableSeqSpeed) And Not m_tableSeqSpeed = 0 Then

                                            Dim colorPalleteIdx : colorPalleteIdx = IncrementUInt8(m_tableSeqDirection(lightsToLeds(syncLight.Idx)),m_tableSeqOffset)
                                            If gametime mod m_tableSeqSpeed = 0 Then
                                                m_tableSeqOffset = m_tableSeqOffset + 1
                                                If m_tableSeqOffset > 255 Then
                                                    m_tableSeqOffset = 0
                                                End If	
                                            End If
                                            If colorPalleteIdx < 0 Then 
                                                colorPalleteIdx = 0
                                            End If
                                            color = Array(m_TableSeqColor(Round(colorPalleteIdx)), Null)
                                            'color = syncLight.Color
                                        Else
                                            color = Array(m_TableSeqColor(m_tableSeqDirection(lightsToLeds(syncLight.Idx))), Null)
                                        End If
                                        
                                    End If
                                End If
                            Else
                                color = Null
                            End If
                    
                            AssignStateForFrame syncLight.name, (new FrameState)(lx.GetInPlayState*100,color, syncLight.Idx)                     
                        End If
                    Next
                End If
            End If

            If m_vpxLightSyncClear = True Then  
                If Not IsNull(m_vpxLightSyncCollection) Then
                    For Each lk in m_vpxLightSyncCollection
                        'sync each light being ran by the vpx LS
                        dim syncClearLight : syncClearLight = Null
                        If m_lights.Exists(lk.name) Then
                            'found a light
                            Set syncClearLight = m_lights(lk.name)
                        End If
                        If Not IsNull(syncClearLight) Then
                            AssignStateForFrame syncClearLight.name, (new FrameState)(0, Null, syncClearLight.idx) 
                        End If
                    Next
                End If
            
                m_vpxLightSyncClear = False
            End If
        End If
        

        If HasKeys(m_currentFrameState) Then
            
            Dim frameStateKey
            For Each frameStateKey in m_currentFrameState.Keys()
                idx = m_currentFrameState(frameStateKey).idx
                
                Dim newColor : newColor = m_currentFrameState(frameStateKey).colors
                Dim bUpdate

                If Not IsNull(newColor) Then
                    'Check current color is the new color coming in, if not, set the new color.
                    
                    Set tmpLight = GetTmpLight(idx)
                    Dim c, cf
                    c = newColor(0)
                    cf= newColor(1)

                    If Not IsNull(c) Then
                        If Not CStr(tmpLight.Color) = CStr(c) Then
                            bUpdate = True
                        End If
                    End If

                    If Not IsNull(cf) Then
                        If Not CStr(tmpLight.ColorFull) = CStr(cf) Then
                            bUpdate = True
                        End If
                    End If
                End If

               
                Dim lm
                If IsArray(Lights(idx)) Then
                    For Each x in Lights(idx)
                        If bUpdate Then 
                            If Not IsNull(c) Then
                                x.color = c
                            End If
                            If Not IsNull(cf) Then
                                x.colorFull = cf
                            End If
                            If m_lightmaps.Exists(x.Name) Then
                                For Each lm in m_lightmaps(x.Name)
                                    If Not IsNull(lm) Then
                                        lm.Color = c
                                    End If
                                Next
                            End If
                        End If
                        x.State = m_currentFrameState(frameStateKey).level/100
                    Next
                Else
                    If bUpdate Then    
                        If Not IsNull(c) Then
                            Lights(idx).color = c
                        End If
                        If Not IsNull(cf) Then
                            Lights(idx).colorFull = cf
                        End If
                        If m_lightmaps.Exists(Lights(idx).Name) Then
                            For Each lm in m_lightmaps(Lights(idx).Name)
                                If Not IsNull(lm) Then
                                    lm.Color = c
                                End If
                            Next
                        End If
                    End If
                    Lights(idx).State = m_currentFrameState(frameStateKey).level/100
                End If
            Next
        End If
        m_currentFrameState.RemoveAll
        m_off.RemoveAll

    End Sub

    Private Function HexToInt(hex)
        HexToInt = CInt("&H" & hex)
    End Function

    Function RGBToHex(r, g, b)
        RGBToHex = Right("0" & Hex(r), 2) & _
            Right("0" & Hex(g), 2) & _
            Right("0" & Hex(b), 2)
    End Function

    Function FadeRGB(light, color1, color2, steps)

    
        Dim r1, g1, b1, r2, g2, b2
        Dim i
        Dim r, g, b
        color1 = clng(color1)
        color2 = clng(color2)
        ' Extract RGB values from the color integers
        r1 = color1 Mod 256
        g1 = (color1 \ 256) Mod 256
        b1 = (color1 \ (256 * 256)) Mod 256

        r2 = color2 Mod 256
        g2 = (color2 \ 256) Mod 256
        b2 = (color2 \ (256 * 256)) Mod 256

        ' Resize the output array
        ReDim outputArray(steps - 1)

        ' Generate the fade
        For i = 0 To steps - 1
            ' Calculate RGB values for this step
            r = r1 + (r2 - r1) * i / (steps - 1)
            g = g1 + (g2 - g1) * i / (steps - 1)
            b = b1 + (b2 - b1) * i / (steps - 1)

            ' Convert RGB to hex and add to output
            outputArray(i) = light & "|100|" & RGBToHex(CInt(r), CInt(g), CInt(b))
        Next
        FadeRGB = outputArray
    End Function

    Public Function CreateColorPalette(startColor, endColor, steps)
        Dim colors()
        ReDim colors(steps)
        
        Dim startRed, startGreen, startBlue, endRed, endGreen, endBlue
        startRed = HexToInt(Left(startColor, 2))
        startGreen = HexToInt(Mid(startColor, 3, 2))
        startBlue = HexToInt(Right(startColor, 2))
        endRed = HexToInt(Left(endColor, 2))
        endGreen = HexToInt(Mid(endColor, 3, 2))
        endBlue = HexToInt(Right(endColor, 2))
        
        Dim redDiff, greenDiff, blueDiff
        redDiff = endRed - startRed
        greenDiff = endGreen - startGreen
        blueDiff = endBlue - startBlue
        
        Dim i
        For i = 0 To steps
            Dim red, green, blue
            red = startRed + (redDiff * (i / steps))
            green = startGreen + (greenDiff * (i / steps))
            blue = startBlue + (blueDiff * (i / steps))
            colors(i) = RGB(red,green,blue)'IntToHex(red, 2) & IntToHex(green, 2) & IntToHex(blue, 2)
        Next
        
        CreateColorPalette = colors
    End Function

    Function CreateColorPaletteWithStops(startColor, endColor, stopPositions, stopColors)

        Dim colors(255)

        Dim fStop : fStop = CreateColorPalette(startColor, stopColors(0), stopPositions(0))
        Dim i, istep
        For i = 0 to stopPositions(0)
            colors(i) = fStop(i)
        Next
        For i = 1 to Ubound(stopColors)
            Dim stopStep : stopStep = CreateColorPalette(stopColors(i-1), stopColors(i), stopPositions(i))
            Dim ii
        ' MsgBox(stopPositions(i) - stopPositions(i-1))
            istep = 0
            For ii = stopPositions(i-1)+1 to stopPositions(i)
            '  MsgBox(ii)
            colors(ii) = stopStep(iStep)
            iStep = iStep + 1
            Next
        Next
        ' MsgBox("Here")
        Dim eStop : eStop = CreateColorPalette(stopColors(UBound(stopColors)), endColor, 255-stopPositions(UBound(stopPositions)))
        'MsgBox(UBound(eStop))
        iStep = 0
        For i = 255-(255-stopPositions(UBound(stopPositions))) to 254
            colors(i) = eStop(iStep)
            iStep = iStep + 1
        Next

        CreateColorPaletteWithStops = colors
    End Function

    Private Function HasKeys(o)
        If Ubound(o.Keys())>-1 Then
            HasKeys = True
        Else
            HasKeys = False
        End If
    End Function

    Private Sub RunLightSeq(seqRunner)

        Dim lcSeq: Set lcSeq = seqRunner.CurrentItem
        dim lsName, isSeqEnd
        If UBound(lcSeq.Sequence)<lcSeq.CurrentIdx Then
            isSeqEnd = True
        Else
            isSeqEnd = False
        End If
        dim lightInSeq
        For each lightInSeq in lcSeq.LightsInSeq
        
            If isSeqEnd Then

                

            'Needs a guard here for something, but i've forgotten. 
            'I remember: Only reset the light if there isn't frame data for the light. 
            'e.g. a previous seq has affected the light, we don't want to clear that here on this frame
                If m_lights.Exists(lightInSeq) = True AND NOT m_currentFrameState.Exists(lightInSeq) Then
                    AssignStateForFrame lightInSeq, (new FrameState)(0, Null, m_lights(lightInSeq).Idx)
                    LightColor m_lights(lightInSeq), m_lights(lightInSeq).BaseColor
                End If
            Else
                


                If m_currentFrameState.Exists(lightInSeq) Then

                    
                    'already frame data for this light.
                    'replace with the last known state from this seq
                    If Not IsNull(lcSeq.LastLightState(lightInSeq)) Then
                        AssignStateForFrame lightInSeq, lcSeq.LastLightState(lightInSeq)
                    End If
                End If

            End If
        Next

        If isSeqEnd Then
            'Check if any seqs need starting via syncms
            If lcSeq.SyncMs > 0 Then
                If m_syncMs.Exists(lcSeq.SyncMs) Then
                    Dim seqToStart
                    For Each seqToStart in m_syncMs(lcSeq.SyncMs).Items
                        m_seqRunners(seqToStart(0)).AddItem seqToStart(1), seqToStart(2), seqToStart(3), seqToStart(4), seqToStart(5), 0
                        RunLightSeq(m_seqRunners(seqToStart(0)))
                    Next
                    m_syncMs(lcSeq.SyncMs).RemoveAll
                End If
            End If
            lcSeq.CurrentIdx = 0
            seqRunner.NextItem()
        End If

        If Not IsNull(seqRunner.CurrentItem) Then
            Dim framesRemaining, seq, color
            Set lcSeq = seqRunner.CurrentItem
            seq = lcSeq.Sequence
            

            Dim name
            Dim ls, x
            If IsArray(seq(lcSeq.CurrentIdx)) Then
                For x = 0 To UBound(seq(lcSeq.CurrentIdx))
                    lsName = Split(seq(lcSeq.CurrentIdx)(x),"|")
                    name = lsName(0)
                    If m_lights.Exists(name) Then
                        Set ls = m_lights(name)
                        
                        If ls.IsRGB Then
                            color = lcSeq.Color

                            If IsNull(color) Then
                                color = ls.Color
                            End If
                        Else
                            Log "NOT RGB:" & ls.Idx
                            color = Null
                            If Ubound(lsName) = 2 Then
                                If lsName(2) <> "" Then
                                    color = Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), Null)        
                                End If
                            End If
                        End If

                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        
                        lcSeq.SetLastLightState name, m_currentFrameState(name) 
                    End If
                Next       
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)
                    
                    If ls.IsRGB Then
                        color = lcSeq.Color

                        If IsNull(color) Then
                            color = ls.Color
                        End If
                    Else
                        color = Null
                        If Ubound(lsName) = 2 Then
                            If lsName(2) <> "" Then
                                color = Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), Null)        
                            End If
                        End If
                    End If

                    AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                    lcSeq.SetLastLightState name, m_currentFrameState(name) 
                End If
            End If

            framesRemaining = lcSeq.Update(m_frameTime)
            If framesRemaining < 0 Then
                lcSeq.ResetInterval()
                lcSeq.NextFrame()
            End If
            
        End If
    
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog "Light Controller", message
        End If
    End Sub
End Class

Class FrameState
    Private m_level, m_colors, m_idx

    Public Property Get Level(): Level = m_level: End Property
    Public Property Let Level(input): m_level = input: End Property

    Public Property Get Colors(): Colors = m_colors: End Property
    Public Property Let Colors(input): m_colors = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public default function init(level, colors, idx)
        m_level = level
        m_colors = colors
        m_idx = idx 

        Set Init = Me
    End Function

    Public Function ColorAt(idx)
        ColorAt = m_colors(idx) 
    End Function
End Class

Class PulseState
    Private m_light, m_pulses, m_idx, m_interval, m_cnt, m_color

    Public Property Get Light(): Set Light = m_light: End Property
    Public Property Let Light(input): Set m_light = input: End Property

    Public Property Get Pulses(): Pulses = m_pulses: End Property
    Public Property Let Pulses(input): m_pulses = input: End Property

    Public Property Get Idx(): Idx = m_idx: End Property
    Public Property Let Idx(input): m_idx = input: End Property

    Public Property Get Interval(): Interval = m_interval: End Property
    Public Property Let Interval(input): m_interval = input: End Property

    Public Property Get Cnt(): Cnt = m_cnt: End Property
    Public Property Let Cnt(input): m_cnt = input: End Property

    Public Property Get Color(): Color = m_color: End Property
    Public Property Let Color(input): m_color = input: End Property		

    Public default function init(light, pulses, idx, interval, cnt, color)
        Set m_light = light
        m_pulses = pulses
        m_idx = idx 
        m_interval = interval
        m_cnt = cnt
        m_color = color

        Set Init = Me
    End Function

    Public Function PulseAt(idx)
        PulseAt = m_pulses(idx) 
    End Function
End Class

Class LCItem
    
    Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y, m_baseColor, m_isRGB

        Public Property Get Idx()
            Idx=m_Idx
        End Property

        Public Property Get IsRGB()
            IsRGB = m_isRGB
        End Property

        Public Property Get Color()
            Color=m_color
        End Property

        Public Property Get BaseColor()
            BaseColor=m_baseColor
        End Property

        Public Property Let Color(input)
            If IsNull(input) Then
                m_Color = Null
            Else
                If Not IsArray(input) Then
                    input = Array(input, null)
                End If
                m_Color = input
            End If
        End Property

        Public Property Let Level(input)
            m_level = input
        End Property

        Public Property Get Level()
            Level=m_level
        End Property

        Public Property Get Name()
            Name=m_name
        End Property

        Public Property Get X()
            X=m_x
        End Property

        Public Property Get Y()
            Y=m_y
        End Property

        Public Property Get Row()
            Row=Round(m_x/20)
        End Property

        Public Property Get Col()
            Col=Round(m_y/20)
        End Property

        Public Sub Init(idx, intervalMs, color, name, x, y)
            m_Idx = idx
            If Not IsArray(color) Then
                m_color = Array(color, null)
            Else
                m_color = color
            End If
            m_baseColor = m_color
            If CLng(color) = vbWhite Then
                m_isRGB = True
            Else
                m_isRGB = False
            End If
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
        End Sub

End Class

Class LCSeq
    
    Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates, m_palette, m_pattern, m_tokens, m_loops, m_syncMs

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
        m_lastLightStates.RemoveAll()
        m_currentIdx = input
    End Property

    Public Property Get Tokens()
        Tokens=m_tokens
    End Property

    Public Property Let Tokens(input)
        If Not IsNull(input) Then
            Set m_tokens = input
        End If
    End Property    

    Public Property Get LightsInSeq()
        LightsInSeq=m_lightsInSeq.Keys()
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property
    
    Public Property Let Sequence(input)
        Dim item, light, lightItem, token
        m_lightsInSeq.RemoveAll
        Dim i, x
        ' Iterate through the input array
        For i = 0 To UBound(input)
            item = input(i)
            If IsArray(item) Then
                For x = 0 To UBound(item)
                    light = item(x)
                    lightItem = Split(light, "|")
                    token = IsToken(lightItem(0))
                    If Not IsNull(token) Then
                        lightItem(0) = m_tokens(token)
                    End If
                    If UBound(lightItem) = 2 Then
                        token = IsToken(lightItem(2))
                        If Not IsNull(token) Then
                            lightItem(2) = m_tokens(token)
                        End If
                    End If
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If
                    light = Join(lightItem, "|")
                    item(x) = light
                Next
                ' Update the input array with modified item array
                input(i) = item
            Else
                lightItem = Split(item, "|")
                token = IsToken(lightItem(0))
                If Not IsNull(token) Then
                    lightItem(0) = m_tokens(token)
                End If
                If UBound(lightItem) = 2 Then
                    token = IsToken(lightItem(2))
                    If Not IsNull(token) Then
                        lightItem(2) = m_tokens(token)
                    End If
                End If
                If Not m_lightsInSeq.Exists(lightItem(0)) Then
                    m_lightsInSeq.Add lightItem(0), True
                End If
                light = Join(lightItem, "|")
                ' Update the input array with modified item string
                input(i) = light
            End If
        Next
        ' Assign the modified input array to m_sequence
        m_sequence = input
    End Property

    Private Function IsToken(mainString)
        ' Check if the string contains an opening parenthesis and ends with a closing parenthesis
        If InStr(mainString, "(") > 0 And Right(mainString, 1) = ")" Then
            ' Extract the substring within the parentheses
            Dim startPos, subString
            startPos = InStr(mainString, "(")
            subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
            IsToken = subString
        Else
            IsToken = Null
        End If
    End Function
    

    Public Property Get LastLightState(light)
        If m_lastLightStates.Exists(light) Then
            dim c : Set c = m_lastLightStates(light)
            Set LastLightState = c
        Else
            LastLightState = Null
        End If
    End Property

    Public Property Let LastLightState(light, input)
        If m_lastLightStates.Exists(light) Then
            m_lastLightStates.Remove light
        End If
        If input.level > 0 Then
            m_lastLightStates.Add light, input
        End If
    End Property

    Public Sub SetLastLightState(light, input)	
        If m_lastLightStates.Exists(light) Then	
            m_lastLightStates.Remove light	
        End If	
        If input.level > 0 Then	
                m_lastLightStates.Add light, input	
        End If	
    End Sub

    Public Property Get Color()
        Color=m_color
    End Property
    
    Public Property Let Color(input)
        If IsNull(input) Then
            m_Color = Null
        Else
            If Not IsArray(input) Then
                input = Array(input, null)
            End If
            m_Color = input
        End If
    End Property

    Public Property Get Palette()
        Palette=m_palette
    End Property
    
    Public Property Let Palette(input)
        If IsNull(input) Then
            m_palette = Null
        Else
            If Not IsArray(input) Then
                m_palette = Null
            Else
                m_palette = input
            End If
        End If
    End Property

    Public Property Get Name()
        Name=m_name
    End Property
    
    Public Property Let Name(input)
        m_name = input
    End Property        

    Public Property Get UpdateInterval()
        UpdateInterval=m_updateInterval
    End Property

    Public Property Let UpdateInterval(input)
        m_updateInterval = input
        'm_Frames = input
    End Property

    Public Property Get Repeat()
        Repeat=m_repeat
    End Property

    Public Property Let Repeat(input)
        m_repeat = input
    End Property

    Public Property Get Loops()
        Loops=m_loops
    End Property

    Public Property Let Loops(input)
        m_loops = input
    End Property

    Public Property Get Pattern()
        Pattern=m_pattern
    End Property

    Public Property Let Pattern(input)
        m_pattern = input
    End Property    

    Public Property Get SyncMs()
        SyncMs=m_syncMs
    End Property
    
    Public Property Let SyncMs(input)
        m_syncMs = input
    End Property
    
    Public Property Get Frames()
        Frames=m_Frames
    End Property
    
    Public Property Let Frames(input)
        m_Frames = input
    End Property

    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Null
        m_updateInterval = 200
        m_repeat = False
        m_loops = 1
        m_Frames = 200
        m_pattern = Null
        m_syncMs = 0
        Set m_lightsInSeq = CreateObject("Scripting.Dictionary")
        Set m_lastLightStates = CreateObject("Scripting.Dictionary")
    End Sub

    Public Property Get Update(framesPassed)
        m_Frames = m_Frames - framesPassed
        Update = m_Frames
    End Property

    Public Sub NextFrame()
        m_currentIdx = m_currentIdx + 1
    End Sub

    Public Sub ResetInterval()
        m_Frames = m_updateInterval
        Exit Sub
    End Sub

End Class

Class LCSeqRunner
    
    Private m_name, m_items, m_currentItemIdx, m_priority

    Public Property Get Name()
        Name=m_name
    End Property
    
    Public Property Let Name(input)
        m_name = input
    End Property

    Public Property Get Priority()
        Priority=m_priority
    End Property

    Public Property Let Priority(input)
        m_priority = input
    End Property

    Public Property Get Items()
        Set Items = m_items
    End Property

    Public Property Get ItemByKey(key)
        Set ItemByKey = m_items(key)
    End Property

    Public Property Get CurrentItemIdx()
        CurrentItemIdx = m_currentItemIdx
    End Property

    Public Property Let CurrentItemIdx(input)
        m_currentItemIdx = input
    End Property

    Public Property Get CurrentItem()
        Dim items: items = m_items.Items()
        If m_currentItemIdx > UBound(items) Then
            m_currentItemIdx = 0
        End If
        If UBound(items) = -1 Then       
            CurrentItem  = Null
        Else
            Set CurrentItem = items(m_currentItemIdx)                
        End If
    End Property

    Private Sub Class_Initialize()    
        Set m_items = CreateObject("Scripting.Dictionary")
        m_currentItemIdx = 0
        m_priority = 1000
    End Sub

    Public Sub AddItem(key, sequence, loops, speed, tokens, syncMs)
        If Not IsNull(sequence) Then

            Dim lSeq : Set lSeq = New LCSeq
            lSeq.Name = key
            lSeq.Tokens = tokens
            lSeq.Sequence = sequence
            lSeq.UpdateInterval = 200/speed
            lSeq.Frames = 200/speed
            lSeq.Loops = loops
            lSeq.SyncMs = syncMs

            If Not m_items.Exists(key) Then
                m_items.Add key, lSeq
            End If
        End If
    End Sub

    Public Sub AddSequenceItem(sequence)
        If Not IsNull(sequence) Then
            Dim loops
            If sequence.Repeat = True Then
                sequence.Loops = -1
            Else
                sequence.Loops = 1
            End If
            If Not m_items.Exists(sequence.Name) Then
                m_items.Add sequence.Name, sequence
            End If
        End If
    End Sub

    Public Sub RemoveAll()
        Dim item
        For Each item in m_items.Keys()
            m_items(item).ResetInterval
            m_items(item).CurrentIdx = 0
            m_items.Remove item
        Next
    End Sub

    Public Sub RemoveItem(key)
        If m_items.Exists(key) Then
            m_items(key).ResetInterval
            m_items(key).CurrentIdx = 0
            m_items.Remove key
        End If
    End Sub

    Public Sub NextItem()
        Dim items: items = m_items.Items
        Dim keys : keys = m_items.Keys
        If items(m_currentItemIdx).Loops = 1 Then
            RemoveItem(keys(m_currentItemIdx))
        Else
            If items(m_currentItemIdx).Loops > 1 Then
                items(m_currentItemIdx).Loops = items(m_currentItemIdx).Loops - 1
            End If
            m_currentItemIdx = m_currentItemIdx + 1
        End If
        
        If m_currentItemIdx > UBound(m_items.Items) Then   
            m_currentItemIdx = 0
        End If
    End Sub

    Public Function HasSeq(key)
        If m_items.Exists(key) Then
            HasSeq = True
        Else
            HasSeq = False
        End If
    End Function

End Class

'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Class GlfDebugLogFile
	Private Filename
	Private TxtFileStream

	Public default Function init()
        Filename = cGameName + "_" & GetTimeStamp & "_debug_log.txt"
	  Set Init = Me
	End Function
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
		GetTimeStamp = _
		LZ(Year(CurrTime),   4) & "-" _
		 & LZ(Month(CurrTime),  2) & "-" _
		 & LZ(Day(CurrTime),	2) & "_" _
		 & LZ(Hour(CurrTime),   2) & "_" _
		 & LZ(Minute(CurrTime), 2) & "_" _
		 & LZ(Second(CurrTime), 2) & "_" _
		 & LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message)
		If glf_debugEnabled = True Then
			Dim FormattedMsg, Timestamp
			
			Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile("logs\"&Filename, 8, True)
			Timestamp = GetTimeStamp
			FormattedMsg = GetTimeStamp + ": " + label + ": " + message
			TxtFileStream.WriteLine FormattedMsg
			TxtFileStream.Close
			Debug.print label & ": " & message
		End If
	End Sub
End Class

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************


Dim glf_lastPinEvent : glf_lastPinEvent = Null

Sub DispatchPinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        glf_debugLog.WriteToLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    glf_debugLog.WriteToLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        glf_debugLog.WriteToLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            GetRef(handler(0))(Array(handler(2), kwargs, e))
        Else
            glf_debugLog.WriteToLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
    Glf_EventBlocks(e).RemoveAll
End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        'glf_debugLog.WriteToLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    'glf_debugLog.WriteToLog "DispatchReplayPinEvent", e
    For Each k In glf_pinEventsOrder(e)
        'glf_debugLog.WriteToLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        kwargs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
    Next
    DispatchRelayPinEvent = kwargs
    Glf_EventBlocks(e).RemoveAll
End Function

Sub AddPinEventListener(evt, key, callbackName, priority, args)
    Dim i, inserted, tempArray
    If Not glf_pinEvents.Exists(evt) Then
        glf_pinEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_pinEvents(evt).Exists(key) Then
        glf_pinEvents(evt).Add key, Array(callbackName, priority, args)
        Sortglf_pinEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePinEventListener(evt, key)
    If glf_pinEvents.Exists(evt) Then
        If glf_pinEvents(evt).Exists(key) Then
            glf_pinEvents(evt).Remove key
            Sortglf_pinEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_pinEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_pinEventsOrder to maintain order based on priority
    If Not glf_pinEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_pinEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_pinEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_pinEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the glf_pinEventsOrder with the newly ordered or modified list
        glf_pinEventsOrder(evt) = tempArray
    End If
End Sub
Function GetPlayerState(key)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If playerState(glf_currentPlayer).Exists(key)  Then
        GetPlayerState = playerState(glf_currentPlayer)(key)
    Else
        GetPlayerState = Null
    End If
End Function

Function GetPlayerScore(player)
    dim p
    Select Case player
        Case 1:
            p = "PLAYER 1"
        Case 2:
            p = "PLAYER 2"
        Case 3:
            p = "PLAYER 3"
        Case 4:
            p = "PLAYER 4"
    End Select

    If playerState.Exists(p) Then
        GetPlayerScore = playerState(p)(SCORE)
    Else
        GetPlayerScore = 0
    End If
End Function

Function Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            Getglf_currentPlayerNumber = 1
        Case "PLAYER 2":
            Getglf_currentPlayerNumber = 2
        Case "PLAYER 3":
            Getglf_currentPlayerNumber = 3
        Case "PLAYER 4":
            Getglf_currentPlayerNumber = 4
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If IsArray(value) Then
        If Join(GetPlayerState(key)) = Join(value) Then
            Exit Function
        End If
    Else
        If GetPlayerState(key) = value Then
            Exit Function
        End If
    End If   
    Dim prevValue
    If playerState(glf_currentPlayer).Exists(key) Then
        prevValue = playerState(glf_currentPlayer)(key)
        playerState(glf_currentPlayer).Remove key
    End If
    playerState(glf_currentPlayer).Add key, value
    
    If glf_playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue
    End If
    
    SetPlayerState = Null
End Function

Sub FirePlayerEventHandlers(evt, value, prevValue)
    If Not glf_playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = glf_playerEvents(evt)
    For Each k In glf_playerEventsOrder(evt)
        GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), Array(evt,value,prevValue)))
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, callbackName, priority, args)
    If Not glf_playerEvents.Exists(evt) Then
        glf_playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_playerEvents(evt).Exists(key) Then
        glf_playerEvents(evt).Add key, Array(callbackName, priority, args)
        Sortglf_playerEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePlayerStateEventListener(evt, key)
    If glf_playerEvents.Exists(evt) Then
        If glf_playerEvents(evt).Exists(key) Then
            glf_playerEvents(evt).Remove key
            Sortglf_playerEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_playerEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_playerEventsOrder to maintain order based on priority
    If Not glf_playerEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_playerEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_playerEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_playerEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the glf_playerEventsOrder with the newly ordered or modified list
        glf_playerEventsOrder(evt) = tempArray
    End If
End Sub

Sub EmitAllglf_playerEvents()
    Dim key
    For Each key in playerState(glf_currentPlayer).Keys()
        FirePlayerEventHandlers key, GetPlayerState(key), GetPlayerState(key)
    Next
End Sub

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
	ballInReleasePostion = True
	UpdateTrough
End Sub
Sub swTrough1_UnHit
	ballInReleasePostion = False
	UpdateTrough
End Sub
Sub swTrough2_Hit
	UpdateTrough
End Sub
Sub swTrough2_UnHit
	UpdateTrough
End Sub
Sub swTrough3_Hit
	UpdateTrough
End Sub
Sub swTrough3_UnHit
	UpdateTrough
End Sub
Sub swTrough4_Hit
	UpdateTrough
End Sub
Sub swTrough4_UnHit
	UpdateTrough
End Sub
Sub swTrough5_Hit
	UpdateTrough
End Sub
Sub swTrough5_UnHit
	UpdateTrough
End Sub
Sub swTrough6_Hit
	UpdateTrough
End Sub
Sub swTrough6_UnHit
	UpdateTrough
End Sub
Sub swTrough7_Hit
	UpdateTrough
End Sub
Sub swTrough7_UnHit
	UpdateTrough
End Sub
Sub swTrough8_Hit
	UpdateTrough
    If glf_gameStarted = True Then
        glf_BIP = glf_BIP - 1
        DispatchRelayPinEvent GLF_BALL_DRAIN, 1
    End If
End Sub
Sub swTrough8_UnHit
	UpdateTrough
End Sub

Sub UpdateTrough
	SetDelay "update_trough", "UpdateTroughDebounced", Null, 100
End Sub

Sub UpdateTroughDebounced(args)
	If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
    If swTrough5.BallCntOver = 0 Then swTrough6.kick 57, 10
    If swTrough6.BallCntOver = 0 Then swTrough7.kick 57, 10
    If swTrough7.BallCntOver = 0 Then swTrough8.kick 57, 10
End Sub
