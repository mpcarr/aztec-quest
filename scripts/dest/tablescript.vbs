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


Dim gameDebugger : Set gameDebugger = new AdvGameDebugger
'*******************************************
'  Constants and Global Variables
'*******************************************

Const UsingROM = False		'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.

Const BallSize = 50			'Ball diameter in VPX units; must be 50
Const BallMass = 1			'Ball mass must be 1
Const tnob = 7				'Total number of balls the table can hold
Const lob = 2				'Locked balls
Const cGameName = "tmntpro"	'The unique alphanumeric name for this table

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP						'Balls in play
BIP = 0
Dim BIPL					'Ball in plunger lane
BIPL = False


Const IMPowerSetting = 50 			'Plunger Power
Const IMTime = 1.1        			'Time in seconds for Full Plunge
Dim plungerIM

Dim lightCtrl : Set lightCtrl = new LStateController
Dim gilvl : gilvl = 0  'General Illumination light state tracked for Dynamic Ball Shadows

'*******************************************
'  Table Initialization and Exiting
'*******************************************

LoadCoreFiles
Sub LoadCoreFiles
	On Error Resume Next
	ExecuteGlobal GetTextFile("core.vbs") 'TODO: drop-in replacement for vpmTimer (maybe vpwQueueManager) and cvpmDictionary (Scripting.Dictionary) to remove core.vbs dependency
	If Err Then MsgBox "Can't open core.vbs"
	On Error GoTo 0
End Sub

Dim tmntproBall1, tmntproBall2, tmntproBall3, tmntproBall4, tmntproBall5, gBOT, tmag, NewtonBall, CaptiveBall

Sub Table1_Init
	Dim i
	
	vpmMapLights alights
	lightCtrl.RegisterLights "VPX"
	'waterfalldiverter.isdropped=1

	'Ball initializations need for physical trough
	Set tmntproBall1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set tmntproBall2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set tmntproBall3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set tmntproBall4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set tmntproBall5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	
	'*** Use gBOT in the script wherever BOT is normally used. Then there is no need for GetBalls calls ***
	gBOT = Array( tmntproBall1, tmntproBall2, tmntproBall3, tmntproBall4, tmntproBall5)
	
	Dim xx
	
	' Add balls to shadow dictionary
	For Each xx In gBOT
		bsDict.Add xx.ID, bsNone
	Next
	
	' Make drop target shadows visible
	For Each xx In ShadowDT
		xx.visible = True
	Next

	Set plungerIM = New cvpmImpulseP
	With plungerIM
		.InitImpulseP swPlunger, IMPowerSetting, IMTime
		.Random 1.5
		.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
		.CreateEvents "plungerIM"
	End With
	PlayVPXSeq
	
	DTDrop 1
End Sub


Sub Table1_Exit
	gameDebugger.Disconnect
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
Set bsDict = New cvpmDictionary
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
        If Err Then MsgBox "Can't start advanced debugger" : m_connected = False
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


' VLM  Arrays - Start
' Arrays per baked part
Dim BP_Cab: BP_Cab=Array(BM_Cab)
Dim BP_PF: BP_PF=Array(BM_PF)
Dim BP_Panther: BP_Panther=Array(BM_Panther)
Dim BP_Parts: BP_Parts=Array(BM_Parts)
Dim BP_sw04: BP_sw04=Array(BM_sw04)
Dim BP_sw05: BP_sw05=Array(BM_sw05)
Dim BP_sw06: BP_sw06=Array(BM_sw06)
Dim BP_sw08: BP_sw08=Array(BM_sw08)
Dim BP_sw09: BP_sw09=Array(BM_sw09)
Dim BP_sw10: BP_sw10=Array(BM_sw10)
Dim BP_sw11: BP_sw11=Array(BM_sw11)
Dim BP_sw12: BP_sw12=Array(BM_sw12)
Dim BP_sw13: BP_sw13=Array(BM_sw13)
Dim BP_sw15: BP_sw15=Array(BM_sw15)
Dim BP_sw16: BP_sw16=Array(BM_sw16)
Dim BP_sw17: BP_sw17=Array(BM_sw17)
Dim BP_targetbank: BP_targetbank=Array(BM_targetbank)
' Arrays per lighting scenario
Dim BL_World: BL_World=Array(BM_Cab, BM_PF, BM_Panther, BM_Parts, BM_sw04, BM_sw05, BM_sw06, BM_sw08, BM_sw09, BM_sw10, BM_sw11, BM_sw12, BM_sw13, BM_sw15, BM_sw16, BM_sw17, BM_targetbank)
' Global arrays
Dim BG_Bakemap: BG_Bakemap=Array(BM_Cab, BM_PF, BM_Panther, BM_Parts, BM_sw04, BM_sw05, BM_sw06, BM_sw08, BM_sw09, BM_sw10, BM_sw11, BM_sw12, BM_sw13, BM_sw15, BM_sw16, BM_sw17, BM_targetbank)
Dim BG_Lightmap: BG_Lightmap=Array()
Dim BG_All: BG_All=Array(BM_Cab, BM_PF, BM_Panther, BM_Parts, BM_sw04, BM_sw05, BM_sw06, BM_sw08, BM_sw09, BM_sw10, BM_sw11, BM_sw12, BM_sw13, BM_sw15, BM_sw16, BM_sw17, BM_targetbank)
' VLM  Arrays - End


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
' This solution improves the physics for drop targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full drop target animation, including deflection on hit and a slight lift when the drop
' targets raise, switch handling, bricking, and popping the ball up if it's over the drop target when it raises.
'
'Add a Timer named DTAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
'DTAnim.interval = 10
'DTAnim.enabled = True

'Sub DTAnim_Timer
'	DoDTAnim
'	DoSTAnim
'End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.

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
  Set DT01 = (new DropTarget)(sw01, sw01a, BM_Panther, 1, 0, True) 
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
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
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
			'controller.Switch(Switchid) = 1
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
			Dim gBOT
			gBOT = GetBalls
			
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
		'controller.Switch(Switchid) = 0
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
  
Sub UpdateDropTargets
	Dim t
	'For Each t in BP_sw38 : t.transz = BM_sw38.transz : t.rotx = BM_sw38.rotx : t.roty = BM_sw38.roty : Next

	'For Each t in BP_sw40 : t.transz = BM_sw40.transz : t.rotx = BM_sw40.rotx : t.roty = BM_sw40.roty : Next

	'For Each t in BP_sw45 : t.transz = BM_sw45.transz : t.rotx = BM_sw45.rotx : t.roty = BM_sw45.roty : Next

	'For Each t in BP_sw46 : t.transz = BM_sw46.transz : t.rotx = BM_sw46.rotx : t.roty = BM_sw46.roty : Next

	'For Each t in BP_sw47 : t.transz = BM_sw47.transz : t.rotx = BM_sw47.rotx : t.roty = BM_sw47.roty : Next
End Sub
  
  


'******************************************************
'		STAND-UP TARGET INITIALIZATION
'******************************************************

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

ST11 = Array(sw11, BM_sw11, 11, 0)
ST12 = Array(sw12, BM_sw12, 12, 0)
ST13 = Array(sw13, BM_sw13, 13, 0)
ST15 = Array(sw15, BM_sw15, 15, 0)
ST16 = Array(sw16, BM_sw16, 16, 0)
ST17 = Array(sw17, BM_sw17, 17, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST11,ST12,ST13, ST15, ST16, ST17)


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
	STArray(i)(3) = STCheckHit(Activeball,STArray(i)(0))
	
	If STArray(i)(3) <> 0 Then
		DTBallPhysics Activeball, STArray(i)(0).orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 To UBound(STArray)
		If STArray(i)(2) = switch Then 
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
		STArray(i)(3) = STAnimate(STArray(i)(0),STArray(i)(1),STArray(i)(2),STArray(i)(3))
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
		primary.uservalue = gametime
	End If
	
	animtime = gametime - primary.uservalue
	
	If animate = 1 Then
		primary.collidable = 0
		prim.transy =  - STMaxOffset
		If UsingROM Then
			vpmTimer.PulseSw switch
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

Sub DTBallPhysics(aBall, angle, mass)
	dim rangle,bangle,calc1, calc2, calc3
	rangle = (angle - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))

	calc1 = cor.BallVel(aball.id) * cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	calc2 = cor.BallVel(aball.id) * sin(bangle - rangle) * cos(rangle + 4*Atn(1)/2)
	calc3 = cor.BallVel(aball.id) * sin(bangle - rangle) * sin(rangle + 4*Atn(1)/2)

	aBall.velx = calc1 * cos(rangle) + calc2
	aBall.vely = calc1 * sin(rangle) + calc3
End Sub
'******************************************************
'		END STAND-UP TARGETS
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

Function RotPoint(x,y,angle)
    dim rx, ry
    rx = x*dCos(angle) - y*dSin(angle)
    ry = x*dSin(angle) + y*dCos(angle)
    RotPoint = Array(rx,ry)
End Function

'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Class DebugLogFile
	Private Filename
	Private TxtFileStream
	
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
		 & LZ(Day(CurrTime),	2) & " " _
		 & LZ(Hour(CurrTime),   2) & ":" _
		 & LZ(Minute(CurrTime), 2) & ":" _
		 & LZ(Second(CurrTime), 2) & ":" _
		 & LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message, code)
		Dim FormattedMsg, Timestamp
		'   Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
		Filename = cGameName + "_debug_log.txt"
		
		Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
		Timestamp = GetTimeStamp
		FormattedMsg = GetTimeStamp + " : " + label + " : " + message
		TxtFileStream.WriteLine FormattedMsg
		TxtFileStream.Close
		Debug.print label & " : " & message
	End Sub
End Class

Sub WriteToLog(label, message)
	If KeepLogs Then
		Dim LogFileObj
		Set LogFileObj = New DebugLogFile
		LogFileObj.WriteToLog label, message, 8
	End If
End Sub

Sub NewLog()
	If KeepLogs Then
		Dim LogFileObj
		Set LogFileObj = New DebugLogFile
		LogFileObj.WriteToLog "NEW LOG", " ", 2
	End If
End Sub

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************


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
	DebugShotTableKeyDownCheck keycode
	
	
	If keycode = LeftFlipperKey Then
		FlipperActivate LeftFlipper, LFPress
		'FlipperActivate LeftFlipper1, LFPress
		SolLFlipper True	'This would be called by the solenoid callbacks if using a ROM
		If gameStarted = True Then 
			DispatchPinEvent SWITCH_LEFT_FLIPPER_DOWN
		End If
	End If
	
	If keycode = RightFlipperKey Then
		FlipperActivate RightFlipper, RFPress
		SolRFlipper True	'This would be called by the solenoid callbacks if using a ROM
		UpRightFlipper.RotateToEnd
		If gameStarted = True Then 
			DispatchPinEvent SWITCH_RIGHT_FLIPPER_DOWN
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
	
	If keycode = StartGameKey Then
		SoundStartButton
		
		If gameStarted = False Then
			AddPlayer()
			StartGame()
		Else
			If canAddPlayers = True Then
				AddPlayer()
			End If		
		End If

	End If
	
	
End Sub



Sub Table1_KeyUp(ByVal keycode)
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
End Sub


'*******************************************
'  Kickers, Saucers
'*******************************************

'************************* VUKs *****************************
Dim KickerBall1

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
	dim rangle
	rangle = PI * (kangle - 90) / 180
    
	kball.z = kball.z + kzlift
	kball.velz = kvelz
	kball.velx = cos(rangle)*kvel
	kball.vely = sin(rangle)*kvel
End Sub



'***********************************************************************************************************************
' Lights State Controller - 8.0.0
'  
' A light state controller for original vpx tables.
'
' Documentation: https://github.com/mpcarr/vpx-light-controller
'
'***********************************************************************************************************************


Class LStateController

    Private m_currentFrameState, m_on, m_off, m_seqRunners, m_lights, m_seqs, m_vpxLightSyncRunning, m_vpxLightSyncClear, m_vpxLightSyncCollection, m_tableSeqColor, m_tableSeqOffset, m_tableSeqSpeed, m_tableSeqDirection, m_tableSeqFadeUp, m_tableSeqFadeDown, m_frametime, m_initFrameTime, m_pulse, m_pulseInterval, useVpxLights, m_lightmaps, m_seqOverrideRunners

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
        m_vpxLightSyncRunning = False
        m_vpxLightSyncCollection = Null
		m_initFrameTime = 0
        m_frameTime = 0
        m_pulseInterval = 26
        m_vpxLightSyncClear = False
        m_tableSeqColor = Null
        m_tableSeqFadeUp = Null
        m_tableSeqFadeDown = Null
        useVpxLights = False
        Set m_lightmaps = CreateObject("Scripting.Dictionary")
    End Sub

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
        Debug.print("Lights YAML File saved to: " & cGameName & "LightShows/lights-"&name&".yaml")
    End Sub

	Dim leds()
	Dim lightsToLeds(255)
	Sub PrintLEDs
		Dim light
        Dim lights : lights = ""
        
		Dim row,col,value
		For row = LBound(leds, 1) To UBound(leds, 1)
			For col = LBound(leds, 2) To UBound(leds, 2)
				' Access the array element and do something with it
				value = leds(row, col)
				lights = lights + cstr(value) & vbTab
			Next
			lights = lights + vbCrLf
		Next

		Dim objFileToWrite : Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(cGameName & "_LightShows/leds.txt",2,true)
	    objFileToWrite.WriteLine(lights)
	    objFileToWrite.Close
	    Set objFileToWrite = Nothing
        Debug.print("Lights File saved to: " & cGameName & "LightShows/leds.txt")

	End Sub

    Public Sub RegisterLights(mode)

        Dim idx,tmp,vpxLight,lcItem
        If mode = "Lampz" Then
            
            For idx = 0 to UBound(Lampz.obj)
                If Lampz.IsLight(idx) Then
                    Set lcItem = new LCItem
                    If IsArray(Lampz.obj(idx)) Then
                        tmp = Lampz.obj(idx)
                        Set vpxLight = tmp(0)
						
                    Else
                        Set vpxLight = Lampz.obj(idx)
                        
                    End If
                    Lampz.Modulate(idx) = 1/100
                    Lampz.FadeSpeedUp(idx) = 100/30 : Lampz.FadeSpeedDown(idx) = 100/120
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next        
        ElseIf mode = "VPX" Then
            useVpxLights = True

			Dim colCount : colCount = Round(tablewidth/20)
			Dim rowCount : rowCount = Round(tableheight/20)
			
			ReDim leds(rowCount,colCount)
				
			dim ledIdx : ledIdx = 0
            For idx = 0 to UBound(Lights)
                vpxLight = Null
                Set lcItem = new LCItem
				debug.print("TRYING TO REGISTER IDX: " & idx)
                If IsArray(Lights(idx)) Then
                    tmp = Lights(idx)
                    Set vpxLight = tmp(0)
					debug.print("TEMP LIGHT NAME for idx:" & idx & ", light: " & vpxLight.name)
                ElseIf IsObject(Lights(idx)) Then
                    Set vpxLight = Lights(idx)
                End If
                If Not IsNull(vpxLight) Then
					Debug.print("Registering Light: "& vpxLight.name)


					Dim r : r = Round(vpxLight.y/20)
					Dim c : c = Round(vpxLight.x/20)
                    If r < rowCount And c < colCount And r >= 0 And c >= 0 Then
                        If Not leds(r,c) = "" Then
                            MsgBox("Move your lights punk: " & idx)
                        End If
                        leds(r,c) = ledIdx
                        lightsToLeds(idx) = ledIdx
                        ledIdx = ledIdx + 1
                    End If
                    Dim e, lmStr: lmStr = "lmArr = Array("    
                    For Each e in GetElements()
                        If InStr(e.Name, "_" & vpxLight.Name) Or InStr(e.Name, "_" & vpxLight.Name & "_") Or InStr(e.Name, "_" & vpxLight.UserValue & "_") Then
                            Debug.Print(e.Name)
                            lmStr = lmStr & e.Name & ","
                        End If
                    Next
                    lmStr = lmStr & "Null)"
                    lmStr = Replace(lmStr, ",Null)", ")")
			        ExecuteGlobal "Dim lmArr : "&lmStr
                    m_lightmaps.Add vpxLight.Name, lmArr
                    Debug.print("Registering Light: "& vpxLight.name) 
                    lcItem.Init idx, vpxLight.BlinkInterval, Array(vpxLight.color, vpxLight.colorFull), vpxLight.name, vpxLight.x, vpxLight.y
                    m_lights.Add vpxLight.Name, lcItem
                    m_seqRunners.Add "lSeqRunner" & CStr(vpxLight.name), new LCSeqRunner
                End If
            Next  
        End If
    End Sub

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
        m_seqRunners.Add "lSeqRunner" & CStr(light.name), new LCSeqRunner
    End Sub

    Public Sub LightState(light, state)
        m_lightOff(light.name)
        If state = 1 Then
            m_lightOn(light.name)
        ElseIF state = 2 Then
            Blink(light)
        End If
    End Sub

    Public Sub LightOn(light)
        m_LightOn(light.name)
    End Sub

    Public Sub LightOnWithColor(light, color)
        m_LightOnWithColor light.name, color
    End Sub

    Public Sub FadeLightToColor(light, color, fadeSpeed)
        If m_lights.Exists(light.name) Then
            dim lightColor, steps
            steps = Round(fadeSpeed/10)
            If steps < 10 Then
                steps = 10
            End If
            lightColor = m_lights(light.name).Color
            Dim seq : Set seq = new LCSeq
            seq.Name = light.name & "Fade"
            seq.Sequence = FadeRGB(light.name, lightColor(0), color, fadeSpeed/10)
            seq.Color = Null
            seq.UpdateInterval = fadeSpeed
            seq.Repeat = False
            m_lights(light.name).Color = color
            m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
            If color = RGB(0,0,0) Then
                m_lightOff(light.name)
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

        If m_lights.Exists(light.name) Then
            m_lights(light.name).Color = color
            'Update internal blink seq for light
            If m_seqs.Exists(light.name & "Blink") Then
                m_seqs(light.name & "Blink").Color = color
            End If

        End If
    End Sub

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
        m_lightOff(light.name)
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
                m_seqs(light.name & "Blink").Sequence = m_buildBlinkSeq(light)
            End If
        End If
    End Sub


    Public Sub AddShot(name, light, color)
        If m_lights.Exists(light.name) Then
            If m_seqs.Exists(name & light.name) Then
                m_seqs(name & light.name).Color = color
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(name & light.name)
            Else
                Dim stateOn : stateOn = light.name&"|100"
                Dim stateOff : stateOff = light.name&"|0"
                Dim seq : Set seq = new LCSeq
                seq.Name = name
                seq.Sequence = Array(stateOn, stateOff,stateOn, stateOff)
                seq.Color = color
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
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
                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem m_seqs(light.name & "Blink")
            Else
                Dim seq : Set seq = new LCSeq
                seq.Name = light.name & "Blink"
                seq.Sequence = m_buildBlinkSeq(light)
                seq.Color = Null
                seq.UpdateInterval = light.BlinkInterval
                seq.Repeat = True

                m_seqRunners("lSeqRunner"&CStr(light.name)).AddItem seq
                m_seqs.Add light.name & "Blink", seq
            End If
            If m_on.Exists(light.name) Then
                m_on.Remove light.name
            End If
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

    Public Sub CreateSeqRunner(name)
        If m_seqRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqRunners.Add name, seqRunner
    End Sub

    Private Sub CreateOverrideSeqRunner(name)
        If m_seqOverrideRunners.Exists(name) Then
            Exit Sub
        End If
        Dim seqRunner : Set seqRunner = new LCSeqRunner
        seqRunner.Name = name
        m_seqOverrideRunners.Add name, seqRunner
    End Sub

    Public Sub AddLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        m_seqRunners(lcSeqRunner).AddItem lcSeq
    End Sub

    Public Sub RemoveLightSeq(lcSeqRunner, lcSeq)
        If Not m_seqRunners.Exists(lcSeqRunner) Then
            Exit Sub
        End If

        Dim light
        For Each light in lcSeq.LightsInSeq
            If(m_lights.Exists(light)) Then
                    AssignStateForFrame light, (new FrameState)(0, Null, m_lights(light).Idx)
            End If
        Next

        m_seqRunners(lcSeqRunner).RemoveItem lcSeq
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
        m_seqOverrideRunners(name).AddItem lcSeq
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
            If m_lightmaps.Exists(light) Then
                For Each lm in m_lightmaps(light)
                    dim color : color = m_lights(light).Color
                    If not IsNull(lm) Then
						lm.Color = color(0)
					End If
                Next
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
	Public Sub SetVpxSyncLightGradientColor(gradient, direction, speed)
		m_tableSeqColor = gradient
		m_tableSeqDirection = direction
		m_tableSeqSpeed = speed
	End Sub

    Public Sub SetTableSequenceFade(fadeUp, fadeDown)
		m_tableSeqFadeUp = fadeUp
        m_tableSeqFadeDown = fadeDown
	End Sub

    Public Sub UseToolkitColoredLightMaps()
        If useVpxLights = True Then
            Exit Sub
        End If

        Dim sUpdateLightMap
        sUpdateLightMap = "Sub UpdateLightMap(idx, lightmap, intensity, ByVal aLvl)" + vbCrLf    
        sUpdateLightMap = sUpdateLightMap + "   if Lampz.UseFunc then aLvl = Lampz.FilterOut(aLvl)	'Callbacks don't get this filter automatically" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   lightmap.Opacity = aLvl * intensity" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   If IsArray(Lampz.obj(idx) ) Then" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.Color = Lampz.obj(idx)(0).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   Else" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "       lightmap.color = Lampz.obj(idx).color" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "   End If" + vbCrLf
        sUpdateLightMap = sUpdateLightMap + "End Sub" + vbCrLf

        ExecuteGlobal sUpdateLightMap

        Dim x
        For x=0 to Ubound(Lampz.cCallback)
            Lampz.cCallback(x) = Replace(Lampz.cCallback(x), "UpdateLightMap ", "UpdateLightMap " & x & ",")
            Lampz.Callback(x) = "" 'Force Callback Sub to be build
        Next
    End Sub

    Private Function m_buildBlinkSeq(light)
        Dim i, buff : buff = Array()
        ReDim buff(Len(light.BlinkPattern)-1)
        For i = 0 To Len(light.BlinkPattern)-1
            
            If Mid(light.BlinkPattern, i+1, 1) = 1 Then
                buff(i) = light.name & "|100"
            Else
                buff(i) = light.name & "|0"
            End If
        Next
        m_buildBlinkSeq=buff
    End Function

    Private Function GetTmpLight(idx)
        If useVpxLights = True Then
          If IsArray(Lights(idx) ) Then	'if array
                Set GetTmpLight = Lights(idx)(0)
            Else
                Set GetTmpLight = Lights(idx)
            End If
        Else
            If IsArray(Lampz.obj(idx) ) Then	'if array
                Set GetTmpLight = Lampz.obj(idx)(0)
            Else
                Set GetTmpLight = Lampz.obj(idx)
            End If
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

    Public Sub Update()

		m_frameTime = gametime - m_initFrameTime : m_initFrameTime = gametime
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
                    AssignStateForFrame lightKey, (new FrameState)(lcItem.level, m_on(lightKey).Color, m_on(lightKey).Idx)
                Next
            End If

            If HasKeys(m_pulse) Then   
                For Each lightKey in m_pulse.Keys()
					Dim pulseColor : pulseColor = m_pulse(lightKey).Color
					If IsNull(pulseColor) Then
						AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).Light.Color, m_pulse(lightKey).light.Idx)
					Else
						AssignStateForFrame lightKey, (new FrameState)(m_pulse(lightKey).PulseAt(m_pulse(lightKey).idx), m_pulse(lightKey).Color, m_pulse(lightKey).light.Idx)
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

                            If IsNull(m_tableSeqColor) Then
                                color = syncLight.Color
                            Else
                                If Not IsArray(m_tableSeqColor) Then
                                    color = Array(m_tableSeqColor, Null)
                                Else
									If Not IsNull(m_tableSeqSpeed) And Not m_tableSeqSpeed = 0 Then
										'dim step : step = m_tableSeqSpeed(m_tableSeqOffset)
										
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
							

                            'TODO - Fix VPX Fade
                            If Not useVpxLights = True Then
                                If Not IsNull(m_tableSeqFadeUp) Then
                                    Lampz.FadeSpeedUp(syncLight.Idx) = m_tableSeqFadeUp
                                End If
                                If Not IsNull(m_tableSeqFadeDown) Then
                                    Lampz.FadeSpeedDown(syncLight.Idx) = m_tableSeqFadeDown
                                End If
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
                            'TODO - Only do fade speed for lampz
                            If Not useVpxLights = True Then
                                Lampz.FadeSpeedUp(syncClearLight.Idx) = 100/30
                                Lampz.FadeSpeedDown(syncClearLight.Idx) = 100/120
                            End If
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

                If useVpxLights = False Then
                    If bUpdate Then
                        'Update lamp color
                        If IsArray(Lampz.obj(idx)) Then
                            for each x in Lampz.obj(idx)
                                If Not IsNull(c) Then
                                    x.color = c
                                End If
                                If Not IsNull(cf) Then
                                    x.colorFull = cf
                                End If
                            Next
                        Else
                            If Not IsNull(c) Then
                                Lampz.obj(idx).color = c
                            End If
                            If Not IsNull(cf) Then
                                Lampz.obj(idx).colorFull = cf
                            End If
                        End If
                        If Lampz.UseCallBack(idx) then Proc Lampz.name & idx,Lampz.Lvl(idx)*Lampz.Modulate(idx)	'Force Callbacks Proc
                    End If
                    Lampz.state(idx) = CInt(m_currentFrameState(frameStateKey).level) 'Lampz will handle redundant updates
                Else
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
                                        lm.Color = c
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

	Public Function GetGradientColors(startColor, endColor)
	  Dim colors()
	  ReDim colors(255)
	  
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
	  For i = 0 To 255
		Dim red, green, blue
		red = startRed + (redDiff * (i / 255))
		green = startGreen + (greenDiff * (i / 255))
		blue = startBlue + (blueDiff * (i / 255))
		colors(i) = RGB(red,green,blue)'IntToHex(red, 2) & IntToHex(green, 2) & IntToHex(blue, 2)
	  Next
	  
	  GetGradientColors = colors
	End Function


	Function GetGradientColorsWithStops(startColor, endColor, stopPositions, stopColors)

	  Dim colors(255)

	  Dim fStop : fStop = GetGradientColors(startColor, stopColors(0))
	  Dim i, istep
	  For i = 0 to stopPositions(0)
		colors(i) = fStop(i)
	  Next
	  For i = 1 to Ubound(stopColors)
		Dim stopStep : stopStep = GetGradientColors(stopColors(i-1), stopColors(i))
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
	  Dim eStop : eStop = GetGradientColors(stopColors(UBound(stopColors)), endColor)
	  'MsgBox(UBound(stopPositions))
	  iStep = 0
	  For i = (255-stopPositions(UBound(stopPositions))) to 254
		colors(i) = eStop(iStep)
		iStep = iStep + 1
	  Next

	  GetGradientColorsWithStops = colors
	End Function

    Private Function HasKeys(o)
        Dim Success
        Success = False

        On Error Resume Next
            o.Keys()
            Success = (Err.Number = 0)
        On Error Goto 0
        HasKeys = Success
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
                        
						color = lcSeq.Color

                        If IsNull(color) Then
							color = ls.Color
                        End If
						
                        If Ubound(lsName) = 2 Then
							If lsName(2) = "" Then
                                AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                            Else
                                AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                            End If
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        End If
                        lcSeq.SetLastLightState name, m_currentFrameState(name) 
                    End If
                Next       
            Else
                lsName = Split(seq(lcSeq.CurrentIdx),"|")
                name = lsName(0)
                If m_lights.Exists(name) Then
                    Set ls = m_lights(name)
                    
					color = lcSeq.Color
                    If IsNull(color) Then
                        color = ls.Color
                    End If
                    If Ubound(lsName) = 2 Then
                        If lsName(2) = "" Then
                            AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                        Else
                            AssignStateForFrame name, (new FrameState)(lsName(1), Array( RGB( HexToInt(Left(lsName(2), 2)), HexToInt(Mid(lsName(2), 3, 2)), HexToInt(Right(lsName(2), 2)) ), RGB(0,0,0)), ls.Idx)
                        End If
                    Else
                        AssignStateForFrame name, (new FrameState)(lsName(1), color, ls.Idx)
                    End If
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
		'debug.Print(Join(Pulses))
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
	
	Private m_Idx, m_State, m_blinkSeq, m_color, m_name, m_level, m_x, m_y

        Public Property Get Idx()
            Idx=m_Idx
        End Property

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
            m_name = name
            m_level = 100
            m_x = x
            m_y = y
	    End Sub

End Class

Class LCSeq
	
	Private m_currentIdx, m_sequence, m_name, m_image, m_color, m_updateInterval, m_Frames, m_repeat, m_lightsInSeq, m_lastLightStates

    Public Property Get CurrentIdx()
        CurrentIdx=m_currentIdx
    End Property

    Public Property Let CurrentIdx(input)
		m_lastLightStates.RemoveAll()
        m_currentIdx = input
    End Property

    Public Property Get LightsInSeq()
        LightsInSeq=m_lightsInSeq.Keys()
    End Property

    Public Property Get Sequence()
        Sequence=m_sequence
    End Property
    
	Public Property Let Sequence(input)
		m_sequence = input
        dim item, light, lightItem
        for each item in input
            If IsArray(item) Then
                for each light in item
                    lightItem = Split(light,"|")
                    If Not m_lightsInSeq.Exists(lightItem(0)) Then
                        m_lightsInSeq.Add lightItem(0), True
                    End If    
                next
            Else
                lightItem = Split(item,"|")
                If Not m_lightsInSeq.Exists(lightItem(0)) Then
                    m_lightsInSeq.Add lightItem(0), True
                End If
            End If
        next
	End Property

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

    Private Sub Class_Initialize()
        m_currentIdx = 0
        m_color = Array(Null, Null)
        m_updateInterval = 180
        m_repeat = False
        m_Frames = 180
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
	
	Private m_name, m_items,m_currentItemIdx

    Public Property Get Name()
        Name=m_name
    End Property
    
	Public Property Let Name(input)
		m_name = input
	End Property

    Public Property Get Items()
		Set Items = m_items
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
    End Sub

    Public Sub AddItem(item)
        If Not IsNull(item) Then
            If Not m_items.Exists(item.Name) Then
                m_items.Add item.Name, item
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

    Public Sub RemoveItem(item)
        If Not IsNull(item) Then
            If m_items.Exists(item.Name) Then
                    item.ResetInterval
                    item.CurrentIdx = 0
                    m_items.Remove item.Name
            End If
        End If
    End Sub

    Public Sub NextItem()
        Dim items: items = m_items.Items
        If items(m_currentItemIdx).Repeat = False Then
            RemoveItem(items(m_currentItemIdx))
        Else
            m_currentItemIdx = m_currentItemIdx + 1
        End If
        
        If m_currentItemIdx > UBound(m_items.Items) Then   
            m_currentItemIdx = 0
        End If
    End Sub

    Public Function HasSeq(name)
        If m_items.Exists(name) Then
            HasSeq = True
        Else
            HasSeq = False
        End If
    End Function

End Class

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
	UpdateDropTargets
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


'TROUGH 
Sub swTrough1_Hit
	UpdateTrough
End Sub
Sub swTrough1_UnHit
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

'************************* VUKs *****************************
Dim KickerBall39, KickerBall40

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
	dim rangle
	rangle = PI * (kangle - 90) / 180
    
	kball.z = kball.z + kzlift
	kball.velz = kvelz
	kball.velx = cos(rangle)*kvel
	kball.vely = sin(rangle)*kvel
End Sub
'***********************************************************************************************************************
'*****     GAME LOGIC START                                                 	                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************


Dim canAddPlayers : canAddPlayers = True
Dim currentPlayer : currentPlayer = Null
Dim autoPlunge : autoPlunge = False
Dim ballSaver : ballSaver = False
Dim gameStarted : gameStarted = False
Dim pinEvents : Set pinEvents = CreateObject("Scripting.Dictionary")
Dim playerEvents : Set playerEvents = CreateObject("Scripting.Dictionary")
Dim playerState : Set playerState = CreateObject("Scripting.Dictionary")



'******************************************************
'*****   Pin Events                                ****
'******************************************************

Const START_GAME = "Start Game"
Const NEXT_PLAYER = "Next Player"
Const BALL_DRAIN = "Ball Drain"
Const BALL_SAVE = "Ball Save"
Const ADD_BALL = "Add Ball"
Const GAME_OVER = "Game Over"

Const SWITCH_LEFT_FLIPPER_DOWN = "Switches Left Flipper Down"
Const SWITCH_RIGHT_FLIPPER_DOWN = "Switches Right Flipper Down"


'***********************************************************************************************************************
'*****     LIGHTS LOGIC START                                                 	                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************

Dim coordsX : coordsX  = Array(  0, 18, 213, 237, 115, 43, 67, 85, 109, 128, 152, 170, 24, 24, 30, 43, 49, 49, 61, 67, 67, 109, 128, 115, 200, 213, 200, 36, 0, 0, 36, 30, 24, 73, 85, 97, 55, 73, 85, 103, 115, 121, 128, 79, 103, 121, 109, 140, 146, 109, 109, 146, 152, 134, 73, 73, 67, 170, 225, 219, 219, 213, 176, 188, 182, 213, 219, 225, 0, 237, 134, 255 )
Dim coordsY : coordsY  = Array(  186, 183, 183, 186, 255, 175, 183, 180, 180, 180, 180, 183, 155, 147, 141, 158, 150, 141, 161, 152, 144, 158, 158, 147, 155, 166, 172, 130, 83, 72, 89, 75, 55, 97, 97, 91, 108, 108, 105, 103, 100, 94, 89, 116, 114, 108, 80, 78, 83, 72, 61, 67, 55, 55, 39, 28, 19, 116, 72, 78, 86, 91, 103, 116, 125, 119, 108, 100, 130, 122, 42, 0 )
Dim angles : angles =  Array( 215, 212, 169, 166, 191, 209, 202, 198, 192, 187, 181, 178, 219, 223, 225, 214, 215, 220, 208, 209, 212, 192, 186, 190, 163, 165, 169, 231, 17, 24, 19, 29, 37, 20, 27, 46, 253, 252, 4, 32, 72, 77, 79, 232, 213, 141, 62, 80, 87, 62, 63, 78, 77, 71, 53, 55, 54, 142, 103, 106, 111, 114, 122, 139, 148, 138, 129, 122, 238, 138, 70, 86 )
Dim radii : radii  = Array( 162, 150, 153, 167, 255, 129, 136, 128, 126, 127, 130, 139, 108, 98, 88, 104, 89, 78, 101, 86, 74, 88, 89, 69, 108, 129, 132, 72, 97, 106, 67, 85, 112, 35, 27, 29, 45, 31, 22, 10, 12, 23, 33, 31, 14, 8, 45, 54, 48, 60, 79, 74, 93, 90, 121, 139, 154, 48, 106, 97, 91, 83, 51, 62, 63, 82, 83, 89, 97, 101, 113, 215 )




'****************************
' Stat Of Game
' Event Listeners:  
AddPinEventListener START_GAME,    "GIStartOfGame"
'
'*****************************
Sub GIStartOfGame()
    Dim x
    For Each x in GI
        lightCtrl.LightOn x
    Next
End Sub

'****************************
' End Of Game
' Event Listeners:  
AddPinEventListener GAME_OVER,    "GIEndOfGame"
'
'*****************************
Sub GIEndOfGame()
    Dim x
    For Each x in GI
        lightCtrl.LightOff x
    Next
End Sub


Sub PlayVPXSeq()
	LightSeq.Play SeqCircleOutOn, 20, 1
	lightCtrl.SyncWithVpxLights LightSeq
	'lightCtrl.SetVpxSyncLightGradientColor MakeGradident, coordsX, 80
End Sub

Sub LightSeq_PlayDone()
    lightCtrl.StopSyncWithVpxLights()
End Sub

Function MakeGradident()
    ' Define the start and end colors
    Dim startColor
    Dim endColor
    startColor = "993400"  ' Red
    endColor = "FF0000"    ' Green

    ' Define the stop positions and colors
    Dim stopPositions(3)
    Dim stopColors(3)
    stopPositions(0) = 0    ' Start at 0%
    stopColors(0) = "993400" ' Red
    stopPositions(1) = 25   ' Yellow at 50%
    stopColors(1) = "FFA500" ' Yellow
    stopPositions(2) = 50   ' Orange at 75%
    stopColors(2) = "FF0000" ' Orange
	stopPositions(3) = 75   ' Orange at 75%
    stopColors(3) = "0080ff" ' Orange

    ' Call the GetGradientColorsWithStops function to generate the gradient colors
    MakeGradident = lightCtrl.GetGradientColorsWithStops(startColor, endColor, stopPositions, stopColors)

End Function


'******************************************************
'*****   GAME MODE LOGIC START                     ****
'******************************************************

Sub StartGame()
    gameStarted = True
    SetPlayerState BALL_SAVE_ENABLED, True
    DispatchPinEvent START_GAME
End Sub

'****************************
' End Of Game
' Event Listeners:  
    AddPinEventListener GAME_OVER,    "EndOfGame"
'
'*****************************
Sub EndOfGame()
    
End Sub


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

'******************************************************
'*****   End of Ball                               ****
'******************************************************

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener BALL_DRAIN, "EndOfBall"
'
'*****************************
Sub EndOfBall()
    debug.print("Ball Saver" & ballSaver)
    If ballSaver = True Then
        DispatchPinEvent BALL_SAVE
    ElseIf BIP - GetPlayerState(BALLS_LOCKED) = 0 Then
        SetPlayerState CURRENT_BALL, GetPlayerState(CURRENT_BALL) + 1

        Select Case currentPlayer
            Case "PLAYER 1":
                If UBound(playerState.Keys()) > 0 Then
                    currentPlayer = "PLAYER 2"
                End If
            Case "PLAYER 2":
                If UBound(playerState.Keys()) > 1 Then
                    currentPlayer = "PLAYER 3"
                Else
                    currentPlayer = "PLAYER 1"
                End If
            Case "PLAYER 3":
                If UBound(playerState.Keys()) > 2 Then
                    currentPlayer = "PLAYER 4"
                Else
                    currentPlayer = "PLAYER 1"
                End If
            Case "PLAYER 4":
                currentPlayer = "PLAYER 1"
        End Select

        If GetPlayerState(CURRENT_BALL) > BALLS_PER_GAME Then
            DispatchPinEvent GAME_OVER
            gameStarted = False
            currentPlayer = Null
            playerState.RemoveAll()
        Else
            SetPlayerState BALL_SAVE_ENABLED, True 
            DispatchPinEvent NEXT_PLAYER
        End If
    End If
End Sub


'******************************************************
'*****   Add Score                                 ****
'******************************************************

Sub AddScore(v)
    SetPlayerState SCORE, GetPlayerState(SCORE) + v
End Sub


'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener START_GAME,    "ReleaseBall"
AddPinEventListener NEXT_PLAYER,   "ReleaseBall"
'
'*****************************
Sub ReleaseBall()
    swTrough1.kick 90, 10
    BIP = BIP + 1
    RandomSoundBallRelease swTrough1
    PuPlayer.LabelSet   pBackglass, "lblBall",      "Ball " & GetPlayerState(CURRENT_BALL),                        1,  "{}"
End Sub


Sub InitPupLabels
    if PUPStatus=false then Exit Sub
    
    PuPlayer.LabelInit pBackglass
    Dim pupFont:pupFont=""

    Dim fontColor : fontColor = RGB(255,255,255)
    'syntax - PuPlayer.LabelNew <screen# or pDMD>,<Labelname>,<fontName>,<size%>,<colour>,<rotation>,<xAlign>,<yAlign>,<xpos>,<ypos>,<PageNum>,<visible>
    '				    Scrn        LblName                 Fnt         Size	        Color	 		    R   Ax    Ay    X       Y           pagenum     Visible 
    
    PuPlayer.LabelNew   pBackglass, "lblTitle",             pupFont,    8,           fontColor,  0,  1,    1,    0,      0,          1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer1",           pupFont,    6,           fontColor,  0,  0,    0,    10,     80,         1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer2",           pupFont,    6,           fontColor,  0,  0,    0,    30,     80,         1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer3",           pupFont,    6,           fontColor,  0,  0,    0,    50,     80,         1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer4",           pupFont,    6,           fontColor,  0,  0,    0,    70,     80,         1,          1

    PuPlayer.LabelNew   pBackglass, "lblPlayer1Score",           pupFont,    6,           fontColor,  0,  0,    0,    10,     90,         1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer2Score",           pupFont,    6,           fontColor,  0,  0,    0,    30,     90,         1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer3Score",           pupFont,    6,           fontColor,  0,  0,    0,    50,     90,         1,          1
    PuPlayer.LabelNew   pBackglass, "lblPlayer4Score",           pupFont,    6,           fontColor,  0,  0,    0,    70,     90,         1,          1

    PuPlayer.LabelNew   pBackglass, "lblBall",              pupFont,    6,           fontColor,  0,  0,    0,    63,     33,         1,          1
    PuPlayer.LabelSet   pBackglass, "lblTitle",     "tmntpro",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer1",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer2",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer3",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer4",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblBall",      "",                        1,  "{}"

    PuPlayer.LabelSet   pBackglass, "lblPlayer1Score",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer2Score",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer3Score",    "",                        1,  "{}"
    PuPlayer.LabelSet   pBackglass, "lblPlayer4Score",    "",                        1,  "{}"
        
    
    
End Sub


'***********************************************************************************************************************
'*****  PIN UP UPDATE SCORES                                   	                                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************
Dim ScoreSize(4)

Sub pupDMDupdate_Timer()
	If gameStarted Then
		pUpdateScores
	End If
End Sub

'************************ called during gameplay to update Scores ***************************
Sub pUpdateScores  'call this ONLY on timer 300ms is good enough
    
	if PUPStatus=false then Exit Sub

	Dim StatusStr
    Dim Size(4)
    Dim ScoreTag(4)
	
	StatusStr = ""

	Dim lenScore:lenScore = Len(GetPlayerState(SCORE) & "")
	Dim scoreScale:scoreScale=0.6
	If lenScore>6 Then
		scoreScale=scoreScale - ((lenScore-6)/20)
	End If
	
	ScoreSize(0) = 6*scoreScale
	ScoreTag(0)="{'mt':2,'size':"& ScoreSize(0) &"}"
	
	Select Case currentPlayer
		Case "PLAYER 1":
			puPlayer.LabelSet pBackglass,"lblPlayer1Score", FormatScore(GetPlayerState(SCORE))	 ,1,ScoreTag(0)
		Case "PLAYER 2":
			puPlayer.LabelSet pBackglass,"lblPlayer2Score", FormatScore(GetPlayerState(SCORE))	 ,1,ScoreTag(0)
		Case "PLAYER 3":
			puPlayer.LabelSet pBackglass,"lblPlayer3Score", FormatScore(GetPlayerState(SCORE))	 ,1,ScoreTag(0)
		Case "PLAYER 4":
			puPlayer.LabelSet pBackglass,"lblPlayer4Score", FormatScore(GetPlayerState(SCORE))	 ,1,ScoreTag(0)
		End Select

end Sub

Function FormatScore(ByVal Num) 'it returns a string with commas
    dim i
    dim NumString

    NumString = CStr(abs(Num))

	If NumString = "0" Then
		NumString = "00"
	Else
		For i = Len(NumString)-3 to 1 step -3
			if IsNumeric(mid(NumString, i, 1))then
				NumString = left(NumString, i) & "," & right(NumString, Len(NumString)-i)
			end if
		Next
	End If
    FormatScore = NumString
End function

'***********************************************************************************************************************
'*****  PIN UP                                                 	                                                    ****
'*****                                                                                                              ****
'***********************************************************************************************************************

' COPY EVERYTHING BELOW TO THE TOP OF YOUR TABLE SCRIPT UNDER OPTION EXPLICIT

'****** PuP Variables ******

Dim cPuPPack: Dim PuPlayer: Dim PUPStatus: PUPStatus=false ' dont edit this line!!!
Dim pBackglass:pBackglass=2

'*************************** PuP Settings for this table ********************************

cPuPPack = "tmntpro"    ' name of the PuP-Pack / PuPVideos folder for this table

'//////////////////// PINUP PLAYER: STARTUP & CONTROL SECTION //////////////////////////

' This is used for the startup and control of Pinup Player

Sub PuPStart(cPuPPack)
    If PUPStatus=true then Exit Sub
    
    Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")
    If PuPlayer is Nothing Then
        PUPStatus=false
        MsgBox("Could not start Pup")
    Else
        PuPlayer.B2SInit "",cPuPPack 'start the Pup-Pack
        PUPStatus=true
        InitPupLabels
    End If

End Sub

Sub pupevent(EventNum)
    if PUPStatus=false then Exit Sub
    PuPlayer.B2SData "E"&EventNum,1  'send event to Pup-Pack
End Sub

' ******* How to use PUPEvent to trigger / control a PuP-Pack *******

' Usage: pupevent(EventNum)

' EventNum = PuP Exxx trigger from the PuP-Pack

' Example: pupevent 102

' This will trigger E102 from the table's PuP-Pack

' DO NOT use any Exxx triggers already used for DOF (if used) to avoid any possible confusion

'************ PuP-Pack Startup **************

PuPStart(cPuPPack) 'Check for PuP - If found, then start Pinup Player / PuP-Pack

'***********************************************************************************************************************



'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub AddPlayer()
    Select Case UBound(playerState.Keys())
        Case -1:
            playerState.Add "PLAYER 1", InitNewPlayer()
            currentPlayer = "PLAYER 1"
            PuPlayer.LabelSet   pBackglass, "lblPlayer1",             "Player 1",                        1,  "{}"
            PuPlayer.LabelSet   pBackglass, "lblPlayer1Score",        "00",                        1,  "{}"
        Case 0:     
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 2", InitNewPlayer()
                PuPlayer.LabelSet   pBackglass, "lblPlayer2",         "Player 2",                        1,  "{}"
                PuPlayer.LabelSet   pBackglass, "lblPlayer2Score",    "00",                        1,  "{}"
            End If
        Case 1:
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 3", InitNewPlayer()
                PuPlayer.LabelSet   pBackglass, "lblPlayer3",         "Player 3",                        1,  "{}"
                PuPlayer.LabelSet   pBackglass, "lblPlayer3Score",    "00",                        1,  "{}"
            End If     
        Case 2:   
            If GetPlayerState(CURRENT_BALL) = 1 Then
                playerState.Add "PLAYER 4", InitNewPlayer()
                PuPlayer.LabelSet   pBackglass, "lblPlayer4",         "Player 4",                        1,  "{}"
                PuPlayer.LabelSet   pBackglass, "lblPlayer4Score",    "00",                        1,  "{}"
            End If  
            canAddPlayers = False
    End Select
End Sub

Function InitNewPlayer()

    Dim state: Set state=CreateObject("Scripting.Dictionary")

    state.Add SCORE, 0
    state.Add PLAYER_NAME, ""
    state.Add CURRENT_BALL, 1

    state.Add LANE_1,   0
    state.Add LANE_2,   0
    state.Add LANE_3,   0
    state.Add LANE_4,   0

    state.Add BALLS_LOCKED, 0

    state.Add BALL_SAVE_ENABLED, False

    state.Add TURTLE, ""
    state.Add PIZZA_INGREDIENT_1, ""
    state.Add PIZZA_INGREDIENT_2, ""
    state.Add PIZZA_INGREDIENT_3, ""
    state.Add PIZZA_INGREDIENT_4, ""
    state.Add PIZZA_INGREDIENT_5, ""
    state.Add PIZZA_INGREDIENT_6, ""
    state.Add PIZZA_INGREDIENT_7, ""
    state.Add PIZZA_INGREDIENT_8, ""
    state.Add CURRENT_MODE, 0
    state.Add MODE_SELECT_TURTLE, False
    
    Set InitNewPlayer = state

End Function


'****************************
' Setup Player
' Event Listeners:  
    AddPinEventListener START_GAME,    "SetupPlayer"
    AddPinEventListener NEXT_PLAYER,   "SetupPlayer"
'
'*****************************
Sub SetupPlayer()
    EmitAllPlayerEvents()
End Sub


'******************************************************
'*****  Game State                                 ****
'******************************************************

' Balls Per Game
Const BALLS_PER_GAME = 3

'Base Points
Const POINTS_BASE = 750

Const BALL_SAVER_GRACE = 3000


'***********************************************************************************
'***** Player State                                                     	    ****
'***********************************************************************************

'Score 
Const SCORE = "Player Score"
Const PLAYER_NAME = "Player Name"
'Ball
Const CURRENT_BALL = "Current Ball"
'Lanes
Const LANE_1 = "Lane 1"
Const LANE_2 = "Lane 2"
Const LANE_3 = "Lane 3"
Const LANE_4 = "Lane 4"
'Ball Save
Const BALL_SAVE_ENABLED = "Ball Save Enabled"
'Locked Balls
Const BALLS_LOCKED = "Balls Locked"

'Game
Const TURTLE = "Turtle"
Const PIZZA_INGREDIENT_1 = "Pizza Ingredient 1"
Const PIZZA_INGREDIENT_2 = "Pizza Ingredient 2"
Const PIZZA_INGREDIENT_3 = "Pizza Ingredient 3"
Const PIZZA_INGREDIENT_4 = "Pizza Ingredient 4"
Const PIZZA_INGREDIENT_5 = "Pizza Ingredient 5"
Const PIZZA_INGREDIENT_6 = "Pizza Ingredient 6"
Const PIZZA_INGREDIENT_7 = "Pizza Ingredient 7"
Const PIZZA_INGREDIENT_8 = "Pizza Ingredient 8"
Const CURRENT_MODE = "Current Mode"
Const MODE_SELECT_TURTLE = "Mode Select Turtle"


Function GetPlayerState(key)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If playerState(currentPlayer).Exists(key)  Then
        GetPlayerState = playerState(currentPlayer)(key)
    Else
        GetPlayerState = Null
    End If
End Function

Function SetPlayerState(key, value)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If playerState(currentPlayer).Exists(key)  Then
        playerState(currentPlayer)(key) = value
    Else
        playerState(currentPlayer).Add key, value
    End If
    gameDebugger.SendPlayerState key, value
    If playerEvents.Exists(key) Then
        Dim x
        For Each x in playerEvents(key).Keys()
            If playerEvents(key)(x) = True Then
                ExecuteGlobal x
            End If
        Next
    End If
    
    SetPlayerState = Null
End Function

Sub AddStateListener(e, v)
    If Not playerEvents.Exists(e) Then
        playerEvents.Add e, CreateObject("Scripting.Dictionary")
    End If
    playerEvents(e).Add v, True
End Sub

Sub AddPinEventListener(e, v)
    If Not pinEvents.Exists(e) Then
        pinEvents.Add e, CreateObject("Scripting.Dictionary")
    End If
    pinEvents(e).Add v, True
End Sub

Sub EmitAllPlayerEvents()
    Dim key
    For Each key in playerState(currentPlayer).Keys()
        gameDebugger.SendPlayerState key, playerState(currentPlayer)(key)
        If playerEvents.Exists(key) Then
            Dim x
            For Each x in playerEvents(key).Keys()
                If playerEvents(key)(x) = True Then
                    ExecuteGlobal x
                End If
            Next
        End If
    Next
End Sub

Sub DispatchPinEvent(e)
    If Not pinEvents.Exists(e) Then
        Exit Sub
    End If
    Dim x
    gameDebugger.SendPinEvent e
    For Each x in pinEvents(e).Keys()
        If pinEvents(e)(x) = True Then
            ExecuteGlobal x
        End If
    Next
End Sub


'***********************************************************************************
'***** Switches                                                         	    ****
'***********************************************************************************

Sub sw11_Hit()
    STHit 11
End Sub

Sub sw12_Hit()
    STHit 12
End Sub

Sub sw13_Hit()
    STHit 13
End Sub

Sub sw15_Hit()
    STHit 15
End Sub

Sub sw16_Hit()
    STHit 16
End Sub

Sub sw17_Hit()
    STHit 17
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
    DTHit 45
End Sub

Sub sw99_Hit()
    DTRaise 1
    lightCtrl.pulse l01, 3
End Sub

'******************************************************
'*****  Plunger Lane                               ****
'******************************************************

Sub BIPL_Hit()
    BIPL = True
    If autoPlunge = True Then
        AutoPlungerDelay.Interval = 300
	    AutoPlungerDelay.Enabled = True
    End If
End Sub

Sub BIPL_Top_Hit()
    BIPL = False
    autoPlunge = False
    If GetPlayerState(BALL_SAVE_ENABLED) = True Then
        EnableBallSaver 10
        SetPlayerState BALL_SAVE_ENABLED, False
    End If
End Sub

Sub AutoPlungerDelay_Timer
	plungerIM.Strength = 45
	plungerIM.AutoFire
	AutoPlungerDelay.Enabled = False
End Sub


'****************************
' Auto Plunge Ball
' Event Listeners:  
    AddPinEventListener BALL_SAVE,  "AutoPlungeBall"
    AddPinEventListener ADD_BALL,   "AutoPlungeBall"
'
'*****************************
Sub AutoPlungeBall()
    If BIPL = False And swTrough1.BallCntOver = 1 Then
        ReleaseBall()
        autoPlunge = True
    Else
        ballsInQ = ballsInQ + 1
        BallReleaseTimer.Enabled = True
    End If
End Sub

Dim ballsInQ : ballsInQ = 0
Sub BallReleaseTimer_Timer()
    If BIPL = False And ballsInQ > 0 AND swTrough1.BallCntOver = 1 Then
        ReleaseBall()
        autoPlunge = True
        ballsInQ = ballsInQ - 1
        If ballsInQ = 0 Then
            BallReleaseTimer.Enabled = False
        End If
    End If
End Sub


'******************************************************
'*****  Drain                                      ****
'******************************************************

Sub Drain_Hit 
    BIP = BIP - 1
	Drain.kick 57, 20
    DispatchPinEvent BALL_DRAIN
End Sub

Sub Drain_UnHit : UpdateTrough : End Sub
