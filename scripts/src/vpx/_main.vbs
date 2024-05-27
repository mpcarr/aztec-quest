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
Dim debugLog : Set debugLog = (new DebugLogFile)()
Dim debugEnabled : debugEnabled = True
'*******************************************
'  Constants and Global Variables
'*******************************************

Const UsingROM = False		'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.

Const BallSize = 50			'Ball diameter in VPX units; must be 50
Const BallMass = 1			'Ball mass must be 1
Const tnob = 7				'Total number of balls the table can hold
Const lob = 2				'Locked balls
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

Dim aztecqball1, aztecqball2, aztecqball3, aztecqball4, aztecqball5, gBOT, tmag, NewtonBall, CaptiveBall

Sub Table1_Init
	Dim i
	
	vpmMapLights alights
	lightCtrl.RegisterLights "VPX"
	

	'Ball initializations need for physical trough
	Set aztecqball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set aztecqball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set aztecqball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set aztecqball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	Set aztecqball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass)
	
	'*** Use gBOT in the script wherever BOT is normally used. Then there is no need for GetBalls calls ***
	gBOT = Array( aztecqball1, aztecqball2, aztecqball3, aztecqball4, aztecqball5)
	
	Dim xx
	
	' Add balls to shadow dictionary
	For Each xx In gBOT
		bsDict.Add xx.ID, bsNone
	Next
	
	' Make drop target shadows visible
	For Each xx In ShadowDT
		xx.visible = True
	Next

	'Set plungerIM = New cvpmImpulseP
	'With plungerIM
	'	.InitImpulseP sw_plunger, IMPowerSetting, IMTime
	'	.Random 1.5
	'	.InitExitSnd SoundFX("fx_kicker", DOFContactors), SoundFX("fx_solenoid", DOFContactors)
	'	.CreateEvents "plungerIM"
	'End With
	'PlayVPXSeq

	LeftSlingShot_Timer
	RightSlingShot_Timer
	lightCtrl.SyncLightMapColors
	
	DTDrop 1
	DTDrop 2
	PuPInit
	

	If AllLightsOnMode Then
		SetLightsOn
	End If

	If useBCP = True Then
		ConnectToBCPMediaController
	End If

End Sub


Sub Table1_Exit
	'gameDebugger.Disconnect
	If Not IsNull(bcpController) Then
		bcpController.Disconnect
		Set bcpController = Nothing
	End If
End Sub
