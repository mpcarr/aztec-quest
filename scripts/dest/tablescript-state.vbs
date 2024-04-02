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
Const SWITCH_LEFT_FLIPPER_UP = "Switches Left Flipper Up"
Const SWITCH_RIGHT_FLIPPER_UP = "Switches Right Flipper Up"
Const SWITCH_BOTH_FLIPPERS_PRESSED = "Switches Both Flippers Pressed"


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
