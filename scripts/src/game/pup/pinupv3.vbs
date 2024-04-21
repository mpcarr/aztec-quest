
'********************* START OF PUPDMD FRAMEWORK v3.0 BETA *************************
'******************************************************************************
'*****   Create a PUPPack within PUPPackEditor for layout config!!!  **********
'******************************************************************************
'
'
'  Quick Steps:
'      1>  create a folder in PUPVideos with Starter_PuPPack.zip and call the folder "yourgame"
'      2>  above set global variable pGameName="yourgame"
'      3>  copy paste the settings section above to top of table script for user changes.
'      4>  on Table you need to create ONE timer only called pupDMDUpdate and set it to 250 ms enabled on startup.
'      5>  go to your table1_init or table first startup function and call PUPINIT function
'      6>  Go to bottom on framework here and setup game to call the appropriate events like pStartGame (call that in your game code where needed)...etc
'      7>  attractmodenext at bottom is setup for you already,  just go to each case and add/remove as many as you want and setup the messages to show.  
'      8>  Have fun and use pDMDDisplay(xxxx)  sub all over where needed.  remember its best to make a bunch of mp4 with text animations... looks the best for sure!
'
'
'Note:  for *Future Pinball* "pupDMDupdate_Timer()" timer needs to be renamed to "pupDMDupdate_expired()"  and then all is good.
'       and for future pinball you need to add the follow lines near top
'Need to use BAM and have com idll enabled.
'				Dim icom : Set icom = xBAM.Get("icom") ' "icom" is name of "icom.dll" in BAM\Plugins dir
'				if icom is Nothing then MSGBOX "Error cannot run without icom.dll plugin"
'				Function CreateObject(className)       
'   					Set CreateObject = icom.CreateObject(className)   
'				End Function


'**************************
'   PinUp Player USER Config
'**************************

Dim pGameName       : pGameName="aztecquest"  'pupvideos foldername, probably set to cGameName in realworld

Const pTopper=0
Const pDMD=5
Const pBackglass=2
Const pMusic=4
Const pCallouts=6
Const pMusic2=7
Const pMusic3=8
Const pPopUP=11

'pages
Const pDMDBlank=0
Const pScores=1
Const pBigLine=2
Const pThreeLines=3
Const pTwoLines=4
Const pTargerLetters=5

dim clRed:   clRed = rgb(255,0,0)
dim clGreen: clGreen = rgb(0,255,0)
dim clBlue:  clBlue = rgb(0,0,255)
dim clWhite: clWhite = rgb(255,255,255)
dim clBlack: clBlack = rgb(0,0,0)
dim clOrange: clOrange = rgb(232,96,9)
dim clTmntText: clTmntText = rgb(249,196,56)

Dim PuPlayer
dim PUPDMDObject  'for realtime mirroring.
Dim pDMDlastchk: pDMDLastchk= -1    'performance of updates
Dim pDMDCurPage: pDMDCurPage= 0     'default page is empty.
Dim pInAttract : pInAttract=false   'pAttract mode
Dim pInstantInfo: pInstantInfo=false
Dim pFrameSizeX: pFrameSizeX=1920     'DO NOT CHANGE, this is pupdmd author framesize
Dim pFrameSizeY: pFrameSizeY=1080     'DO NOT CHANGE, this is pupdmd author framesize
Dim pUseFramePos : pUseFramePos=1     'DO NOT CHANGE, this is pupdmd author setting


'*************  starts PUP system,  must be called AFTER b2s/controller running so put in last line of table1_init
Sub PuPInit

    If haspup = False Then 
        Exit Sub
    End If
    If haspup = True then Set PuPlayer = CreateObject("PinUpPlayer.PinDisplay")   
    If haspup = True then PuPlayer.B2SInit "", pGameName

    PuPlayer.LabelInit pDMD
    pForceFrameRescale pDMD,1920,1080
    pDMDStartUP
End Sub 'end PUPINIT

Sub pDMDStartUP
    pSetPageLayouts
    pDMDStartGame
End Sub

Sub pSetPageLayouts

    DIM dmddef
    DIM dmdalt
    DIM dmdscr
    DIM dmdfixed
    DIM dmdTMNTtext

	dmdalt="Gameplay"    
    dmdfixed="Instruction"
	dmdscr="Impact"  'main score font
	dmddef="Impact"
	dmdTMNTtext="CCZoinks"
	
    pDMDAlwaysPAD		'we pad all text with space before and after for shadow clipping/etc

    pupCreateLabel      "Ball","",dmddef,40,clWhite,960,60,1,1
    pDMDLabelSetBorder 	"Ball",clBlack,3,3,1
    pupCreateLabel      "Credits","",dmddef,40,clWhite,960,20,1,1
    pDMDLabelSetBorder 	"Credits",clBlack,3,3,1
    pupCreateLabel      "Play1","",dmddef,60,clBlack,100,965,1,1
    pupCreateLabel      "Play2","",dmddef,50,clBlack,685,990,1,1
    pupCreateLabel      "Play3","",dmddef,50,clBlack,1140,990,1,1
    pupCreateLabel      "Play4","",dmddef,50,clBlack,1560,990,1,1
    pupCreateLabel      "CurrentPlayer","",dmdTMNTtext,75,clBlack,330,45,1,1
    
    pupCreateLabelImage "tmntoverlay","PupOverlays\\TMNTOverlay.png",1920/2,1080/2,1920,1080,1,0
    pupCreateLabelImage "tmnt_turtle_frame","",1920/2,1080/2,1920,1080,1,0
    
    pupCreateLabelImage "GreenBox5","PupOverlays\\GreenBox5.png",1680,88,475,178,1,0
    pupCreateLabelImage "GreenBox3","PupOverlays\\GreenBox3.png",1590,88,295,175,1,0

    'Slices Eaten
    pupCreateLabel      "Slices","",dmdTMNTtext,60,clTmntText,1688,200,1,1
    pupCreateLabel      "SlicesCount","",dmdscr,60,clWhite,1850,200,1,1
    pDMDLabelSetBorder 	"Slices",clBlack,3,3,1
    pDMDLabelSetBorder 	"SlicesCount",clBlack,3,3,1	

    'Large Cutting Board
    pupCreateLabelImage "Ingredient_1","",1500,85,61,112,1,0 'Pos 1
    pupCreateLabelImage "Ingredient_2","",1590,85,77,123,1,0 'Pos 2
    pupCreateLabelImage "Ingredient_3","",1680,85,77,123,1,0 'Pos 3
    pupCreateLabelImage "Ingredient_4","",1770,85,77,123,1,0 'Pos 4
    pupCreateLabelImage "Ingredient_5","",1860,85,77,123,1,0 'Pos 5
    'Small Cutting Board
    pupCreateLabelImage "Ingredient_6","",1300,76,59,91,1,1 'Pos 1    
    pupCreateLabelImage "Ingredient_7","",1365,76,59,91,1,1 'Pos 2    
    pupCreateLabelImage "Ingredient_8","",1425,76,59,91,1,1 'Pos 3

    'Episodes Labels
    pupCreateLabelImage "EpState_1","",1865,293,80,83,1,1 'Pos 1
    pupCreateLabelImage "EpState_2","",1865,385,80,83,1,1 'Pos 2
    pupCreateLabelImage "EpState_3","",1865,477,80,83,1,1 'Pos 3
    pupCreateLabelImage "EpState_4","",1865,569,80,83,1,1 'Pos 4
    pupCreateLabelImage "EpState_5","",1865,661,80,83,1,1 'Pos 5
    pupCreateLabelImage "EpState_6","",1865,753,80,83,1,1 'Pos 6
    pupCreateLabelImage "EpState_7","",1865,845,80,83,1,1 'Pos 7
    pupCreateLabelImage "EpState_8","",1865,938,80,83,1,1 'Pos 8
    'Item Labels
    pupCreateLabelImage "EpItem_1","",1870,293,80,83,1,1 'Pos 1
    pupCreateLabelImage "EpItem_2","",1870,385,80,83,1,1 'Pos 2
    pupCreateLabelImage "EpItem_3","",1870,472,80,83,1,1 'Pos 3
    pupCreateLabelImage "EpItem_4","",1870,565,80,83,1,1 'Pos 4
    pupCreateLabelImage "EpItem_5","",1870,660,80,83,1,1 'Pos 5
    pupCreateLabelImage "EpItem_6","",1870,755,80,83,1,1 'Pos 6
    pupCreateLabelImage "EpItem_7","",1870,845,80,83,1,1 'Pos 7
    pupCreateLabelImage "EpItem_8","",1870,935,80,83,1,1 'Pos 8
    'Lock Label
    pupCreateLabelImage "LockEP_7","",1870,845,80,83,1,1 'Pos 7
    pupCreateLabelImage "LockEP_8","",1870,935,80,83,1,1 'Pos 8

    'Small Ingredients

    'pupCreateLabelImage "CaptureCountMultiplayer","DMDmisc\\CaptureCount.png",991.25,180.3,90,194,1,1
    'pupCreateLabel      "Play1","P1-",dmddef,60,clWhite,395,1019,1,1
    'pupCreateLabel      "Play2","P2-",dmddef,60,clWhite,635,1019,1,1
    'pupCreateLabel      "Play3","P3-",dmddef,60,clWhite,850,1019,1,1
    'pupCreateLabel      "Play4","P4-",dmddef,60,clWhite,1110,1019,1,1

    pupCreateLabel   "curScorePos1","",dmddef,100,clWhite,300,1040,1,1
    pupCreateLabel   "curScorePos2","",dmddef,100,clWhite,960,885,1,1
    '2 Player Layout
    pupCreateLabel   "curScore2P1","",dmddef,110,clWhite,470,1035,1,1
    pupCreateLabel   "curScore2P2","",dmddef,110,clWhite,1440,1035,1,1
    '3 Player Layout
    pupCreateLabel   "curScore3P1","",dmddef,110,clWhite,350,1035,1,1
    pupCreateLabel   "curScore3P2","",dmddef,110,clWhite,990,1035,1,1
    pupCreateLabel   "curScore3P3","",dmddef,110,clWhite,1600,1035,1,1
    '4 Player Layout
    pupCreateLabel   "curScore4P1","",dmddef,110,clWhite,270,1030,1,1
    pupCreateLabel   "curScore4P2","",dmddef,110,clWhite,740,1030,1,1
    pupCreateLabel   "curScore4P3","",dmddef,110,clWhite,1180,1030,1,1	
    pupCreateLabel   "curScore4P4","",dmddef,110,clWhite,1680,1030,1,1
    'Autosize Scoring labels
    pDMDLabelSetBorder "WorkerLeft",clBlack,6,6,1
    pDMDLabelSetBorder "Player",clBlack,3,3,1
    pDMDLabelSetAutoSize "curscore",980,980
    pupCreateLabel      "Splash","",dmddef,300,clWhite,960,350,1,0
    pDMDLabelSetBorder "Splash",clBlack,3,3,1
    pupCreateLabel     "CallOuts","",dmddef,180,clWhite,960,175,10,0
    pDMDLabelSetBorder "Player",clWhite,3,3,1
    pDMDLabelSetBorder "Ball",clWhite,3,3,1
    pDMDLabelSetBorder "Credits",clWhite,3,3,1
    pDMDLabelSetShadow "curScore",RGB(0,0,0),2,2,1

End Sub






'************************ Helper Subs ***************************


'************************ called during gameplay to update Scores ***************************
Dim CurTestScore:CurTestScore=100000
Sub pUpdateScores3  'call this ONLY on timer 300ms is good enough
'if pDMDCurPage <> pScores then Exit Sub
'PuPlayer.LabelSet pDMD,"Credits","CREDITS " & ""& Credits ,1,""
puPlayer.LabelSet pDMD,"Play1","Player 1",1,""
puPlayer.LabelSet pDMD,"Play2","Player 2",1,""
puPlayer.LabelSet pDMD,"Play3","Player 3",1,""
puPlayer.LabelSet pDMD,"Play4","Player 4",1,""
puPlayer.LabelSet pDMD,"CurrentPlayer","PLAYER " & CurrentPlayer,1,""
puPlayer.LabelSet pDMD,"Slices","SLICES EATEN:",1,""
puPlayer.LabelSet pDMD,"SlicesCount","0",1,""
'puPlayer.LabelSet pDMD,"Ball"," "&pDMDCurPriority ,1,""
puPlayer.LabelSet pDMD,"CurScore","" & FormatNumber(Score(CurrentPlayer),0),1,""
puPlayer.LabelSet pDMD,"Player","Player " & CurrentPlayer,1,""
puPlayer.LabelSet pDMD,"Ball","Ball  " & Balls,1,""
puPlayer.LabelSet pDMD,"Credits","FREE PLAY "  ,1,""
pDMDLabelShow "GreenBox5"

								   
											  

		If CurrentPlayer = 1 Then
			PuPlayer.LabelSet pDMD,"curscorePos1","" & FormatNumber(Score(CurrentPlayer),0),1,""
			PuPlayer.LabelSet pDMD,"Play1","Player 1",1,""
			'pDMDSetBackFrame "P1.png" 
			'make other scores red (inactive)
		End If
		If CurrentPlayer = 2 Then
			pDMDLabelHide "Curscore"
			pDMDLabelHide "CurscorePos1"
			PuPlayer.LabelSet pDMD,"curscorePos2","" & FormatNumber(Score(CurrentPlayer),0),1,""
			PuPlayer.LabelSet pDMD,"curscore2P1","" & FormatNumber(Score(1),0),1,""
			PuPlayer.LabelSet pDMD,"curscore2P2","" & FormatNumber(Score(2),0),1,""
			'PuPlayer.LabelSet pDMD,"Play2","Player 2",1,"{'mt':2,'color':16777215 }"
			
			'pDMDSetBackFrame "P2.png"

		End If
		If CurrentPlayer = 3 Then
			pDMDLabelHide "CurscorePos1"
			pdmdLabelhide "Curscore2P1"
			pdmdLabelhide "Curscore2P2"
			PuPlayer.LabelSet pDMD,"curscorePos2","" & FormatNumber(Score(CurrentPlayer),0),1,""
			PuPlayer.LabelSet pDMD,"curscore3P1","" & FormatNumber(Score(1),0),1,""
			PuPlayer.LabelSet pDMD,"curscore3P2","" & FormatNumber(Score(2),0),1,""
			PuPlayer.LabelSet pDMD,"curscore3P3","" & FormatNumber(Score(3),0),1,""
			'pDMDSetBackFrame "P3.png"
		End If

		If CurrentPlayer = 4 Then
			PuPlayer.LabelSet pDMD,"Play4","Player 4",1,"{'mt':2,'color':16777215 }"
			pDMDLabelHide "CurscorePos1"
			pdmdLabelhide "Curscore3P1"
			pdmdLabelhide "Curscore3P2"
			pdmdLabelhide "Curscore3P3"
			PuPlayer.LabelSet pDMD,"curscorePos2","" & FormatNumber(Score(CurrentPlayer),0),1,""
			PuPlayer.LabelSet pDMD,"curscore4P1","" & FormatNumber(Score(1),0),1,""
			PuPlayer.LabelSet pDMD,"curscore4P2","" & FormatNumber(Score(2),0),1,""
			PuPlayer.LabelSet pDMD,"curscore4P3","" & FormatNumber(Score(3),0),1,""
			PuPlayer.LabelSet pDMD,"curscore4P4","" & FormatNumber(Score(3),0),1,""
			'pDMDSetBackFrame "P4.png"
End If

End Sub


Sub pDMDStartGame
    pDMDSetPage(pScores)
    PuPlayer.playlistplayex pDMD,"DMDBackground","Lair.mp4",100,1
    pDMDSetloop 1
End Sub	