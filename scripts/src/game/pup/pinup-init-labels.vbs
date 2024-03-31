
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
