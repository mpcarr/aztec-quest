
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
