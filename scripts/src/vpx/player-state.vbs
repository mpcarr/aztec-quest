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


Function GetCurrentPlayerNumber()
    
    Select Case currentPlayer
        Case "PLAYER 1":
            GetCurrentPlayerNumber = 1
        Case "PLAYER 2":
            GetCurrentPlayerNumber = 2
        Case "PLAYER 3":
            GetCurrentPlayerNumber = 3
        Case "PLAYER 4":
            GetCurrentPlayerNumber = 4
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(currentPlayer) Then
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
    
    If playerState(currentPlayer).Exists(key) Then
       playerState(currentPlayer).Remove key
    End If
    playerState(currentPlayer).Add key, value

    If IsArray(value) Then
        gameDebugger.SendPlayerState key, Join(value)
    Else
        gameDebugger.SendPlayerState key, value
    End If
    If playerEvents.Exists(key) Then
        FirePlayerEventCallback key
    End If
    
    SetPlayerState = Null
End Function

Sub RegisterPlayerStateEvent(e, v)
    If Not playerEvents.Exists(e) Then
        playerEvents.Add e, CreateObject("Scripting.Dictionary")
    End If
    playerEvents(e).Add v, True
End Sub

Sub EmitAllPlayerEvents()
    Dim key
    For Each key in playerState(currentPlayer).Keys()
        FirePlayerEventCallback key
    Next
End Sub

Sub BuildPlayerEventSelectCase()
    Dim eventName, functionName, caseString, innerDict,BuildSelectCase
    ' Initialize the Select Case string
    BuildSelectCase = "Sub FirePlayerEventCallback(eventName)" & vbCrLf
    BuildSelectCase = BuildSelectCase & "    Select Case eventName" & vbCrLf
    
    ' Iterate over the outer dictionary (playerEvents)
    For Each eventName In playerEvents.Keys
        ' Start the Case clause for this event
        caseString = "        Case """ & eventName & """:" & vbCrLf
        
        ' Get the sub-dictionary for this event
        Set innerDict = playerEvents(eventName)
        
        ' Iterate over the sub-dictionary to append function names
        For Each functionName In innerDict.Keys
            ' Only append if the value is True (as per your description)
            If innerDict(functionName) = True Then
                caseString = caseString & "            " & functionName & vbCrLf
            End If
        Next
        
        ' Append this case to the overall Select Case string
        BuildSelectCase = BuildSelectCase & caseString
    Next
    
    ' Close the Select Case statement
    BuildSelectCase = BuildSelectCase & "    End Select" & vbCrLf
    BuildSelectCase = BuildSelectCase & "End Sub"
    'debug.print(BuildSelectCase)
    ExecuteGlobal BuildSelectCase
End Sub