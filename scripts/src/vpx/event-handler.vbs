
Dim BlockAllPinEvents : BlockAllPinEvents = False
Dim AllowPinEventsList : Set AllowPinEventsList = CreateObject("Scripting.Dictionary")
Dim lastPinEvent : lastPinEvent = Null
Sub DispatchPinEvent(e)
    If Not pinEvents.Exists(e) Then
        Exit Sub
    End If
    If GameTilted = True And Not e = BALL_DRAIN Then
        Exit Sub
    End If
    Dim x
    If e=SWITCH_LEFT_FLIPPER_DOWN or _
    e=SWITCH_RIGHT_FLIPPER_DOWN or _
    e=SWITCH_LEFT_FLIPPER_UP or _
    e=SWITCH_RIGHT_FLIPPER_UP or _
    e=SWITCH_BOTH_FLIPPERS_PRESSED Then
    Else
        'SetTimer "BallSearch", "BallSearch", 6000
    End If
    If BlockAllPinEvents = False Or (BlockAllPinEvents=True And AllowPinEventsList.Exists(e)) Then
        lastPinEvent = e
        gameDebugger.SendPinEvent e
        FirePinEventCallback e
    End If
End Sub

Sub AddPinEventListener(e, v)
    If Not pinEvents.Exists(e) Then
        pinEvents.Add e, CreateObject("Scripting.Dictionary")
    End If
    pinEvents(e).Add v, True
End Sub

Sub BuildPinEventSelectCase()
    Dim eventName, functionName, caseString, innerDict,BuildSelectCase
    ' Initialize the Select Case string
    BuildSelectCase = "Sub FirePinEventCallback(eventName)" & vbCrLf
    BuildSelectCase = BuildSelectCase & "    Select Case eventName" & vbCrLf
    
    ' Iterate over the outer dictionary (playerEvents)
    For Each eventName In pinEvents.Keys
        ' Start the Case clause for this event
        caseString = "        Case """ & eventName & """:" & vbCrLf
        
        ' Get the sub-dictionary for this event
        Set innerDict = pinEvents(eventName)
        
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