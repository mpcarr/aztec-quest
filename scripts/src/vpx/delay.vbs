

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
    executionTime = AlignToQuarterSecond(gametime + delayInMs)
    
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

    debugLog.WriteToLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc
    delayQueue(executionTime).Add name, (new DelayObject)(name, callbackFunc, executionTime, args)
    delayQueueMap.Add name, executionTime
    
End Sub

Function AlignToQuarterSecond(timeMs)
    AlignToQuarterSecond = Int(timeMs / 125) * 125
End Function

Function RemoveDelay(name)
    If delayQueueMap.Exists(name) Then
        If delayQueue.Exists(delayQueueMap(name)) Then
            delayQueue(delayQueueMap(name)).Remove name
            delayQueueMap.Remove name
            RemoveDelay = True
            debugLog.WriteToLog "Delay", "Removing delay for " & name
            Exit Function
        End If
    End If
    RemoveDelay = False
End Function

Sub DelayTick()
    Dim key, delayObject

    Dim executionTime
    executionTime = AlignToQuarterSecond(gametime)
    If delayQueue.Exists(executionTime) Then
        For Each key In delayQueue(executionTime).Keys()
            Set delayObject = delayQueue(executionTime)(key)
            debugLog.WriteToLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
            GetRef(delayObject.Callback)(delayObject.Args)
        Next
        delayQueue.Remove executionTime
    End If
End Sub