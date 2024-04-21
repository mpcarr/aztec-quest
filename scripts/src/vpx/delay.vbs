

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
Dim delayCallbacks : Set delayCallbacks = CreateObject("Scripting.Dictionary")

Sub SetDelay(name, callbackFunc, args, delayInMs)
    Dim executionTime
    executionTime = gametime + delayInMs
    
    If delayQueue.Exists(name) Then
        delayQueue.Remove(name)
    End If
    debugLog.WriteToLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc
    delayQueue.Add name, (new DelayObject)(name, callbackFunc, executionTime, args) 
End Sub

Sub RemoveDelay(name)
    If delayQueue.Exists(name) Then
        delayQueue.Remove(name)
    End If
End Sub

Sub DelayTick()
    Dim key, delayObject
    
    For Each key In delayQueue.Keys()
        Set delayObject = delayQueue(key)
        If delayObject.TTL <= gametime Then
            delayQueue.remove(key)
            debugLog.WriteToLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
            GetRef(delayObject.Callback)(delayObject.Args)
        End If
    Next
End Sub