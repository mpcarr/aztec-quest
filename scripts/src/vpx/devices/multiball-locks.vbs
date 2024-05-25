
Class MultiballLocks

    Private m_name
    Private m_priority
    Private m_mode
    Private m_enable_events
    Private m_disable_events
    Private m_balls_to_lock
    Private m_balls_locked
    Private m_lock_devices
    Private m_debug

    Private m_count

    Public Property Let EnableEvents(value) : m_enable_events = value : End Property
    Public Property Let DisableEvents(value) : m_disable_events = value : End Property
    Public Property Let BallsToLock(value) : m_balls_to_lock = value : End Property
    Public Property Let LockDevices(value) : m_lock_devices = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name, mode, balls_to_lock, lock_devices)
        m_name = "multiball_lock_" & name
        m_mode = mode.Name
        m_balls_to_lock = balls_to_lock
        m_lock_devices = lock_devices
        m_balls_locked = 0

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "MultiballLocksHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "MultiballLocksHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MultiballLocksHandler", m_priority, Array("enable", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_enable_events
            RemovePinEventListener evt, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_lock_devices
            AddPinEventListener evt & "_ball_entered", m_name & "_ball_locked", "MultiballLocksHandler", m_priority, Array("lock", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_count_events
            RemovePinEventListener evt, m_name & "_count"
        Next
    End Sub

    Public Sub Lock()
        m_balls_locked = m_balls_locked + 1
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MultiballLocksHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    Select Case evt
        Case "activate"
            multiball.Activate
        Case "deactivate"
            multiball.Deactivate
        Case "enable"
            multiball.Enable
        Case "disable"
            multiball.Disable
        Case "lock"
            multiball.Lock
    End Select
    MultiballLocksHandler = kwargs
End Function