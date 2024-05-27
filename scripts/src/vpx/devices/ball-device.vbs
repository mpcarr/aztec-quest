Class BallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeouts
    Private m_ball
    Private m_eject_angle
    Private m_eject_strength
    Private m_eject_direction
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_debug

	Public Property Get HasBall(): HasBall = Not IsNull(m_ball): End Property
    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    Public Property Let EjectAllEvents(value) : m_eject_all_events = value : End Property
        
	Public default Function init(name, ball_switches, player_controlled_eject_event, eject_timeouts, default_device, debug_on)
        m_ball_switches = ball_switches
        m_player_controlled_eject_event = player_controlled_eject_event
        m_eject_timeouts = eject_timeouts * 1000
        m_eject_all_events = Array()
        m_name = "balldevice_" & name
        m_ball=False
        m_debug = debug_on
        m_default_device = default_device
        If default_device = True Then
            Set PlungerDevice = Me
        End If
        Dim evt
        For Each evt in m_ball_switches
            AddPinEventListener evt&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_enter", Me)
            AddPinEventListener evt&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me)
        Next
        
	  Set Init = Me
	End Function

    Public Sub BallEnter(ball)
        RemoveDelay m_name & "_eject_timeout"
        SoundSaucerLock()
        Set m_ball = ball
        Log "Ball Entered" 
        DispatchPinEvent m_name & "_ball_entered", Null
        If m_default_device = False Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), m_ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball)
        DispatchPinEvent m_name & "_ball_exiting", Null
        SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_ball), m_eject_timeouts
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        DispatchPinEvent m_name & "_ball_eject_success", Null
        m_ball = Null
        Log "Ball successfully exited"
    End Sub

    Public Sub Eject
        Log "Ejecting."
        GetRef(m_eject_callback)(m_ball)
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Sub BallDeviceEventHandler(args)
    Dim ownProps, ball : ownProps = args(0) : Set ball = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballDevice : Set ballDevice = ownProps(1)
    Select Case evt
        Case "ball_enter"
            ballDevice.BallEnter ball
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_exiting"
            ballDevice.BallExiting ball
        Case "eject_timeout"
            ballDevice.BallExitSuccess ball
    End Select
End Sub