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
    Private m_debug

	Public Property Get HasBall(): HasBall = Not IsNull(m_ball): End Property
  
	Public default Function init(name, ball_switches, player_controlled_eject_event, eject_timeouts, default_device, eject_angle, eject_strength, eject_direction, debug_on)
        m_ball_switches = ball_switches
        m_player_controlled_eject_event = player_controlled_eject_event
        m_eject_timeouts = eject_timeouts * 1000
        m_name = "balldevice_"&name
        m_eject_angle = eject_angle
        m_eject_strength = eject_strength
        m_eject_direction = eject_direction
        m_ball=False
        m_debug = debug_on
        m_default_device = default_device
        If default_device = True Then
            Set PlungerDevice = Me
        End If
        AddPinEventListener m_ball_switches&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_enter", Me)
        AddPinEventListener m_ball_switches&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me)
	  Set Init = Me
	End Function

    Public Sub BallEnter(ball)
        RemoveDelay m_name & "_eject_timeout"
        Set m_ball = ball
        Log "Ball Entered"        
        If m_default_device = False Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), m_ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball)
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
        dim rangle
	    rangle = PI * (m_eject_angle - 90) / 180
        Select Case m_eject_direction
            Case "y-up"
                m_ball.vely = sin(rangle)*m_eject_strength
            Case "z-up"
                m_ball.z = m_ball.z + 30
                m_ball.velz = m_eject_strength        
        End Select
        SoundSaucerKick 1, m_ball
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