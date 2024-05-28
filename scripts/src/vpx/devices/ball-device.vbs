Class BallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeouts
    Private m_balls
    Private m_eject_angle
    Private m_eject_strength
    Private m_eject_direction
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_mechcanical_eject
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
	Public Property Get HasBall(): HasBall = Not IsNull(m_balls(0)): End Property
    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    Public Property Let EjectAllEvents(value) : m_eject_all_events = value : End Property
    Public Property Let MechcanicalEject(value) : m_mechcanical_eject = value : End Property
        
	Public default Function init(name, ball_switches, player_controlled_eject_event, eject_timeouts, default_device, debug_on)
        m_ball_switches = ball_switches
        m_player_controlled_eject_event = player_controlled_eject_event
        m_eject_timeouts = eject_timeouts * 1000
        m_eject_all_events = Array()
        m_name = "balldevice_" & name
        m_balls = Array(Ubound(ball_switches))
        m_debug = debug_on
        m_default_device = default_device
        If default_device = True Then
            Set PlungerDevice = Me
        End If
        Dim x
        For x=0 to UBound(ball_switches)
            AddPinEventListener ball_switches(x)&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_entering", Me, x)
            AddPinEventListener ball_switches(x)&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, x)
        Next
        
	  Set Init = Me
	End Function

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_eject_timeout"
        SoundSaucerLockAtBall ball
        Set m_balls(switch) = ball
        Log "Ball Entered" 
        DispatchPinEvent m_name & "_ball_entered", Null
        If m_default_device = False And switch = 0 Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        m_balls(switch) = Null
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechcanical_eject = True Then
            SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), ball), m_eject_timeouts
        End If
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        DispatchPinEvent m_name & "_ball_eject_success", Null
        Log "Ball successfully exited"
    End Sub

    Public Sub Eject
        Log "Ejecting."
        SetDelay m_name & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_balls(0)), m_eject_timeouts
        GetRef(m_eject_callback)(m_balls(0))
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
    Dim switch
    Select Case evt
        Case "ball_entering"
            switch = ownProps(2)
            SetDelay ballDevice.Name & "_" & switch & "_ball_enter", "BallDeviceEventHandler", Array(Array("ball_enter", ballDevice, switch), ball), 500
        Case "ball_enter"
            switch = ownProps(2)
            ballDevice.BallEnter ball, switch
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_exiting"
            switch = ownProps(2)
            If RemoveDelay(ballDevice.Name & "_" & switch & "_ball_enter") = False Then
                ballDevice.BallExiting ball, switch
            End If
        Case "eject_timeout"
            ballDevice.BallExitSuccess ball
    End Select
End Sub