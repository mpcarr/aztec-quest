Class BallSave

    Private m_name
    Private m_active_time
    Private m_grace_period
    Private m_enable_events
    Private m_timer_start_events
    Private m_auto_launch
    Private m_balls_to_save
    Private m_enabled
    Private m_timer_started
    Private m_tick
    Private m_in_grace
    Private m_in_hurry_up
    Private m_hurry_up_time
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AutoLaunch(): AutoLaunch = m_auto_launch: End Property
  

	Public default Function init(name, active_time, grace_period, hurry_up_time, enable_events, timer_start_events, auto_launch, balls_to_save, debug_on)
        m_name = "ball_saves_"&name
        m_active_time = active_time*1000
	    m_grace_period = grace_period*1000
        m_hurry_up_time = hurry_up_time*1000
        m_enable_events = enable_events
        m_timer_start_events = timer_start_events
	    m_auto_launch = auto_launch
	    m_balls_to_save = balls_to_save
        m_enabled = False
        m_timer_started = False
        m_debug = debug_on
        AddPinEventListener m_enable_events, m_name & "_enable", "BallSaveEventHandler", 1000, Array("enable", Me)
        AddPinEventListener m_timer_start_events, m_name & "_timer_start", "BallSaveEventHandler", 1000, Array("timer_start", Me)
	  Set Init = Me
	End Function

    Public Sub Enable
        If m_enabled = True Then
            Exit Sub
        End If
        m_enabled = True
        Log "Enabling. Auto launch: "&m_auto_launch&", Balls to save: "&m_balls_to_save&", Active time: "& m_active_time&"ms"
        AddPinEventListener "ball_drain", m_name & "_ball_drain", "BallSaveEventHandler", 1000, Array("drain", Me)
        DispatchPinEvent m_name&"_enabled", Null
    End Sub

    Public Sub Disable
        'Disable ball save
        If m_enabled = False Then
            Exit Sub
        End If
        m_enabled = False
        m_timer_started = False
        Log "Disabling..."
        RemovePinEventListener "ball_drain", m_name & "_ball_drain"
        RemoveDelay "_ball_saves_"&m_name&"_disable"
        RemoveDelay m_name&"_grace_period"
        RemoveDelay m_name&"_hurry_up_time"
        DispatchPinEvent m_name&"_disabled", Null
    End Sub

    Sub Drain(ballsToSave)
        If m_enabled = True And ballsToSave > 0 Then
            Log "Ball(s) drained while active. Requesting new one(s). Auto launch: "& m_auto_launch
            DispatchPinEvent m_name&"_saving_ball", Null
            SetDelay m_name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", Me),Null), 1000
        End If
    End Sub

    Public Sub TimerStart
        'Start the timer.
        'This is usually called after the ball was ejected while the ball save may have been enabled earlier.
        If m_timer_started=True Or m_enabled=False Then
            Exit Sub
        End If
        m_timer_started=True
        DispatchPinEvent m_name&"_timer_start", Null
        If m_active_time > 0 Then
            Log "Starting ball save timer: " & m_active_time
            Log "gametime: "& gametime & ". disabled at: " & gametime+m_active_time+m_grace_period
            SetDelay m_name&"_disable", "BallSaveEventHandler" , Array(Array("disable", Me),Null), m_active_time+m_grace_period
            SetDelay m_name&"_grace_period", "BallSaveEventHandler", Array(Array("grace_period", Me),Null), m_active_time
            SetDelay m_name&"_hurry_up_time", "BallSaveEventHandler", Array(Array("hurry_up_time", Me), Null), m_active_time-m_hurry_up_time
        End If
    End Sub

    Public Sub GracePeriod
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Sub HurryUpTime
        DispatchPinEvent m_name & "_hurry_up_time", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallSaveEventHandler(args)
    Dim ownProps, ballsToSave : ownProps = args(0) : ballsToSave = args(1) 
    Dim evt : evt = ownProps(0)
    Dim ballSave : Set ballSave = ownProps(1)
    Select Case evt
        Case "enable"
            ballSave.Enable
        Case "disable"
            ballSave.Disable
        Case "grace_period"
            ballSave.GracePeriod
        Case "hurry_up_time"
            ballSave.HurryUpTime
        Case "drain"
            If ballsToSave > 0 Then
                ballSave.Drain ballsToSave
                ballsToSave = ballsToSave - 1
            End If
        Case "timer_start"
            ballSave.TimerStart
        Case "queue_release"
            If PlungerDevice.HasBall = False And ballInReleasePostion = True Then
                ReleaseBall(Null)
                If ballSave.AutoLaunch = True Then
                    SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave),Null), 500
                End If
            Else
                SetDelay ballSave.Name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", ballSave), Null), 1000
            End If
        Case "auto_launch"
            If PlungerDevice.HasBall = True Then
                'PlungerDevice.Eject
            Else
                SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave), Null), 500
            End If
    End Select
    BallSaveEventHandler = ballsToSave
End Function