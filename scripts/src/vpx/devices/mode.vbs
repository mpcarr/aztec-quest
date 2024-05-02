
Class Mode

    Private m_name
    Private m_start_events
    Private m_stop_events
    Private m_debug

	Public default Function init(name, start_events, stop_events, debug_on)
        m_name = "mode_"&name
        m_start_events = start_events
        m_stop_events = stop_events
        
        m_debug = debug_on
        Dim evt
        For Each evt in m_start_events
            AddPinEventListener evt, m_name & "_start", "ModeEventHandler", 1000, Array("start", Me)
        Next
        For Each evt in m_stop_events
            AddPinEventListener evt, m_name & "_stop", "ModeEventHandler", 1000, Array("stop", Me)
        Next
        Set Init = Me
	End Function

    Public Sub Start()
        Log "Starting"
        Dim evt
    End Sub

    Public Sub Stop()
        Log "Stopping"
        Dim evt
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function ModeEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim mode : Set mode = ownProps(1)
    Select Case evt
        Case "start"
            mode.Enable
        Case "stop"
            mode.Disable
    End Select
    ModeEventHandler = kwargs
End Function