

Class EventPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug

    Private m_value

    Public Property Let Events(value) : m_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        
        AddPinEventListener m_mode & "_starting", m_name & "_activate", "EventPlayerEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "EventPlayerEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        
    End Sub

    Public Sub Deactivate()
        Dim evt
        
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function EventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            eventPlayer.Activate
        Case "deactivate"
            eventPlayer.Deactivate
    End Select
    EventPlayerEventHandler = kwargs
End Function