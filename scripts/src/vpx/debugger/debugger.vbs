

'*****************************************************************************************************************************************
'  Advance Game Debugger by flux
'*****************************************************************************************************************************************
Class AdvGameDebugger

    Private m_advDebugger, m_connected

    Private Sub Class_Initialize()
        On Error Resume Next
        Set m_advDebugger = CreateObject("vpx_adv_debugger.VPXAdvDebugger")
        m_advDebugger.Connect()
        m_connected = True
        If Err Then Debug.print("Can't start advanced debugger") : m_connected = False
    End Sub

	Public Sub SendPlayerState(key, value)
		If m_connected Then
            m_advDebugger.SendPlayerState key, value
        End If
	End Sub

    Public Sub SendPinEvent(evt)
		If m_connected Then
            m_advDebugger.SendPinEvent evt
        End If
	End Sub

    Public Sub Disconnect()
        If m_connected Then
            m_advDebugger.Disconnect()
        End If
    End Sub
End Class

'*****************************************************************************************************************************************
'  Advance Game Debugger by flux
'*****************************************************************************************************************************************
