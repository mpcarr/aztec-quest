

'*****************************************************************************************************************************************
'  Vpx Bcp Controller
'*****************************************************************************************************************************************

Class VpxBcpController

    Private m_bcpController, m_connected

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_server.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        bcpUpdate.Enabled = True
        If Err Then Debug.print("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

	Public Sub Send(commandMessage)
		If m_connected Then
            m_bcpController.Send commandMessage
        End If
	End Sub

    Public Function GetMessages
		If m_connected Then
            GetMessages = m_bcpController.GetMessages
        End If
	End Function

    Public Sub Reset()
		If m_connected Then
            m_bcpController.Send "reset"
        End If
	End Sub
    
    Public Sub PlaySlide(slide, context, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""" & slide & """: {""action"": ""play"", ""expire"": 0}}, ""context"": """ & context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub SendPlayerVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & GetCurrentPlayerNumber
            '06:34:34.644 : VERBOSE : BCP : Received BCP command: ball_start?player_num=int:1&ball=int:1
        End If
	End Sub

    Private Function EncodeVariable(value)
        Dim retValue
        Select Case VarType(value)
            Case vbInteger, vbLong
                retValue = "int:" & value
            Case vbSingle, vbDouble
                retValue = "float:" & value
            Case vbString
                retValue = "string:" & value
            Case vbBoolean
                retValue = "bool:" & CStr(value)
            Case Else
                retValue = "NoneType:"
        End Select
        EncodeVariable = retValue
    End Function
    
    Private Function VariableVariance(v1, v2)
        Dim retValue
        Select Case VarType(v1)
            Case vbInteger, vbLong, vbSingle, vbDouble
                retValue = Abs(v1 - v2)
            Case Else
                retValue = True 
        End Select
        VariableVariance = retValue
    End Function

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
            bcpUpdate.Enabled = False
        End If
    End Sub
End Class

Sub BcpSendPlayerVar(args)
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim player_var : player_var = kwargs(0)
    Dim value : value = kwargs(1)
    Dim prevValue : prevValue = kwargs(2)
    bcpController.SendPlayerVariable player_var, value, prevValue
End Sub

Sub BcpAddPlayer(playerNum)
    If useBcp Then
        bcpController.Send("player_added?player_num=int:"&playerNum)
    End If
End Sub

Sub bcpUpdate_Timer()
    Dim messages : messages = bcpController.GetMessages()
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter
        For Each message in messages
            Select Case message.Command
                case "hello"
                    bcpController.Reset
                case "monitor_start"
                    Dim category : category = message.GetValue("category")
                    If category = "player_vars" Then
                        AddPlayerStateEventListener SCORE, "bcp_player_var_score", "BcpSendPlayerVar", 1000, Null
                        AddPlayerStateEventListener CURRENT_BALL, "bcp_player_var_ball", "BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    Dim eventName : eventName = message.GetValue("event")
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Bcp Controller
'*****************************************************************************************************************************************
