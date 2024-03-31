
'***********************************************************************************
'***** Player State                                                     	    ****
'***********************************************************************************

'Score 
Const SCORE = "Player Score"
Const PLAYER_NAME = "Player Name"
'Ball
Const CURRENT_BALL = "Current Ball"
'Lanes
Const LANE_1 = "Lane 1"
Const LANE_2 = "Lane 2"
Const LANE_3 = "Lane 3"
Const LANE_4 = "Lane 4"
'Ball Save
Const BALL_SAVE_ENABLED = "Ball Save Enabled"
'Locked Balls
Const BALLS_LOCKED = "Balls Locked"

'Game
Const TURTLE = "Turtle"
Const PIZZA_INGREDIENT_1 = "Pizza Ingredient 1"
Const PIZZA_INGREDIENT_2 = "Pizza Ingredient 2"
Const PIZZA_INGREDIENT_3 = "Pizza Ingredient 3"
Const PIZZA_INGREDIENT_4 = "Pizza Ingredient 4"
Const PIZZA_INGREDIENT_5 = "Pizza Ingredient 5"
Const PIZZA_INGREDIENT_6 = "Pizza Ingredient 6"
Const PIZZA_INGREDIENT_7 = "Pizza Ingredient 7"
Const PIZZA_INGREDIENT_8 = "Pizza Ingredient 8"
Const CURRENT_MODE = "Current Mode"
Const MODE_SELECT_TURTLE = "Mode Select Turtle"


Function GetPlayerState(key)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If playerState(currentPlayer).Exists(key)  Then
        GetPlayerState = playerState(currentPlayer)(key)
    Else
        GetPlayerState = Null
    End If
End Function

Function SetPlayerState(key, value)
    If IsNull(currentPlayer) Then
        Exit Function
    End If

    If playerState(currentPlayer).Exists(key)  Then
        playerState(currentPlayer)(key) = value
    Else
        playerState(currentPlayer).Add key, value
    End If
    gameDebugger.SendPlayerState key, value
    If playerEvents.Exists(key) Then
        Dim x
        For Each x in playerEvents(key).Keys()
            If playerEvents(key)(x) = True Then
                ExecuteGlobal x
            End If
        Next
    End If
    
    SetPlayerState = Null
End Function

Sub AddStateListener(e, v)
    If Not playerEvents.Exists(e) Then
        playerEvents.Add e, CreateObject("Scripting.Dictionary")
    End If
    playerEvents(e).Add v, True
End Sub

Sub AddPinEventListener(e, v)
    If Not pinEvents.Exists(e) Then
        pinEvents.Add e, CreateObject("Scripting.Dictionary")
    End If
    pinEvents(e).Add v, True
End Sub

Sub EmitAllPlayerEvents()
    Dim key
    For Each key in playerState(currentPlayer).Keys()
        gameDebugger.SendPlayerState key, playerState(currentPlayer)(key)
        If playerEvents.Exists(key) Then
            Dim x
            For Each x in playerEvents(key).Keys()
                If playerEvents(key)(x) = True Then
                    ExecuteGlobal x
                End If
            Next
        End If
    Next
End Sub

Sub DispatchPinEvent(e)
    If Not pinEvents.Exists(e) Then
        Exit Sub
    End If
    Dim x
    gameDebugger.SendPinEvent e
    For Each x in pinEvents(e).Keys()
        If pinEvents(e)(x) = True Then
            ExecuteGlobal x
        End If
    Next
End Sub
