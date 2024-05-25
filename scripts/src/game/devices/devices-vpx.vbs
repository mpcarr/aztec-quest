
'Set up ball devices
Function PlungerKickBall(ball)
    dim rangle
    rangle = PI * (0 - 90) / 180
    ball.vely = sin(rangle)*50
    SoundSaucerKick 1, ball
End Function

Function CaveKickBall(ball)
    dim rangle
    rangle = PI * (0 - 90) / 180
    ball.z = ball.z + 30
    ball.velz = 60        
    SoundSaucerKick 1, ball
End Function

bd_plunger.EjectCallback = "PlungerKickBall"
bd_cave_scoop.EjectCallback = "CaveKickBall"

'Set up diverters

dv_panther.ActionCallback = "MovePanther"
Sub MovePanther(enabled)
    If enabled Then
        DTRaise 1
    Else
        DTDrop 1
    End If
End Sub

dv_leftorbit.ActionCallback = "MoveLeftOrbitDiverter"
Sub MoveLeftOrbitDiverter(enabled)
    waterfalldiverter.isdropped=enabled
End Sub

dv_waterfall.ActionCallback = "WaterfallRelease"
Sub WaterfallRelease(enabled)
    sw44.isdropped=enabled
End Sub

Dim DT01, DT02
Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True, Null) 
Set DT02 = (new DropTarget)(sw02, sw02a, BM_sw02, 2, 0, True, Null)

Dim DTArray
DTArray = Array(DT01, DT02, dt_map1, dt_map2, dt_map3, dt_map4, dt_map5, dt_map6)