
'Set up ball devices
Function PlungerKickBall(ball)
    dim rangle
    rangle = PI * (0 - 90) / 180
    ball.vely = sin(rangle)*50
    SoundSaucerKick 1, ball
End Function

Function CaveKickBall(ball)
    ball.z = ball.z + 30
    ball.velz = 60        
    SoundSaucerKick 1, ball
End Function

Function WaterfallVukKickBall(ball)
    'ball.z = ball.z + 30
    'ball.velz = 1        
    SoundSaucerKick 1, ball
    sw46.Kick 0, 45, 1.36
End Function


'Set up diverters
Sub MovePanther(enabled)
    If enabled Then
        DTRaise 1
    Else
        DTDrop 1
    End If
End Sub

Sub MoveLeftOrbitDiverter(enabled)
    waterfalldiverter.isdropped=enabled
End Sub

Sub WaterfallRelease(enabled)
    sw44.isdropped=enabled
End Sub

