
Sub PlayVPXSeq()
	LightSeq.Play SeqCircleOutOn, 20, 1
	lightCtrl.SyncWithVpxLights LightSeq
	'lightCtrl.SetVpxSyncLightGradientColor MakeGradident, coordsX, 80
End Sub

Sub LightSeq_PlayDone()
    lightCtrl.StopSyncWithVpxLights()
End Sub

Function MakeGradident()
    ' Define the start and end colors
    Dim startColor
    Dim endColor
    startColor = "993400"  ' Red
    endColor = "FF0000"    ' Green

    ' Define the stop positions and colors
    Dim stopPositions(3)
    Dim stopColors(3)
    stopPositions(0) = 0    ' Start at 0%
    stopColors(0) = "993400" ' Red
    stopPositions(1) = 25   ' Yellow at 50%
    stopColors(1) = "FFA500" ' Yellow
    stopPositions(2) = 50   ' Orange at 75%
    stopColors(2) = "FF0000" ' Orange
	stopPositions(3) = 75   ' Orange at 75%
    stopColors(3) = "0080ff" ' Orange

    ' Call the GetGradientColorsWithStops function to generate the gradient colors
    MakeGradident = lightCtrl.GetGradientColorsWithStops(startColor, endColor, stopPositions, stopColors)

End Function
