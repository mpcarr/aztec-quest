
'Set up ball devices

bd_plunger.EjectAngle = 0
bd_plunger.EjectStrength = 50
bd_plunger.EjectDirection = "y-up"

bd_cave_scoop.EjectAngle = 0
bd_cave_scoop.EjectStrength = 60
bd_cave_scoop.EjectDirection = "z-up"

'Set up diverters

dv_panther.ActionCallback = "MovePanther"
Sub MovePanther(enabled)
    If enabled Then
        DTRaise 1
    Else
        DTDrop 1
    End If
End Sub

Dim DT01, DT02
Set DT01 = (new DropTarget)(sw01, sw01a, BM_sw01, 1, 0, True, Null) 
Set DT02 = (new DropTarget)(sw02, sw02a, BM_sw02, 2, 0, True, Null)

Dim DTArray
DTArray = Array(DT01, DT02, dt_map1, dt_map2, dt_map3, dt_map4, dt_map5, dt_map6)