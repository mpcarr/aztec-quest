
Dim mode_beasts : Set mode_beasts = (new Mode)("beasts", 100) 
With mode_beasts
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    .Debug = False
End With

Dim counter_beasts : Set counter_beasts = (new Counter)("beasts", mode_beasts)
With counter_beasts
    .EnableEvents = Array("mode_beasts_started", "sw01_inactive")
    .CountEvents = Array("sw99_active")
    .CountCompleteValue = 2
    .DisableOnComplete = True
    .ResetOnComplete = True
    .EventsWhenComplete = Array("activate_panther")
    .PersistState = True
    .Debug = False
End With

Dim timer_beasts_panther : Set timer_beasts_panther = (new ModeTimer)("beasts_panther", mode_beasts)
With timer_beasts_panther
    .StartEvents = Array("sw01_active")
    .StopEvents = Array("sw01_inactive")
    .Direction = "down"
    .StartValue = 10
    .EndValue = 0
    .Debug = True
End With

Dim event_player_beasts : Set event_player_beasts = (New EventPlayer)(mode_beasts)
Dim event_player_beasts_events : Set event_player_beasts_events = CreateObject("Scripting.Dictionary")
event_player_beasts_events.Add "timer_beasts_panther_complete", Array("deactivate_panther")
With event_player_beasts
    .Events = event_player_beasts_events
    .Debug = False
End With


Dim beasts_show : beasts_show = Array( _
(New ShowPlayerLightStep)(0, Array(Array(l01,rgb(255,255,255), 200), Array(l02,rgb(255,255,255), 200))), _ 
(New ShowPlayerLightStep)(2000, Array(Array(l01,rgb(255,0,255), 200), Array(l02,rgb(255,0,255), 200))) _ 
)

Dim show_player_beasts_item : Set show_player_beasts_item = (New ShowPlayerItem)("flash", mode_beasts, beasts_show)
With show_player_beasts_item
   .Speed = 1
   .Tokens = ""
   .Debug = False
End With

Dim show_player_beasts_events : Set show_player_beasts_events = CreateObject("Scripting.Dictionary")
show_player_beasts_events.Add "sw01_active", show_player_beasts_item

Dim show_player_beasts : Set show_player_beasts = (New ShowPlayer)(mode_beasts)
With show_player_beasts
   .Events = show_player_beasts_events
   .Debug = True
End With

'Dim light_player_beasts_events : Set light_player_beasts_events = CreateObject("Scripting.Dictionary")
'light_player_beasts_events.Add "sw01_active", Array(Array(l01,rgb(255,255,255)), Array(l02,rgb(255,255,255)))
'light_player_beasts_events.Add "sw01_inactive", Array(Array(l01,"off"), Array(l02,"off"))

'Dim light_player_beasts : Set light_player_beasts = (New LightPlayer)(mode_beasts)
'With light_player_beasts
'   .Events = light_player_beasts_events
'   .Debug = True
'End With
