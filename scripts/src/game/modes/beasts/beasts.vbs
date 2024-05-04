
Dim mode_beasts : Set mode_beasts = (new Mode)("beasts", 100) 
With mode_beasts
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    .Debug =  True
End With

Dim counter_beasts : Set counter_beasts = (new Counter)("beasts", mode_beasts)
With counter_beasts
    .EnableEvents = Array("mode_beasts_started", "sw01_inactive")
    .CountEvents = Array("s_left_ramp_opto_active")
    .CountCompleteValue = 2
    .DisableOnComplete = True
    .ResetOnComplete = True
    .EventsWhenComplete = Array("activate_panther")
    .PersistState = True
    .Debug = True
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
    .Debug = True
End With
