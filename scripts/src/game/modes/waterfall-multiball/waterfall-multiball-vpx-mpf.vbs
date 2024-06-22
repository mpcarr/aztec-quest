
Dim mode_waterfall_mb : Set mode_waterfall_mb = (new Mode)("waterfall_mb", 100) 
With mode_waterfall_mb
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
End With

Dim counter_waterfall_mb : Set counter_waterfall_mb = (new Counter)("waterfall_mb", mode_waterfall_mb)
With counter_waterfall_mb
    .EnableEvents = Array("mode_waterfall_mb_started", "sw01_inactive")
    .CountEvents = Array("sw45_active")
    .CountCompleteValue = 5
    .DisableOnComplete = True
    .ResetOnComplete = True
    .EventsWhenComplete = Array("enable_waterfall")
    .PersistState = True
End With

Dim waterfall_mb_locks : Set waterfall_mb_locks = (new MultiballLocks)("waterfall", mode_waterfall_mb) 
With waterfall_mb_locks
    .EnableEvents = Array("enable_waterfall")
    .BallsToLock = 3
    .LockEvents = Array("balldevice_bd_waterfall_vuk_ball_eject_success")
    .ResetEvents = Array("multiball_waterfall_started")
End With

Dim waterfall_mb : Set waterfall_mb = (new Multiball)("waterfall", "multiball_locks_waterfull" ,mode_waterfall_mb) 
With waterfall_mb
    .EnableEvents = Array("mode_waterfall_mb_started")
    .StartEvents = Array("multiball_locks_waterfall_full")
End With

Dim waterfall_mb_ball_save : Set waterfall_mb_ball_save = (new BallSave)("waterfall", mode_waterfall_mb)
With waterfall_mb_ball_save
    .EnableEvents = Array("waterfall_multiball_started")
    .ActiveTime = 10
    .HurryUpTime = 3
    .GracePeriod = 5
    .BallsToSave = -1
    .AutoLaunch = True
End With