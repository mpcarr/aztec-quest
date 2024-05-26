
Dim mode_waterfall_mb : Set mode_waterfall_mb = (new Mode)("waterfall_mb", 100) 
With mode_waterfall_mb
    .StartEvents = Array("ball_started")
    .StopEvents = Array("ball_ended")
    .Debug = True
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
    .Debug = False
End With

Dim waterfall_mb_locks : Set waterfall_mb_locks = (new MultiballLocks)("waterfall_mb_locks", mode_waterfall_mb) 
With waterfall_mb_locks
    .EnableEvents = Array("enable_waterfall")
    .BallsToLock = 3
    .LockEvents = Array("balldevice_bd_waterfall_vuk_ball_eject_success")
    .ResetEvents = Array("multiball_waterfall_started")
    .Debug = True
End With

Dim waterfall_mb : Set waterfall_mb = (new Multiball)("waterfall", mode_waterfall_mb) 
With waterfall_mb
    .EnableEvents = Array("mode_waterfall_mb_started")
    .StartEvents = Array("multiball_lock_waterfall_mb_locks_full")
    .Debug = True
End With