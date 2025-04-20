

Sub CreateBaseMode()

	With CreateGlfMode("base", 100)
		.StartEvents = Array("ball_started")
		.StopEvents = Array("ball_ended")   
	End With
	
	With .Tilt
		.WarningsToTilt = TiltWarnings
		.SettleTime = 5000  ' 5000 milliseconds (5 seconds) settle time
		.MultipleHitWindow = 1000  ' 1000 milliseconds (1 second) multiple hit window
		.ResetWarningEvents = Array("reset_warnings_event")
		.TiltWarningEvents = Array("tilt_warning_event")
		.TiltEvents = Array("tilt_event")
		.TiltSlamTiltEvents = Array("slam_tilt_event")
	End With

End Sub

Public Sub CreateGIMode()

	With CreateGlfMode("gi_control", 1000)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ended") 
		With .LightPlayer()
			With .EventName("mode_gi_control_started")
				With .Lights("GI")
					.Color = "ffffff"
				End With
			End With
		End With
	End With

End Sub