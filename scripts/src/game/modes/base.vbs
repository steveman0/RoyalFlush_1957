

Sub CreateBaseMode()

	With CreateGlfMode("base", 100)
		.StartEvents = Array("ball_started")
		.StopEvents = Array("ball_ended")
		
		With .Tilt
			.WarningsToTilt = TiltWarnings
			.SettleTime = 5000  ' 5000 milliseconds (5 seconds) settle time
			.MultipleHitWindow = 1000  ' 1000 milliseconds (1 second) multiple hit window
			.ResetWarningEvents = Array("reset_warnings_event")
		End With
	End With
End Sub

Public Sub CreateGIMode()
	Dim i
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

		With .DOFPlayer()
			With .EventName("mode_gi_control_started")
				.Action = "DOF_ON"
				.DOFEvent = 1 ' Backglass lights
			End With

			For i=1 to 9
				With .EventName("mode_gi_control_started." & i)
					.Action = "DOF_OFF"
					.DOFEvent = 10 + i
				End With
				With .EventName("mode_gi_control_started.1" & i)
					.Action = "DOF_OFF"
					.DOFEvent = 20 + i
				End With
				With .EventName("mode_gi_control_started.2" & i)
					.Action = "DOF_OFF"
					.DOFEvent = 30 + i
				End With
			Next
			With .EventName("mode_gi_control_started.10")
				.Action = "DOF_ON"
				.DOFEvent = 10
			End With

		End With
	End With

End Sub