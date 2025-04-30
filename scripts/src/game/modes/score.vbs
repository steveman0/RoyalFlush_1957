
Public Sub CreateScoreMode()
	Dim i
	With CreateGlfMode("score", 2000)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ended")
		.Debug = true
		With .EventPlayer()
			.Add "BumperTL_active", Array("score_50k")
			.Add "BumperTR_active", Array("score_50k")
			.Add "BumperBL_active", Array("score_50k")
			.Add "BumperBR_active", Array("score_50k")
			.Add "timer_score_50k_timer_tick", Array("score_10k")
			.Add "timer_score_500k_timer_tick", Array("score_100k")
		End With
		
		With .SoundPlayer
			With .EventName("score_10k")
				.Sound = "10pts"
				.Action = "play"
			End With
			With .EventName("score_100k")
				.Sound = "10pts"
				.Action = "play"
			End With
			With .EventName("score_1m")
				.Sound = "10pts"
				.Action = "play"
			End With
		End With
		
		With .VariablePlayer()
			With .EventName("score_10k") 
				With .Variable("score")
					.Action = "add"
					.Int = "10000"
				End With
			End With
			.Debug = true
			With .EventName("score_100k") 
				With .Variable("score")
					.Action = "add"
					.Int = "100000"
				End With
			End With
			With .EventName("score_1m") 
				With .Variable("score")
					.Action = "add"
					.Int = "1000000"
				End With
			End With
		End With

		With .DOFPlayer()
			For i=0 to 9
				With .EventName("score_10k.1" & i & "{(current_player.score Mod 100000) / 10000 == i}")
					.Action = "DOF_ON"
					.DOFEvent = 11 + i
				End With
				With .EventName("score_10k.2" & i & "{(current_player.score Mod 100000) / 10000 == i}")
					.Action = "DOF_OFF"
					.DOFEvent = 10 + i
				End With
			Next
			.Debug = true
		End With
		
		' 50k points awarded by ticking 10k 5 times
		With .Timers("score_50k_timer")
			With .ControlEvents()
				.EventName = "score_50k"
				.Action = "start"
			End With
			' Only reset on completion / don't restart while active
			With .ControlEvents()
				.EventName = "timer_score_50k_timer_complete"
				.Action = "reset"
			End With
			.Direction = "down"
			.StartValue = 5
			.EndValue = 0
			.TickInterval = 100    ' In ms
		End With
		
		' 500k points awarded by ticking 100k 5 times
		With .Timers("score_500k_timer")
			With .ControlEvents()
				.EventName = "score_500k"
				.Action = "start"
			End With
			' Only reset on completion / don't restart while active
			With .ControlEvents()
				.EventName = "timer_score_500k_timer_complete"
				.Action = "reset"
			End With
			.Direction = "down"
			.StartValue = 1
			.EndValue = 0
			.TickInterval = 100    ' In ms
		End With
	End With
End Sub