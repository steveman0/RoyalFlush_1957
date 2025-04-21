
Public Sub CreateScoreMode()

	With CreateGlfMode("score", 2000)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ended")
		
		With .EventPlayer()
			.Add "BumperTL_active", Array("score_50k")
			.Add "BumperTR_active", Array("score_50k")
			.Add "BumperBL_active", Array("score_50k")
			.Add "BumperBR_active", Array("score_50k")
			.Add "score_50k_timer_tick", Array("score_10k")
			.Add "score_500k_timer_tick", Array("score_100k")
		End With
		
		' With .SoundPlayer
			' With .EventName("score_10k")
				' .Sound = "10pts"
				' .Action = "play"
			' End With
			' With .EventName("score_100k")
				' .Sound = "10pts"
				' .Action = "play"
			' End With
			' With .EventName("score_1m")
				' .Sound = "10pts"
				' .Action = "play"
			' End With
		' End With
		
		With .VariablePlayer()
			With .EventName("score_10k") 
				With .Variable("score")
					.Action = "add"
					.Int = "10000"
				End With
			End With
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
		
		' 50k points awarded by ticking 10k 5 times
		With .Timers("score_50k_timer")
			' Configure start and stop events
			With .ControlEvents()
				.EventName = "score_50k"
				.Action = "start"
			End With
			.Direction = "down"
			.StartValue = 1
			.EndValue = 0
			.TickInterval = 200    ' In ms
		End With
		
		' 500k points awarded by ticking 100k 5 times
		With .Timers("score_500k_timer")
			' Configure start and stop events
			With .ControlEvents()
				.EventName = "score_500k"
				.Action = "start"
			End With
			.Direction = "down"
			.StartValue = 1
			.EndValue = 0
			.TickInterval = 200    ' In ms
		End With
	End With
End Sub