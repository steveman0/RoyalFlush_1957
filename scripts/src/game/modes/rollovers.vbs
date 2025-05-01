

Sub CreateRolloversMode()

	With CreateGlfMode("rollovers", 500)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ended")   
		.Debug = true
		With .EventPlayer()
			.Add "TriggerLane1_active", Array("score_10k")
			.Add "TriggerLane2_active", Array("score_10k")
			.Add "TriggerLane3_active", Array("score_10k")
			.Add "TriggerLane4_active", Array("score_10k")
			.Add "TriggerLane5_active", Array("score_10k")
			.Add "TriggerLane1_active{current_player.rollover3_hit == 1}", Array("rollover1_3_hit")
			.Add "TriggerLane3_active{current_player.rollover1_hit == 1}", Array("rollover1_3_hit")
			.Add "TriggerLane1_active{current_player.rollover2_hit == 1 && current_player.rollover3_hit == 1 && current_player.rollover4_hit == 1 && current_player.rollover5_hit == 1}", Array("all_rollovers_hit")
			.Add "TriggerLane2_active{current_player.rollover1_hit == 1 && current_player.rollover3_hit == 1 && current_player.rollover4_hit == 1 && current_player.rollover5_hit == 1}", Array("all_rollovers_hit")
			.Add "TriggerLane3_active{current_player.rollover2_hit == 1 && current_player.rollover1_hit == 1 && current_player.rollover4_hit == 1 && current_player.rollover5_hit == 1}", Array("all_rollovers_hit")
			.Add "TriggerLane4_active{current_player.rollover2_hit == 1 && current_player.rollover3_hit == 1 && current_player.rollover1_hit == 1 && current_player.rollover5_hit == 1}", Array("all_rollovers_hit")
			.Add "TriggerLane5_active{current_player.rollover2_hit == 1 && current_player.rollover3_hit == 1 && current_player.rollover4_hit == 1 && current_player.rollover1_hit == 1}", Array("all_rollovers_hit")
		End With
		
		With .LightPlayer()
			With .EventName("mode_rollovers_started")
				With .Lights("I1")
					.Color = "ffffff"
				End With
				With .Lights("I2")
					.Color = "ffffff"
				End With
				With .Lights("I3")
					.Color = "ffffff"
				End With
				With .Lights("I4")
					.Color = "ffffff"
				End With
				With .Lights("I5")
					.Color = "ffffff"
				End With
			End With
			
			With .EventName("TriggerLane1_active")
				With .Lights("I1")
					.Color = "000000"
				End With
			End With
			With .EventName("TriggerLane2_active")
				With .Lights("I2")
					.Color = "000000"
				End With
			End With
			With .EventName("TriggerLane3_active")
				With .Lights("I3")
					.Color = "000000"
				End With
			End With
			With .EventName("TriggerLane4_active")
				With .Lights("I4")
					.Color = "000000"
				End With
			End With
			With .EventName("TriggerLane5_active")
				With .Lights("I5")
					.Color = "000000"
				End With
			End With
			.Debug = true
		End With
		
		With .VariablePlayer
			.Debug = true
			With .EventName("mode_rollovers_started")
				With .Variable("rollover1_hit")
					.Action = "set"
					.Int = 0
				End With
				With .Variable("rollover2_hit")
					.Action = "set"
					.Int = 0
				End With
				With .Variable("rollover3_hit")
					.Action = "set"
					.Int = 0
				End With
				With .Variable("rollover4_hit")
					.Action = "set"
					.Int = 0
				End With
				With .Variable("rollover5_hit")
					.Action = "set"
					.Int = 0
				End With
				With .Variable("all_rollovers")
					.Action = "set"
					.Int = 0
				End With
			End With
			
			With .EventName("TriggerLane1_active")
				With .Variable("rollover1_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("TriggerLane2_active")
				With .Variable("rollover2_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("TriggerLane3_active")
				With .Variable("rollover3_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("TriggerLane4_active")
				With .Variable("rollover4_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("TriggerLane5_active")
				With .Variable("rollover5_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("all_rollovers_hit")
				With .Variable("all_rollovers")
					.Action = "set"
					.Int = 1
				End With
			End With
		End With
	End With
End Sub