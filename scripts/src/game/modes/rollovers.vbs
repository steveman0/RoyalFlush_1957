

Sub CreateRolloversMode()

	With CreateGlfMode("rollovers", 500)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ending")

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
			.Add "rollover_1_started{current_player.all_rollovers == 1}", Array("light_is1")
			.Add "rollover_2_started{current_player.all_rollovers == 1}", Array("light_is2")
			.Add "rollover_3_started{current_player.all_rollovers == 1}", Array("light_is3")
			.Add "rollover_4_started{current_player.all_rollovers == 1}", Array("light_is4")
			.Add "rollover_5_started{current_player.all_rollovers == 1}", Array("light_is5")
			.Add "TriggerLane1_active{current_player.all_rollovers == 1 && devices.state_machines.rollover_special.state==""rollover_1""}", Array("award_special")
			.Add "TriggerLane2_active{current_player.all_rollovers == 1 && devices.state_machines.rollover_special.state==""rollover_2""}", Array("award_special")
			.Add "TriggerLane3_active{current_player.all_rollovers == 1 && devices.state_machines.rollover_special.state==""rollover_3""}", Array("award_special")
			.Add "TriggerLane4_active{current_player.all_rollovers == 1 && devices.state_machines.rollover_special.state==""rollover_4""}", Array("award_special")
			.Add "TriggerLane5_active{current_player.all_rollovers == 1 && devices.state_machines.rollover_special.state==""rollover_5""}", Array("award_special")
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

			With .EventName("light_is1")
				With .Lights("IS1")
					.Color = "ffffff"
				End With
				With .Lights("IS5")
					.Color = "000000"
				End With
			End With
			With .EventName("light_is2")
				With .Lights("IS2")
					.Color = "ffffff"
				End With
				With .Lights("IS1")
					.Color = "000000"
				End With
			End With
			With .EventName("light_is3")
				With .Lights("IS3")
					.Color = "ffffff"
				End With
				With .Lights("IS2")
					.Color = "000000"
				End With
			End With
			With .EventName("light_is4")
				With .Lights("IS4")
					.Color = "ffffff"
				End With
				With .Lights("IS3")
					.Color = "000000"
				End With
			End With
			With .EventName("light_is5")
				With .Lights("IS5")
					.Color = "ffffff"
				End With
				With .Lights("IS4")
					.Color = "000000"
				End With
			End With
		End With
		
		With .VariablePlayer
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

		' Simple rotating state machine on 10k point scoring
		With .StateMachines("rollover_special")
			.PersistState = True
			.StartingState = "rollover_1"

			With .States("rollover_1")
				.Label = "Rollover 1"
				.EventsWhenStarted = Array("rollover_1_started")
			End With
			With .States("rollover_2")
				.Label = "Rollover 2"
				.EventsWhenStarted = Array("rollover_2_started")
			End With
			With .States("rollover_3")
				.Label = "Rollover 3"
				.EventsWhenStarted = Array("rollover_3_started")
			End With
			With .States("rollover_4")
				.Label = "Rollover 4"
				.EventsWhenStarted = Array("rollover_4_started")
			End With
			With .States("rollover_5")
				.Label = "Rollover 5"
				.EventsWhenStarted = Array("rollover_5_started")
			End With

			With .Transitions
				.Source = Array("rollover_1")
				.Target = "rollover_2"
				.Events = Array("score_10k")
        	End With
			With .Transitions
				.Source = Array("rollover_2")
				.Target = "rollover_3"
				.Events = Array("score_10k")
        	End With
			With .Transitions
				.Source = Array("rollover_3")
				.Target = "rollover_4"
				.Events = Array("score_10k")
        	End With
			With .Transitions
				.Source = Array("rollover_4")
				.Target = "rollover_5"
				.Events = Array("score_10k")
        	End With
			With .Transitions
				.Source = Array("rollover_5")
				.Target = "rollover_1"
				.Events = Array("score_10k")
        	End With
		End With
	End With
End Sub