
Public Sub CreateLitModes()

	' Left bumper and slignshots lit for 100k / right gobble for 1m
	With CreateGlfMode("left_lit", 500)
		.StartEvents = Array("light_left","game_started")
		.StopEvents = Array("light_right")

		With .EventPlayer()
			.Add "Bumper1_active", Array("score_100k")
			.Add "LeftSlingShot_active", Array("score_100k")
			.Add "Bumper3_active", Array("score_10k")
			.Add "RightSlingShot_active", Array("score_10k")
			.Add "score_10k", Array("light_right")
			' Opposite side gobble lit with lanes 1 and 3 hit
			.Add "mode_left_lit_started{current_player.rollover1_hit == 1 && current_player.rollover3_hit == 1}", Array("light_right_gobble")
			.Add "rollover1_3_hit", Array("light_right_gobble")
			' Score gobble holes based on lighted status
			.Add "GobbleLsw_active", Array("score_500k")
			.Add "GobbleRsw_active{current_player.rollover1_hit == 0 Or current_player.rollover3_hit == 0}", Array("score_500k")
			.Add "GobbleRsw_active{current_player.rollover1_hit == 1 && current_player.rollover3_hit == 1}", Array("score_1m")
		End With
		
		With .LightPlayer()
			With .EventName("mode_left_lit_started")
				With .Lights("GI_LEFT")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_right_gobble")
				With .Lights("RGL")
					.Color = "ffffff"
				End With
			End With
		End With
	End With
	
	' Right bumper and slignshots lit for 100k / left gobble for 1m
	With CreateGlfMode("right_lit", 500)
		.StartEvents = Array("light_right")
		.StopEvents = Array("light_left")

		With .EventPlayer()
			.Add "Bumper1_active", Array("score_10k")
			.Add "LeftSlingShot_active", Array("score_10k")
			.Add "Bumper3_active", Array("score_100k")
			.Add "RightSlingShot_active", Array("score_100k")
			.Add "score_10k", Array("light_left")
			' Opposite side gobble lit with lanes 1 and 3 hit
			.Add "mode_right_lit_started{current_player.rollover1_hit == 1 && current_player.rollover3_hit == 1}", Array("light_left_gobble")
			.Add "rollover1_3_hit", Array("light_left_gobble")
			' Score gobble holes based on lighted status
			.Add "GobbleRsw_active", Array("score_500k")
			.Add "GobbleLsw_active{current_player.rollover1_hit == 0 Or current_player.rollover3_hit == 0}", Array("score_500k")
			.Add "GobbleLsw_active{current_player.rollover1_hit == 1 && current_player.rollover3_hit == 1}", Array("score_1m")
		End With
		
		With .LightPlayer()
			With .EventName("mode_right_lit_started")
				With .Lights("GI_RIGHT")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_left_gobble")
				With .Lights("LGL")
					.Color = "ffffff"
				End With
			End With
		End With
	End With

End Sub