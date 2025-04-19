
Public Sub CreateLitModes()

	' Left bumper and slignshots lit for 100k
	With CreateGlfMode("left_lit", 500)
		.StartEvents = Array("light_left","game_started")
		.StopEvents = Array("light_right")
		.Debug = true

		With .EventPlayer()
			.Add "Bumper1_active", Array("score_100k")
			.Add "LeftSlingShot_active", Array("score_100k")
			.Add "Bumper3_active", Array("score_10k")
			.Add "RightSlingShot_active", Array("score_10k")
			.Add "score_10k", Array("light_right")
		End With
		
		With .LightPlayer()
			With .EventName("light_left")
				With .Lights("GI_LEFT")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_right")
				With .Lights("GI_LEFT")
					.Color = "000000"
				End With
			End With
		End With
	End With
	
	' Right bumper and slignshots lit for 100k
	With CreateGlfMode("right_lit", 500)
		.StartEvents = Array("light_right")
		.StopEvents = Array("light_left")
		.Debug = true

		With .EventPlayer()
			.Add "Bumper1_active", Array("score_10k")
			.Add "LeftSlingShot_active", Array("score_10k")
			.Add "Bumper3_active", Array("score_100k")
			.Add "RightSlingShot_active", Array("score_100k")
			.Add "score_10k", Array("light_left")
		End With
		
		With .LightPlayer()
			With .EventName("light_right")
				With .Lights("GI_RIGHT")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_left")
				With .Lights("GI_RIGHT")
					.Color = "000000"
				End With
			End With
		End With
	End With

End Sub