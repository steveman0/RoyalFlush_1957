
Public Sub CreateLitModes()

	' Left bumper and slignshots lit for 100k
	With CreateGlfMode("left_lit", 500)
		.StartEvents = Array("light_left","ball_started")
		.StopEvents = Array("light_right")

		With .EventPlayer()
			.Add "Bumper1_active", Array("score_100k")
			.Add "LeftSlingShot_active", Array("score_100k")
			.Add "Bumper3_active", Array("score_10k")
			.Add "RightSlingShot_active", Array("score_10k")
			.Add "RubberBand_active", Array("light_right")
			.Add "RubberBand001_active", Array("light_right")
			.Add "RubberBand002_active", Array("light_right")
			.Add "RubberBand003_active", Array("light_right")
			.Add "RubberBand010_active", Array("light_right")
			.Add "RubberBand011_active", Array("light_right")
		End With
		
		' With .LightPlayer()
			' With .EventName("light_left")
				' With .Lights("GI_LEFT")
					' .Color = "ffffff"
				' End With
			' ' Turn out lights for other side
				' With .Lights("GI_RIGHT_BUMP")
					' .Color = "000000"
				' End With
				' With .Lights("GI_RIGHT_SLING")
					' .Color = "000000"
				' End With
			' End With
		' End With
        
	End With
	
	' Right bumper and slignshots lit for 100k
	With CreateGlfMode("right_lit", 500)
		.StartEvents = Array("light_right")
		.StopEvents = Array("light_left")

		With .EventPlayer()
			.Add "Bumper1_active", Array("score_10k")
			.Add "LeftSlingShot_active", Array("score_10k")
			.Add "Bumper3_active", Array("score_100k")
			.Add "RightSlingShot_active", Array("score_100k")
			.Add "RubberBand_active", Array("light_left")
			.Add "RubberBand001_active", Array("light_left")
			.Add "RubberBand002_active", Array("light_left")
			.Add "RubberBand003_active", Array("light_left")
			.Add "RubberBand010_active", Array("light_left")
			.Add "RubberBand011_active", Array("light_left")
		End With
		
		' With .LightPlayer()
			' With .EventName("light_right")
				' With .Lights("GI_RIGHT_BUMP")
					' .Color = "ffffff"
				' End With
			' End With
			' With .EventName("light_right")
				' With .Lights("GI_RIGHT_SLING")
					' .Color = "ffffff"
				' End With
			' End With
			
			' ' Turn out lights for other side
			' With .EventName("light_right")
				' With .Lights("GI_LEFT_BUMP")
					' .Color = "000000"
				' End With
			' End With
			' With .EventName("light_right")
				' With .Lights("GI_LEFT_SLING")
					' .Color = "000000"
				' End With
			' End With
		' End With
        
	End With

End Sub