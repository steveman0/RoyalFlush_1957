
Public Sub CreateLitModes()

	' Left bumper and slignshots lit for 100k
	With CreateGlfMode("left_lit", 500)
		.StartEvents = Array("light_left")
		.StopEvents = Array("light_right")

		With .EventPlayer()
			.Add "Bumper1_hit", Array("score_100k")
			.Add "LeftSlingShot_hit", Array("score_100k")
			.Add "Bumper3_hit", Array("score_10k")
			.Add "RightSlingShot_hit", Array("score_10k")
			.Add "RubberSWx_hit", Array("light_right") ' TODO: hook up individual rubber switches 
		End With
		
		With .LightPlayer()
			With .EventName("light_left")
				With .Lights("GI_LEFT_BUMP")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_left")
				With .Lights("GI_LEFT_SLING")
					.Color = "ffffff"
				End With
			End With
			
			' Turn out lights for other side
			With .EventName("light_left")
				With .Lights("GI_RIGHT_BUMP")
					.Color = "000000"
				End With
			End With
			With .EventName("light_left")
				With .Lights("GI_RIGHT_SLING")
					.Color = "000000"
				End With
			End With
		End With
        
	End With
	
	' Right bumper and slignshots lit for 100k
	With CreateGlfMode("right_lit", 500)
		.StartEvents = Array("light_right")
		.StopEvents = Array("light_left")

		With .EventPlayer()
			.Add "Bumper1_hit", Array("score_10k")
			.Add "LeftSlingShot_hit", Array("score_10k")
			.Add "Bumper3_hit", Array("score_100k")
			.Add "RightSlingShot_hit", Array("score_100k")
			.Add "RubberSWx_hit", Array("light_left") ' TODO: hook up individual rubber switches 
		End With
		
		With .LightPlayer()
			With .EventName("light_right")
				With .Lights("GI_RIGHT_BUMP")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_right")
				With .Lights("GI_RIGHT_SLING")
					.Color = "ffffff"
				End With
			End With
			
			' Turn out lights for other side
			With .EventName("light_right")
				With .Lights("GI_LEFT_BUMP")
					.Color = "000000"
				End With
			End With
			With .EventName("light_right")
				With .Lights("GI_LEFT_SLING")
					.Color = "000000"
				End With
			End With
		End With
        
	End With

End Sub