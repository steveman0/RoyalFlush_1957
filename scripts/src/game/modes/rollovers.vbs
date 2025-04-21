

Sub CreateRolloversMode()

	With CreateGlfMode("rollovers", 500)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ended")   
		
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
			
			With .EventName("rollover1_active")
				With .Lights("I1")
					.Color = "000000"
				End With
			End With
			With .EventName("rollover2_active")
				With .Lights("I2")
					.Color = "000000"
				End With
			End With
			With .EventName("rollover3_active")
				With .Lights("I3")
					.Color = "000000"
				End With
			End With
			With .EventName("rollover4_active")
				With .Lights("I4")
					.Color = "000000"
				End With
			End With
			With .EventName("rollover5_active")
				With .Lights("I5")
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
			End With
			
			With .EventName("rollover1_active")
				With .Variable("rollover1_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("rollover2_active")
				With .Variable("rollover2_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("rollover3_active")
				With .Variable("rollover3_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("rollover4_active")
				With .Variable("rollover4_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
			With .EventName("rollover5_active")
				With .Variable("rollover5_hit")
					.Action = "set"
					.Int = 1
				End With
			End With
		End With
	End With
End Sub