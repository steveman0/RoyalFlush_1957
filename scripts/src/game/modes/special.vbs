Public Sub CreateSpecialMode()
    With CreateGlfMode("special", 3000)
		.StartEvents = Array("game_started")
		.StopEvents = Array("game_ended")

        With .EventPlayer()
            ' Credit wheel maxes out at 26, only add when below this limit
            .Add "award_special{machine.credits < 26}", Array("add_credit")
        End With

        With .VariablePlayer()
            With .EventName("add_credit") 
                With .Variable("credits")
                    .Action = "add_machine"
                    .Int = 1
                End With
            End With
        End With

        With .DOFPlayer()
            With .EventName("award_special")
                .Action = "DOF_PULSE"
                .DOFEvent = 108
            End With
        End With
    End With
End Sub

