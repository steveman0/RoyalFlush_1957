Public Sub CreateResetMode()
    With CreateGlfMode("reset", 5000)
		.StartEvents = Array("game_started")
		.StopEvents = Array("reset_complete")
        Dim i

        With .EventPlayer()
            .Add "mode_reset_started{machine.ace_card == 1 && machine.king_card == 1 && machine.queen_card == 1 && machine.jack_card == 1 && machine.ten_card == 1 && machine.joker_card == 1}", Array("reset_ace_card", "reset_king_card", "reset_queen_card", "reset_jack_card", "reset_ten_card", "reset_joker_card", "pick_starter_card")
            .Add "mode_reset_started{machine.ace_card == 0 && machine.king_card == 0 && machine.queen_card == 0 && machine.jack_card == 0 && machine.ten_card == 0 && machine.joker_card == 0}", Array("pick_starter_card")
            .Add "mode_reset_started", Array("match_off")
            ' The actual player score is reset at the end of the last game
            ' Use a stand-in to continue updating the B2S until reset is complete
'            .Add "timer_score_reset_timer_tick{machine.last_score Mod 100000 != 0}", Array("sim_score_tick", "play_10pts")
'            .Add "timer_score_reset_timer_tick{machine.last_score Mod 100000 == 0}", Array("reset_score", "update_b2s", "reset_complete")
        End With

        With .VariablePlayer()
            With .EventName("reset_score") 
                With .Variable("score")
                    .Action = "set"
                    .Int = 0
                End With
            End With
            ' With .EventName("sim_score_tick") 
            '     With .Variable("last_score")
            '         .Action = "machine_add"
            '         .Int = 10000
            '     End With
            ' End With
        End With

        With .DOFPlayer()
            With .EventName("reset_ace_card")
                .Action = "DOF_OFF"
                .DOFEvent = 3
            End With
            With .EventName("reset_king_card")
                .Action = "DOF_OFF"
                .DOFEvent = 4
            End With
            With .EventName("reset_queen_card")
                .Action = "DOF_OFF"
                .DOFEvent = 5
            End With
            With .EventName("reset_jack_card")
                .Action = "DOF_OFF"
                .DOFEvent = 6
            End With
            With .EventName("reset_ten_card")
                .Action = "DOF_OFF"
                .DOFEvent = 7
            End With
            With .EventName("reset_joker_card")
                .Action = "DOF_OFF"
                .DOFEvent = 8
            End With
            For i = 0 To 9
                With .EventName("match_off." & i)
                    .Action = "DOF_OFF"
                    .DOFEvent = (40 + i)
                End With
            Next
        End With

        With .RandomEventPlayer
            With .EventName("pick_starter_card")
                .Add "set_joker_card", 1
                .Add "set_ace_card", 1
                .Add "set_king_card", 1
                .Add "set_queen_card", 1
                .Add "set_jack_card", 1
                .Add "set_ten_card", 1
            End With
        End With

        ' Tick up to 9 times to roll 10k score back to 0
		With .Timers("score_reset_timer")
			With .ControlEvents()
				.EventName = "mode_reset_started"
				.Action = "start"
			End With
			' Only reset on completion / don't restart while active
			With .ControlEvents()
				.EventName = "timer_score_reset_timer_complete"
				.Action = "reset"
			End With
			.Direction = "down"
			.StartValue = 9
			.EndValue = 0
			.TickInterval = 100    ' In ms
		End With
    End With
End Sub

