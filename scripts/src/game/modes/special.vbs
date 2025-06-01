Public Sub CreateSpecialMode()
    With CreateGlfMode("special", 3000)
		.StartEvents = Array("mode_reset_stopped") ' Start after reset to avoid score ticks awarding a special
		.StopEvents = Array("match_complete")

        With .EventPlayer()
            ' Credit wheel maxes out at 26, only add when below this limit
            .Add "award_special{machine.credits < 26}", Array("add_credit")
            ' The sound effect event occurs following score addition for all scoring events
            .Add "play_10pts{current_player.score >= 5200000 && current_player.score_special_1 == 0}", Array("award_score_special_1", "award_special")
            .Add "play_10pts{current_player.score >= 5500000 && current_player.score_special_2 == 0}", Array("award_score_special_2", "award_special")
            .Add "play_10pts{current_player.score >= 6000000 && current_player.score_special_3 == 0}", Array("award_score_special_3", "award_special")
            .Add "play_10pts{current_player.score >= 6500000 && current_player.score_special_4 == 0}", Array("award_score_special_4", "award_special")
            .Add "play_10pts{current_player.score >= 7000000 && current_player.score_special_5 == 0}", Array("award_score_special_5", "award_special")
        End With

        With .VariablePlayer()
            With .EventName("add_credit") 
                With .Variable("credits")
                    .Action = "add_machine"
                    .Int = 1
                End With
            End With
            With .EventName("mode_special_started")
                With .Variable("score_special_1")
                    .Action = "set"
                    .Int = 0
                End With
                With .Variable("score_special_2")
                    .Action = "set"
                    .Int = 0
                End With
                With .Variable("score_special_3")
                    .Action = "set"
                    .Int = 0
                End With
                With .Variable("score_special_4")
                    .Action = "set"
                    .Int = 0
                End With
                With .Variable("score_special_5")
                    .Action = "set"
                    .Int = 0
                End With
            End With
            With .EventName("award_score_special_1")
                With .Variable("score_special_1")
                    .Action = "set"
                    .Int = 1
                End With
            End With
            With .EventName("award_score_special_2")
                With .Variable("score_special_2")
                    .Action = "set"
                    .Int = 1
                End With
            End With
            With .EventName("award_score_special_3")
                With .Variable("score_special_3")
                    .Action = "set"
                    .Int = 1
                End With
            End With
            With .EventName("award_score_special_4")
                With .Variable("score_special_4")
                    .Action = "set"
                    .Int = 1
                End With
            End With
            With .EventName("award_score_special_5")
                With .Variable("score_special_5")
                    .Action = "set"
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

