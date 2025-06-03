Public Sub CreateMatchMode()
    With CreateGlfMode("match", 4000)
		.StartEvents = Array("game_ending")
		.StopEvents = Array("match_complete")
        Dim i

        With .EventPlayer()
            ' Game ending event resets the player score, use persisted machine variable instead
            .Add "match_0{machine.player1_score Mod 100000 == 0}", Array("award_special")
            .Add "match_1{machine.player1_score Mod 100000 == 10000}", Array("award_special")
            .Add "match_2{machine.player1_score Mod 100000 == 20000}", Array("award_special")
            .Add "match_3{machine.player1_score Mod 100000 == 30000}", Array("award_special")
            .Add "match_4{machine.player1_score Mod 100000 == 40000}", Array("award_special")
            .Add "match_5{machine.player1_score Mod 100000 == 50000}", Array("award_special")
            .Add "match_6{machine.player1_score Mod 100000 == 60000}", Array("award_special")
            .Add "match_7{machine.player1_score Mod 100000 == 70000}", Array("award_special")
            .Add "match_8{machine.player1_score Mod 100000 == 80000}", Array("award_special")
            .Add "match_9{machine.player1_score Mod 100000 == 90000}", Array("award_special")
            For i = 0 To 9
                .Add "match_" & i, Array("match_complete")
            Next
        End With

        With .RandomEventPlayer
            With .EventName("mode_match_started")
                For i = 0 to 9
                    .Add "match_" & i, 1
                Next
            End With
        End With

        With .DOFPlayer()
            For i = 0 To 9
                With .EventName("match_" & i)
                    .Action = "DOF_ON"
                    .DOFEvent = (40 + i)
                End With
            Next
        End With
    End With
End Sub

