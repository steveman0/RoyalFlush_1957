Sub CreateRotoTargetMode()

	With CreateGlfMode("rototarget", 500)
		.StartEvents = Array("ball_started")
		.StopEvents = Array("ball_ended")
		
        ' Roto target index mapping
        ' 1: Joker
        ' 2: Ace
        ' 3: King
        ' 4: 10
        ' 5: Joker
        ' 6: Ace
        ' 7: King
        ' 8: Queen
        ' 9: 10
        ' 10: Jack
        ' 11: Queen
        ' 12: Joker
        ' 13: Ace
        ' 14: King
        ' 15: Jack

        ' Totals
        ' Ace: 3
        ' King: 3
        ' Queen: 2
        ' Jack: 2
        ' 10: 2
        ' Joker: 3

        With .EventPlayer()
            .Add "RotoTarget_active", Array("score_100k")
			.Add "BumperTL_active", Array("pick_new_card")
			.Add "BumperTR_active", Array("pick_new_card")
			.Add "BumperBL_active", Array("pick_new_card")
			.Add "BumperBR_active", Array("pick_new_card")
            .Add "RotoTarget_active{machine.roto_index == 1 && machine.joker_card == 0}", Array("set_joker_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 5 && machine.joker_card == 0}", Array("set_joker_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 12 && machine.joker_card == 0}", Array("set_joker_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 2 && machine.ace_card == 0}", Array("set_ace_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 6 && machine.ace_card == 0}", Array("set_ace_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 13 && machine.ace_card == 0}", Array("set_ace_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 3 && machine.king_card == 0}", Array("set_king_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 7 && machine.king_card == 0}", Array("set_king_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 14 && machine.king_card == 0}", Array("set_king_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 8 && machine.queen_card == 0}", Array("set_queen_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 11 && machine.queen_card == 0}", Array("set_queen_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 10 && machine.jack_card == 0}", Array("set_jack_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 15 && machine.jack_card == 0}", Array("set_jack_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 4 && machine.ten_card == 0}", Array("set_ten_card", "set_new_card")
            .Add "RotoTarget_active{machine.roto_index == 9 && machine.ten_card == 0}", Array("set_ten_card", "set_new_card")
            ' On first activating the last card the special light is lit but also awards the special
            ' Subsequent hits also trigger special while the game continues
            .Add "set_new_card{machine.ace_card == 1 && machine.king_card == 1 && machine.queen_card == 1 && machine.jack_card == 1 && machine_.ten_card == 1 && machine.joker_card == 1}", Array("light_roto_special", "award_special")
            .Add "RotoTarget_active{machine.ace_card == 1 && machine.king_card == 1 && machine.queen_card == 1 && machine.jack_card == 1 && machine_.ten_card == 1 && machine.joker_card == 1}", Array("award_special")
		End With

        Dim i
        With .RandomEventPlayer
            With .EventName("pick_new_card")
                For i = 1 to 15
                    .Add "roto_" & i, 1
                Next
            End With
        End With

        With .VariablePlayer()
            For i = 1 to 15
                With .EventName("roto_" & i) 
                    With .Variable("roto_index")
                        .Action = "set_machine"
                        .Int = i
                    End With
                End With
            Next
            With .EventName("set_joker_card") 
                With .Variable("joker_card")
                    .Action = "set_machine"
                    .Int = 1
                End With
            End With
            With .EventName("set_ace_card") 
                With .Variable("ace_card")
                    .Action = "set_machine"
                    .Int = 1
                End With
            End With
            With .EventName("set_king_card") 
                With .Variable("king_card")
                    .Action = "set_machine"
                    .Int = 1
                End With
            End With
            With .EventName("set_queen_card") 
                With .Variable("queen_card")
                    .Action = "set_machine"
                    .Int = 1
                End With
            End With
            With .EventName("set_jack_card") 
                With .Variable("jack_card")
                    .Action = "set_machine"
                    .Int = 1
                End With
            End With
            With .EventName("set_ten_card") 
                With .Variable("ten_card")
                    .Action = "set_machine"
                    .Int = 1
                End With
            End With
		End With
		
        With .DOFPlayer()
            With .EventName("set_ace_card")
                .Action = "DOF_ON"
                .DOFEvent = 3
            End With
            With .EventName("set_king_card")
                .Action = "DOF_ON"
                .DOFEvent = 4
            End With
            With .EventName("set_queen_card")
                .Action = "DOF_ON"
                .DOFEvent = 5
            End With
            With .EventName("set_jack_card")
                .Action = "DOF_ON"
                .DOFEvent = 6
            End With
            With .EventName("set_ten_card")
                .Action = "DOF_ON"
                .DOFEvent = 7
            End With
            With .EventName("set_joker_card")
                .Action = "DOF_ON"
                .DOFEvent = 8
            End With
        End With

        With .LightPlayer()
			With .EventName("light_roto_special")
				With .Lights("ISTarget")
					.Color = "ffffff"
				End With
			End With
		End With
	End With
End Sub