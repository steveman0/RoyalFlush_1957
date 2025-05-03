Sub CreateRotoTargetMode()

	With CreateGlfMode("rototarget", 500)
		.StartEvents = Array("ball_started")
		.StopEvents = Array("ball_ended")
		
        With .EventPlayer()
            .Add "RotoTarget_active", Array("score_100k")
			.Add "BumperTL_active", Array("pick_new_card")
			.Add "BumperTR_active", Array("pick_new_card")
			.Add "BumperBL_active", Array("pick_new_card")
			.Add "BumperBR_active", Array("pick_new_card")
		End With

        Dim i
        With .RandomEventPlayer
            With .EventName("pick_new_card")
                For i = 1 to 15
                    .Add "roto_" & i, 1
                Next
            End With
        End With
		
	End With
End Sub