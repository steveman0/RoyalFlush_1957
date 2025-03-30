
Public Sub CreateScoreMode()

    With CreateGlfMode("score", 2000)
        .StartEvents = Array("game_started")
        .StopEvents = Array("game_ended")

        With .VariablePlayer()
            With .EventName("score_10k") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "10000"
                End With
            End With
            With .EventName("score_100k") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "100000"
                End With
            End With
            With .EventName("score_500k") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "500000"
                End With
            End With
            With .EventName("score_1m") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "1000000"
                End With
            End With

		End With

        
    End With

End Sub