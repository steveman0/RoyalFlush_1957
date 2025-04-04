

Sub CreateBaseMode()

    With CreateGlfMode("base", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_ended")   
    End With

End Sub

Public Sub CreateGIMode()

    With CreateGlfMode("gi_control", 1000)
        .StartEvents = Array("game_started")
        .StopEvents = Array("game_ended") 
        With .LightPlayer()
            With .EventName("mode_gi_control_started")
                With .Lights("GI")
                    .Color = "ffffff"
                End With
            End With
        End With
    End With

End Sub