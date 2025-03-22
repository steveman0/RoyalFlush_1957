

Sub CreateBaseMode()

    With CreateGlfMode("base", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_ended")

        
    End With

End Sub