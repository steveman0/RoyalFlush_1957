
'******************************************************
'	ZGCF:  GLF Configurations
'******************************************************

Sub ConfigureGlfDevices

    
    ' Plunger
    With CreateGlfBallDevice("plunger")
        .BallSwitches = Array("s_Plunger")
        .EjectTimeout = 200
        .MechanicalEject = True
        .DefaultDevice = True
        .EjectCallback = "PlungerEjectCallback"
    End With


    CreateBaseMode()

End Sub
