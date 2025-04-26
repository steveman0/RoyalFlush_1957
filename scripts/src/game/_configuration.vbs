
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
	
	' Flippers
	With CreateGlfFlipper("left")
		.Switch = "LeftFlipper"
		.ActionCallback = "LeftFlipperAction"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started", "enable_flippers")
	End With

	With CreateGlfFlipper("right")
		.Switch = "RightFlipper"
		.ActionCallback = "RightFlipperAction"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started", "enable_flippers")
	End With
	
	' Slingshots
	With CreateGlfAutoFireDevice("left_sling")
		.Switch = "LeftSlingShot"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With

	With CreateGlfAutoFireDevice("right_sling")
		.Switch = "RightSlingShot"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With
	
	' Bumpers
	With CreateGlfAutoFireDevice("left_bumper")
		.Switch = "Bumper1"
		.ActionCallback = "Bumper1Action"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With

	With CreateGlfAutoFireDevice("right_bumper")
		.Switch = "Bumper3"
		.ActionCallback = "Bumper3Action"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With
	
	With CreateGlfAutoFireDevice("top_left_bumper")
		.Switch = "BumperTL"
		.ActionCallback = "BumperTLAction"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With
	
	With CreateGlfAutoFireDevice("top_right_bumper")
		.Switch = "BumperTR"
		.ActionCallback = "BumperTRAction"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With
	
	With CreateGlfAutoFireDevice("bottom_left_bumper")
		.Switch = "BumperBL"
		.ActionCallback = "BumperBLAction"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With
	
	With CreateGlfAutoFireDevice("bottom_right_bumper")
		.Switch = "BumperBR"
		.ActionCallback = "BumperBRAction"
		.DisableEvents = Array("kill_flippers")
		.EnableEvents = Array("ball_started","enable_flippers")
	End With
	
	' Rubber band switches
	With CreateGlfAutoFireDevice("rubber_band")
		.Switch = "RubberBand"
		.EnableEvents = Array("ball_started")
	End With
	With CreateGlfAutoFireDevice("rubber_band1")
		.Switch = "RubberBand001"
		.EnableEvents = Array("ball_started")
	End With
	With CreateGlfAutoFireDevice("rubber_band2")
		.Switch = "RubberBand002"
		.EnableEvents = Array("ball_started")
	End With
	With CreateGlfAutoFireDevice("rubber_band3")
		.Switch = "RubberBand003"
		.EnableEvents = Array("ball_started")
	End With
	With CreateGlfAutoFireDevice("rubber_band10")
		.Switch = "RubberBand010"
		.EnableEvents = Array("ball_started")
	End With
	With CreateGlfAutoFireDevice("rubber_band11")
		.Switch = "RubberBand011"
		.EnableEvents = Array("ball_started")
	End With

	' Rollover Switches TriggerLane1
	With CreateGlfAutoFireDevice("rollover1")
		.Switch = "TriggerLane1"
		.EnableEvents = Array("ball_started")
	End With
		With CreateGlfAutoFireDevice("rollover2")
		.Switch = "TriggerLane2"
		.EnableEvents = Array("ball_started")
	End With
		With CreateGlfAutoFireDevice("rollover3")
		.Switch = "TriggerLane3"
		.EnableEvents = Array("ball_started")
	End With
		With CreateGlfAutoFireDevice("rollover4")
		.Switch = "TriggerLane4"
		.EnableEvents = Array("ball_started")
	End With
		With CreateGlfAutoFireDevice("rollover5")
		.Switch = "TriggerLane5"
		.EnableEvents = Array("ball_started")
	End With
	
	With CreateGlfSound("10pts")
		.File = "10pts" 'Name in VPX Sound Manager
		.Bus = "sfx" ' Sound bus to play on
		'.Volume = 0.6 'Override bus volume
		.Duration = 0.5 * 1000
		'.EventsWhenStopped = Array("10pts_stopped")
	End With

	With CreateGlfSoundBus("sfx")
		.SimultaneousSounds = 8
		.Volume = 0.5
	End With
	
	CreateBaseMode()
	CreateGIMode()
	CreateScoreMode()
	CreateLitModes()
	CreateRolloversMode()
	
End Sub

Sub LeftFlipperAction(Enabled)
	If Enabled Then
		LeftFlipper.RotateToEnd
	Else
		LeftFlipper.RotateToStart
	End If
End Sub

Sub RightFlipperAction(Enabled)
	If Enabled Then
		RightFlipper.RotateToEnd
	Else
		RightFlipper.RotateToStart
	End If
End Sub
