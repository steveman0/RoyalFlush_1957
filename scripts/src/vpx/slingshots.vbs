'****************************************************************
'	ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShot_Slingshot(args)
	Dim enabled, ball : enabled = args(0)
	If enabled then
		If Not IsNull(args(1)) Then
			RS.VelocityCorrect(args(1))
		End If
		Addscore 10000
		RSling1.Visible = 1
		Sling1.TransY =  - 20   'Sling Metal Bracket
		RStep = 0
		RightSlingShot_Timer
		RightSlingShot.TimerEnabled = 1
		RightSlingShot.TimerInterval = 17
		RandomSoundSlingshotRight Sling1
		'DOF 104, DOFPulse
	End If
End Sub

Sub RightSlingShot_Timer
	Select Case RStep
		Case 3
			RSLing1.Visible = 0
			RSLing2.Visible = 1
			Sling1.TransY =  - 10
		Case 4
			RSLing2.Visible = 0
			Sling1.TransY = 0
			RightSlingShot.TimerEnabled = 0
	End Select
	RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot(args)
	Dim enabled, ball : enabled = args(0)
	If enabled then
		If Not IsNull(args(1)) Then
			LS.VelocityCorrect(args(1))
		End If
		Addscore 10000
		LSling1.Visible = 1
		Sling2.TransY =  - 20   'Sling Metal Bracket		
		LStep = 0
		LeftSlingShot_Timer
		LeftSlingShot.TimerEnabled = 1
		LeftSlingShot.TimerInterval = 17
		RandomSoundSlingshotLeft Sling2
		'DOF 103, DOFPulse
	End If
End Sub

Sub LeftSlingShot_Timer
	Select Case LStep
		Case 3
			LSLing1.Visible = 0
			LSLing2.Visible = 1
			Sling2.TransY =  - 10
		Case 4
			LSLing2.Visible = 0
			Sling2.TransY = 0
			LeftSlingShot.TimerEnabled = 0
	End Select
	LStep = LStep + 1
End Sub

Sub TestSlingShot_Slingshot
	TS.VelocityCorrect(ActiveBall)
End Sub
