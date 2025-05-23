'****************************************************************
'	ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShotAction(args)
	Dim enabled, ball : enabled = args(0)
	If enabled then
		If Not IsNull(args(1)) Then
			RS.VelocityCorrect(args(1))
		End If
		RStep = 0
		RightSlingShot_Timer
		RightSlingShot.TimerEnabled = 1
		RightSlingShot.TimerInterval = 17
		RandomSoundSlingshotRight Sling1
		DOF 104, DOFPulse
	End If
End Sub

Sub RightSlingShot_Timer
	Dim BL
	Dim x1, x2, y : x1 = True : x2 = False : y = 25	
	Select Case RStep
		Case 3: x1 = False : x2 = True : y = 15
		Case 4: x1 = False : x2 = False : y = 0
		RightSlingShot.TimerEnabled = 0
	End Select
	For Each BL in BP_RSling2 : BL.Visible = x1 : Next
	For Each BL in BP_RSling3 : BL.Visible = x2 : Next
	For Each BL in BP_SlingArmR : BL.transx = y : Next
	RStep = RStep + 1
End Sub

Sub LeftSlingShotAction(args)
	Dim enabled, ball : enabled = args(0)
	If enabled then
		If Not IsNull(args(1)) Then
			LS.VelocityCorrect(args(1))
		End If	
		LStep = 0
		LeftSlingShot_Timer
		LeftSlingShot.TimerEnabled = 1
		LeftSlingShot.TimerInterval = 17
		RandomSoundSlingshotLeft Sling2
		DOF 103, DOFPulse
	End If
End Sub

Sub LeftSlingShot_Timer
	Dim BL
	Dim x1, x2, y : x1 = True : x2 = False : y = 25	
	Select Case LStep
		Case 3: x1 = False : x2 = True : y = 15
		Case 4: x1 = False : x2 = False : y = 0
			LeftSlingShot.TimerEnabled = 0
	End Select
	For Each BL in BP_LSling2 : BL.Visible = x1 : Next
	For Each BL in BP_LSling3 : BL.Visible = x2 : Next
	For Each BL in BP_SlingArmL : BL.transx = y : Next
	LStep = LStep + 1
End Sub

Sub TestSlingShot_Slingshot
	TS.VelocityCorrect(ActiveBall)
End Sub
