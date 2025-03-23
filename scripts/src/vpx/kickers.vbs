'*******************************************
'	ZKIC: Kickers, Saucers
'*******************************************

'To include some randomness in the Kicker's kick, use the following parmeters
Const KickerAngleTol = 2	'Number of degrees the kicker angle varies around its intended direction
Const KickerStrengthTol = 1 'Number of strength units the kicker varies around its intended strength
Dim ballInKicker1
ballInKicker1 = False

Sub Kicker1_Hit
	ballInKicker1 = True
	Addscore 5000
	SoundSaucerLock
	
	' Determine drop target bonus
	Dim dropsDropped
	dropsDropped = 0
	Dim i
	For i = 0 To UBound(DTArray)
		If DTDropped(DTArray(i).sw) Then dropsDropped = dropsDropped + 1
	Next
	Select Case (dropsDropped)
		Case 0
			queue.Add "dropBonus0", "ShowScene flexScenes(5), FlexDMD_RenderMode_DMD_RGB, 6", 2, 0, 0, 2000, 0, False
		Case 1
			Addscore 5000
			queue.Add "dropBonus1", "ShowScene flexScenes(6), FlexDMD_RenderMode_DMD_RGB, 7", 2, 0, 0, 3000, 0, False
		Case 2
			Addscore 10000
			queue.Add "dropBonus2", "ShowScene flexScenes(7), FlexDMD_RenderMode_DMD_RGB, 8", 2, 0, 0, 3000, 0, False
		Case 3
			Addscore 20000
			queue.Add "dropBonus3", "ShowScene flexScenes(8), FlexDMD_RenderMode_DMD_RGB, 9", 2, 0, 0, 3000, 0, False
	End Select
End Sub

Sub Kicker1_Timer
	SoundSaucerKick 1, Kicker1
	Kicker1.Kick 160 + RndNum( - KickerAngleTol,KickerAngleTol), 20 + RndNum( - KickerStrengthTol,KickerStrengthTol)
	Kicker1.timerenabled = False
End Sub
