'*******************************************
'	ZBMP: Bumpers
'*******************************************

Sub Bumper1_Hit
	Addscore 250
	ToggleGI 0
	RandomSoundBumperTop Bumper1
	FlBumperFadeTarget(1) = 1   'Flupper bumper demo
	Bumper1.timerenabled = True
End Sub

Sub Bumper1_Timer
	FlBumperFadeTarget(1) = 0
End Sub

Sub Bumper2_Hit
	Addscore 250
	RandomSoundBumperMiddle Bumper2
	FlBumperFadeTarget(2) = 1   'Flupper bumper demo
	Bumper2.timerenabled = True
End Sub

Sub Bumper2_Timer
	FlBumperFadeTarget(2) = 0
End Sub

Sub Bumper3_Hit
	Addscore 250
	RandomSoundBumperBottom Bumper3
	FlBumperFadeTarget(3) = 1   'Flupper bumper demo
	Bumper3.timerenabled = True
End Sub

Sub Bumper3_Timer
	FlBumperFadeTarget(3) = 0
End Sub

Sub Bumper4_Hit
	Addscore 250
	RandomSoundBumperTop Bumper4
	FlBumperFadeTarget(4) = 1   'Flupper bumper demo
	Bumper4.timerenabled = True
End Sub

Sub Bumper4_Timer
	FlBumperFadeTarget(4) = 0
End Sub

Sub Bumper5_Hit
	Addscore 250
	RandomSoundBumperMiddle Bumper5
	FlBumperFadeTarget(5) = 1   'Flupper bumper demo
	Bumper5.timerenabled = True
End Sub

Sub Bumper5_Timer
	FlBumperFadeTarget(5) = 0
End Sub
