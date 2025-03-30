'*******************************************
'	ZBMP: Bumpers
'*******************************************

Sub Bumper1_Hit
'	Addscore 250
	RandomSoundBumperMiddle Bumper1
'	FlBumperFadeTarget(1) = 1   'Flupper bumper demo
'	Bumper1.timerenabled = True
End Sub

' Sub Bumper1_Timer
	' FlBumperFadeTarget(1) = 0
' End Sub

Sub Bumper2_Hit
	RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper3_Hit
	RandomSoundBumperMiddle Bumper3
End Sub

Sub BumperTL_Hit
	RandomSoundBumperTop BumperTL
End Sub

Sub BumperTR_Hit
	RandomSoundBumperTop BumperTR
End Sub

Sub BumperBL_Hit
	RandomSoundBumperBottom BumperBL
End Sub

Sub BumperBR_Hit
	RandomSoundBumperBottom BumperBR
End Sub