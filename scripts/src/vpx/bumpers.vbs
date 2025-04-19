'*******************************************
'	ZBMP: Bumpers
'*******************************************

Sub Bumper1Action(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperMiddle Bumper1
	End If
End Sub

' Sub Bumper1_Timer
	' FlBumperFadeTarget(1) = 0
' End Sub

Sub Bumper2_Hit
	RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper3Action(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperMiddle Bumper3
	End If
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