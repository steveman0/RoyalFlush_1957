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
	DispatchPinEvent "score_10k", ActiveBall
End Sub

Sub Bumper3Action(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperMiddle Bumper3
	End If
End Sub

' Spin Bumpers
Sub BumperTLAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperTop BumperTL
		SpinRoto
	End If
End Sub

Sub BumperTRAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperTop BumperTR
		SpinRoto
	End If
End Sub

Sub BumperBLAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperBottom BumperBL
		SpinRoto
	End If
End Sub

Sub BumperBRAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperBottom BumperBR
		SpinRoto
	End If
End Sub

Dim RotoSpinning : RotoSpinning = false
Sub SpinRoto
	DispatchPinEvent "score_50k", ActiveBall
	If RotoSpinning Then
		Exit Sub
	End If
End Sub