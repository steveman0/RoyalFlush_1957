'*******************************************
'	ZBMP: Bumpers
'*******************************************

Sub Bumper1Action(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperMiddle Bumper1
		DOF 105, DOFPulse
	End If
End Sub

' Sub Bumper1_Timer
	' FlBumperFadeTarget(1) = 0
' End Sub

Sub Bumper2_Hit
	RandomSoundBumperMiddle Bumper2
	DispatchPinEvent "score_10k", ActiveBall
	DOF 106, DOFPulse
End Sub

Sub Bumper3Action(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		RandomSoundBumperMiddle Bumper3
		DOF 107, DOFPulse
	End If
End Sub

' Spin Bumpers
Sub BumperTLAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		SpinRoto
	End If
End Sub

Sub BumperTRAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		SpinRoto
	End If
End Sub

Sub BumperBLAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		SpinRoto
	End If
End Sub

Sub BumperBRAction(args)
	Dim enabled : enabled = args(0)
	If enabled Then
		SpinRoto
	End If
End Sub

Dim RotoSpinning : RotoSpinning = false
Sub SpinRoto
	If RotoSpinning Then
		Exit Sub
	End If

	' Both or shaker motor only
	If RotoDofMode = 0 Or RotoDofMode = 1 Then
		DOF 109, DOFPulse
	End If
	' Both or gear motor only
	If RotoDofMode = 0 Or RotoDofMode = 2 Then
		DOF 110, DOFPulse
	End If
End Sub