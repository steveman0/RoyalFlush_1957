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

	LastRotoTime = GameTime
	TotalRotoTime = 0
	RotoTargetAngle = 1000
	RotoAngleToGo = -1000
	RotoSpinTimer.Enabled = True

	' Both or shaker motor only
	If RotoDofMode = 0 Or RotoDofMode = 1 Then
		DOF 109, DOFPulse
	End If
	' Both or gear motor only
	If RotoDofMode = 0 Or RotoDofMode = 2 Then
		DOF 110, DOFPulse
	End If
End Sub

Const DriveSpan = 100 ' Time in ms
Const SpinDownEndTime = 500
Const LockInEndTime = 600
Const MaxSpeed = 1000 ' Speed in degrees / s
Const DriveAcc = 10000 ' Acceleration in degrees / s^2
Const SpinMinOvershoot = 7 ' Degrees
Const SpinMaxOvershoot = 19
Const LockInMinError = -2
Const LockInMaxError = 2

Dim LastRotoTime
Dim TotalRotoTime
Dim RotoSpeed : RotoSpeed = 0
Dim RotoAngle : RotoAngle = RotoTargetPrim.RotY
Dim RotoTargetAngle
Dim RotoStartAngle
Dim RotoAngleToGo
Dim RotoLerpTarget
Dim RotoSpinPhase
Sub RotoSpinTimer_timer()
	Dim timeStep : timeStep = GameTime - LastRotoTime
	TotalRotoTime = TotalRotoTime + timeStep
	LastRotoTime = GameTime

	Dim angle
	If TotalRotoTime <= DriveSpan Then ' Drive phase
		RotoSpinPhase = 1
		IncrementRotoAngle RotoSpeed * timeStep / 1000
		RotoSpeed = RotoSpeed + DriveAcc * timeStep / 1000
		If RotoSpeed > MaxSpeed Then
			RotoSpeed = MaxSpeed
		End If
	ElseIf TotalRotoTime <= SpinDownEndTime Then ' Spin Down phase
		If RotoSpinPhase = 1 Then
			RotoSpinPhase = 2
			GetAngleToGo
		End If
		angle = EasedLerp(RotoStartAngle, RotoLerpTarget, (1 - (TotalRotoTime - DriveSpan) / 400))
		SetRotoAngle angle
	Else ' Lock In
		If RotoSpinPhase = 2 Then
			RotoSpinPhase = 3
			GetAngleToGo
		End If
		angle = EasedLerp(RotoStartAngle, RotoLerpTarget, (1 - (TotalRotoTime - SpinDownEndTime) / 100))
		SetRotoAngle angle
		If TotalRotoTime > LockInEndTime Then
			RotoSpinTimer.Enabled = False
		End If
	End If
End Sub

Function EasedLerp(start, target, t)
	EasedLerp = start + (target - start) * (1 - (t * t))

	' Debug.Print "Lerp start: " & start & " Lerp target: " & target & " t: " & t & " lerp: " & EasedLerp
End Function

Sub IncrementRotoAngle(angle)
	RotoAngle = RotoAngle + angle
	RotoAngleToGo = RotoAngleToGo - angle
	If (RotoAngle > 360) Then
		RotoAngle = RotoAngle - 360
	End If
	If (RotoAngle < 0) Then
		RotoAngle = RotoAngle + 360
	End If
	RotoTargetPrim.RotY = RotoAngle
End Sub

' For use by direct set through lerp
Sub SetRotoAngle(angle)
	' Restore roto angle to 0-360 range
	RotoAngle = angle Mod 360
	If RotoAngle < 0 Then
		RotoAngle = RotoAngle + 360
	End If
	RotoTargetPrim.RotY = RotoAngle
End Sub

' Lerp based model
Sub GetAngleToGo()
	' Overshoot the target final angle with some error
	' The longer spin down phase has more error/overshoot
	Dim overshoot
	If RotoSpinPhase = 2 Then
		overshoot = RndNum(SpinMinOvershoot, SpinMaxOvershoot)
	ElseIf RotoSpinPhase = 3 Then
		overshoot = RndNum(LockInMinError, LockInMaxError)
	End If

	' Start and target angles are in absolute angles on a 0-360 degree scale
	RotoStartAngle = RotoAngle
	RotoLerpTarget = RotoTargetAngle + overshoot

	' The relative angles need to be converted to a -Inf to +Inf scale to apply the lerp easily
	' Spin down phase is in positive direction, target should be greater than start
	' If it is less than, this means the target is wrapped around
	If RotoSpinPhase = 2 And RotoLerpTarget < RotoStartAngle Then
		RotoLerpTarget = 360 + RotoLerpTarget
	End If
	If RotoSpinPhase = 3 And RotoLerpTarget > RotoStartAngle Then
		RotoLerpTarget = RotoLerpTarget - 360
	End If
	' This now gives the absolute angular distance to travel
	' This isn't fully tracked beyond this point when lerping
	RotoAngleToGo = RotoLerpTarget - RotoStartAngle
	
	' Allow extra rotation in the positive spin down phase if travel is short
	If RotoSpinPhase = 2 And RotoAngleToGo < (RotoSpeed / 2 - 360) Then
		RotoLerpTarget = RotoLerpTarget + 360
		RotoAngleToGo = RotoLerpTarget - RotoStartAngle
	End If

	' Debug.Print "RotoStartAngle: " & RotoStartAngle & " RotoLerpTarget: " & RotoLerpTarget & " RotoAngleToGo: " & RotoAngleToGo & " True target: " & RotoTargetAngle
End Sub