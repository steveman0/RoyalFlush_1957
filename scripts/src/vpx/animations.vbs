'******************************************************
'  ZANI: Misc Animations
'******************************************************

' Flippers
Sub LeftFlipper_Animate
	Dim a: a = LeftFlipper.CurrentAngle
	FlipperLSh.RotZ = a

	Dim v, BP 
	v = 255.0 * (121.0 - LeftFlipper.CurrentAngle) / 52

	For Each BP in BP_LFlipper
		BP.RotZ = a - 91
		BP.visible = v < 128.0
	Next
	For Each BP in BP_LFlipperUp
		BP.RotZ = a - 91
		BP.visible = v >= 128.0
	Next	
End Sub

Sub RightFlipper_Animate
	Dim a: a = RightFlipper.CurrentAngle
	FlipperRSh.RotZ = a

	Dim v, BP 
	v = 255.0 * (121 - RightFlipper.CurrentAngle) / 52

	For Each BP in BP_RFlipper
		BP.RotZ = a - 91
		BP.visible = v < 128.0
	Next
	For Each BP in BP_RFlipperUp
		BP.RotZ = a - 91
		BP.visible = v >= 128.0
	Next	
End Sub

Sub Gate_Animate
	Dim a: a = Gate.CurrentAngle 
	Dim BP : For Each BP in BP_Gate_Wire: BP.RotX = -a: Next
End Sub

Sub Bumper1_Animate
	Dim z, BP
	z = Bumper1.CurrentRingOffset
	For Each BP in BP_Bumper_Ring_004 : BP.transz = z: Next
End Sub

Sub Bumper2_Animate
	Dim z, BP
	z = Bumper2.CurrentRingOffset
	For Each BP in BP_Bumper_Ring_006 : BP.transz = z: Next
End Sub

Sub Bumper3_Animate
	Dim z, BP
	z = Bumper3.CurrentRingOffset
	For Each BP in BP_Bumper_Ring_005 : BP.transz = z: Next
End Sub

Sub TriggerLane1_Animate
	Dim z : z = TriggerLane1.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLane1 : BP.transz = z: Next
End Sub

Sub TriggerLane2_Animate
	Dim z : z = TriggerLane2.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLane2 : BP.transz = z: Next
End Sub

Sub TriggerLane3_Animate
	Dim z : z = TriggerLane3.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLane3 : BP.transz = z: Next
End Sub

Sub TriggerLane4_Animate
	Dim z : z = TriggerLane4.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLane4 : BP.transz = z: Next
End Sub

Sub TriggerLane5_Animate
	Dim z : z = TriggerLane5.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLane5 : BP.transz = z: Next
End Sub

Sub TriggerLaneL_Animate
	Dim z : z = TriggerLaneL.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLaneL : BP.transz = z: Next
End Sub

Sub TriggerLaneR_Animate
	Dim z : z = TriggerLaneR.CurrentAnimOffset
	Dim BP : For Each BP in BP_TriggerLaneR : BP.transz = z: Next
End Sub
