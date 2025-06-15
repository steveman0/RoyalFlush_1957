'*******************************************
'	ZDRN: Drain, Trough, and Ball Release
'*******************************************

' Original code retained for reference - largely replaced by GLF trough management

' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this,
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to
'	 the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 300 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


'TROUGH
' Sub swTrough1_Hit
	' UpdateTrough
' End Sub
' Sub swTrough1_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough2_Hit
	' UpdateTrough
' End Sub
' Sub swTrough2_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough3_Hit
	' UpdateTrough
' End Sub
' Sub swTrough3_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough4_Hit
	' UpdateTrough
' End Sub
' Sub swTrough4_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough5_Hit
	' UpdateTrough
' End Sub
' Sub swTrough5_UnHit
	' UpdateTrough
' End Sub

Sub TroughTeleport_Hit
	Dim b : Set b = GetBallNear(TroughTeleport)
	If Not b Is Nothing Then 
		TroughTeleport.kick 90, 10
		b.X = LaneKicker.X
		b.Y = LaneKicker.Y
		KickBall b, -90, 2, 0, 0
	End If
End Sub

' Trough current sits at Z=0, pop ball from subway to drain
Sub DrainTeleport_Hit
	Dim b : Set b = GetBallNear(DrainTeleport)
	If Not b Is Nothing Then 
		DrainTeleport.kick 90, 10
		b.X = Drain.X
		b.Y = Drain.Y
		b.Z = 0
		'KickBall b, -90, 2, 0, 0
	End If
	If ub5 = 0 Then
		DisplaySupply.CreateBall 'SizedballWithMass Ballsize / 2, Ballmass
		DisplaySupply.kick 80, 10
	End If
End Sub

Function GetBallNear(obj)
	Dim b
	For b = 0 To UBound(gBOT)
		If abs(obj.X - gBOT(b).X) < 25 And abs(obj.Y - gBOT(b).Y) < 25 Then
			Set GetBallNear = gBOT(b)
		End If
	Next
End Function

' Display balls played as a ball counter
Sub Drainsw_hit
	If ub5 = 0 Then
		DisplaySupply.CreateBall 'SizedballWithMass Ballsize / 2, Ballmass
		DisplaySupply.kick 80, 10
	End If
End Sub

' Game initialization will have balls in all display slots
Dim ub5, ub4, ub3, ub2, ub1
ub5 = 1
ub4 = 1
ub3 = 1
ub2 = 1
ub1 = 1

Sub UsedBall5_Hit
	ub5 = 1
	If ub4 = 0 Then
		UsedBall5.kick 80, 10
	End If
End Sub

Sub UsedBall5_Unhit
	ub5 = 0
End Sub

Sub UsedBall4_Hit
	ub4 = 1
	If ub3 = 0 Then
		UsedBall4.kick 80, 10
	End If
End Sub

Sub UsedBall4_Unhit
	ub4 = 0
End Sub

Sub UsedBall3_Hit
	ub3 = 1
	If ub2 = 0 Then
		UsedBall3.kick 80, 10
	End If
End Sub

Sub UsedBall3_Unhit
	ub3 = 0
End Sub

Sub UsedBall2_Hit
	ub2 = 1
	If ub1 = 0 Then
		UsedBall2.kick 80, 10
	End If
End Sub

Sub UsedBall2_Unhit
	ub2 = 0
End Sub

Sub UsedBall1_Hit
	ub1 = 1
End Sub

Sub UsedBallTimer_timer()
	If ub1 = 1 Then
		UsedBall1.kick 80, 10
	End If
	If ub2 = 1 Then
		UsedBall2.kick 80, 10
	End If
	If ub3 = 1 Then
		UsedBall3.kick 80, 10
	End If
	If ub4 = 1 Then
		UsedBall4.kick 80, 10
	End If
	If ub5 = 1 Then
		UsedBall5.kick 80, 10
	End If

	If ub1 = 0 And ub2 = 0 And ub3 = 0 And ub4 = 0 And ub5 = 0 Then
		UsedBallTimer.Enabled = False
	End If
End Sub

Sub UsedBall1_Unhit
	ub1 = 0
End Sub

Sub DisplayDestroyer_Hit
	DisplayDestroyer.DestroyBall
End Sub