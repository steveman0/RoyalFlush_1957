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

Function GetBallNear(obj)
	Dim b
	For b = 0 To UBound(gBOT)
		If abs(obj.X - gBOT(b).X) < 25 And abs(obj.Y - gBOT(b).Y) < 25 Then
			Set GetBallNear = gBOT(b)
		End If
	Next
End Function

' Sub UpdateTrough
	' UpdateTroughTimer.Interval = 300
	' UpdateTroughTimer.Enabled = 1
' End Sub

' Sub UpdateTroughTimer_Timer
	' If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	' If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	' If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	' If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	' Me.Enabled = 0
' End Sub

' ' DRAIN & RELEASE
' Sub Drain_Hit
	' BIP = BIP - 1
	' DMDBigText "DRAIN BLOCK",77,1
	' RandomSoundDrain Drain
	' vpmTimer.AddTimer 500, "Drain.kick 57, 20'"
' End Sub

' Sub SolRelease(enabled)
	' If enabled Then
		' BIP = BIP + 1
		' swTrough1.kick 90, 10
		' RandomSoundBallRelease swTrough1
	' End If
' End Sub
