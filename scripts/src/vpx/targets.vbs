'********************************************
'	ZTAR: Targets
'********************************************

Sub sw11_Hit
	STHit 11
End Sub

Sub sw11o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw12_Hit
	STHit 12
End Sub

Sub sw12o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw13_Hit
	STHit 13
End Sub

Sub sw13o_Hit
	TargetBouncer ActiveBall, 1
End Sub

'********************************************
'  Drop Target Controls
'********************************************

' Drop targets
Sub sw1_Hit
	DTHit 1
End Sub

Sub sw2_Hit
	DTHit 2
End Sub

Sub sw3_Hit
	DTHit 3
End Sub

' If the drop targets can be reset individually, use specific solenoid subs for each like below
' These subroutines would be called by the solenoid callbacks if using a ROM

Sub SolDT1(enabled) ' Drop Target 1 Solenoid
	If enabled Then
		RandomSoundDropTargetReset sw1p
		DTRaise 1
	End If
End Sub

Sub SolDT2(enabled) ' Drop Target 2 Solenoid
	If enabled Then
		RandomSoundDropTargetReset  sw2p
		DTRaise 2
	End If
End Sub

Sub SolDT3(enabled) ' Drop Target 3 Solenoid
	If enabled Then
		RandomSoundDropTargetReset  sw3p
		DTRaise 3
	End If
End Sub

' If a whole bank of drop targets can be reset at once, use sub like below

Sub SolDTBank123(enabled)
	Dim xx
	If enabled Then
		RandomSoundDropTargetReset sw2p
		DTRaise 1
		DTRaise 2
		DTRaise 3
		For Each xx In ShadowDT
			xx.visible = True
		Next
	End If
End Sub
