'*******************************************
'	ZSCR: Scoring
'*******************************************

Sub Addscore (value)
	PlayerScore(CurrentPlayer) = PlayerScore(CurrentPlayer) + value
	If value > 99999 Then DMDBGFlash = 15
	If value > 499999 Then DMDFire = Flexframe + 50
	
	' Add chimes based on score amount here
End Sub
