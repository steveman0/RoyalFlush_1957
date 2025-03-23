'*******************************************
'	ZSCR: Scoring
'*******************************************

Sub Addscore (value)
	PlayerScore(CurrentPlayer) = PlayerScore(CurrentPlayer) + value
	If value > 199 Then DMDBGFlash = 15
	If value > 4999 Then DMDFire = Flexframe + 50
End Sub
