'**********************************************************************************************************
' 	ZCRD:  Instruction Card Zoom
'**********************************************************************************************************

Dim CardCounter, ScoreCard

Sub CardTimer_Timer
	If scorecard = 1 Then
		CardCounter = CardCounter + 2
		If CardCounter > 50 Then CardCounter = 50
	Else
		CardCounter = CardCounter - 4
		If CardCounter < 0 Then CardCounter = 0
	End If
	InstructionCard.transX = CardCounter * 6
	InstructionCard.transY = CardCounter * 6
	InstructionCard.transZ =  - cardcounter * 2
	'   InstructionCard.objRotX = -cardcounter/2
	InstructionCard.size_x = 1 + CardCounter / 25
	InstructionCard.size_y = 1 + CardCounter / 25
	If CardCounter = 0 Then
		CardTimer.Enabled = False
		InstructionCard.visible = 0
	Else
		InstructionCard.visible = 1
	End If
End Sub

'**********************************************************************************************************
'***  Instruction Card Zoom
'**********************************************************************************************************
