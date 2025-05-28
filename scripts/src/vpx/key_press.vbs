'*******************************************
'	ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal keycode)
	If keycode = StartGameKey Then
		Dim cred : cred = glf_machine_vars("credits").GetValue()
		If cred = 0 Then
			Exit Sub
		End If
	End If
	Glf_KeyDown(keycode)
	If keycode = PlungerKey Then 
		Plunger.Pullback
		SoundPlungerPull
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If KeyCode = PlungerKey Then
		Plunger.Fire
		SoundPlungerReleaseBall
	End If
	If KeyCode = AddCreditKey Then
		If AddCredit() Then
			RandomSoundCoin
			DOF 108, DOFPulse
		End If
	End If
	Glf_KeyUp(keycode)
End Sub

Function AddCredit()
	Dim cred
	cred = glf_machine_vars("credits").GetValue()
	If cred < 26 Then
		glf_machine_vars("credits").Value = cred + 1
		UpdateCreditReel
		AddCredit = True
	Else
		AddCredit = False
	End If
End Function

Sub UpdateCreditReel()
	Dim cred, ones, tens
	cred = glf_machine_vars("credits").GetValue()
	ones = cred mod 10
	tens = (cred - ones) / 10
	Controller.B2SSetCredits 27, ones
	Controller.B2SSetCredits 26, tens
End Sub
