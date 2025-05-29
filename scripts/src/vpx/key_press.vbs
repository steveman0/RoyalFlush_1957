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
		UpdateCreditReelDirect
		AddCredit = True
	Else
		AddCredit = False
	End If
End Function

Function SubtractCredit(args)
	Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If     


    Dim cred
	cred = glf_machine_vars("credits").GetValue()
	If cred > 0 Then
		glf_machine_vars("credits").Value = cred - 1
		UpdateCreditReelDirect
	End If
    
	'Keep this return call.
	If IsObject(args(1)) Then
		Set SubtractCredit= kwargs
    Else
        SubtractCredit= kwargs
    End If
End Function

Function UpdateCreditReel(args)
	Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If  

	UpdateCreditReelDirect

	'Keep this return call.
	If IsObject(args(1)) Then
		Set UpdateCreditReel= kwargs
    Else
        UpdateCreditReel= kwargs
    End If
End Function

Sub UpdateCreditReelDirect()
	Dim cred, ones, tens
	cred = glf_machine_vars("credits").GetValue()
	ones = cred mod 10
	tens = (cred - ones) / 10
	Controller.B2SSetCredits 27, ones
	Controller.B2SSetCredits 26, tens
End Sub
