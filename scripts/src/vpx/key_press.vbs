'*******************************************
'	ZKEY: Key Press Handling
'*******************************************

Sub Table1_KeyDown(ByVal keycode)
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
	Glf_KeyUp(keycode)
End Sub
