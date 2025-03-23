'****************************************************************
'	ZGII: GI
'****************************************************************

Dim gilvl   'General Illumination light state tracked for Ball Shadows
gilvl = 1

Sub ToggleGI(Enabled)
	Dim xx
	If enabled Then
		For Each xx In GI
			xx.state = 1
		Next
		PFShadowsGION.visible = 1
		gilvl = 1
	Else
		For Each xx In GI
			xx.state = 0
		Next
		PFShadowsGION.visible = 0
		GITimer.enabled = True
		gilvl = 0
	End If
	Sound_GI_Relay enabled, bumper1
End Sub

Sub GITimer_Timer()
	Me.enabled = False
	ToggleGI 1
End Sub
