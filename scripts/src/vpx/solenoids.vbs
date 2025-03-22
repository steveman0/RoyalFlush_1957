
' '*******************************************
' '  ZSOL : Other Solenoids
' '*******************************************



'  Kickers, Saucers
'*******************************************

Sub KickBall(kball, kangle, kvel, kvelz, kzlift)
	dim rangle
	rangle = PI * (kangle - 90) / 180
    
	kball.z = kball.z + kzlift
	kball.velz = kvelz
	kball.velx = cos(rangle)*kvel
	kball.vely = sin(rangle)*kvel
End Sub


Sub PlungerEjectCallback(ball)
	If s_Plunger.BallCntOver > 0 Then
		KickBall ball, 0, 60, 0, 0
		SoundSaucerKick 1, s_Plunger
	Else
		SoundSaucerKick 0, s_Plunger
	End If
End Sub