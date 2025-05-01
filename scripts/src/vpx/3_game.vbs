'*******************************************
'	ZGAM: Non-GLF gameplay events
'*******************************************

Sub TriggerLaneL_Hit
	DispatchPinEvent "score_100k", ActiveBall
End Sub

Sub TriggerLaneR_Hit
	DispatchPinEvent "score_100k", ActiveBall
End Sub

Sub RubberBand_Hit
	DispatchPinEvent "score_10k", ActiveBall
End Sub

Sub RubberBand001_Hit
	DispatchPinEvent "score_10k", ActiveBall
End Sub

Sub RubberBand002_Hit
	DispatchPinEvent "score_10k", ActiveBall
End Sub

Sub RubberBand003_Hit
	DispatchPinEvent "score_10k", ActiveBall
End Sub

Sub RubberBand010_Hit
	DispatchPinEvent "score_10k", ActiveBall
End Sub

Sub RubberBand011_Hit
	DispatchPinEvent "score_10k", ActiveBall
End Sub

' Responsible for B2S lighting states for score
Function UpdateBackglassScore(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If     

	Dim s
	Dim k
	Dim kk
	Dim m
	s = GetPlayerStateForPlayer("0", "score")
	k = s Mod 100000
	kk = s Mod 1000000 - k
	m = s - (s Mod 1000000)

	' Debug.Print "k on val " & (10 + k / 10000)
	' Debug.Print "k off val " & (10 + k / 10000 - 1)
	' Debug.Print "kk on val " & (20 + kk / 100000)
	' Debug.Print "kk off val " & (20 + kk / 100000 - 1)
	' Debug.Print "m on val " & (30 + m / 1000000)
	' Debug.Print "m off val " & (30 + m / 1000000 - 1)

	Controller.B2SSetData (10 + k / 10000 - 1), DOFOff
	Controller.B2SSetData (20 + kk / 100000 - 1), DOFOff
	Controller.B2SSetData (30 + m / 1000000 - 1), DOFOff
	
	' Disable 9's on rollover case
	If k Mod 100000 = 0 Then 
		Debug.Print "Turning 19 off"
		Controller.B2SSetData 19, DOFOff
	End If
	If kk Mod 1000000 = 0 Then
		Debug.Print "Turning 29 off"
		Controller.B2SSetData 29, DOFOff
	End If

	Controller.B2SSetData (10 + k / 10000), DOFOn
    Controller.B2SSetData (20 + kk / 100000), DOFOn
	Controller.B2SSetData (30 + m / 1000000), DOFOn

	'Keep this return call.
	If IsObject(args(1)) Then
		Set UpdateBackglassScore= kwargs
    Else
        UpdateBackglassScore= kwargs
    End If
End Function