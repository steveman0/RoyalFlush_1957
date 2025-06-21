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

	Controller.B2SSetData (10 + k / 10000 - 1), DOFOff
	Controller.B2SSetData (20 + kk / 100000 - 1), DOFOff
	Controller.B2SSetData (30 + m / 1000000 - 1), DOFOff
	
	' Disable 9's on rollover case
	If k Mod 100000 = 0 Then 
		Controller.B2SSetData 19, DOFOff
	End If
	If kk Mod 1000000 = 0 Then
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

' Responsible for interim rendering of roto target visuals
Function UpdateTargetWalls(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If     

	' Roto target index mapping
        ' 1: Joker
        ' 2: Ace
        ' 3: King
        ' 4: 10
        ' 5: Joker
        ' 6: Ace
        ' 7: King
        ' 8: Queen
        ' 9: 10
        ' 10: Jack
        ' 11: Queen
        ' 12: Joker
        ' 13: Ace
        ' 14: King
        ' 15: Jack

	Dim i
	i = glf_machine_vars("roto_index").GetValue()
	'Debug.Print "Update visuals for roto index " & i
	Select case i
		case 1 : VisibleTarget.Image = "Joker"
		case 2 : VisibleTarget.Image = "Ace"
		case 3 : VisibleTarget.Image = "King"
		case 4 : VisibleTarget.Image = "Ten"
		case 5 : VisibleTarget.Image = "Joker"
		case 6 : VisibleTarget.Image = "Ace"
		case 7 : VisibleTarget.Image = "King"
		case 8 : VisibleTarget.Image = "Queen"
		case 9 : VisibleTarget.Image = "Ten"
		case 10 : VisibleTarget.Image = "Jack"
		case 11 : VisibleTarget.Image = "Queen"
		case 12 : VisibleTarget.Image = "Joker"
		case 13 : VisibleTarget.Image = "Ace"
		case 14 : VisibleTarget.Image = "King"
		case 15 : VisibleTarget.Image = "Jack"
	End Select

	' Update rotational angle target of the real roto
	RotoTargetAngle = (182 + (i - 1) * 24) Mod 360

	'Keep this return call.
	If IsObject(args(1)) Then
		Set UpdateTargetWalls= kwargs
    Else
        UpdateTargetWalls= kwargs
    End If
End Function