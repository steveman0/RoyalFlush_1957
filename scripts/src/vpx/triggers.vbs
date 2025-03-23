'*******************************************
'	ZTRI: Triggers
'*******************************************

Sub Trigger1_Hit
	BIPL = True
End Sub

Sub Trigger1_UnHit
	BIPL = False
End Sub

Sub Spinner_Spin
	Addscore 110
	SoundSpinner Spinner
	Flash4 True 'Demo of the flasher
	vpmTimer.AddTimer 150,"Flash4 False'"   'Disable the flash after short time, just like a ROM would do
End Sub

Sub LeftInlane_Hit: leftInlaneSpeedLimit: End Sub
Sub RighttInlane_Hit: rightInlaneSpeedLimit: End Sub

' Inlane switch speedlimit code

Sub leftInlaneSpeedLimit
	'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
	if activeball.vely < 0 then exit sub 							'don't affect upwards movement
    activeball.AngMomZ = -abs(activeball.AngMomZ) * RndNum(3,6)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

    if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub


Sub rightInlaneSpeedLimit
	'Wylte's implementation
'    debug.print "Spin in: "& activeball.AngMomZ
'    debug.print "Speed in: "& activeball.vely
	if activeball.vely < 0 then exit sub 							'don't affect upwards movement
    activeball.AngMomZ = abs(activeball.AngMomZ) * RndNum(2,4)
    If abs(activeball.AngMomZ) > 60 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If abs(activeball.AngMomZ) > 80 Then activeball.AngMomZ = 0.8 * activeball.AngMomZ
    If activeball.AngMomZ > 100 Then activeball.AngMomZ = RndNum(80,100)
    If activeball.AngMomZ < -100 Then activeball.AngMomZ = RndNum(-80,-100)

	if abs(activeball.vely) > 5 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 10 then activeball.vely = 0.8 * activeball.vely
    if abs(activeball.vely) > 15 then activeball.vely = 0.8 * activeball.vely
    if activeball.vely > 16 then activeball.vely = RndNum(14,16)
    if activeball.vely < -16 then activeball.vely = RndNum(-14,-16)
'    debug.print "Spin out: "& activeball.AngMomZ
'    debug.print "Speed out: "& activeball.vely
End Sub

'*******************************************
'  Ramp Triggers
'*******************************************
Sub ramptrigger01_hit()
	WireRampOn True	 'Play Plastic Ramp Sound
End Sub

Sub ramptrigger02_hit()
	WireRampOff	 'Turn off the Plastic Ramp Sound
End Sub

Sub ramptrigger02_unhit()
	Addscore 10000
	WireRampOn False	'On Wire Ramp, Play Wire Ramp Sound
	queue.Add "flexJackpot", "ShowScene flexScenes(4), FlexDMD_RenderMode_DMD_RGB, 5", 5, 0, 0, 3000, 3000, False   'Jackpot animation becomes irrelevant if not played within 3 seconds
End Sub

Sub ramptrigger03_hit()
	WireRampOff	 'Exiting Wire Ramp Stop Playing Sound
End Sub

Sub ramptrigger03_unhit()
	RandomSoundRampStop ramptrigger03
End Sub
