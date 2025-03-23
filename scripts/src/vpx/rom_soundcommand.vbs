' Not ROM-based maybe unnecessary? Will remove if no errors...
' '********************************************************
' '	ZLIS: ROM SoundCommand Listener
' '********************************************************
' 'Originally by ...?
' 'LotR code by Apophis
' '2-in-1 sub by Wylte

' 'Examples from Robot (simple If-Then use) and LotR (complicated state combinations)

' ' *** Sound Commands ***

' 'It's good to note them here for reference

' 'In Robot
' ' D4,D4,54,	- Challenge (New Game)
' ' D2,D2,52,	- Mission Completed (End Game)
' ' C4,C4,44,	- Whoa Whoa (Tilt)
' ' C3,C3,43,	- Ow (Tilt)
' ' E3,E3,63,	- New Ball

' 'In LotR
' ' FD1D = start of destroy the ring mode
' ' FD2C = successfully completed DTR
' ' FD0B = Ring wraithes mission 1
' ' FD0C = Ring wraithes mission 2
' ' FD0D = Ring wraithes mission 3
' ' FD09 = shelob mission starts
' ' FD04 = end of ball
' ' FD03 = main theme 2 (not in a mode)
' ' FD02 = main theme 1 (not in a mode)

' ' *** Constants ***

' ' The commands from the ROM come through in hex
' ' so add constants for the numbers you're listening for

' 'Const hexFD = 253
' 'Const hexE3 = 227
' 'Const hexD4 = 212
' 'Const hexD2 = 210
' 'Const hexC3 = 196
' 'Const hexC4 = 195
' 'Const hex7F = 127
' 'Const hex54 = 84
' 'Const hex52 = 82
' 'Const hex2C = 44
' 'Const hex1D = 29
' 'Const hex0D = 13
' 'Const hex0C = 12
' 'Const hex0B = 11
' 'Const hex09 = 9
' 'Const hex08 = 8
' 'Const hex04 = 4
' 'Const hex03 = 3
' 'Const hex02 = 2

' Dim SndCmdStr
' Dim LastSnd
' LastSnd = 0

' ' *** Listener ***

' ' WHAT YOU NEED:

' ' Add the SoundCmdListener call to a fast timer that always runs
' ' Copy in the 10 "Listener###" textboxes from the backglass view
' ' Play the game, writing down the values for the event you need
' ' The event you want may have multiple sounds including sounds played at other times
' ' If unsure, you can note the values and watch to see if they are used elsewhere
' ' You may want to record your screen if it's moving fast

' Sub SoundCmdListener
	' Dim NewSounds,ii,Snd
	' NewSounds = Controller.NewSoundCommands	 ' Listening to the ROM
	' If Not IsEmpty(NewSounds) Then
		' '* Finding the Values *
		' ' Comment out or delete this part once you know the commands
		' ' And delete the text boxes
		' SndCmdStr = ""	  ' Empty the string
		' For ii = 0 To UBound(NewSounds)	 ' Setting the new to a string to display
			' Snd = NewSounds(ii,0)
			' If Snd <> 0 And Snd <> 255 Then
				' SndCmdStr = SndCmdStr & Hex(Snd) & ","
			' End If
		' Next
		
		' Listener010.Text = Listener009.Text				 ' Display sound commands in textboxes
		' Listener009.Text = Listener008.Text				 ' ^^^
		' Listener008.Text = Listener007.Text				 ' ^^^
		' Listener007.Text = Listener006.Text				 ' ^^^
		' Listener006.Text = Listener005.Text				 ' ^^^
		' Listener005.Text = Listener004.Text				 ' ^^^
		' Listener004.Text = Listener003.Text				 ' ^^^
		' Listener003.Text = Listener002.Text				 ' ^^^
		' Listener002.Text = Listener001.Text				 ' ^^^
		' Listener001.Text = SndCmdStr						' Arrange them vertically so the sound commands scroll up the screen
		
		' '* Using them to do stuff *
		
		' For ii = 0 To UBound(NewSounds) 'Listen for specific commands; if a specified combination occurs, do a thing
			' Snd = NewSounds(ii,0)
			' 'From Robot:
			' '			If LastSnd = hexD2 And Snd = hexD2 Then													'End Game, Flippers Off and Dropped
			' '				FlippersDisabled = True
			' '				LeftFlipper.RotateToStart
			' '				RightFlipper.RotateToStart
			' '			End If
			' 'End Robot
			
			' 'From LotR:
			' '			If LastSnd = hexFD Then
			' '				Select Case Snd
			' '					Case hex1D: ClearSndFlags : Snd_DTRstart = True : TORcnt = 1 : TheOneRingUpdates.Enabled=True ': debug.print "Snd_DTRstart = True"
			' '					Case hex2C: ClearSndFlags : Snd_DTRfinal = True ': debug.print "Snd_DTRfinal = True"
			' '					Case hex0D: ClearSndFlags : Snd_Wraithes = True ': debug.print "hex0D: Snd_Wraithes = True"
			' '					Case hex0C: ClearSndFlags : Snd_Wraithes = True ': debug.print "hex0C: Snd_Wraithes = True"
			' '					Case hex0B: ClearSndFlags : Snd_Wraithes = True ': debug.print "hex0B: Snd_Wraithes = True"
			' '					Case hex08: ClearSndFlags ': debug.print "hex08: Snd_Wraithes = False"
			' '					Case hex09: ClearSndFlags : Snd_Shelob = True : Setlamp 102,1 ': debug.print "Snd_Shelob = True"
			' '					Case hex04: ClearSndFlags : Snd_EOB = True ': debug.print "Snd_EOB = True"
			' '					Case hex03: ClearSndFlags : Snd_Main = True ': debug.print "Snd_Main = True"
			' '					Case hex02: ClearSndFlags : Snd_Main = True ': debug.print "Snd_Main = True"
			' '				End Select
			' '			End If
			' 'End LotR
			
			' LastSnd = Snd	   ' Remembering the previous bit of hex
			' '			debug.print "LastSnd: " & LastSnd
		' Next
	' End If
' End Sub

' 'LotR Specific code (here for reference)

' 'Sub ClearSndFlags
' '	'debug.print "ClearSndFlags"
' '	Snd_DTRstart = False
' '	Snd_DTRfinal = False
' '	Snd_Shelob = False
' '	Snd_Wraithes = False
' '	Snd_EOB = False
' '	Snd_Main = False
' '
' '	Setlamp 102,0
' 'End Sub

' '********************************************************
' '  END ROM SoundCommand Listener
' '********************************************************
