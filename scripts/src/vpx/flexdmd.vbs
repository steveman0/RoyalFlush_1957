'*******************************************
'  ZDMD: FlexDMD
'*******************************************
'
' DMDTimer @17ms
' StartFlex with the intro @ Table1_Init
' Startgame ( KeyDown : PlungerKey ) calls the Game DMD to start if intro is on
' Make a copy of VPWExampleTableDMD folder, rename and paste into Visual Pinball\Tables\"InsertTableNameDMD"
' Update .ProjectFolder = ".\VPWExampleTableDMD\" to DMD folder name from previous step
' Update DMDTimer_Timer Sub to allow DMD to update remaining balls and credit: find ("Ball") and ("Credit")

' Commands :
'	 DMDBigText "LAUNCH",77,1  : display text instead of score : "text",frames(x17ms),effect  0=solid 1=blink
'	DMDBGFlash=15 : will light up background with new image for xx frames
'	DMDFire=Flexframe+50  : will animate the font for 50 frames
'
' For another demo, see the following from the FlexDMD developer:
'   https://github.com/vbousquet/flexdmd/tree/master/FlexDemo


Dim FlexDMD	 'This is the FlexDMD object
Dim FlexMode	'This is use for specifying the state of the DMD
Dim FlexFrame   'This is the current Frame count. It increments every time DMDTimer_Timer is run

Const FlexDMD_RenderMode_DMD_GRAY = 0, _
FlexDMD_RenderMode_DMD_GRAY_4 = 1, _
FlexDMD_RenderMode_DMD_RGB = 2, _
FlexDMD_RenderMode_SEG_2x16Alpha = 3, _
FlexDMD_RenderMode_SEG_2x20Alpha = 4, _
FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num = 5, _
FlexDMD_RenderMode_SEG_2x7Alpha_2x7Num_4x1Num = 6, _
FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num = 7, _
FlexDMD_RenderMode_SEG_2x7Num_2x7Num_10x1Num = 8, _
FlexDMD_RenderMode_SEG_2x7Num_2x7Num_4x1Num_gen7 = 9, _
FlexDMD_RenderMode_SEG_2x7Num10_2x7Num10_4x1Num = 10, _
FlexDMD_RenderMode_SEG_2x6Num_2x6Num_4x1Num = 11, _
FlexDMD_RenderMode_SEG_2x6Num10_2x6Num10_4x1Num = 12, _
FlexDMD_RenderMode_SEG_4x7Num10 = 13, _
FlexDMD_RenderMode_SEG_6x4Num_4x1Num = 14, _
FlexDMD_RenderMode_SEG_2x7Num_4x1Num_1x16Alpha = 15, _
FlexDMD_RenderMode_SEG_1x16Alpha_1x16Num_1x7Num = 16

Const FlexDMD_Align_TopLeft = 0, _
FlexDMD_Align_Top = 1, _
FlexDMD_Align_TopRight = 2, _
FlexDMD_Align_Left = 3, _
FlexDMD_Align_Center = 4, _
FlexDMD_Align_Right = 5, _
FlexDMD_Align_BottomLeft = 6, _
FlexDMD_Align_Bottom = 7, _
FlexDMD_Align_BottomRight = 8


Dim FontScoreInactive
Dim FontScoreActive
Dim FontBig1
Dim FontBig2
Dim FontBig3

Sub Flex_Init
	If UseFlexDMD = 0 Then Exit Sub
	Set FlexDMD = CreateObject("FlexDMD.FlexDMD")
	If FlexDMD Is Nothing Then
		MsgBox "No FlexDMD found. This table will Not run without it."
		Exit Sub
	End If
	SetLocale(1033)
	With FlexDMD
		.GameName = cGameName
		.TableFile = Table1.Filename & ".vpx"
		.Color = RGB(255, 88, 32)
		.RenderMode = FlexDMD_RenderMode_DMD_RGB
		.Width = 128
		.Height = 32
		.ProjectFolder = "./VPWExampleTableDMD/"
		.Clear = True
		.Run = True
	End With
	
	'----- Build flex scenes -----
	
	' Stern Score Scene
	Dim i
	Set FontScoreActive = FlexDMD.NewFont("TeenyTinyPixls5.fnt", vbWhite, vbWhite, 0)
	Set FontScoreInactive = FlexDMD.NewFont("TeenyTinyPixls5.fnt", RGB(100, 100, 100), vbWhite, 0)
	Set FontBig1 = FlexDMD.NewFont("sys80.fnt", vbWhite, vbBlack, 0)
	Set FontBig2 = FlexDMD.NewFont("sys80_1.fnt", vbWhite, vbBlack, 0)
	Set FontBig3 = FlexDMD.NewFont("sys80.fnt", RGB ( 10,10,10) ,vbBlack, 0)
	Set FlexScenes(0) = FlexDMD.NewGroup("Score")
	With FlexScenes(0)
		.AddActor FlexDMD.NewImage("bg","bgdarker.png")
		.Getimage("bg").visible = True ' False
		.AddActor FlexDMD.NewImage("bg2","bg.png")
		.Getimage("bg2").visible = False
		For i = 1 To 4
			.AddActor FlexDMD.NewLabel("Score_" & i, FontScoreInactive, "0")
		Next
		.AddActor FlexDMD.NewFrame("VSeparator")
		.GetFrame("VSeparator").Thickness = 1
		.GetFrame("VSeparator").SetBounds 45, 0, 1, 32
		.AddActor FlexDMD.NewGroup("Content")
		.GetGroup("Content").Clip = True
		.GetGroup("Content").SetBounds 47, 0, 81, 32
	End With
	Dim title
	Set title = FlexDMD.NewLabel("TitleScroller", FontScoreActive, ">>> Flex DMD <<<")
	Dim af
	Set af = title.ActionFactory
	Dim list
	Set list = af.Sequence()
	list.Add af.MoveTo(128, 2, 0)
	list.Add af.Wait(0.5)
	list.Add af.MoveTo( - 128, 2, 5.0)
	list.Add af.Wait(3.0)
	title.AddAction af.Repeat(list, - 1)
	FlexScenes(0).GetGroup("Content").AddActor title
	Set title = FlexDMD.NewLabel("Title2", FontBig3, " ")
	title.SetAlignedPosition 42, 16, FlexDMD_Align_Center
	FlexScenes(0).GetGroup("Content").AddActor title
	Set title = FlexDMD.NewLabel("Title", FontBig1, " ")
	title.SetAlignedPosition 42, 16, FlexDMD_Align_Center
	FlexScenes(0).GetGroup("Content").AddActor title
	FlexScenes(0).GetGroup("Content").AddActor FlexDMD.NewLabel("Ball", FontScoreActive, "Ball 1")
	FlexScenes(0).GetGroup("Content").AddActor FlexDMD.NewLabel("Credit", FontScoreActive, "Credit 5")
	
	' Welcome animation
	Set FlexScenes(1) = FlexDMD.NewGroup("Welcome")
	With FlexScenes(1)
		.AddActor FlexDMD.Newvideo ("test","spinner.gif")
		.Getvideo("test").visible = True
		.AddActor FlexDMD.NewImage("logo","VPWLogo32.png")
		.Getimage("logo").visible = False
	End With
	
	' Bonus X animation
	Set FlexScenes(2) = FlexDMD.NewGroup("BonusX")
	FlexScenes(2).AddActor FlexDMD.Newvideo ("bonusX","bonusx.gif")
	
	' Multiball
	Set FlexScenes(3) = FlexDMD.NewGroup("multiball")
	FlexScenes(3).AddActor FlexDMD.Newvideo ("multiball","multiball.gif")
	
	' Jackpot
	Set FlexScenes(4) = FlexDMD.NewGroup("jackpot")
	FlexScenes(4).AddActor FlexDMD.Newvideo ("jackpot","jackpot.gif")
	
	' Drop target bonus
	Set FlexScenes(5) = FlexDMD.NewGroup("nobonus")
	FlexScenes(5).AddActor FlexDMD.Newvideo ("nobonus","nobonus.gif")
	Set FlexScenes(6) = FlexDMD.NewGroup("bonus1")
	FlexScenes(6).AddActor FlexDMD.Newvideo ("bonus1","bonus1.gif")
	Set FlexScenes(7) = FlexDMD.NewGroup("bonus2")
	FlexScenes(7).AddActor FlexDMD.Newvideo ("bonus2","bonus2.gif")
	Set FlexScenes(8) = FlexDMD.NewGroup("bonus3")
	FlexScenes(8).AddActor FlexDMD.Newvideo ("bonus3","bonus3.gif")
End Sub

'--------------------------------------------
' Easy wrapper to play a FlexDMD scene
'
' flexScene: FlexDMD.Group to play
' render: The render mode to use
' mode: The mode number used in the DMDTimer
'--------------------------------------------
Sub ShowScene(flexScene, render, mode) 'Easy wrapper to play a FlexDMD scene
	If UseFlexDMD = 0 Then Exit Sub
	
	FlexDMD.LockRenderThread
	FlexDMD.RenderMode = render
	FlexDMD.Stage.RemoveAll
	FlexDMD.Stage.AddActor flexScene
	If VRroom > 0 Or FlexONPlayfield Then FlexDMD.Show = False Else FlexDMD.Show = True
	FlexDMD.UnlockRenderThread
	
	FlexMode = mode
End Sub

Dim DMDTextOnScore
Dim DMDTextDisplayTime
Dim DMDTextEffect
Sub DMDBigText(text,Time,effect)
	If UseFlexDMD = 0 Then Exit Sub
	DMDTextOnScore = text
	DMDTextDisplayTime = FLEXframe + Time
	
	DMDTextEffect = effect
End Sub

Sub FlexFlasher 'Flex on vrroom and playfield runs this one
	Dim DMDp
	DMDp = FlexDMD.DMDColoredPixels
	If Not IsEmpty(DMDp) Then
		DMDWidth = FlexDMD.Width
		DMDHeight = FlexDMD.Height
		DMDColoredPixels = DMDp
	End If
End Sub

Dim DMDFire
Dim DMDBGFlash
Sub DMDTimer_Timer 'Main FlexDMD Timer
	If UseFlexDMD = 0 Then Exit Sub
	If VRroom > 0 Or FlexONPlayfield Then FlexFlasher
	
	Dim i, n, x, y, label
	FlexFrame = FlexFrame + 1
	FlexDMD.LockRenderThread
	
	Select Case FlexMode
		Case 1 'scoreboard
			'If (FlexFrame Mod 64) = 0 Then CurrentPlayer = 1 + (CurrentPlayer Mod 4)
			If (FlexFrame Mod 16) = 0 Then
				For i = 1 To 4
					Set label = FlexDMD.Stage.GetLabel("Score_" & i)
					If i = CurrentPlayer Then
						label.Font = FontScoreActive
					Else
						label.Font = FontScoreInactive
					End If
					label.Text = FormatNumber(PlayerScore(i), 0)
					label.SetAlignedPosition 45, 1 + (i - 1) * 6, FlexDMD_Align_TopRight
				Next
			End If
			
			If DMDBGFlash > 0 Then
				DMDBGFlash = DMDBGFlash - 1
				FlexDMD.Stage.GetImage("bg2").visible = True
			Else
				FlexDMD.Stage.GetImage("bg2").visible = False
			End If
			
			If DMDfire > FLEXframe And (FlexFrame Mod 8) > 3 Then
				FlexDMD.Stage.GetLabel("Title").font = FontBig2
			Else
				FlexDMD.Stage.GetLabel("Title").font = FontBig1
			End If
			
			If DMDTextDisplayTime > FLEXframe Then
				If DMDTextEffect = 1 And (FLEXframe Mod 20) > 10 Then
					FlexDMD.Stage.GetLabel("Title").Text = " "
					FlexDMD.Stage.GetLabel("Title2").Text = " "
				Else
					FlexDMD.Stage.GetLabel("Title").Text = DMDTextOnScore
					FlexDMD.Stage.GetLabel("Title2").Text = DMDTextOnScore
				End If
			Else
				FlexDMD.Stage.GetLabel("Title").Text = FormatNumber(PlayerScore(CurrentPlayer), 0)
				FlexDMD.Stage.GetLabel("Title2").Text = FormatNumber(PlayerScore(CurrentPlayer), 0)
			End If
			
			FlexDMD.Stage.GetLabel("Title").SetAlignedPosition 42, 16, FlexDMD_Align_Center
			FlexDMD.Stage.GetLabel("Title2").SetAlignedPosition 43, 17, FlexDMD_Align_Center
			FlexDMD.Stage.GetLabel("Ball").SetAlignedPosition 0, 33, FlexDMD_Align_BottomLeft
			FlexDMD.Stage.GetLabel("Credit").SetAlignedPosition 81, 33, FlexDMD_Align_BottomRight
			'Update with your own code for BallsRemaining
			'   FlexDMD.Stage.GetLabel("Ball").Text = "Ball " & 4 - BallsRemaining(CurrentPlayer)
			'Update with your own code for Credits
			'   FlexDMD.Stage.GetLabel("Credit").Text = "Credit " & (Credits(CurrentPlayer)) - 1
			
		Case 2 'dmdintro
			If FlexFrame = 88 Then FlexDMD.Stage.Getimage("logo").visible = True
			
			If FlexFrame > 110 Then
				If (FlexFrame Mod 32) = 10 Then FlexDMD.Stage.Getimage("logo").visible = True
				If (FlexFrame Mod 32) = 1 Then FlexDMD.Stage.Getimage("logo").visible = False
			End If
	End Select
	FlexDMD.UnlockRenderThread
End Sub
