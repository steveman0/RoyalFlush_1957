' __      _________          __  ______                           _        _______    _     _
' \ \    / /  __ \ \        / / |  ____|                         | |      |__   __|  | |   | |
'  \ \  / /| |__) \ \  /\  / /  | |__  __  ____ _ _ __ ___  _ __ | | ___     | | __ _| |__ | | ___
'   \ \/ / |  ___/ \ \/  \/ /   |  __| \ \/ / _` | '_ ` _ \| '_ \| |/ _ \    | |/ _` | '_ \| |/ _ \
'    \  /  | |      \  /\  /    | |____ >  < (_| | | | | | | |_) | |  __/    | | (_| | |_) | |  __/
'     \/   |_|       \/  \/     |______/_/\_\__,_|_| |_| |_| .__/|_|\___|    |_|\__,_|_.__/|_|\___|   V1.6
'                                                          | |
'                                                          |_|
'
'
'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
'	ZTUT: VPW Tutorial Videos
'	ZOPT: User Options
'	ZCON: Constants and Global Variables
'	ZDMD: FlexDMD
'	ZTIM: Timers
'	ZINI: Table Initialization and Exiting
' 	ZMAT: General Math Functions
'	ZANI: MISC ANIMATIONS
'	ZDRN: Drain, Trough, and Ball Release
'	ZSCR: Scoring
'	ZKEY: Key Press Handling
'	ZFLP: Flippers
'	ZBMP: Bumpers
'	ZGII: GI
'	ZSLG: Slingshots
'	ZKIC: Kickers, Saucers
' 	ZTRI: Triggers
'	ZTAR: Targets
'	ZSOL: Other Solenoids
'	ZLIS: ROM SoundCommand Listener
'	ZSHA: AMBIENT BALL SHADOWS
'	ZPHY: GNEREAL ADVICE ON PHYSICS
'	ZNFF: FLIPPER CORRECTIONS
' 	ZDMP: RUBBER DAMPENERS
' 	ZBOU: VPW TargetBouncer for targets and posts
'	ZSSC: SLINGSHOT CORRECTION
' 	ZRDT: DROP TARGETS
'	ZRST: STAND-UP TARGETS
' 	ZBRL: BALL ROLLING AND DROP SOUNDS
' 	ZRRL: RAMP ROLLING SFX
' 	ZFLE: FLEEP MECHANICAL SOUNDS
' 	Z3DI: 3D INSERTS
' 	ZFLD: FLUPPER DOMES
' 	ZFLB: FLUPPER BUMPERS
' 	ZTST: Debug Shot Tester
' 	ZQUE: VPIN WORKSHOP ADVANCED QUEUING SYSTEM
'  	ZLOG: ERROR LOGS
' 	ZCRD: Instruction Card Zoom
' 	ZVRR: VR Room / VR Cabinet
'
'*********************************************************************************************************************************

' CONTENT CONTRIBUTION CREDITS
' Dynamic Ball Shadows: iaakki, apophis, Wylte
' Rubberizer: iaakki
' Target Bouncer: iaakki, wrd1972, apophis
' Flipper and physics corrections: nFozzy, Rothbauerw
' Sound effects package: Fleep
' Ramp rolling sounds: nFozzy
' Lampz: nFozzy
' Bumpers: Flupper
' Flasher domes: Flupper
' 3D inserts: Flupper, Benji
' Drop targets: Rothbauerw
' VR Cabinet & Room: Sixtoe, Flupper, 3rdaxis
' FlexDMD: oqqsan
' Queuing System & Coding Standard: Arelyel
' Error logs: baldgeek
'
'
' NOTES. PLEASE READ
'
' - This table contains many of the tools, resources, and best practices that have been developed by the VPX community over the years.
'   These techniques continue to be developed, so this table attempts to capture their latest versions.
' - Most VPX projects are recreations and use a ROM to control the game. However, this example table does not use a ROM so in some cases
'   it does not fully describe how to implement features that interact with the ROM.
' - This table provides an example of a "physical trough" where balls are only created once at table initialization and never destroyed.
'   This approach is different than traditional, but it is considered a best practice by VPW as it forces the table to behave more realistically.
'   Part of the implementation of this includes the use of a global ball array called "gBOT". If you want to use the more typical scripting
'   techniques that allow for ball destruction, you will want to replace all instances of "gBOT" with "BOT" and use "GetBalls" calls where needed.
'
'
' CHANGE LOG
'001 apophis -	Included: nFozzy flippers, dampeners, and physics; Fleep sound package; Roth drop targets; VPW dynamic shadows; VPW rubberizer and targetbouncer
'002 Sixtoe -	Added basic VR cabinet and room and associated scripting, primitive sideblades, primitive lockdown and siderails, ocd tidied up some bits, unhidden screws, added slingshot animations, added low poly versions of bolt prims, added locknut prims to slingshot plastics,
'003 fluffhead35 - Added Ramp Rolling SFX subroutine by nFozzy
'004 iaakki -	very simple nFozzy Lampz light fading example added
'005 apophis -	Updated ramp roll sounds so volume is function of ballspeed
'006 fluffhead35 - Removed BallPitch subroutine as not used anymore after 005 was applied
'007 apophis -	Added Flupper Dome flashers example
'008 Wylte -	Added explanations for Dynamic Ball Shadow code, deprecated RtxBScnt and CntDwn code (superceded by currentShadowCount), renamed collection to DynamicSources (working on removing rtx from most places)
'009 apophis -	Added flasherblooms, and updated Flupper dome installation instructions
'010 apophis -	Added 3D inserts examples. Added more installation instructions throughout.
'011 apophis -	Added Flupper bumpers example
'012 iaakki -	some insert adjustments
'013 apophis -	Added drop target shadows example. Added drop target textures.
'014 apophis -	Fixed some layer assignments. Changed the pincab and backbox textures.
'015 apophis -	Added lightning shaped 3D inserts. Scaled bumpers to be effectively 1.5 inches tall (thanks gtxjoe). Fixed Csng error in some Fleep functions (thanks iaakki).
'016 gtxjoe -	Added debug shot tester.  Press 2 to block drain and outlanes. Press and hold a debug key (W,E,R,Y,U,I,P,A,S,F,G) to capture ball. Release key to shoot ball
'017 apophis -	Adjusted blue and purple insert materials opacity from 0.6 to 0.99. Was not being rendered well on some systems.
'018 apophis -	Added orange Flupper dome flasher textures and made flasherbloom colors initialize correctly. Minor tweaks to ball shadow sub. Updated TargetBouncer to conserver energy.
'019 apophis -	Fixed new TargetBouncer
'020 apophis -	Added a second rubberizer option. Fixed divide by zero issue in new TargetBouncer
'021 Wylte -	Small update to shadow instructions, added z code back in + instructions (commented by default), added required functions to segment (commented), added toggle for ambient shadow
'022 apophis -	Added multiple target bouncer options
'023 oqqsan -	Added Flex DMD intro and scoring with some effects
'024 apophis -	Some final table and physics tweaks before release. Added new backglass image. Updated Lampz comments.
'025 apophis -	Rubberizer logic updated
'026 Devious626 - Added player options for FlexDMD and a some extra information in instructions
'1.0 apophis -	Final script formatting adjustments made for release.
'1.01 oqqsan -	added my rules magnifier, such a goonie before i rememberd to get it on here
'1.02 oqqsan -	added Flex Flasher DMD only when in vrroom
'1.03 apophis -	Added UsingROM option and removed b2son variable. Added some notes to the Dynamic Shadow instructions.
'1.04 Wylte -	Shadow code optimized...then complicated again (still should be better though), flasher shadows added, more ambient shadow options given, changelog formatted to give me less of a headache (damn you long-named people!)
'1.04b Wylte -	Dynamic Ball Shadow instructions overhauled
'1.05 fluffhead35 - Updating RampRoll Sounds code as well well as adding BallRolling and RampRolling Amplification Sounds
'1.06 fluffhead35 - Updating RampRoll Sound Logic after discussions with apophis
'1.07 apophis -	Updated drop target code with help from rothbauerw (thanks!) to handle collisions with ball when target is raised.
'1.1 apophis -	Initialized LFState and RFState variables. Ready for release.
'1.11 rothbauerw - Bug fix to drop target raise animation, added Fleep drop target and relay sounds, added GI On/Off when hitting bumper 1,
'	 updated flasher domes to automatically also place flasher, light, and assign correct bloom images, fixed flipper nudge bug, added stand up target code (similar to drop targets),
'	 added DTAction and STAction to correctly account for target hits in non-ROM based tables, updated rubberizer to most recent code, adjusted target bouncer to work when
'	 hitting secondary drop target and stand up target objects instead of primary objects.
'1.12 iaakki -  DynamicShadows to use arrays to improve performance. Merging Roth's implementation. Fixed lamptimer so fading speed is not linked to fps.
'1.13 Wylte -	Small bugfix, updated instructions for shadows, changed collection slightly (fewer top lanes, added 1 outlane light each side)
'1.14 apophis -	Updated ball rolling and ramp rolling to use only loudest sound effect with a volume option. Changed nudge strength settings to value 1. Rubberizer option removed, now using rothbauerw's version. Physics options moved away from player options area.
'1.15 scampa123 - Updated the BallRolling, RampLoop, WireLoop sounds in Sound Manager and updated the code to address the name changes (removed _amp9).  Added some additional comments to the table.
'1.2 apophis -	Minor comment edits. Removed old target bouncer code. Other cleanup.
'1.21 apophis -	Sling angle code test revA
'1.22 Wylte -	Fixed standups
'1.23 apophis - Added better installation instructions for slingshot corrections. Updated the correction curve.
'1.24 apophis - Added physical trough. Added mechanical tilt stuff (from oqqsan). Kicker randomness (from Thal).
'1.25 Wylte	-	Update shadow code (thanks vbousquet), better comments, switched all getballs to gBOT
'1.26 apophis - Added fix to table1_exit logic and ST collidable prim positions. Updated flasher domes to use ObjTargetLevel (iaakki). Updated Lampz class to support modulated signals, per TZ example (Niwak).
'1.27 apophis - Added SoundFX call to SoundFlipperUpAttackLeft and SoundFlipperUpAttackRight
'1.3 apophis -  Release
'1.31 Wylte	-	Fix shadows, ramp DB, added comments and reorganized shadow code for (hopefully) more clarity
'1.31a Wylte -	Couple more fixes/comments after doing a split pf
'1.32 apophis - Released with DMD folder
'1.33 Wylte -	ROM SoundCommandListener, ball on ramp shadows improved, update efficiency improved
'1.4 Arelyel -  VPin Workshop Advanced Queuing System 1.1.0 + examples; added additional property to DTArray elements for determining if a drop target is dropped
'1.41 Wylte -   Implement fluffhead35's bsRampType definitions for accurate shadows on ramps, small tweaks to ramp geometry
'1.5 Arelyel -  Implemented the VPW Coding Standard; fixed the script with the new standard; Breaking: renamed Lampz.UseFunc to Lampz.UseFunc; Added baldgeek error logs from Blood Machines
'1.5.1 apophis -Converted four spaces to tabs. Updated lampz to support original tables. Renamed and organized layers.
'1.5.2 Arelyel - Minor bug fixes in DTArrayID and STArrayID; useful tip for tabbing blocks of code added
'1.5.3 apophis - Updated flipper correction code per nFozzy's update. Added rendered playfield shadows. Updated rtx shadow image (bsrtx8) with softer edges. Updated drop target primitives and textures. Reduced target elasticity falloff (per nFozzy). Cleaned up slingshot objects. Added table of contents in the comments. Automated VRRoom.
'1.5.4 apophis - Updated all the video links to our new tutorial videos on YouTube
'1.5.5 Arelyel - Updated coding standards; added link to auto-formatting tool for VPX scripts
'1.5.6 apophis - Updated roth DT and ST code to be compatible with VPX standalone implementations
'1.5.7 Arelyel - Added function DTDropped; Code format
'1.5.8 jsm174 - Script updates to support standalone VPX builds
'1.5.9 apophis - Added new flipper trick FlipperCradleCollision. Tuned on VPW's Godzilla (Sega)
'1.5.10 apophis - Added min function. Fixed VolumeDial application in Fleep code. Commented new flipper trick.
'1.5.11 mcarter78 - Add random wire ramp stop sounds
'1.5.12 rothbauerw - Drop Target fix for edge case discovered on Paragon, new recommendations for EOSTNew, moved FlipperActivate and FlipperDeactivate to the flipper Sol subs
'1.5.13 rothbauerw - Added code for stand up and drop targets to handle multiple switches with same number, added instructions for creating prims for stand-up and drop targets, added updated live catch code, added updated flipper trajectory code, flipper physics recommendations, and new polarity and velocity curves.
'1.5.14 apophis - Updated tutorial vid links. Removed coding standards section. Added Inlane switch speed limit code. Added Flipper_Animate examples. Moved most timer stuff to FrameTimer (improved performance). Reorg some script. Added table rules and description with markdown formatting examples (Table > Table Info ...)
'1.5.15 mcarter78
' - Add tweak menu examples, including LUTs (Table1_OptionEvent). 
' - Removed old dynamic shadows in favor of new built-in dynamic shadows.
' - Updated ambient shadow code.
' - Remove Lampz in favor of VPX light handler (use VPX "Incandescent" fader for lights)
' - Add light_animate subs for all 3D inserts (to control their primitive blenddisablelighting)
'1.5.16 mcarter78 - set insert lights to incandescent fader
'1.5.17 RobbyKingPin - Added new improved primitives for the Flupper Domes
'1.5.18 apophis - Insert light fade times longer (so fading is more obvious). Updated pf-shadows image.
'1.5.19 apophis - Disabled "hide parts behind" for ball and flipper shadow primitives.
'1.5.20 apophis - Added correction to aBall.velz in dampener code
'1.5.21 apophis - Added examples of correctly sized flippers to the Flippers layer.
'1.5.22 apophis - Updated DisableStaticPreRendering functionality to be consistent with VPX 10.8.1 API
'1.6.1 apophis - Fixed issue with flipper dampening code.

'*********************************************************************************************************************************
' === ZTUT: VPW TUTORIAL VIDEOS  ===
'
' Watch all the VPW videos at:     									https://www.youtube.com/@vpinworkshop
'
' The following videos pertain to examples presented in this table:
'
' Audio : Adding Fleep Part 1										https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2										https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3										https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy
' Adding nFozzy roth physics : pt1 rubber dampeners 				https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics 					https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements 					https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
' Shadows : How to make rendered playfield shadows					https://youtu.be/rYI-gI7-RyE?si=DRVDlf-II5LMh1K-
' Basic VP kicker holes in Blender 2.8								https://youtu.be/Zzg89u8j0zA?si=Z7LfV_C8ts5C6x9G
' Intermediate VP mesh playfield in Blender 2.8						https://youtu.be/khq5elVVmw4?si=4A1m6op5IXbjO7-J
'
'*********************************************************************************************************************************



Option Explicit
Randomize

On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the controller.vbs In order To run this table, available In the vp10 package"
On Error GoTo 0


'*******************************************
'  ZCON: Constants and Global Variables
'*******************************************

Const UsingROM = False			  'The UsingROM flag is to indicate code that requires ROM usage. Mostly for instructional purposes only.

Const BallSize = 50				 'Ball diameter in VPX units; must be 50
Const BallMass = 1				  'Ball mass must be 1
Const tnob = 5					  'Total number of balls the table can hold
Const lob = 0					   'Locked balls
Const cGameName = "RoyalFlush_1957"		 'The unique alphanumeric name for this table
Dim gBOT

'Detect if VPX is rendering in VR and then make sure the VR Room Chioce is used
Dim VRRoom
If RenderingMode = 2 Then
	VRRoom = VRRoomChoice
Else
	VRRoom = 0
End If

Dim tablewidth
tablewidth = Table1.width
Dim tableheight
tableheight = Table1.height
Dim BIP							 'Balls in play
BIP = 0
Dim BIPL							'Ball in plunger lane
BIPL = False

Dim PlayerScore(4)
Dim CurrentPlayer
Dim BonusX(4)					   'Bonus X tracking at bumper inlanes

Dim queue						   'Main queue system via vpwQueueManager
Set queue = New vpwQueueManager

Dim FlexScenes(100)				 'Array of FlexDMD scenes

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

'*******************************************
'	ZTIM: Timers
'*******************************************

'The FrameTimer interval should be -1, so executes at the display frame rate
'The frame timer should be used to update anything visual, like some animations, shadows, etc.
'However, a lot of animations will be handled in their respective _animate subroutines.

Dim FrameTime, InitFrameTime
InitFrameTime = 0

FrameTimer.Interval = -1
Sub FrameTimer_Timer() 'The frame timer interval should be -1, so executes at the display frame rate
	FrameTime = GameTime - InitFrameTime
	InitFrameTime = GameTime	'Count frametime
	'Add animation stuff here
	RollingUpdate   		'update rolling sounds
	DoDTAnim				'handle drop target animations
	DoSTAnim				'handle stand up target animations
	queue.Tick	      		'handle the queue system
	BSUpdate
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub



'*******************************************
'	ZINI: Table Initialization and Exiting
'*******************************************

LoadCoreFiles
Sub LoadCoreFiles
	On Error Resume Next
	ExecuteGlobal GetTextFile("core.vbs") 'TODO: drop-in replacement for vpmTimer (maybe vpwQueueManager) and cvpmDictionary (Scripting.Dictionary) to remove core.vbs dependency
	If Err Then MsgBox "Can't open core.vbs"
	On Error GoTo 0
End Sub

Sub Table1_Init
	ConfigureGlfDevices()
    Glf_Init()
End Sub


Sub Table1_Exit
	'Close flexDMD
	If UseFlexDMD = 0 Then Exit Sub
	If Not FlexDMD Is Nothing Or VRRoom = 0 Then
		FlexDMD.Show = False
		FlexDMD.Run = False
		FlexDMD = Null
	End If

	Glf_Exit()
End Sub




'**********************************
' 	ZMAT: General Math Functions
'**********************************
' These get used throughout the script. 

Dim PI
PI = 4 * Atn(1)

Function dSin(degrees)
	dsin = Sin(degrees * Pi / 180)
End Function

Function dCos(degrees)
	dcos = Cos(degrees * Pi / 180)
End Function

Function Atn2(dy, dx)
	If dx > 0 Then
		Atn2 = Atn(dy / dx)
	ElseIf dx < 0 Then
		If dy = 0 Then
			Atn2 = pi
		Else
			Atn2 = Sgn(dy) * (pi - Atn(Abs(dy / dx)))
		End If
	ElseIf dx = 0 Then
		If dy = 0 Then
			Atn2 = 0
		Else
			Atn2 = Sgn(dy) * pi / 2
		End If
	End If
End Function

Function ArcCos(x)
	If x = 1 Then
		ArcCos = 0/180*PI
	ElseIf x = -1 Then
		ArcCos = 180/180*PI
	Else
		ArcCos = Atn(-x/Sqr(-x * x + 1)) + 2 * Atn(1)
	End If
End Function

Function max(a,b)
	If a > b Then
		max = a
	Else
		max = b
	End If
End Function

Function min(a,b)
	If a > b Then
		min = b
	Else
		min = a
	End If
End Function

' Used for drop targets
Function InRect(px,py,ax,ay,bx,by,cx,cy,dx,dy) 'Determines if a Points (px,py) is inside a 4 point polygon A-D in Clockwise/CCW order
	Dim AB, BC, CD, DA
	AB = (bx * py) - (by * px) - (ax * py) + (ay * px) + (ax * by) - (ay * bx)
	BC = (cx * py) - (cy * px) - (bx * py) + (by * px) + (bx * cy) - (by * cx)
	CD = (dx * py) - (dy * px) - (cx * py) + (cy * px) + (cx * dy) - (cy * dx)
	DA = (ax * py) - (ay * px) - (dx * py) + (dy * px) + (dx * ay) - (dy * ax)
	
	If (AB <= 0 And BC <= 0 And CD <= 0 And DA <= 0) Or (AB >= 0 And BC >= 0 And CD >= 0 And DA >= 0) Then
		InRect = True
	Else
		InRect = False
	End If
End Function

Function InRotRect(ballx,bally,px,py,angle,ax,ay,bx,by,cx,cy,dx,dy)
	Dim rax,ray,rbx,rby,rcx,rcy,rdx,rdy
	Dim rotxy
	rotxy = RotPoint(ax,ay,angle)
	rax = rotxy(0) + px
	ray = rotxy(1) + py
	rotxy = RotPoint(bx,by,angle)
	rbx = rotxy(0) + px
	rby = rotxy(1) + py
	rotxy = RotPoint(cx,cy,angle)
	rcx = rotxy(0) + px
	rcy = rotxy(1) + py
	rotxy = RotPoint(dx,dy,angle)
	rdx = rotxy(0) + px
	rdy = rotxy(1) + py
	
	InRotRect = InRect(ballx,bally,rax,ray,rbx,rby,rcx,rcy,rdx,rdy)
End Function

Function RotPoint(x,y,angle)
	Dim rx, ry
	rx = x * dCos(angle) - y * dSin(angle)
	ry = x * dSin(angle) + y * dCos(angle)
	RotPoint = Array(rx,ry)
End Function




'******************************************************
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
	dim a: a = LeftFlipper.CurrentAngle
	FlipperLSh.RotZ = a
	LFLogo.RotZ = a
	'Add any left flipper related animations here
End Sub

Sub RightFlipper_Animate
	dim a: a = RightFlipper.CurrentAngle
	FlipperRSh.RotZ = a
	RFlogo.RotZ = a
	'Add any right flipper related animations here
End Sub




'*******************************************
'	ZDRN: Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this,
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to
'	 the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 300 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


'TROUGH
Sub swTrough1_Hit
	UpdateTrough
End Sub
Sub swTrough1_UnHit
	UpdateTrough
End Sub
Sub swTrough2_Hit
	UpdateTrough
End Sub
Sub swTrough2_UnHit
	UpdateTrough
End Sub
Sub swTrough3_Hit
	UpdateTrough
End Sub
Sub swTrough3_UnHit
	UpdateTrough
End Sub
Sub swTrough4_Hit
	UpdateTrough
End Sub
Sub swTrough4_UnHit
	UpdateTrough
End Sub
Sub swTrough5_Hit
	UpdateTrough
End Sub
Sub swTrough5_UnHit
	UpdateTrough
End Sub

Sub UpdateTrough
	UpdateTroughTimer.Interval = 300
	UpdateTroughTimer.Enabled = 1
End Sub

Sub UpdateTroughTimer_Timer
	If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	Me.Enabled = 0
End Sub

' DRAIN & RELEASE
Sub Drain_Hit
	BIP = BIP - 1
	DMDBigText "DRAIN BLOCK",77,1
	RandomSoundDrain Drain
	vpmTimer.AddTimer 500, "Drain.kick 57, 20'"
End Sub

Sub SolRelease(enabled)
	If enabled Then
		BIP = BIP + 1
		swTrough1.kick 90, 10
		RandomSoundBallRelease swTrough1
	End If
End Sub

'*******************************************
'	ZSCR: Scoring
'*******************************************

Sub Addscore (value)
	PlayerScore(CurrentPlayer) = PlayerScore(CurrentPlayer) + value
	If value > 199 Then DMDBGFlash = 15
	If value > 4999 Then DMDFire = Flexframe + 50
End Sub

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


'*******************************************
'	ZFLP: Flippers
'*******************************************

Const ReflipAngle = 20

' Flipper Solenoid Callbacks (these subs mimics how you would handle flippers in ROM based tables)
Sub SolLFlipper(Enabled) 'Left flipper solenoid callback
	If Enabled Then
		FlipperActivate LeftFlipper, LFPress
		LF.Fire  'leftflipper.rotatetoend
		
		If leftflipper.currentangle < leftflipper.endangle + ReflipAngle Then
			RandomSoundReflipUpLeft LeftFlipper
		Else
			SoundFlipperUpAttackLeft LeftFlipper
			RandomSoundFlipperUpLeft LeftFlipper
		End If
	Else
		FlipperDeActivate LeftFlipper, LFPress
		LeftFlipper.RotateToStart
		If LeftFlipper.currentangle < LeftFlipper.startAngle - 5 Then
			RandomSoundFlipperDownLeft LeftFlipper
		End If
		FlipperLeftHitParm = FlipperUpSoundLevel
	End If
End Sub

Sub SolRFlipper(Enabled) 'Right flipper solenoid callback
	If Enabled Then
		FlipperActivate RightFlipper, RFPress
		RF.Fire 'rightflipper.rotatetoend
		
		If rightflipper.currentangle > rightflipper.endangle - ReflipAngle Then
			RandomSoundReflipUpRight RightFlipper
		Else
			SoundFlipperUpAttackRight RightFlipper
			RandomSoundFlipperUpRight RightFlipper
		End If
	Else
		FlipperDeActivate RightFlipper, RFPress
		RightFlipper.RotateToStart
		If RightFlipper.currentangle > RightFlipper.startAngle + 5 Then
			RandomSoundFlipperDownRight RightFlipper
		End If
		FlipperRightHitParm = FlipperUpSoundLevel
	End If
End Sub

' Flipper collide subs
Sub LeftFlipper_Collide(parm)
	CheckLiveCatch ActiveBall, LeftFlipper, LFCount, parm
	LF.ReProcessBalls ActiveBall
	LeftFlipperCollide parm
End Sub

Sub RightFlipper_Collide(parm)
	CheckLiveCatch ActiveBall, RightFlipper, RFCount, parm
	RF.ReProcessBalls ActiveBall
	RightFlipperCollide parm
End Sub

'*******************************************
'	ZBMP: Bumpers
'*******************************************

Sub Bumper1_Hit
	Addscore 250
	ToggleGI 0
	RandomSoundBumperTop Bumper1
	FlBumperFadeTarget(1) = 1   'Flupper bumper demo
	Bumper1.timerenabled = True
End Sub

Sub Bumper1_Timer
	FlBumperFadeTarget(1) = 0
End Sub

Sub Bumper2_Hit
	Addscore 250
	RandomSoundBumperMiddle Bumper2
	FlBumperFadeTarget(2) = 1   'Flupper bumper demo
	Bumper2.timerenabled = True
End Sub

Sub Bumper2_Timer
	FlBumperFadeTarget(2) = 0
End Sub

Sub Bumper3_Hit
	Addscore 250
	RandomSoundBumperBottom Bumper3
	FlBumperFadeTarget(3) = 1   'Flupper bumper demo
	Bumper3.timerenabled = True
End Sub

Sub Bumper3_Timer
	FlBumperFadeTarget(3) = 0
End Sub

Sub Bumper4_Hit
	Addscore 250
	RandomSoundBumperTop Bumper4
	FlBumperFadeTarget(4) = 1   'Flupper bumper demo
	Bumper4.timerenabled = True
End Sub

Sub Bumper4_Timer
	FlBumperFadeTarget(4) = 0
End Sub

Sub Bumper5_Hit
	Addscore 250
	RandomSoundBumperMiddle Bumper5
	FlBumperFadeTarget(5) = 1   'Flupper bumper demo
	Bumper5.timerenabled = True
End Sub

Sub Bumper5_Timer
	FlBumperFadeTarget(5) = 0
End Sub

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

'****************************************************************
'	ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShot_Slingshot
	RS.VelocityCorrect(ActiveBall)
	Addscore 10
	RSling1.Visible = 1
	Sling1.TransY =  - 20   'Sling Metal Bracket
	RStep = 0
	RightSlingShot.TimerEnabled = 1
	RightSlingShot.TimerInterval = 10
	'   vpmTimer.PulseSw 52	'Slingshot Rom Switch
	RandomSoundSlingshotRight Sling1
End Sub

Sub RightSlingShot_Timer
	Select Case RStep
		Case 3
			RSLing1.Visible = 0
			RSLing2.Visible = 1
			Sling1.TransY =  - 10
		Case 4
			RSLing2.Visible = 0
			Sling1.TransY = 0
			RightSlingShot.TimerEnabled = 0
	End Select
	RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
	LS.VelocityCorrect(ActiveBall)
	Addscore 10
	LSling1.Visible = 1
	Sling2.TransY =  - 20   'Sling Metal Bracket
	LStep = 0
	LeftSlingShot.TimerEnabled = 1
	LeftSlingShot.TimerInterval = 10
	'   vpmTimer.PulseSw 51	'Slingshot Rom Switch
	RandomSoundSlingshotLeft Sling2
End Sub

Sub LeftSlingShot_Timer
	Select Case LStep
		Case 3
			LSLing1.Visible = 0
			LSLing2.Visible = 1
			Sling2.TransY =  - 10
		Case 4
			LSLing2.Visible = 0
			Sling2.TransY = 0
			LeftSlingShot.TimerEnabled = 0
	End Select
	LStep = LStep + 1
End Sub

Sub TestSlingShot_Slingshot
	TS.VelocityCorrect(ActiveBall)
End Sub

'*******************************************
'	ZKIC: Kickers, Saucers
'*******************************************

'To include some randomness in the Kicker's kick, use the following parmeters
Const KickerAngleTol = 2	'Number of degrees the kicker angle varies around its intended direction
Const KickerStrengthTol = 1 'Number of strength units the kicker varies around its intended strength
Dim ballInKicker1
ballInKicker1 = False

Sub Kicker1_Hit
	ballInKicker1 = True
	Addscore 5000
	SoundSaucerLock
	
	' Determine drop target bonus
	Dim dropsDropped
	dropsDropped = 0
	Dim i
	For i = 0 To UBound(DTArray)
		If DTDropped(DTArray(i).sw) Then dropsDropped = dropsDropped + 1
	Next
	Select Case (dropsDropped)
		Case 0
			queue.Add "dropBonus0", "ShowScene flexScenes(5), FlexDMD_RenderMode_DMD_RGB, 6", 2, 0, 0, 2000, 0, False
		Case 1
			Addscore 5000
			queue.Add "dropBonus1", "ShowScene flexScenes(6), FlexDMD_RenderMode_DMD_RGB, 7", 2, 0, 0, 3000, 0, False
		Case 2
			Addscore 10000
			queue.Add "dropBonus2", "ShowScene flexScenes(7), FlexDMD_RenderMode_DMD_RGB, 8", 2, 0, 0, 3000, 0, False
		Case 3
			Addscore 20000
			queue.Add "dropBonus3", "ShowScene flexScenes(8), FlexDMD_RenderMode_DMD_RGB, 9", 2, 0, 0, 3000, 0, False
	End Select
End Sub

Sub Kicker1_Timer
	SoundSaucerKick 1, Kicker1
	Kicker1.Kick 160 + RndNum( - KickerAngleTol,KickerAngleTol), 20 + RndNum( - KickerStrengthTol,KickerStrengthTol)
	Kicker1.timerenabled = False
End Sub

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

'********************************************
'	ZTAR: Targets
'********************************************

Sub sw11_Hit
	STHit 11
End Sub

Sub sw11o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw12_Hit
	STHit 12
End Sub

Sub sw12o_Hit
	TargetBouncer ActiveBall, 1
End Sub

Sub sw13_Hit
	STHit 13
End Sub

Sub sw13o_Hit
	TargetBouncer ActiveBall, 1
End Sub

'********************************************
'  Drop Target Controls
'********************************************

' Drop targets
Sub sw1_Hit
	DTHit 1
End Sub

Sub sw2_Hit
	DTHit 2
End Sub

Sub sw3_Hit
	DTHit 3
End Sub

' If the drop targets can be reset individually, use specific solenoid subs for each like below
' These subroutines would be called by the solenoid callbacks if using a ROM

Sub SolDT1(enabled) ' Drop Target 1 Solenoid
	If enabled Then
		RandomSoundDropTargetReset sw1p
		DTRaise 1
	End If
End Sub

Sub SolDT2(enabled) ' Drop Target 2 Solenoid
	If enabled Then
		RandomSoundDropTargetReset  sw2p
		DTRaise 2
	End If
End Sub

Sub SolDT3(enabled) ' Drop Target 3 Solenoid
	If enabled Then
		RandomSoundDropTargetReset  sw3p
		DTRaise 3
	End If
End Sub

' If a whole bank of drop targets can be reset at once, use sub like below

Sub SolDTBank123(enabled)
	Dim xx
	If enabled Then
		RandomSoundDropTargetReset sw2p
		DTRaise 1
		DTRaise 2
		DTRaise 3
		For Each xx In ShadowDT
			xx.visible = True
		Next
	End If
End Sub

'********************************************************
'	ZLIS: ROM SoundCommand Listener
'********************************************************
'Originally by ...?
'LotR code by Apophis
'2-in-1 sub by Wylte

'Examples from Robot (simple If-Then use) and LotR (complicated state combinations)

' *** Sound Commands ***

'It's good to note them here for reference

'In Robot
' D4,D4,54,	- Challenge (New Game)
' D2,D2,52,	- Mission Completed (End Game)
' C4,C4,44,	- Whoa Whoa (Tilt)
' C3,C3,43,	- Ow (Tilt)
' E3,E3,63,	- New Ball

'In LotR
' FD1D = start of destroy the ring mode
' FD2C = successfully completed DTR
' FD0B = Ring wraithes mission 1
' FD0C = Ring wraithes mission 2
' FD0D = Ring wraithes mission 3
' FD09 = shelob mission starts
' FD04 = end of ball
' FD03 = main theme 2 (not in a mode)
' FD02 = main theme 1 (not in a mode)

' *** Constants ***

' The commands from the ROM come through in hex
' so add constants for the numbers you're listening for

'Const hexFD = 253
'Const hexE3 = 227
'Const hexD4 = 212
'Const hexD2 = 210
'Const hexC3 = 196
'Const hexC4 = 195
'Const hex7F = 127
'Const hex54 = 84
'Const hex52 = 82
'Const hex2C = 44
'Const hex1D = 29
'Const hex0D = 13
'Const hex0C = 12
'Const hex0B = 11
'Const hex09 = 9
'Const hex08 = 8
'Const hex04 = 4
'Const hex03 = 3
'Const hex02 = 2

Dim SndCmdStr
Dim LastSnd
LastSnd = 0

' *** Listener ***

' WHAT YOU NEED:

' Add the SoundCmdListener call to a fast timer that always runs
' Copy in the 10 "Listener###" textboxes from the backglass view
' Play the game, writing down the values for the event you need
' The event you want may have multiple sounds including sounds played at other times
' If unsure, you can note the values and watch to see if they are used elsewhere
' You may want to record your screen if it's moving fast

Sub SoundCmdListener
	Dim NewSounds,ii,Snd
	NewSounds = Controller.NewSoundCommands	 ' Listening to the ROM
	If Not IsEmpty(NewSounds) Then
		'* Finding the Values *
		' Comment out or delete this part once you know the commands
		' And delete the text boxes
		SndCmdStr = ""	  ' Empty the string
		For ii = 0 To UBound(NewSounds)	 ' Setting the new to a string to display
			Snd = NewSounds(ii,0)
			If Snd <> 0 And Snd <> 255 Then
				SndCmdStr = SndCmdStr & Hex(Snd) & ","
			End If
		Next
		
		Listener010.Text = Listener009.Text				 ' Display sound commands in textboxes
		Listener009.Text = Listener008.Text				 ' ^^^
		Listener008.Text = Listener007.Text				 ' ^^^
		Listener007.Text = Listener006.Text				 ' ^^^
		Listener006.Text = Listener005.Text				 ' ^^^
		Listener005.Text = Listener004.Text				 ' ^^^
		Listener004.Text = Listener003.Text				 ' ^^^
		Listener003.Text = Listener002.Text				 ' ^^^
		Listener002.Text = Listener001.Text				 ' ^^^
		Listener001.Text = SndCmdStr						' Arrange them vertically so the sound commands scroll up the screen
		
		'* Using them to do stuff *
		
		For ii = 0 To UBound(NewSounds) 'Listen for specific commands; if a specified combination occurs, do a thing
			Snd = NewSounds(ii,0)
			'From Robot:
			'			If LastSnd = hexD2 And Snd = hexD2 Then													'End Game, Flippers Off and Dropped
			'				FlippersDisabled = True
			'				LeftFlipper.RotateToStart
			'				RightFlipper.RotateToStart
			'			End If
			'End Robot
			
			'From LotR:
			'			If LastSnd = hexFD Then
			'				Select Case Snd
			'					Case hex1D: ClearSndFlags : Snd_DTRstart = True : TORcnt = 1 : TheOneRingUpdates.Enabled=True ': debug.print "Snd_DTRstart = True"
			'					Case hex2C: ClearSndFlags : Snd_DTRfinal = True ': debug.print "Snd_DTRfinal = True"
			'					Case hex0D: ClearSndFlags : Snd_Wraithes = True ': debug.print "hex0D: Snd_Wraithes = True"
			'					Case hex0C: ClearSndFlags : Snd_Wraithes = True ': debug.print "hex0C: Snd_Wraithes = True"
			'					Case hex0B: ClearSndFlags : Snd_Wraithes = True ': debug.print "hex0B: Snd_Wraithes = True"
			'					Case hex08: ClearSndFlags ': debug.print "hex08: Snd_Wraithes = False"
			'					Case hex09: ClearSndFlags : Snd_Shelob = True : Setlamp 102,1 ': debug.print "Snd_Shelob = True"
			'					Case hex04: ClearSndFlags : Snd_EOB = True ': debug.print "Snd_EOB = True"
			'					Case hex03: ClearSndFlags : Snd_Main = True ': debug.print "Snd_Main = True"
			'					Case hex02: ClearSndFlags : Snd_Main = True ': debug.print "Snd_Main = True"
			'				End Select
			'			End If
			'End LotR
			
			LastSnd = Snd	   ' Remembering the previous bit of hex
			'			debug.print "LastSnd: " & LastSnd
		Next
	End If
End Sub

'LotR Specific code (here for reference)

'Sub ClearSndFlags
'	'debug.print "ClearSndFlags"
'	Snd_DTRstart = False
'	Snd_DTRfinal = False
'	Snd_Shelob = False
'	Snd_Wraithes = False
'	Snd_EOB = False
'	Snd_Main = False
'
'	Setlamp 102,0
'End Sub

'********************************************************
'  END ROM SoundCommand Listener
'********************************************************

'***************************************************************
'	ZSHA: Ambient ball shadows
'***************************************************************

' For dynamic ball shadows, Check the "Raytraced ball shadows" box for the specific light. 
' Also make sure the light's z position is around 25 (mid ball)

'Ambient (Room light source)
Const AmbientBSFactor = 0.9    '0 To 1, higher is darker
Const AmbientMovement = 1	   '1+ higher means more movement as the ball moves left and right
Const offsetX = 0			   'Offset x position under ball (These are if you want to change where the "room" light is for calculating the shadow position,)
Const offsetY = 0			   'Offset y position under ball (^^for example 5,5 if the light is in the back left corner)

' *** Trim or extend these to match the number of balls/primitives/flashers on the table!  (will throw errors if there aren't enough objects)
Dim objBallShadow(7)

'Initialization
BSInit

Sub BSInit()
	Dim iii
	'Prepare the shadow objects before play begins
	For iii = 0 To tnob - 1
		Set objBallShadow(iii) = Eval("BallShadow" & iii)
		objBallShadow(iii).material = "BallShadow" & iii
		UpdateMaterial objBallShadow(iii).material,1,0,0,0,0,0,AmbientBSFactor,RGB(0,0,0),0,0,False,True,0,0,0,0
		objBallShadow(iii).Z = 3 + iii / 1000
		objBallShadow(iii).visible = 0
	Next
End Sub


Sub BSUpdate
	Dim s: For s = lob To UBound(gBOT)
		' *** Normal "ambient light" ball shadow
		
		'Primitive shadow on playfield, flasher shadow in ramps
		'** If on main and upper pf
		If gBOT(s).Z > 20 And gBOT(s).Z < 30 Then
			objBallShadow(s).visible = 1
			objBallShadow(s).X = gBOT(s).X + (gBOT(s).X - (tablewidth / 2)) / (Ballsize / AmbientMovement) + offsetX
			objBallShadow(s).Y = gBOT(s).Y + offsetY
			'objBallShadow(s).Z = gBOT(s).Z + s/1000 + 1.04 - 25	

		'** No shadow if ball is off the main playfield (this may need to be adjusted per table)
		Else
			objBallShadow(s).visible = 0
		End If
	Next
End Sub


'******************************************************
'	ZPHY:  GNEREAL ADVICE ON PHYSICS
'******************************************************
'
' It's advised that flipper corrections, dampeners, and general physics settings should all be updated per these
' examples as all of these improvements work together to provide a realistic physics simulation.
'
' Tutorial videos provided by Bord
' Adding nFozzy roth physics : pt1 rubber dampeners 				https://youtu.be/AXX3aen06FM?si=Xqd-rcaqTlgEd_wx
' Adding nFozzy roth physics : pt2 flipper physics 					https://youtu.be/VSBFuK2RCPE?si=i8ne8Ao2co8rt7fy
' Adding nFozzy roth physics : pt3 other elements 					https://youtu.be/JN8HEJapCvs?si=hvgMOk-ej1BEYjJv
'
' Note: BallMass must be set to 1. BallSize should be set to 50 (in other words the ball radius is 25)
'
' Recommended Table Physics Settings
' | Gravity Constant             | 0.97      |
' | Playfield Friction           | 0.15-0.25 |
' | Playfield Elasticity         | 0.25      |
' | Playfield Elasticity Falloff | 0         |
' | Playfield Scatter            | 0         |
' | Default Element Scatter      | 2         |
'
' Bumpers
' | Force         | 12-15    |
' | Hit Threshold | 1.6-2    |
' | Scatter Angle | 2        |
'
' Slingshots
' | Hit Threshold      | 2    |
' | Slingshot Force    | 3-5  |
' | Slingshot Theshold | 2-3  |
' | Elasticity         | 0.85 |
' | Friction           | 0.8  |
' | Scatter Angle      | 1    |





'******************************************************
'	ZNFF:  FLIPPER CORRECTIONS by nFozzy
'******************************************************
'
' There are several steps for taking advantage of nFozzy's flipper solution.  At a high level we'll need the following:
'	1. flippers with specific physics settings
'	2. custom triggers for each flipper (TriggerLF, TriggerRF)
'	3. and, special scripting
'
' TriggerLF and RF should now be 27 vp units from the flippers. In addition, 3 degrees should be added to the end angle
' when creating these triggers.
'
' RF.ReProcessBalls Activeball and LF.ReProcessBalls Activeball must be added the flipper_collide subs.
'
' A common mistake is incorrect flipper length.  A 3-inch flipper with rubbers will be about 3.125 inches long.
' This translates to about 147 vp units.  Therefore, the flipper start radius + the flipper length + the flipper end
' radius should  equal approximately 147 vp units. Another common mistake is is that sometimes the right flipper
' angle was set with a large postive value (like 238 or something). It should be using negative value (like -122).
'
' The following settings are a solid starting point for various eras of pinballs.
' |                    | EM's           | late 70's to mid 80's | mid 80's to early 90's | mid 90's and later |
' | ------------------ | -------------- | --------------------- | ---------------------- | ------------------ |
' | Mass               | 1              | 1                     | 1                      | 1                  |
' | Strength           | 500-1000 (750) | 1400-1600 (1500)      | 2000-2600              | 3200-3300 (3250)   |
' | Elasticity         | 0.88           | 0.88                  | 0.88                   | 0.88               |
' | Elasticity Falloff | 0.15           | 0.15                  | 0.15                   | 0.15               |
' | Fricition          | 0.8-0.9        | 0.9                   | 0.9                    | 0.9                |
' | Return Strength    | 0.11           | 0.09                  | 0.07                   | 0.055              |
' | Coil Ramp Up       | 2.5            | 2.5                   | 2.5                    | 2.5                |
' | Scatter Angle      | 0              | 0                     | 0                      | 0                  |
' | EOS Torque         | 0.4            | 0.4                   | 0.375                  | 0.375              |
' | EOS Torque Angle   | 4              | 4                     | 6                      | 6                  |
'

'******************************************************
' Flippers Polarity (Select appropriate sub based on era)
'******************************************************

Dim LF : Set LF = New FlipperPolarity
Dim RF : Set RF = New FlipperPolarity

InitPolarity

'
''*******************************************
'' Late 70's to early 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'	for each x in a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 80
'		x.DebugOn=False ' prints some info in debugger
'
'
'        x.AddPt "Polarity", 0, 0, 0
'        x.AddPt "Polarity", 1, 0.05, - 2.7
'        x.AddPt "Polarity", 2, 0.16, - 2.7
'        x.AddPt "Polarity", 3, 0.22, - 0
'        x.AddPt "Polarity", 4, 0.25, - 0
'        x.AddPt "Polarity", 5, 0.3, - 1
'        x.AddPt "Polarity", 6, 0.4, - 2
'        x.AddPt "Polarity", 7, 0.5, - 2.7
'        x.AddPt "Polarity", 8, 0.65, - 1.8
'        x.AddPt "Polarity", 9, 0.75, - 0.5
'        x.AddPt "Polarity", 10, 0.81, - 0.5
'        x.AddPt "Polarity", 11, 0.88, 0
'        x.AddPt "Polarity", 12, 1.3, 0
'
'		x.AddPt "Velocity", 0, 0, 0.85
'		x.AddPt "Velocity", 1, 0.15, 0.85
'		x.AddPt "Velocity", 2, 0.2, 0.9
'		x.AddPt "Velocity", 3, 0.23, 0.95
'		x.AddPt "Velocity", 4, 0.41, 0.95
'		x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'		x.AddPt "Velocity", 6, 0.62, 1.0
'		x.AddPt "Velocity", 7, 0.702, 0.968
'		x.AddPt "Velocity", 8, 0.95,  0.968
'		x.AddPt "Velocity", 9, 1.03,  0.945
'		x.AddPt "Velocity", 10, 1.5,  0.945
'
'	Next
'
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
'
'
''*******************************************
'' Mid 80's
'
'Sub InitPolarity()
'   dim x, a : a = Array(LF, RF)
'	for each x in a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 80
'		x.DebugOn=False ' prints some info in debugger
'
'		x.AddPt "Polarity", 0, 0, 0
'		x.AddPt "Polarity", 1, 0.05, - 3.7
'		x.AddPt "Polarity", 2, 0.16, - 3.7
'		x.AddPt "Polarity", 3, 0.22, - 0
'		x.AddPt "Polarity", 4, 0.25, - 0
'		x.AddPt "Polarity", 5, 0.3, - 2
'		x.AddPt "Polarity", 6, 0.4, - 3
'		x.AddPt "Polarity", 7, 0.5, - 3.7
'		x.AddPt "Polarity", 8, 0.65, - 2.3
'		x.AddPt "Polarity", 9, 0.75, - 1.5
'		x.AddPt "Polarity", 10, 0.81, - 1
'		x.AddPt "Polarity", 11, 0.88, 0
'		x.AddPt "Polarity", 12, 1.3, 0
'
'		x.AddPt "Velocity", 0, 0, 0.85
'		x.AddPt "Velocity", 1, 0.15, 0.85
'		x.AddPt "Velocity", 2, 0.2, 0.9
'		x.AddPt "Velocity", 3, 0.23, 0.95
'		x.AddPt "Velocity", 4, 0.41, 0.95
'		x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'		x.AddPt "Velocity", 6, 0.62, 1.0
'		x.AddPt "Velocity", 7, 0.702, 0.968
'		x.AddPt "Velocity", 8, 0.95,  0.968
'		x.AddPt "Velocity", 9, 1.03,  0.945
'		x.AddPt "Velocity", 10, 1.5,  0.945
'
'	Next
'
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'    LF.SetObjects "LF", LeftFlipper, TriggerLF
'    RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub
'
''*******************************************
''  Late 80's early 90's
'
'Sub InitPolarity()
'	dim x, a : a = Array(LF, RF)
'	for each x in a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 60
'		x.DebugOn=False ' prints some info in debugger
'
'		x.AddPt "Polarity", 0, 0, 0
'		x.AddPt "Polarity", 1, 0.05, - 5
'		x.AddPt "Polarity", 2, 0.16, - 5
'		x.AddPt "Polarity", 3, 0.22, - 0
'		x.AddPt "Polarity", 4, 0.25, - 0
'		x.AddPt "Polarity", 5, 0.3, - 2
'		x.AddPt "Polarity", 6, 0.4, - 3
'		x.AddPt "Polarity", 7, 0.5, - 4.0
'		x.AddPt "Polarity", 8, 0.7, - 3.5
'		x.AddPt "Polarity", 9, 0.75, - 3.0
'		x.AddPt "Polarity", 10, 0.8, - 2.5
'		x.AddPt "Polarity", 11, 0.85, - 2.0
'		x.AddPt "Polarity", 12, 0.9, - 1.5
'		x.AddPt "Polarity", 13, 0.95, - 1.0
'		x.AddPt "Polarity", 14, 1, - 0.5
'		x.AddPt "Polarity", 15, 1.1, 0
'		x.AddPt "Polarity", 16, 1.3, 0
'
'		x.AddPt "Velocity", 0, 0, 0.85
'		x.AddPt "Velocity", 1, 0.15, 0.85
'		x.AddPt "Velocity", 2, 0.2, 0.9
'		x.AddPt "Velocity", 3, 0.23, 0.95
'		x.AddPt "Velocity", 4, 0.41, 0.95
'		x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
'		x.AddPt "Velocity", 6, 0.62, 1.0
'		x.AddPt "Velocity", 7, 0.702, 0.968
'		x.AddPt "Velocity", 8, 0.95,  0.968
'		x.AddPt "Velocity", 9, 1.03,  0.945
'		x.AddPt "Velocity", 10, 1.5,  0.945

'	Next
'
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'	LF.SetObjects "LF", LeftFlipper, TriggerLF
'	RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

'*******************************************
' Early 90's and after

Sub InitPolarity()
	Dim x, a
	a = Array(LF, RF)
	For Each x In a
		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 60
		x.DebugOn=False ' prints some info in debugger

		x.AddPt "Polarity", 0, 0, 0
		x.AddPt "Polarity", 1, 0.05, - 5.5
		x.AddPt "Polarity", 2, 0.16, - 5.5
		x.AddPt "Polarity", 3, 0.20, - 0.75
		x.AddPt "Polarity", 4, 0.25, - 1.25
		x.AddPt "Polarity", 5, 0.3, - 1.75
		x.AddPt "Polarity", 6, 0.4, - 3.5
		x.AddPt "Polarity", 7, 0.5, - 5.25
		x.AddPt "Polarity", 8, 0.7, - 4.0
		x.AddPt "Polarity", 9, 0.75, - 3.5
		x.AddPt "Polarity", 10, 0.8, - 3.0
		x.AddPt "Polarity", 11, 0.85, - 2.5
		x.AddPt "Polarity", 12, 0.9, - 2.0
		x.AddPt "Polarity", 13, 0.95, - 1.5
		x.AddPt "Polarity", 14, 1, - 1.0
		x.AddPt "Polarity", 15, 1.05, -0.5
		x.AddPt "Polarity", 16, 1.1, 0
		x.AddPt "Polarity", 17, 1.3, 0

		x.AddPt "Velocity", 0, 0, 0.85
		x.AddPt "Velocity", 1, 0.23, 0.85
		x.AddPt "Velocity", 2, 0.27, 1
		x.AddPt "Velocity", 3, 0.3, 1
		x.AddPt "Velocity", 4, 0.35, 1
		x.AddPt "Velocity", 5, 0.6, 1 '0.982
		x.AddPt "Velocity", 6, 0.62, 1.0
		x.AddPt "Velocity", 7, 0.702, 0.968
		x.AddPt "Velocity", 8, 0.95,  0.968
		x.AddPt "Velocity", 9, 1.03,  0.945
		x.AddPt "Velocity", 10, 1.5,  0.945

	Next
	
	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
	LF.SetObjects "LF", LeftFlipper, TriggerLF
	RF.SetObjects "RF", RightFlipper, TriggerRF
End Sub

'******************************************************
'  FLIPPER CORRECTION FUNCTIONS
'******************************************************

' modified 2023 by nFozzy
' Removed need for 'endpoint' objects
' Added 'createvents' type thing for TriggerLF / TriggerRF triggers.
' Removed AddPt function which complicated setup imo
' made DebugOn do something (prints some stuff in debugger)
'   Otherwise it should function exactly the same as before\
' modified 2024 by rothbauerw
' Added Reprocessballs for flipper collisions (LF.Reprocessballs Activeball and RF.Reprocessballs Activeball must be added to the flipper collide subs
' Improved handling to remove correction for backhand shots when the flipper is raised

Class FlipperPolarity
	Public DebugOn, Enabled
	Private FlipAt		'Timer variable (IE 'flip at 723,530ms...)
	Public TimeDelay		'delay before trigger turns off and polarity is disabled
	Private Flipper, FlipperStart, FlipperEnd, FlipperEndY, LR, PartialFlipCoef, FlipStartAngle
	Private Balls(20), balldata(20)
	Private Name
	
	Dim PolarityIn, PolarityOut
	Dim VelocityIn, VelocityOut
	Dim YcoefIn, YcoefOut
	Public Sub Class_Initialize
		ReDim PolarityIn(0)
		ReDim PolarityOut(0)
		ReDim VelocityIn(0)
		ReDim VelocityOut(0)
		ReDim YcoefIn(0)
		ReDim YcoefOut(0)
		Enabled = True
		TimeDelay = 50
		LR = 1
		Dim x
		For x = 0 To UBound(balls)
			balls(x) = Empty
			Set Balldata(x) = new SpoofBall
		Next
	End Sub
	
	Public Sub SetObjects(aName, aFlipper, aTrigger)
		
		If TypeName(aName) <> "String" Then MsgBox "FlipperPolarity: .SetObjects error: first argument must be a String (And name of Object). Found:" & TypeName(aName) End If
		If TypeName(aFlipper) <> "Flipper" Then MsgBox "FlipperPolarity: .SetObjects error: Second argument must be a flipper. Found:" & TypeName(aFlipper) End If
		If TypeName(aTrigger) <> "Trigger" Then MsgBox "FlipperPolarity: .SetObjects error: third argument must be a trigger. Found:" & TypeName(aTrigger) End If
		If aFlipper.EndAngle > aFlipper.StartAngle Then LR = -1 Else LR = 1 End If
		Name = aName
		Set Flipper = aFlipper
		FlipperStart = aFlipper.x
		FlipperEnd = Flipper.Length * Sin((Flipper.StartAngle / 57.295779513082320876798154814105)) + Flipper.X ' big floats for degree to rad conversion
		FlipperEndY = Flipper.Length * Cos(Flipper.StartAngle / 57.295779513082320876798154814105)*-1 + Flipper.Y
		
		Dim str
		str = "Sub " & aTrigger.name & "_Hit() : " & aName & ".AddBall ActiveBall : End Sub'"
		ExecuteGlobal(str)
		str = "Sub " & aTrigger.name & "_UnHit() : " & aName & ".PolarityCorrect ActiveBall : End Sub'"
		ExecuteGlobal(str)
		
	End Sub
	
	' Legacy: just no op
	Public Property Let EndPoint(aInput)
		
	End Property
	
	Public Sub AddPt(aChooseArray, aIDX, aX, aY) 'Index #, X position, (in) y Position (out)
		Select Case aChooseArray
			Case "Polarity"
				ShuffleArrays PolarityIn, PolarityOut, 1
				PolarityIn(aIDX) = aX
				PolarityOut(aIDX) = aY
				ShuffleArrays PolarityIn, PolarityOut, 0
			Case "Velocity"
				ShuffleArrays VelocityIn, VelocityOut, 1
				VelocityIn(aIDX) = aX
				VelocityOut(aIDX) = aY
				ShuffleArrays VelocityIn, VelocityOut, 0
			Case "Ycoef"
				ShuffleArrays YcoefIn, YcoefOut, 1
				YcoefIn(aIDX) = aX
				YcoefOut(aIDX) = aY
				ShuffleArrays YcoefIn, YcoefOut, 0
		End Select
	End Sub
	
	Public Sub AddBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If IsEmpty(balls(x)) Then
				Set balls(x) = aBall
				Exit Sub
			End If
		Next
	End Sub
	
	Private Sub RemoveBall(aBall)
		Dim x
		For x = 0 To UBound(balls)
			If TypeName(balls(x) ) = "IBall" Then
				If aBall.ID = Balls(x).ID Then
					balls(x) = Empty
					Balldata(x).Reset
				End If
			End If
		Next
	End Sub
	
	Public Sub Fire()
		Flipper.RotateToEnd
		processballs
	End Sub
	
	Public Property Get Pos 'returns % position a ball. For debug stuff.
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x)) Then
				pos = pSlope(Balls(x).x, FlipperStart, 0, FlipperEnd, 1)
			End If
		Next
	End Property
	
	Public Sub ProcessBalls() 'save data of balls in flipper range
		FlipAt = GameTime
		Dim x
		For x = 0 To UBound(balls)
			If Not IsEmpty(balls(x)) Then
				balldata(x).Data = balls(x)
			End If
		Next
		FlipStartAngle = Flipper.currentangle
		PartialFlipCoef = ((Flipper.StartAngle - Flipper.CurrentAngle) / (Flipper.StartAngle - Flipper.EndAngle))
		PartialFlipCoef = abs(PartialFlipCoef-1)
	End Sub

	Public Sub ReProcessBalls(aBall) 'save data of balls in flipper range
		If FlipperOn() Then
			Dim x
			For x = 0 To UBound(balls)
				If Not IsEmpty(balls(x)) Then
					if balls(x).ID = aBall.ID Then
						If isempty(balldata(x).ID) Then
							balldata(x).Data = balls(x)
						End If
					End If
				End If
			Next
		End If
	End Sub

	'Timer shutoff for polaritycorrect
	Private Function FlipperOn()
		If GameTime < FlipAt+TimeDelay Then
			FlipperOn = True
		End If
	End Function
	
	Public Sub PolarityCorrect(aBall)
		If FlipperOn() Then
			Dim tmp, BallPos, x, IDX, Ycoef, BalltoFlip, BalltoBase, NoCorrection, checkHit
			Ycoef = 1
			
			'y safety Exit
			If aBall.VelY > -8 Then 'ball going down
				RemoveBall aBall
				Exit Sub
			End If
			
			'Find balldata. BallPos = % on Flipper
			For x = 0 To UBound(Balls)
				If aBall.id = BallData(x).id And Not IsEmpty(BallData(x).id) Then
					idx = x
					BallPos = PSlope(BallData(x).x, FlipperStart, 0, FlipperEnd, 1)
					BalltoFlip = DistanceFromFlipperAngle(BallData(x).x, BallData(x).y, Flipper, FlipStartAngle)
					If ballpos > 0.65 Then  Ycoef = LinearEnvelope(BallData(x).Y, YcoefIn, YcoefOut)								'find safety coefficient 'ycoef' data
				End If
			Next
			
			If BallPos = 0 Then 'no ball data meaning the ball is entering and exiting pretty close to the same position, use current values.
				BallPos = PSlope(aBall.x, FlipperStart, 0, FlipperEnd, 1)
				If ballpos > 0.65 Then  Ycoef = LinearEnvelope(aBall.Y, YcoefIn, YcoefOut)												'find safety coefficient 'ycoef' data
				NoCorrection = 1
			Else
				checkHit = 50 + (20 * BallPos) 

				If BalltoFlip > checkHit or (PartialFlipCoef < 0.5 and BallPos > 0.22) Then
					NoCorrection = 1
				Else
					NoCorrection = 0
				End If
			End If
			
			'Velocity correction
			If Not IsEmpty(VelocityIn(0) ) Then
				Dim VelCoef
				VelCoef = LinearEnvelope(BallPos, VelocityIn, VelocityOut)
				
				'If partialflipcoef < 1 Then VelCoef = PSlope(partialflipcoef, 0, 1, 1, VelCoef)
				
				If Enabled Then aBall.Velx = aBall.Velx*VelCoef
				If Enabled Then aBall.Vely = aBall.Vely*VelCoef
			End If
			
			'Polarity Correction (optional now)
			If Not IsEmpty(PolarityIn(0) ) Then
				Dim AddX
				AddX = LinearEnvelope(BallPos, PolarityIn, PolarityOut) * LR
				
				If Enabled and NoCorrection = 0 Then aBall.VelX = aBall.VelX + 1 * (AddX*ycoef*PartialFlipcoef*VelCoef)
			End If
			If DebugOn Then debug.print "PolarityCorrect" & " " & Name & " @ " & GameTime & " " & Round(BallPos*100) & "%" & " AddX:" & Round(AddX,2) & " Vel%:" & Round(VelCoef*100)
		End If
		RemoveBall aBall
	End Sub
End Class

'******************************************************
'  FLIPPER POLARITY AND RUBBER DAMPENER SUPPORTING FUNCTIONS
'******************************************************

' Used for flipper correction and rubber dampeners
Sub ShuffleArray(ByRef aArray, byVal offset) 'shuffle 1d array
	Dim x, aCount
	aCount = 0
	ReDim a(UBound(aArray) )
	For x = 0 To UBound(aArray)		'Shuffle objects in a temp array
		If Not IsEmpty(aArray(x) ) Then
			If IsObject(aArray(x)) Then
				Set a(aCount) = aArray(x)
			Else
				a(aCount) = aArray(x)
			End If
			aCount = aCount + 1
		End If
	Next
	If offset < 0 Then offset = 0
	ReDim aArray(aCount-1+offset)		'Resize original array
	For x = 0 To aCount-1				'set objects back into original array
		If IsObject(a(x)) Then
			Set aArray(x) = a(x)
		Else
			aArray(x) = a(x)
		End If
	Next
End Sub

' Used for flipper correction and rubber dampeners
Sub ShuffleArrays(aArray1, aArray2, offset)
	ShuffleArray aArray1, offset
	ShuffleArray aArray2, offset
End Sub

' Used for flipper correction, rubber dampeners, and drop targets
Function BallSpeed(ball) 'Calculates the ball speed
	BallSpeed = Sqr(ball.VelX^2 + ball.VelY^2 + ball.VelZ^2)
End Function

' Used for flipper correction and rubber dampeners
Function PSlope(Input, X1, Y1, X2, Y2)		'Set up line via two points, no clamping. Input X, output Y
	Dim x, y, b, m
	x = input
	m = (Y2 - Y1) / (X2 - X1)
	b = Y2 - m*X2
	Y = M*x+b
	PSlope = Y
End Function

' Used for flipper correction
Class spoofball
	Public X, Y, Z, VelX, VelY, VelZ, ID, Mass, Radius
	Public Property Let Data(aBall)
		With aBall
			x = .x
			y = .y
			z = .z
			velx = .velx
			vely = .vely
			velz = .velz
			id = .ID
			mass = .mass
			radius = .radius
		End With
	End Property
	Public Sub Reset()
		x = Empty
		y = Empty
		z = Empty
		velx = Empty
		vely = Empty
		velz = Empty
		id = Empty
		mass = Empty
		radius = Empty
	End Sub
End Class

' Used for flipper correction and rubber dampeners
Function LinearEnvelope(xInput, xKeyFrame, yLvl)
	Dim y 'Y output
	Dim L 'Line
	'find active line
	Dim ii
	For ii = 1 To UBound(xKeyFrame)
		If xInput <= xKeyFrame(ii) Then
			L = ii
			Exit For
		End If
	Next
	If xInput > xKeyFrame(UBound(xKeyFrame) ) Then L = UBound(xKeyFrame)		'catch line overrun
	Y = pSlope(xInput, xKeyFrame(L-1), yLvl(L-1), xKeyFrame(L), yLvl(L) )
	
	If xInput <= xKeyFrame(LBound(xKeyFrame) ) Then Y = yLvl(LBound(xKeyFrame) )		 'Clamp lower
	If xInput >= xKeyFrame(UBound(xKeyFrame) ) Then Y = yLvl(UBound(xKeyFrame) )		'Clamp upper
	
	LinearEnvelope = Y
End Function

'******************************************************
'  FLIPPER TRICKS
'******************************************************
' To add the flipper tricks you must
'	 - Include a call to FlipperCradleCollision from within OnBallBallCollision subroutine
'	 - Include a call the CheckLiveCatch from the LeftFlipper_Collide and RightFlipper_Collide subroutines
'	 - Include FlipperActivate and FlipperDeactivate in the Flipper solenoid subs

RightFlipper.timerinterval = 1
Rightflipper.timerenabled = True

Sub RightFlipper_timer()
	FlipperTricks LeftFlipper, LFPress, LFCount, LFEndAngle, LFState
	FlipperTricks RightFlipper, RFPress, RFCount, RFEndAngle, RFState
	FlipperNudge RightFlipper, RFEndAngle, RFEOSNudge, LeftFlipper, LFEndAngle
	FlipperNudge LeftFlipper, LFEndAngle, LFEOSNudge,  RightFlipper, RFEndAngle
End Sub

Dim LFEOSNudge, RFEOSNudge

Sub FlipperNudge(Flipper1, Endangle1, EOSNudge1, Flipper2, EndAngle2)
	Dim b
	'   Dim BOT
	'   BOT = GetBalls
	
	If Flipper1.currentangle = Endangle1 And EOSNudge1 <> 1 Then
		EOSNudge1 = 1
		'   debug.print Flipper1.currentangle &" = "& Endangle1 &"--"& Flipper2.currentangle &" = "& EndAngle2
		If Flipper2.currentangle = EndAngle2 Then
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper1) Then
					'Debug.Print "ball in flip1. exit"
					Exit Sub
				End If
			Next
			For b = 0 To UBound(gBOT)
				If FlipperTrigger(gBOT(b).x, gBOT(b).y, Flipper2) Then
					gBOT(b).velx = gBOT(b).velx / 1.3
					gBOT(b).vely = gBOT(b).vely - 0.5
				End If
			Next
		End If
	Else
		If Abs(Flipper1.currentangle) > Abs(EndAngle1) + 30 Then EOSNudge1 = 0
	End If
End Sub


Dim FCCDamping: FCCDamping = 0.4

Sub FlipperCradleCollision(ball1, ball2, velocity)
	if velocity < 0.7 then exit sub		'filter out gentle collisions
    Dim DoDamping, coef
    DoDamping = false
    'Check left flipper
    If LeftFlipper.currentangle = LFEndAngle Then
		If FlipperTrigger(ball1.x, ball1.y, LeftFlipper) OR FlipperTrigger(ball2.x, ball2.y, LeftFlipper) Then DoDamping = true
    End If
    'Check right flipper
    If RightFlipper.currentangle = RFEndAngle Then
		If FlipperTrigger(ball1.x, ball1.y, RightFlipper) OR FlipperTrigger(ball2.x, ball2.y, RightFlipper) Then DoDamping = true
    End If
    If DoDamping Then
		coef = FCCDamping
        ball1.velx = ball1.velx * coef: ball1.vely = ball1.vely * coef: ball1.velz = ball1.velz * coef
        ball2.velx = ball2.velx * coef: ball2.vely = ball2.vely * coef: ball2.velz = ball2.velz * coef
    End If
End Sub
	




'*************************************************
'  Check ball distance from Flipper for Rem
'*************************************************

Function Distance(ax,ay,bx,by)
	Distance = Sqr((ax - bx) ^ 2 + (ay - by) ^ 2)
End Function

Function DistancePL(px,py,ax,ay,bx,by) 'Distance between a point and a line where point Is px,py
	DistancePL = Abs((by - ay) * px - (bx - ax) * py + bx * ay - by * ax) / Distance(ax,ay,bx,by)
End Function

Function Radians(Degrees)
	Radians = Degrees * PI / 180
End Function

Function AnglePP(ax,ay,bx,by)
	AnglePP = Atn2((by - ay),(bx - ax)) * 180 / PI
End Function

Function DistanceFromFlipper(ballx, bally, Flipper)
	DistanceFromFlipper = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Flipper.currentangle + 90)) + Flipper.x, Sin(Radians(Flipper.currentangle + 90)) + Flipper.y)
End Function

Function DistanceFromFlipperAngle(ballx, bally, Flipper, Angle)
	DistanceFromFlipperAngle = DistancePL(ballx, bally, Flipper.x, Flipper.y, Cos(Radians(Angle + 90)) + Flipper.x, Sin(Radians(angle + 90)) + Flipper.y)
End Function

Function FlipperTrigger(ballx, bally, Flipper)
	Dim DiffAngle
	DiffAngle = Abs(Flipper.currentangle - AnglePP(Flipper.x, Flipper.y, ballx, bally) - 90)
	If DiffAngle > 180 Then DiffAngle = DiffAngle - 360
	
	If DistanceFromFlipper(ballx,bally,Flipper) < 48 And DiffAngle <= 90 And Distance(ballx,bally,Flipper.x,Flipper.y) < Flipper.Length Then
		FlipperTrigger = True
	Else
		FlipperTrigger = False
	End If
End Function

'*************************************************
'  End - Check ball distance from Flipper for Rem
'*************************************************

Dim LFPress, RFPress, LFCount, RFCount
Dim LFState, RFState
Dim EOST, EOSA,Frampup, FElasticity,FReturn
Dim RFEndAngle, LFEndAngle

Const FlipperCoilRampupMode = 0 '0 = fast, 1 = medium, 2 = slow (tap passes should work)

LFState = 1
RFState = 1
EOST = leftflipper.eostorque
EOSA = leftflipper.eostorqueangle
Frampup = LeftFlipper.rampup
FElasticity = LeftFlipper.elasticity
FReturn = LeftFlipper.return
'Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
Const EOSAnew = 1
Const EOSRampup = 0
Dim SOSRampup
Select Case FlipperCoilRampupMode
	Case 0
		SOSRampup = 2.5
	Case 1
		SOSRampup = 6
	Case 2
		SOSRampup = 8.5
End Select

Const LiveCatch = 16
Const LiveElasticity = 0.45
Const SOSEM = 0.815
'   Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
Const EOSReturn = 0.035  'mid 80's to early 90's
'   Const EOSReturn = 0.025  'mid 90's and later

LFEndAngle = Leftflipper.endangle
RFEndAngle = RightFlipper.endangle

Sub FlipperActivate(Flipper, FlipperPress)
	FlipperPress = 1
	Flipper.Elasticity = FElasticity
	
	Flipper.eostorque = EOST
	Flipper.eostorqueangle = EOSA
End Sub

Sub FlipperDeactivate(Flipper, FlipperPress)
	FlipperPress = 0
	Flipper.eostorqueangle = EOSA
	Flipper.eostorque = EOST * EOSReturn / FReturn
	
	If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 0.1 Then
		Dim b', BOT
		'		BOT = GetBalls
		
		For b = 0 To UBound(gBOT)
			If Distance(gBOT(b).x, gBOT(b).y, Flipper.x, Flipper.y) < 55 Then 'check for cradle
				If gBOT(b).vely >= - 0.4 Then gBOT(b).vely =  - 0.4
			End If
		Next
	End If
End Sub

Sub FlipperTricks (Flipper, FlipperPress, FCount, FEndAngle, FState)
	Dim Dir
	Dir = Flipper.startangle / Abs(Flipper.startangle) '-1 for Right Flipper
	
	If Abs(Flipper.currentangle) > Abs(Flipper.startangle) - 0.05 Then
		If FState <> 1 Then
			Flipper.rampup = SOSRampup
			Flipper.endangle = FEndAngle - 3 * Dir
			Flipper.Elasticity = FElasticity * SOSEM
			FCount = 0
			FState = 1
		End If
	ElseIf Abs(Flipper.currentangle) <= Abs(Flipper.endangle) And FlipperPress = 1 Then
		If FCount = 0 Then FCount = GameTime
		
		If FState <> 2 Then
			Flipper.eostorqueangle = EOSAnew
			Flipper.eostorque = EOSTnew
			Flipper.rampup = EOSRampup
			Flipper.endangle = FEndAngle
			FState = 2
		End If
	ElseIf Abs(Flipper.currentangle) > Abs(Flipper.endangle) + 0.01 And FlipperPress = 1 Then
		If FState <> 3 Then
			Flipper.eostorque = EOST
			Flipper.eostorqueangle = EOSA
			Flipper.rampup = Frampup
			Flipper.Elasticity = FElasticity
			FState = 3
		End If
	End If
End Sub

Const LiveDistanceMin = 5  'minimum distance In vp units from flipper base live catch dampening will occur
Const LiveDistanceMax = 114 'maximum distance in vp units from flipper base live catch dampening will occur (tip protection)
Const BaseDampen = 0.55

Sub CheckLiveCatch(ball, Flipper, FCount, parm) 'Experimental new live catch
    Dim Dir, LiveDist
    Dir = Flipper.startangle / Abs(Flipper.startangle)    '-1 for Right Flipper
    Dim LiveCatchBounce   'If live catch is not perfect, it won't freeze ball totally
    Dim CatchTime
    CatchTime = GameTime - FCount
    LiveDist = Abs(Flipper.x - ball.x)

    If CatchTime <= LiveCatch And parm > 3 And LiveDist > LiveDistanceMin And LiveDist < LiveDistanceMax Then
        If CatchTime <= LiveCatch * 0.5 Then   'Perfect catch only when catch time happens in the beginning of the window
            LiveCatchBounce = 0
        Else
            LiveCatchBounce = Abs((LiveCatch / 2) - CatchTime)  'Partial catch when catch happens a bit late
        End If
        
        If LiveCatchBounce = 0 And ball.velx * Dir > 0 And LiveDist > 30 Then ball.velx = 0

        If ball.velx * Dir > 0 And LiveDist < 30 Then
            ball.velx = BaseDampen * ball.velx
            ball.vely = BaseDampen * ball.vely
            ball.angmomx = BaseDampen * ball.angmomx
            ball.angmomy = BaseDampen * ball.angmomy
            ball.angmomz = BaseDampen * ball.angmomz
        Elseif LiveDist > 30 Then
            ball.vely = LiveCatchBounce * (32 / LiveCatch) ' Multiplier for inaccuracy bounce
            ball.angmomx = 0
            ball.angmomy = 0
            ball.angmomz = 0
        End If
    Else
        If Abs(Flipper.currentangle) <= Abs(Flipper.endangle) + 1 Then FlippersD.Dampenf ActiveBall, parm
    End If
End Sub

'******************************************************
'****  END FLIPPER CORRECTIONS
'******************************************************





'******************************************************
' 	ZDMP:  RUBBER  DAMPENERS
'******************************************************
' These are data mined bounce curves,
' dialed in with the in-game elasticity as much as possible to prevent angle / spin issues.
' Requires tracking ballspeed to calculate COR

Sub dPosts_Hit(idx)
	RubbersD.dampen ActiveBall
	TargetBouncer ActiveBall, 1
End Sub

Sub dSleeves_Hit(idx)
	SleevesD.Dampen ActiveBall
	TargetBouncer ActiveBall, 0.7
End Sub

Dim RubbersD				'frubber
Set RubbersD = New Dampener
RubbersD.name = "Rubbers"
RubbersD.debugOn = False	'shows info in textbox "TBPout"
RubbersD.Print = False	  'debug, reports In debugger (In vel, out cor); cor bounce curve (linear)

'for best results, try to match in-game velocity as closely as possible to the desired curve
'   RubbersD.addpoint 0, 0, 0.935   'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 0, 0, 1.1		 'point# (keep sequential), ballspeed, CoR (elasticity)
RubbersD.addpoint 1, 3.77, 0.97
RubbersD.addpoint 2, 5.76, 0.967	'dont take this as gospel. if you can data mine rubber elasticitiy, please help!
RubbersD.addpoint 3, 15.84, 0.874
RubbersD.addpoint 4, 56, 0.64	   'there's clamping so interpolate up to 56 at least

Dim SleevesD	'this is just rubber but cut down to 85%...
Set SleevesD = New Dampener
SleevesD.name = "Sleeves"
SleevesD.debugOn = False	'shows info in textbox "TBPout"
SleevesD.Print = False	  'debug, reports In debugger (In vel, out cor)
SleevesD.CopyCoef RubbersD, 0.85

'######################### Add new FlippersD Profile
'######################### Adjust these values to increase or lessen the elasticity

Dim FlippersD
Set FlippersD = New Dampener
FlippersD.name = "Flippers"
FlippersD.debugOn = False
FlippersD.Print = False
FlippersD.addpoint 0, 0, 1.1
FlippersD.addpoint 1, 3.77, 0.99
FlippersD.addpoint 2, 6, 0.99

Class Dampener
	Public Print, debugOn   'tbpOut.text
	Public name, Threshold  'Minimum threshold. Useful for Flippers, which don't have a hit threshold.
	Public ModIn, ModOut
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
	End Sub
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Dampen(aBall)
		If threshold Then
			If BallSpeed(aBall) < threshold Then Exit Sub
		End If
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If debugOn Then str = name & " In vel:" & Round(cor.ballvel(aBall.id),2 ) & vbNewLine & "desired cor: " & Round(desiredcor,4) & vbNewLine & _
		"actual cor: " & Round(realCOR,4) & vbNewLine & "ballspeed coef: " & Round(coef, 3) & vbNewLine
		If Print Then Debug.print Round(cor.ballvel(aBall.id),2) & ", " & Round(desiredcor,3)
		
		aBall.velx = aBall.velx * coef
		aBall.vely = aBall.vely * coef
		aBall.velz = aBall.velz * coef
		If debugOn Then TBPout.text = str
	End Sub
	
	Public Sub Dampenf(aBall, parm) 'Rubberizer is handle here
		Dim RealCOR, DesiredCOR, str, coef
		DesiredCor = LinearEnvelope(cor.ballvel(aBall.id), ModIn, ModOut )
		RealCOR = BallSpeed(aBall) / (cor.ballvel(aBall.id) + 0.0001)
		coef = desiredcor / realcor
		If Abs(aball.velx) < 2 And aball.vely < 0 And aball.vely >  - 3.75 Then
			aBall.velx = aBall.velx * coef
			aBall.vely = aBall.vely * coef
		End If
	End Sub
	
	Public Sub CopyCoef(aObj, aCoef) 'alternative addpoints, copy with coef
		Dim x
		For x = 0 To UBound(aObj.ModIn)
			addpoint x, aObj.ModIn(x), aObj.ModOut(x) * aCoef
		Next
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
End Class

'******************************************************
'  TRACK ALL BALL VELOCITIES
'  FOR RUBBER DAMPENER AND DROP TARGETS
'******************************************************

Dim cor
Set cor = New CoRTracker

Class CoRTracker
	Public ballvel, ballvelx, ballvely
	
	Private Sub Class_Initialize
		ReDim ballvel(0)
		ReDim ballvelx(0)
		ReDim ballvely(0)
	End Sub
	
	Public Sub Update()	'tracks in-ball-velocity
		Dim str, b, AllBalls, highestID
		allBalls = GetBalls
		
		For Each b In allballs
			If b.id >= HighestID Then highestID = b.id
		Next
		
		If UBound(ballvel) < highestID Then ReDim ballvel(highestID)	'set bounds
		If UBound(ballvelx) < highestID Then ReDim ballvelx(highestID)	'set bounds
		If UBound(ballvely) < highestID Then ReDim ballvely(highestID)	'set bounds
		
		For Each b In allballs
			ballvel(b.id) = BallSpeed(b)
			ballvelx(b.id) = b.velx
			ballvely(b.id) = b.vely
		Next
	End Sub
End Class

' Note, cor.update must be called in a 10 ms timer. The example table uses the GameTimer for this purpose, but sometimes a dedicated timer call RDampen is used.
'
'Sub RDampen_Timer
'	Cor.Update
'End Sub

'******************************************************
'****  END PHYSICS DAMPENERS
'******************************************************



'******************************************************
' 	ZBOU: VPW TargetBouncer for targets and posts by Iaakki, Wrd1972, Apophis
'******************************************************

Const TargetBouncerEnabled = 1	  '0 = normal standup targets, 1 = bouncy targets
Const TargetBouncerFactor = 0.9	 'Level of bounces. Recommmended value of 0.7-1

Sub TargetBouncer(aBall,defvalue)
	Dim zMultiplier, vel, vratio
	If TargetBouncerEnabled = 1 And aball.z < 30 Then
		'   debug.print "velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		vel = BallSpeed(aBall)
		If aBall.velx = 0 Then vratio = 1 Else vratio = aBall.vely / aBall.velx
		Select Case Int(Rnd * 6) + 1
			Case 1
				zMultiplier = 0.2 * defvalue
			Case 2
				zMultiplier = 0.25 * defvalue
			Case 3
				zMultiplier = 0.3 * defvalue
			Case 4
				zMultiplier = 0.4 * defvalue
			Case 5
				zMultiplier = 0.45 * defvalue
			Case 6
				zMultiplier = 0.5 * defvalue
		End Select
		aBall.velz = Abs(vel * zMultiplier * TargetBouncerFactor)
		aBall.velx = Sgn(aBall.velx) * Sqr(Abs((vel ^ 2 - aBall.velz ^ 2) / (1 + vratio ^ 2)))
		aBall.vely = aBall.velx * vratio
		'   debug.print "---> velx: " & aball.velx & " vely: " & aball.vely & " velz: " & aball.velz
		'   debug.print "conservation check: " & BallSpeed(aBall)/vel
	End If
End Sub

'Add targets or posts to the TargetBounce collection if you want to activate the targetbouncer code from them
Sub TargetBounce_Hit(idx)
	TargetBouncer ActiveBall, 1
End Sub



'******************************************************
'	ZSSC: SLINGSHOT CORRECTION FUNCTIONS by apophis
'******************************************************
' To add these slingshot corrections:
'	 - On the table, add the endpoint primitives that define the two ends of the Slingshot
'	 - Initialize the SlingshotCorrection objects in InitSlingCorrection
'	 - Call the .VelocityCorrect methods from the respective _Slingshot event sub

Dim LS
Set LS = New SlingshotCorrection
Dim RS
Set RS = New SlingshotCorrection

InitSlingCorrection

Sub InitSlingCorrection
	LS.Object = LeftSlingshot
	LS.EndPoint1 = EndPoint1LS
	LS.EndPoint2 = EndPoint2LS
	
	RS.Object = RightSlingshot
	RS.EndPoint1 = EndPoint1RS
	RS.EndPoint2 = EndPoint2RS
	
	'Slingshot angle corrections (pt, BallPos in %, Angle in deg)
	' These values are best guesses. Retune them if needed based on specific table research.
	AddSlingsPt 0, 0.00, - 4
	AddSlingsPt 1, 0.45, - 7
	AddSlingsPt 2, 0.48,	0
	AddSlingsPt 3, 0.52,	0
	AddSlingsPt 4, 0.55,	7
	AddSlingsPt 5, 1.00,	4
End Sub

Sub AddSlingsPt(idx, aX, aY)		'debugger wrapper for adjusting flipper script In-game
	Dim a
	a = Array(LS, RS)
	Dim x
	For Each x In a
		x.addpoint idx, aX, aY
	Next
End Sub

'' The following sub are needed, however they may exist somewhere else in the script. Uncomment below if needed
'Dim PI: PI = 4*Atn(1)
'Function dSin(degrees)
'	dsin = sin(degrees * Pi/180)
'End Function
'Function dCos(degrees)
'	dcos = cos(degrees * Pi/180)
'End Function
'
'Function RotPoint(x,y,angle)
'	dim rx, ry
'	rx = x*dCos(angle) - y*dSin(angle)
'	ry = x*dSin(angle) + y*dCos(angle)
'	RotPoint = Array(rx,ry)
'End Function

Class SlingshotCorrection
	Public DebugOn, Enabled
	Private Slingshot, SlingX1, SlingX2, SlingY1, SlingY2
	
	Public ModIn, ModOut
	
	Private Sub Class_Initialize
		ReDim ModIn(0)
		ReDim Modout(0)
		Enabled = True
	End Sub
	
	Public Property Let Object(aInput)
		Set Slingshot = aInput
	End Property
	
	Public Property Let EndPoint1(aInput)
		SlingX1 = aInput.x
		SlingY1 = aInput.y
	End Property
	
	Public Property Let EndPoint2(aInput)
		SlingX2 = aInput.x
		SlingY2 = aInput.y
	End Property
	
	Public Sub AddPoint(aIdx, aX, aY)
		ShuffleArrays ModIn, ModOut, 1
		ModIn(aIDX) = aX
		ModOut(aIDX) = aY
		ShuffleArrays ModIn, ModOut, 0
		If GameTime > 100 Then Report
	End Sub
	
	Public Sub Report() 'debug, reports all coords in tbPL.text
		If Not debugOn Then Exit Sub
		Dim a1, a2
		a1 = ModIn
		a2 = ModOut
		Dim str, x
		For x = 0 To UBound(a1)
			str = str & x & ": " & Round(a1(x),4) & ", " & Round(a2(x),4) & vbNewLine
		Next
		TBPout.text = str
	End Sub
	
	
	Public Sub VelocityCorrect(aBall)
		Dim BallPos, XL, XR, YL, YR
		
		'Assign right and left end points
		If SlingX1 < SlingX2 Then
			XL = SlingX1
			YL = SlingY1
			XR = SlingX2
			YR = SlingY2
		Else
			XL = SlingX2
			YL = SlingY2
			XR = SlingX1
			YR = SlingY1
		End If
		
		'Find BallPos = % on Slingshot
		If Not IsEmpty(aBall.id) Then
			If Abs(XR - XL) > Abs(YR - YL) Then
				BallPos = PSlope(aBall.x, XL, 0, XR, 1)
			Else
				BallPos = PSlope(aBall.y, YL, 0, YR, 1)
			End If
			If BallPos < 0 Then BallPos = 0
			If BallPos > 1 Then BallPos = 1
		End If
		
		'Velocity angle correction
		If Not IsEmpty(ModIn(0) ) Then
			Dim Angle, RotVxVy
			Angle = LinearEnvelope(BallPos, ModIn, ModOut)
			'   debug.print " BallPos=" & BallPos &" Angle=" & Angle
			'   debug.print " BEFORE: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
			RotVxVy = RotPoint(aBall.Velx,aBall.Vely,Angle)
			If Enabled Then aBall.Velx = RotVxVy(0)
			If Enabled Then aBall.Vely = RotVxVy(1)
			'   debug.print " AFTER: aBall.Velx=" & aBall.Velx &" aBall.Vely" & aBall.Vely
			'   debug.print " "
		End If
	End Sub
End Class






'******************************************************
' 	ZRDT:  DROP TARGETS by Rothbauerw
'******************************************************
' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' to move through the target enabling the ability to score more than one target with a well placed shot.
' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
'
' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' DTAnim.interval = 10
' DTAnim.enabled = True

' Sub DTAnim_Timer
' 	DoDTAnim
'	DoSTAnim
' End Sub

' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' target mesh in this table's example. It uses the same texture map as the VP drop targets.
'
' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' The target should animate backwards using transy.
'
' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' (Other methods will work as well, but this is easy for even non-blender users to do)
' 1) Open a new blank table. Delete everything off the table in editor.
' 2) Copy and paste the VP target from your table into this blank table.
' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' 4) Under the file menu, select Export "OBJ Mesh"
' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also 
'    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
'    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
'    The primitive will use the same target texture as the VP target object. 
'
' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' The 100 and 200 will be removed when setting the switch value for the target.

'******************************************************
'  DROP TARGETS INITIALIZATION
'******************************************************

Class DropTarget
  Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  Public Property Let Secondary(input): Set m_secondary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  Public Property Let IsDropped(input): m_isDropped = input: End Property

  Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    Set m_primary = primary
    Set m_secondary = secondary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate
    m_isDropped = isDropped

    Set Init = Me
  End Function
End Class

'Define a variable for each drop target
Dim DT1, DT2, DT3

'Set array with drop target objects
'
'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
'   primary:	primary target wall to determine drop
'   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
'   prim:	   primitive target used for visuals and animation
'				   IMPORTANT!!!
'				   rotz must be used for orientation
'				   rotx to bend the target back
'				   transz to move it up and down
'				   the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
'   switch:	 ROM switch number
'   animate:	Array slot for handling the animation instrucitons, set to 0
'				   Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
'   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
'					Use the function DTDropped(switchid) to check a target's drop status.

Set DT1 = (new DropTarget)(sw1, sw1a, sw1p, 1, 0, False)
Set DT2 = (new DropTarget)(sw2, sw2a, sw2p, 2, 0, False)
Set DT3 = (new DropTarget)(sw3, sw3a, sw3p, 3, 0, False)

Dim DTArray
DTArray = Array(DT1, DT2, DT3)

'Configure the behavior of Drop Targets.
Const DTDropSpeed = 90 'in milliseconds
Const DTDropUpSpeed = 40 'in milliseconds
Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
Const DTMaxBend = 8 'max degrees primitive rotates when hit
Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

'******************************************************
'  DROP TARGETS FUNCTIONS
'******************************************************

Sub DTHit(switch)
	Dim i
	i = DTArrayID(switch)
	
	PlayTargetSound
	DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
	If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
		DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
	End If
	DoDTAnim
End Sub

Sub DTRaise(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate =  - 1
	DoDTAnim
End Sub

Sub DTDrop(switch)
	Dim i
	i = DTArrayID(switch)
	
	DTArray(i).animate = 1
	DoDTAnim
End Sub

Function DTArrayID(switch)
	Dim i
	For i = 0 To UBound(DTArray)
		If DTArray(i).sw = switch Then
			DTArrayID = i
			Exit Function
		End If
	Next
End Function

Sub DTBallPhysics(aBall, angle, mass)
	Dim rangle,bangle,calc1, calc2, calc3
	rangle = (angle - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	
	calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
	calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)
	
	aBall.velx = calc1 * Cos(rangle) + calc2
	aBall.vely = calc1 * Sin(rangle) + calc3
End Sub

'Check if target is hit on it's face or sides and whether a 'brick' occurred
Function DTCheckBrick(aBall, dtprim)
	Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
	rangle = (dtprim.rotz - 90) * 3.1416 / 180
	rangle2 = dtprim.rotz * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)
	
	Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
	Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)
	
	cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)
	
	perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	If perpvel > 0 And  perpvelafter <= 0 Then
		If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
			DTCheckBrick = 3
		Else
			DTCheckBrick = 1
		End If
	ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		DTCheckBrick = 4
	Else
		DTCheckBrick = 0
	End If
End Function

Sub DoDTAnim()
	Dim i
	For i = 0 To UBound(DTArray)
		DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
	Next
End Sub

Function DTAnimate(primary, secondary, prim, switch, animate)
	Dim transz, switchid
	Dim animtime, rangle
	
	switchid = switch
	
	Dim ind
	ind = DTArrayID(switchid)
	
	rangle = prim.rotz * PI / 180
	
	DTAnimate = animate
	
	If animate = 0 Then
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	ElseIf primary.uservalue = 0 Then
		primary.uservalue = GameTime
	End If
	
	animtime = GameTime - primary.uservalue
	
	If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
		primary.collidable = 0
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
		DTAnimate = animate
		Exit Function
	ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
		primary.collidable = 0
		If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
		animate = 2
		SoundDropTargetDrop prim
	End If
	
	If animate = 2 Then
		transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
		If prim.transz >  - DTDropUnits  Then
			prim.transz = transz
		End If
		
		prim.rotx = DTMaxBend * Cos(rangle) / 2
		prim.roty = DTMaxBend * Sin(rangle) / 2
		
		If prim.transz <= - DTDropUnits Then
			prim.transz =  - DTDropUnits
			secondary.collidable = 0
			DTArray(ind).isDropped = True 'Mark target as dropped
			If UsingROM Then
				controller.Switch(Switchid mod 100) = 1
			Else
				DTAction switchid
			End If
			primary.uservalue = 0
			DTAnimate = 0
			Exit Function
		Else
			DTAnimate = 2
			Exit Function
		End If
	End If
	
	If animate = 3 And animtime < DTDropDelay Then
		primary.collidable = 0
		secondary.collidable = 1
		prim.rotx = DTMaxBend * Cos(rangle)
		prim.roty = DTMaxBend * Sin(rangle)
	ElseIf animate = 3 And animtime > DTDropDelay Then
		primary.collidable = 1
		secondary.collidable = 0
		prim.rotx = 0
		prim.roty = 0
		primary.uservalue = 0
		DTAnimate = 0
		Exit Function
	End If
	
	If animate =  - 1 Then
		transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1
		
		If prim.transz =  - DTDropUnits Then
			Dim b
			'Dim gBOT
			'gBOT = GetBalls
			
			For b = 0 To UBound(gBOT)
				If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
					gBOT(b).velz = 20
				End If
			Next
		End If
		
		If prim.transz < 0 Then
			prim.transz = transz
		ElseIf transz > 0 Then
			prim.transz = transz
		End If
		
		If prim.transz > DTDropUpUnits Then
			DTAnimate =  - 2
			prim.transz = DTDropUpUnits
			prim.rotx = 0
			prim.roty = 0
			primary.uservalue = GameTime
		End If
		primary.collidable = 0
		secondary.collidable = 1
		DTArray(ind).isDropped = False 'Mark target as not dropped
		If UsingROM Then controller.Switch(Switchid mod 100) = 0
	End If
	
	If animate =  - 2 And animtime > DTRaiseDelay Then
		prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
		If prim.transz < 0 Then
			prim.transz = 0
			primary.uservalue = 0
			DTAnimate = 0
			
			primary.collidable = 1
			secondary.collidable = 0
		End If
	End If
End Function

Function DTDropped(switchid)
	Dim ind
	ind = DTArrayID(switchid)
	
	DTDropped = DTArray(ind).isDropped
End Function

Sub DTAction(switchid)
	Select Case switchid
		Case 1
			Addscore 1000
			ShadowDT(0).visible = False
			
		Case 2
			Addscore 1000
			ShadowDT(1).visible = False
			
		Case 3
			Addscore 1000
			ShadowDT(2).visible = False
	End Select
End Sub


'******************************************************
'****  END DROP TARGETS
'******************************************************




'******************************************************
'	ZRST: STAND-UP TARGETS by Rothbauerw
'******************************************************

Class StandupTarget
  Private m_primary, m_prim, m_sw, m_animate

  Public Property Get Primary(): Set Primary = m_primary: End Property
  Public Property Let Primary(input): Set m_primary = input: End Property

  Public Property Get Prim(): Set Prim = m_prim: End Property
  Public Property Let Prim(input): Set m_prim = input: End Property

  Public Property Get Sw(): Sw = m_sw: End Property
  Public Property Let Sw(input): m_sw = input: End Property

  Public Property Get Animate(): Animate = m_animate: End Property
  Public Property Let Animate(input): m_animate = input: End Property

  Public default Function init(primary, prim, sw, animate)
    Set m_primary = primary
    Set m_prim = prim
    m_sw = sw
    m_animate = animate

    Set Init = Me
  End Function
End Class

'Define a variable for each stand-up target
Dim ST11, ST12, ST13

'Set array with stand-up target objects
'
'StandupTargetvar = Array(primary, prim, swtich)
'   primary:	vp target to determine target hit
'   prim:	   primitive target used for visuals and animation
'				   IMPORTANT!!!
'				   transy must be used to offset the target animation
'   switch:	 ROM switch number
'   animate:	Arrary slot for handling the animation instrucitons, set to 0
'
'You will also need to add a secondary hit object for each stand up (name sw11o, sw12o, and sw13o on the example Table1)
'these are inclined primitives to simulate hitting a bent target and should provide so z velocity on high speed impacts


Set ST11 = (new StandupTarget)(sw11, psw11,11, 0)
Set ST12 = (new StandupTarget)(sw12, psw12,12, 0)
Set ST13 = (new StandupTarget)(sw13, psw13,13, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
STArray = Array(ST11, ST12, ST13)

'Configure the behavior of Stand-up Targets
Const STAnimStep = 1.5  'vpunits per animation step (control return to Start)
Const STMaxOffset = 9   'max vp units target moves when hit

Const STMass = 0.2	  'Mass of the Stand-up Target (between 0 and 1), higher values provide more resistance

'******************************************************
'				STAND-UP TARGETS FUNCTIONS
'******************************************************

Sub STHit(switch)
	Dim i
	i = STArrayID(switch)
	
	PlayTargetSound
	STArray(i).animate = STCheckHit(ActiveBall,STArray(i).primary)
	
	If STArray(i).animate <> 0 Then
		DTBallPhysics ActiveBall, STArray(i).primary.orientation, STMass
	End If
	DoSTAnim
End Sub

Function STArrayID(switch)
	Dim i
	For i = 0 To UBound(STArray)
		If STArray(i).sw = switch Then
			STArrayID = i
			Exit Function
		End If
	Next
End Function

Function STCheckHit(aBall, target) 'Check if target is hit on it's face
	Dim bangle, bangleafter, rangle, rangle2, perpvel, perpvelafter, paravel, paravelafter
	rangle = (target.orientation - 90) * 3.1416 / 180
	bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	bangleafter = Atn2(aBall.vely,aball.velx)
	
	perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	If perpvel > 0 And  perpvelafter <= 0 Then
		STCheckHit = 1
	ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		STCheckHit = 1
	Else
		STCheckHit = 0
	End If
End Function

Sub DoSTAnim()
	Dim i
	For i = 0 To UBound(STArray)
		STArray(i).animate = STAnimate(STArray(i).primary,STArray(i).prim,STArray(i).sw,STArray(i).animate)
	Next
End Sub

Function STAnimate(primary, prim, switch,  animate)
	Dim animtime
	
	STAnimate = animate
	
	If animate = 0  Then
		primary.uservalue = 0
		STAnimate = 0
		Exit Function
	ElseIf primary.uservalue = 0 Then
		primary.uservalue = GameTime
	End If
	
	animtime = GameTime - primary.uservalue
	
	If animate = 1 Then
		primary.collidable = 0
		prim.transy =  - STMaxOffset
		If UsingROM Then
			vpmTimer.PulseSw switch mod 100
		Else
			STAction switch
		End If
		STAnimate = 2
		Exit Function
	ElseIf animate = 2 Then
		prim.transy = prim.transy + STAnimStep
		If prim.transy >= 0 Then
			prim.transy = 0
			primary.collidable = 1
			STAnimate = 0
			Exit Function
		Else
			STAnimate = 2
		End If
	End If
End Function


Sub STAction(Switch)
	Select Case Switch
		Case 11
			Addscore 1000
			Flash1 True 'Demo of the flasher
			vpmTimer.AddTimer 150,"Flash1 False'"   'Disable the flash after short time, just like a ROM would do
			
		Case 12
			Addscore 1000
			Flash2 True 'Demo of the flasher
			vpmTimer.AddTimer 150,"Flash2 False'"   'Disable the flash after short time, just like a ROM would do
			
		Case 13
			Addscore 1000
			Flash3 True 'Demo of the flasher
			vpmTimer.AddTimer 150,"Flash3 False'"   'Disable the flash after short time, just like a ROM would do
	End Select
End Sub

'******************************************************
'****   END STAND-UP TARGETS
'******************************************************




'******************************************************
'	ZBRL:  BALL ROLLING AND DROP SOUNDS
'******************************************************

' Be sure to call RollingUpdate in a timer with a 10ms interval see the GameTimer_Timer() sub

ReDim rolling(tnob)
InitRolling

Dim DropCount
ReDim DropCount(tnob)

Sub InitRolling
	Dim i
	For i = 0 To tnob
		rolling(i) = False
	Next
End Sub

Sub RollingUpdate()
	Dim b
	'   Dim BOT
	'   BOT = GetBalls
	
	' stop the sound of deleted balls
	For b = UBound(gBOT) + 1 To tnob - 1
		rolling(b) = False
		StopSound("BallRoll_" & b)
	Next
	
	' exit the sub if no balls on the table
	If UBound(gBOT) =  - 1 Then Exit Sub
	
	' play the rolling sound for each ball
	For b = 0 To UBound(gBOT)
		If BallVel(gBOT(b)) > 1 And gBOT(b).z < 30 Then
			rolling(b) = True
			PlaySound ("BallRoll_" & b), - 1, VolPlayfieldRoll(gBOT(b)) * BallRollVolume * VolumeDial, AudioPan(gBOT(b)), 0, PitchPlayfieldRoll(gBOT(b)), 1, 0, AudioFade(gBOT(b))
		Else
			If rolling(b) = True Then
				StopSound("BallRoll_" & b)
				rolling(b) = False
			End If
		End If
		
		' Ball Drop Sounds
		If gBOT(b).VelZ <  - 1 And gBOT(b).z < 55 And gBOT(b).z > 27 Then 'height adjust for ball drop sounds
			If DropCount(b) >= 5 Then
				DropCount(b) = 0
				If gBOT(b).velz >  - 7 Then
					RandomSoundBallBouncePlayfieldSoft gBOT(b)
				Else
					RandomSoundBallBouncePlayfieldHard gBOT(b)
				End If
			End If
		End If
		
		If DropCount(b) < 5 Then
			DropCount(b) = DropCount(b) + 1
		End If
	Next
End Sub

'******************************************************
'****  END BALL ROLLING AND DROP SOUNDS
'******************************************************




'******************************************************
' 	ZRRL: RAMP ROLLING SFX
'******************************************************

'Ball tracking ramp SFX 1.0
'   Reqirements:
'		  * Import A Sound File for each ball on the table for plastic ramps.  Call It RampLoop<Ball_Number> ex: RampLoop1, RampLoop2, ...
'		  * Import a Sound File for each ball on the table for wire ramps. Call it WireLoop<Ball_Number> ex: WireLoop1, WireLoop2, ...
'		  * Create a Timer called RampRoll, that is enabled, with a interval of 100
'		  * Set RampBAlls and RampType variable to Total Number of Balls
'	Usage:
'		  * Setup hit events and call WireRampOn True or WireRampOn False (True = Plastic ramp, False = Wire Ramp)
'		  * To stop tracking ball
'				 * call WireRampOff
'				 * Otherwise, the ball will auto remove if it's below 30 vp units
'

Dim RampMinLoops
RampMinLoops = 4

' RampBalls
' Setup:  Set the array length of x in RampBalls(x,2) Total Number of Balls on table + 1:  if tnob = 5, then RampBalls(6,2)
Dim RampBalls(6,2)
'x,0 = ball x,1 = ID, 2 = Protection against ending early (minimum amount of updates)

'0,0 is boolean on/off, 0,1 unused for now
RampBalls(0,0) = False

' RampType
' Setup: Set this array to the number Total number of balls that can be tracked at one time + 1.  5 ball multiball then set value to 6
' Description: Array type indexed on BallId and a values used to deterimine what type of ramp the ball is on: False = Wire Ramp, True = Plastic Ramp
Dim RampType(6)

Sub WireRampOn(input)
	Waddball ActiveBall, input
	RampRollUpdate
End Sub

Sub WireRampOff()
	WRemoveBall ActiveBall.ID
End Sub

' WaddBall (Active Ball, Boolean)
Sub Waddball(input, RampInput) 'This subroutine is called from WireRampOn to Add Balls to the RampBalls Array
	' This will loop through the RampBalls array checking each element of the array x, position 1
	' To see if the the ball was already added to the array.
	' If the ball is found then exit the subroutine
	Dim x
	For x = 1 To UBound(RampBalls)	'Check, don't add balls twice
		If RampBalls(x, 1) = input.id Then
			If Not IsEmpty(RampBalls(x,1) ) Then Exit Sub	'Frustating issue with BallId 0. Empty variable = 0
		End If
	Next
	
	' This will itterate through the RampBalls Array.
	' The first time it comes to a element in the array where the Ball Id (Slot 1) is empty.  It will add the current ball to the array
	' The RampBalls assigns the ActiveBall to element x,0 and ball id of ActiveBall to 0,1
	' The RampType(BallId) is set to RampInput
	' RampBalls in 0,0 is set to True, this will enable the timer and the timer is also turned on
	For x = 1 To UBound(RampBalls)
		If IsEmpty(RampBalls(x, 1)) Then
			Set RampBalls(x, 0) = input
			RampBalls(x, 1) = input.ID
			RampType(x) = RampInput
			RampBalls(x, 2) = 0
			'exit For
			RampBalls(0,0) = True
			RampRoll.Enabled = 1	 'Turn on timer
			'RampRoll.Interval = RampRoll.Interval 'reset timer
			Exit Sub
		End If
		If x = UBound(RampBalls) Then	 'debug
			Debug.print "WireRampOn error, ball queue Is full: " & vbNewLine & _
			RampBalls(0, 0) & vbNewLine & _
			TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & "type:" & RampType(1) & vbNewLine & _
			TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & "type:" & RampType(2) & vbNewLine & _
			TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & "type:" & RampType(3) & vbNewLine & _
			TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & "type:" & RampType(4) & vbNewLine & _
			TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & "type:" & RampType(5) & vbNewLine & _
			" "
		End If
	Next
End Sub

' WRemoveBall (BallId)
Sub WRemoveBall(ID) 'This subroutine is called from the RampRollUpdate subroutine and is used to remove and stop the ball rolling sounds
	'   Debug.Print "In WRemoveBall() + Remove ball from loop array"
	Dim ballcount
	ballcount = 0
	Dim x
	For x = 1 To UBound(RampBalls)
		If ID = RampBalls(x, 1) Then 'remove ball
			Set RampBalls(x, 0) = Nothing
			RampBalls(x, 1) = Empty
			RampType(x) = Empty
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
		'if RampBalls(x,1) = Not IsEmpty(Rampballs(x,1) then ballcount = ballcount + 1
		If Not IsEmpty(Rampballs(x,1)) Then ballcount = ballcount + 1
	Next
	If BallCount = 0 Then RampBalls(0,0) = False	'if no balls in queue, disable timer update
End Sub

Sub RampRoll_Timer()
	RampRollUpdate
End Sub

Sub RampRollUpdate()	'Timer update
	Dim x
	For x = 1 To UBound(RampBalls)
		If Not IsEmpty(RampBalls(x,1) ) Then
			If BallVel(RampBalls(x,0) ) > 1 Then ' if ball is moving, play rolling sound
				If RampType(x) Then
					PlaySound("RampLoop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitchV(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
					StopSound("wireloop" & x)
				Else
					StopSound("RampLoop" & x)
					PlaySound("wireloop" & x), - 1, VolPlayfieldRoll(RampBalls(x,0)) * RampRollVolume * VolumeDial, AudioPan(RampBalls(x,0)), 0, BallPitch(RampBalls(x,0)), 1, 0, AudioFade(RampBalls(x,0))
				End If
				RampBalls(x, 2) = RampBalls(x, 2) + 1
			Else
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
			End If
			If RampBalls(x,0).Z < 30 And RampBalls(x, 2) > RampMinLoops Then	'if ball is on the PF, remove  it
				StopSound("RampLoop" & x)
				StopSound("wireloop" & x)
				Wremoveball RampBalls(x,1)
			End If
		Else
			StopSound("RampLoop" & x)
			StopSound("wireloop" & x)
		End If
	Next
	If Not RampBalls(0,0) Then RampRoll.enabled = 0
End Sub

' This can be used to debug the Ramp Roll time.  You need to enable the tbWR timer on the TextBox
Sub tbWR_Timer()	'debug textbox
	Me.text = "on? " & RampBalls(0, 0) & " timer: " & RampRoll.Enabled & vbNewLine & _
	"1 " & TypeName(RampBalls(1, 0)) & " ID:" & RampBalls(1, 1) & " type:" & RampType(1) & " Loops:" & RampBalls(1, 2) & vbNewLine & _
	"2 " & TypeName(RampBalls(2, 0)) & " ID:" & RampBalls(2, 1) & " type:" & RampType(2) & " Loops:" & RampBalls(2, 2) & vbNewLine & _
	"3 " & TypeName(RampBalls(3, 0)) & " ID:" & RampBalls(3, 1) & " type:" & RampType(3) & " Loops:" & RampBalls(3, 2) & vbNewLine & _
	"4 " & TypeName(RampBalls(4, 0)) & " ID:" & RampBalls(4, 1) & " type:" & RampType(4) & " Loops:" & RampBalls(4, 2) & vbNewLine & _
	"5 " & TypeName(RampBalls(5, 0)) & " ID:" & RampBalls(5, 1) & " type:" & RampType(5) & " Loops:" & RampBalls(5, 2) & vbNewLine & _
	"6 " & TypeName(RampBalls(6, 0)) & " ID:" & RampBalls(6, 1) & " type:" & RampType(6) & " Loops:" & RampBalls(6, 2) & vbNewLine & _
	" "
End Sub

Function BallPitch(ball) ' Calculates the pitch of the sound based on the ball speed
	BallPitch = pSlope(BallVel(ball), 1, - 1000, 60, 10000)
End Function

Function BallPitchV(ball) ' Calculates the pitch of the sound based on the ball speed Variation
	BallPitchV = pSlope(BallVel(ball), 1, - 4000, 60, 7000)
End Function

Sub RandomSoundRampStop(obj)
	Select Case Int(rnd*3)
		Case 0: PlaySoundAtVol "wireramp_stop1", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
		Case 1: PlaySoundAtVol "wireramp_stop2", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
		Case 2: PlaySoundAtVol "wireramp_stop3", obj, 0.2*VolumeDial:PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6
	End Select
End Sub

'******************************************************
'**** END RAMP ROLLING SFX
'******************************************************





'******************************************************
' 	ZFLE:  FLEEP MECHANICAL SOUNDS
'******************************************************

' This part in the script is an entire block that is dedicated to the physics sound system.
' Various scripts and sounds that may be pretty generic and could suit other WPC systems, but the most are tailored specifically for the TOM table

' Many of the sounds in this package can be added by creating collections and adding the appropriate objects to those collections.
' Create the following new collections:
'	 Metals (all metal objects, metal walls, metal posts, metal wire guides)
'	 Apron (the apron walls and plunger wall)
'	 Walls (all wood or plastic walls)
'	 Rollovers (wire rollover triggers, star triggers, or button triggers)
'	 Targets (standup or drop targets, these are hit sounds only ... you will want to add separate dropping sounds for drop targets)
'	 Gates (plate gates)
'	 GatesWire (wire gates)
'	 Rubbers (all rubbers including posts, sleeves, pegs, and bands)
' When creating the collections, make sure "Fire events for this collection" is checked.
' You'll also need to make sure "Has Hit Event" is checked for each object placed in these collections (not necessary for gates and triggers).
' Once the collections and objects are added, the save, close, and restart VPX.
'
' Many places in the script need to be modified to include the correct sound effect subroutine calls. The tutorial videos linked below demonstrate
' how to make these updates. But in summary the following needs to be updated:
'	- Nudging, plunger, coin-in, start button sounds will be added to the keydown and keyup subs.
'	- Flipper sounds in the flipper solenoid subs. Flipper collision sounds in the flipper collide subs.
'	- Bumpers, slingshots, drain, ball release, knocker, spinner, and saucers in their respective subs
'	- Ball rolling sounds sub
'
' Tutorial videos by Apophis
' Audio : Adding Fleep Part 1					https://youtu.be/rG35JVHxtx4?si=zdN9W4cZWEyXbOz_
' Audio : Adding Fleep Part 2					https://youtu.be/dk110pWMxGo?si=2iGMImXXZ0SFKVCh
' Audio : Adding Fleep Part 3					https://youtu.be/ESXWGJZY_EI?si=6D20E2nUM-xAw7xy


'///////////////////////////////  SOUNDS PARAMETERS  //////////////////////////////
Dim GlobalSoundLevel, CoinSoundLevel, PlungerReleaseSoundLevel, PlungerPullSoundLevel, NudgeLeftSoundLevel
Dim NudgeRightSoundLevel, NudgeCenterSoundLevel, StartButtonSoundLevel, RollingSoundFactor

CoinSoundLevel = 1					  'volume level; range [0, 1]
NudgeLeftSoundLevel = 1				 'volume level; range [0, 1]
NudgeRightSoundLevel = 1				'volume level; range [0, 1]
NudgeCenterSoundLevel = 1			   'volume level; range [0, 1]
StartButtonSoundLevel = 0.1			 'volume level; range [0, 1]
PlungerReleaseSoundLevel = 0.8 '1 wjr   'volume level; range [0, 1]
PlungerPullSoundLevel = 1			   'volume level; range [0, 1]
RollingSoundFactor = 1.1 / 5

'///////////////////////-----Solenoids, Kickers and Flash Relays-----///////////////////////
Dim FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel, FlipperUpAttackLeftSoundLevel, FlipperUpAttackRightSoundLevel
Dim FlipperUpSoundLevel, FlipperDownSoundLevel, FlipperLeftHitParm, FlipperRightHitParm
Dim SlingshotSoundLevel, BumperSoundFactor, KnockerSoundLevel

FlipperUpAttackMinimumSoundLevel = 0.010		'volume level; range [0, 1]
FlipperUpAttackMaximumSoundLevel = 0.635		'volume level; range [0, 1]
FlipperUpSoundLevel = 1.0					   'volume level; range [0, 1]
FlipperDownSoundLevel = 0.45					'volume level; range [0, 1]
FlipperLeftHitParm = FlipperUpSoundLevel		'sound helper; not configurable
FlipperRightHitParm = FlipperUpSoundLevel	   'sound helper; not configurable
SlingshotSoundLevel = 0.95					  'volume level; range [0, 1]
BumperSoundFactor = 4.25						'volume multiplier; must not be zero
KnockerSoundLevel = 1						   'volume level; range [0, 1]

'///////////////////////-----Ball Drops, Bumps and Collisions-----///////////////////////
Dim RubberStrongSoundFactor, RubberWeakSoundFactor, RubberFlipperSoundFactor,BallWithBallCollisionSoundFactor
Dim BallBouncePlayfieldSoftFactor, BallBouncePlayfieldHardFactor, PlasticRampDropToPlayfieldSoundLevel, WireRampDropToPlayfieldSoundLevel, DelayedBallDropOnPlayfieldSoundLevel
Dim WallImpactSoundFactor, MetalImpactSoundFactor, SubwaySoundLevel, SubwayEntrySoundLevel, ScoopEntrySoundLevel
Dim SaucerLockSoundLevel, SaucerKickSoundLevel

BallWithBallCollisionSoundFactor = 3.2		  'volume multiplier; must not be zero
RubberStrongSoundFactor = 0.055 / 5			 'volume multiplier; must not be zero
RubberWeakSoundFactor = 0.075 / 5			   'volume multiplier; must not be zero
RubberFlipperSoundFactor = 0.075 / 5			'volume multiplier; must not be zero
BallBouncePlayfieldSoftFactor = 0.025		   'volume multiplier; must not be zero
BallBouncePlayfieldHardFactor = 0.025		   'volume multiplier; must not be zero
DelayedBallDropOnPlayfieldSoundLevel = 0.8	  'volume level; range [0, 1]
WallImpactSoundFactor = 0.075				   'volume multiplier; must not be zero
MetalImpactSoundFactor = 0.075 / 3
SaucerLockSoundLevel = 0.8
SaucerKickSoundLevel = 0.8

'///////////////////////-----Gates, Spinners, Rollovers and Targets-----///////////////////////

Dim GateSoundLevel, TargetSoundFactor, SpinnerSoundLevel, RolloverSoundLevel, DTSoundLevel

GateSoundLevel = 0.5 / 5			'volume level; range [0, 1]
TargetSoundFactor = 0.0025 * 10	 'volume multiplier; must not be zero
DTSoundLevel = 0.25				 'volume multiplier; must not be zero
RolloverSoundLevel = 0.25		   'volume level; range [0, 1]
SpinnerSoundLevel = 0.5			 'volume level; range [0, 1]

'///////////////////////-----Ball Release, Guides and Drain-----///////////////////////
Dim DrainSoundLevel, BallReleaseSoundLevel, BottomArchBallGuideSoundFactor, FlipperBallGuideSoundFactor

DrainSoundLevel = 0.8				   'volume level; range [0, 1]
BallReleaseSoundLevel = 1			   'volume level; range [0, 1]
BottomArchBallGuideSoundFactor = 0.2	'volume multiplier; must not be zero
FlipperBallGuideSoundFactor = 0.015	 'volume multiplier; must not be zero

'///////////////////////-----Loops and Lanes-----///////////////////////
Dim ArchSoundFactor
ArchSoundFactor = 0.025 / 5			 'volume multiplier; must not be zero

'/////////////////////////////  SOUND PLAYBACK FUNCTIONS  ////////////////////////////
'/////////////////////////////  POSITIONAL SOUND PLAYBACK METHODS  ////////////////////////////
' Positional sound playback methods will play a sound, depending on the X,Y position of the table element or depending on ActiveBall object position
' These are similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
' For surround setup - positional sound playback functions will fade between front and rear surround channels and pan between left and right channels
' For stereo setup - positional sound playback functions will only pan between left and right channels
' For mono setup - positional sound playback functions will not pan between left and right channels and will not fade between front and rear channels

' PlaySound full syntax - PlaySound(string, int loopcount, float volume, float pan, float randompitch, int pitch, bool useexisting, bool restart, float front_rear_fade)
' Note - These functions will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlaySoundAtLevelStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelExistingStatic(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 1, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticLoop(playsoundparams, aVol, tableobj)
	PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelStaticRandomPitch(playsoundparams, aVol, randomPitch, tableobj)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

Sub PlaySoundAtLevelActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLevelExistingActiveBall(playsoundparams, aVol)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0, 0, 1, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtLeveTimerActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 0, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelTimerExistingActiveBall(playsoundparams, aVol, ballvariable)
	PlaySound playsoundparams, 0, min(aVol,1) * VolumeDial, AudioPan(ballvariable), 0, 0, 1, 0, AudioFade(ballvariable)
End Sub

Sub PlaySoundAtLevelRoll(playsoundparams, aVol, pitch)
	PlaySound playsoundparams, - 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), randomPitch, 0, 0, 0, AudioFade(tableobj)
End Sub

' Previous Positional Sound Subs

Sub PlaySoundAt(soundname, tableobj)
	PlaySound soundname, 1, 1 * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtVol(soundname, tableobj, aVol)
	PlaySound soundname, 1, min(aVol,1) * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
	PlaySoundAt soundname, ActiveBall
End Sub

Sub PlaySoundAtBallVol (Soundname, aVol)
	PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 1, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtBallVolM (Soundname, aVol)
	PlaySound soundname, 1,min(aVol,1) * VolumeDial, AudioPan(ActiveBall), 0,0,0, 0, AudioFade(ActiveBall)
End Sub

Sub PlaySoundAtVolLoops(sound, tableobj, Vol, Loops)
	PlaySound sound, Loops, Vol * VolumeDial, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

'******************************************************
'  Fleep  Supporting Ball & Sound Functions
'******************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.y * 2 / tableheight - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioFade = CSng(tmp ^ 10)
	Else
		AudioFade = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
	Dim tmp
	tmp = tableobj.x * 2 / tablewidth - 1
	
	If tmp > 7000 Then
		tmp = 7000
	ElseIf tmp <  - 7000 Then
		tmp =  - 7000
	End If
	
	If tmp > 0 Then
		AudioPan = CSng(tmp ^ 10)
	Else
		AudioPan = CSng( - (( - tmp) ^ 10) )
	End If
End Function

Function Vol(ball) ' Calculates the volume of the sound based on the ball speed
	Vol = CSng(BallVel(ball) ^ 2)
End Function

Function Volz(ball) ' Calculates the volume of the sound based on the ball speed
	Volz = CSng((ball.velz) ^ 2)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
	Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
	BallVel = Int(Sqr((ball.VelX ^ 2) + (ball.VelY ^ 2) ) )
End Function

Function VolPlayfieldRoll(ball) ' Calculates the roll volume of the sound based on the ball speed
	VolPlayfieldRoll = RollingSoundFactor * 0.0005 * CSng(BallVel(ball) ^ 3)
End Function

Function PitchPlayfieldRoll(ball) ' Calculates the roll pitch of the sound based on the ball speed
	PitchPlayfieldRoll = BallVel(ball) ^ 2 * 15
End Function

Function RndInt(min, max) ' Sets a random number integer between min and max
	RndInt = Int(Rnd() * (max - min + 1) + min)
End Function

Function RndNum(min, max) ' Sets a random number between min and max
	RndNum = Rnd() * (max - min) + min
End Function

'/////////////////////////////  GENERAL SOUND SUBROUTINES  ////////////////////////////

Sub SoundStartButton()
	PlaySound ("Start_Button"), 0, StartButtonSoundLevel, 0, 0.25
End Sub

Sub SoundNudgeLeft()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeLeftSoundLevel * VolumeDial, - 0.1, 0.25
End Sub

Sub SoundNudgeRight()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeRightSoundLevel * VolumeDial, 0.1, 0.25
End Sub

Sub SoundNudgeCenter()
	PlaySound ("Nudge_" & Int(Rnd * 2) + 1), 0, NudgeCenterSoundLevel * VolumeDial, 0, 0.25
End Sub

Sub SoundPlungerPull()
	PlaySoundAtLevelStatic ("Plunger_Pull_1"), PlungerPullSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseBall()
	PlaySoundAtLevelStatic ("Plunger_Release_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

Sub SoundPlungerReleaseNoBall()
	PlaySoundAtLevelStatic ("Plunger_Release_No_Ball"), PlungerReleaseSoundLevel, Plunger
End Sub

'/////////////////////////////  KNOCKER SOLENOID  ////////////////////////////

Sub KnockerSolenoid()
	PlaySoundAtLevelStatic SoundFX("Knocker_1",DOFKnocker), KnockerSoundLevel, KnockerPosition
End Sub

'/////////////////////////////  DRAIN SOUNDS  ////////////////////////////

Sub RandomSoundDrain(drainswitch)
	PlaySoundAtLevelStatic ("Drain_" & Int(Rnd * 11) + 1), DrainSoundLevel, drainswitch
End Sub

'/////////////////////////////  TROUGH BALL RELEASE SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBallRelease(drainswitch)
	PlaySoundAtLevelStatic SoundFX("BallRelease" & Int(Rnd * 7) + 1,DOFContactors), BallReleaseSoundLevel, drainswitch
End Sub

'/////////////////////////////  SLINGSHOT SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundSlingshotLeft(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_L" & Int(Rnd * 10) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

Sub RandomSoundSlingshotRight(sling)
	PlaySoundAtLevelStatic SoundFX("Sling_R" & Int(Rnd * 8) + 1,DOFContactors), SlingshotSoundLevel, Sling
End Sub

'/////////////////////////////  BUMPER SOLENOID SOUNDS  ////////////////////////////

Sub RandomSoundBumperTop(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Top_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperMiddle(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Middle_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

Sub RandomSoundBumperBottom(Bump)
	PlaySoundAtLevelStatic SoundFX("Bumpers_Bottom_" & Int(Rnd * 5) + 1,DOFContactors), Vol(ActiveBall) * BumperSoundFactor, Bump
End Sub

'/////////////////////////////  SPINNER SOUNDS  ////////////////////////////

Sub SoundSpinner(spinnerswitch)
	PlaySoundAtLevelStatic ("Spinner"), SpinnerSoundLevel, spinnerswitch
End Sub

'/////////////////////////////  FLIPPER BATS SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  FLIPPER BATS SOLENOID ATTACK SOUND  ////////////////////////////

Sub SoundFlipperUpAttackLeft(flipper)
	FlipperUpAttackLeftSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-L01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

Sub SoundFlipperUpAttackRight(flipper)
	FlipperUpAttackRightSoundLevel = RndNum(FlipperUpAttackMinimumSoundLevel, FlipperUpAttackMaximumSoundLevel)
	PlaySoundAtLevelStatic SoundFX("Flipper_Attack-R01",DOFFlippers), FlipperUpAttackLeftSoundLevel, flipper
End Sub

'/////////////////////////////  FLIPPER BATS SOLENOID CORE SOUND  ////////////////////////////

Sub RandomSoundFlipperUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_L0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperLeftHitParm, Flipper
End Sub

Sub RandomSoundFlipperUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_R0" & Int(Rnd * 9) + 1,DOFFlippers), FlipperRightHitParm, Flipper
End Sub

Sub RandomSoundReflipUpLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_L0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundReflipUpRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_ReFlip_R0" & Int(Rnd * 3) + 1,DOFFlippers), (RndNum(0.8, 1)) * FlipperUpSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownLeft(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Left_Down_" & Int(Rnd * 7) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

Sub RandomSoundFlipperDownRight(flipper)
	PlaySoundAtLevelStatic SoundFX("Flipper_Right_Down_" & Int(Rnd * 8) + 1,DOFFlippers), FlipperDownSoundLevel, Flipper
End Sub

'/////////////////////////////  FLIPPER BATS BALL COLLIDE SOUND  ////////////////////////////

Sub LeftFlipperCollide(parm)
	FlipperLeftHitParm = parm / 10
	If FlipperLeftHitParm > 1 Then
		FlipperLeftHitParm = 1
	End If
	FlipperLeftHitParm = FlipperUpSoundLevel * FlipperLeftHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RightFlipperCollide(parm)
	FlipperRightHitParm = parm / 10
	If FlipperRightHitParm > 1 Then
		FlipperRightHitParm = 1
	End If
	FlipperRightHitParm = FlipperUpSoundLevel * FlipperRightHitParm
	RandomSoundRubberFlipper(parm)
End Sub

Sub RandomSoundRubberFlipper(parm)
	PlaySoundAtLevelActiveBall ("Flipper_Rubber_" & Int(Rnd * 7) + 1), parm * RubberFlipperSoundFactor
End Sub

'/////////////////////////////  ROLLOVER SOUNDS  ////////////////////////////

Sub RandomSoundRollover()
	PlaySoundAtLevelActiveBall ("Rollover_" & Int(Rnd * 4) + 1), RolloverSoundLevel
End Sub

Sub Rollovers_Hit(idx)
	RandomSoundRollover
End Sub

'/////////////////////////////  VARIOUS PLAYFIELD SOUND SUBROUTINES  ////////////////////////////
'/////////////////////////////  RUBBERS AND POSTS  ////////////////////////////
'/////////////////////////////  RUBBERS - EVENTS  ////////////////////////////

Sub Rubbers_Hit(idx)
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 5 Then
		RandomSoundRubberStrong 1
	End If
	If finalspeed <= 5 Then
		RandomSoundRubberWeak()
	End If
End Sub

'/////////////////////////////  RUBBERS AND POSTS - STRONG IMPACTS  ////////////////////////////

Sub RandomSoundRubberStrong(voladj)
	Select Case Int(Rnd * 10) + 1
		Case 1
			PlaySoundAtLevelActiveBall ("Rubber_Strong_1"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 2
			PlaySoundAtLevelActiveBall ("Rubber_Strong_2"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 3
			PlaySoundAtLevelActiveBall ("Rubber_Strong_3"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 4
			PlaySoundAtLevelActiveBall ("Rubber_Strong_4"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 5
			PlaySoundAtLevelActiveBall ("Rubber_Strong_5"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 6
			PlaySoundAtLevelActiveBall ("Rubber_Strong_6"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 7
			PlaySoundAtLevelActiveBall ("Rubber_Strong_7"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 8
			PlaySoundAtLevelActiveBall ("Rubber_Strong_8"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 9
			PlaySoundAtLevelActiveBall ("Rubber_Strong_9"), Vol(ActiveBall) * RubberStrongSoundFactor * voladj
		Case 10
			PlaySoundAtLevelActiveBall ("Rubber_1_Hard"), Vol(ActiveBall) * RubberStrongSoundFactor * 0.6 * voladj
	End Select
End Sub

'/////////////////////////////  RUBBERS AND POSTS - WEAK IMPACTS  ////////////////////////////

Sub RandomSoundRubberWeak()
	PlaySoundAtLevelActiveBall ("Rubber_" & Int(Rnd * 9) + 1), Vol(ActiveBall) * RubberWeakSoundFactor
End Sub

'/////////////////////////////  WALL IMPACTS  ////////////////////////////

Sub Walls_Hit(idx)
	RandomSoundWall()
End Sub

Sub RandomSoundWall()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 5) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_1"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_2"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_5"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_7"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 5
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_9"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 4) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_3"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 4
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 3) + 1
			Case 1
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_4"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 2
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_6"), Vol(ActiveBall) * WallImpactSoundFactor
			Case 3
				PlaySoundAtLevelExistingActiveBall ("Wall_Hit_8"), Vol(ActiveBall) * WallImpactSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  METAL TOUCH SOUNDS  ////////////////////////////

Sub RandomSoundMetal()
	PlaySoundAtLevelActiveBall ("Metal_Touch_" & Int(Rnd * 13) + 1), Vol(ActiveBall) * MetalImpactSoundFactor
End Sub

'/////////////////////////////  METAL - EVENTS  ////////////////////////////

Sub Metals_Hit (idx)
	RandomSoundMetal
End Sub

Sub ShooterDiverter_collide(idx)
	RandomSoundMetal
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE  ////////////////////////////
'/////////////////////////////  BOTTOM ARCH BALL GUIDE - SOFT BOUNCES  ////////////////////////////

Sub RandomSoundBottomArchBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Bounce_" & Int(Rnd * 2) + 1), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
	If finalspeed < 6 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Bounce_Soft_1"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Medium_3"), Vol(ActiveBall) * BottomArchBallGuideSoundFactor
		End Select
	End If
End Sub

'/////////////////////////////  BOTTOM ARCH BALL GUIDE - HARD HITS  ////////////////////////////

Sub RandomSoundBottomArchBallGuideHardHit()
	PlaySoundAtLevelActiveBall ("Apron_Hard_Hit_" & Int(Rnd * 3) + 1), BottomArchBallGuideSoundFactor * 0.25
End Sub

Sub Apron_Hit (idx)
	If Abs(cor.ballvelx(ActiveBall.id) < 4) And cor.ballvely(ActiveBall.id) > 7 Then
		RandomSoundBottomArchBallGuideHardHit()
	Else
		RandomSoundBottomArchBallGuide
	End If
End Sub

'/////////////////////////////  FLIPPER BALL GUIDE  ////////////////////////////

Sub RandomSoundFlipperBallGuide()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 16 Then
		Select Case Int(Rnd * 2) + 1
			Case 1
				PlaySoundAtLevelActiveBall ("Apron_Hard_1"),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
			Case 2
				PlaySoundAtLevelActiveBall ("Apron_Hard_2"),  Vol(ActiveBall) * 0.8 * FlipperBallGuideSoundFactor
		End Select
	End If
	If finalspeed >= 6 And finalspeed <= 16 Then
		PlaySoundAtLevelActiveBall ("Apron_Medium_" & Int(Rnd * 3) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
	If finalspeed < 6 Then
		PlaySoundAtLevelActiveBall ("Apron_Soft_" & Int(Rnd * 7) + 1),  Vol(ActiveBall) * FlipperBallGuideSoundFactor
	End If
End Sub

'/////////////////////////////  TARGET HIT SOUNDS  ////////////////////////////

Sub RandomSoundTargetHitStrong()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 5,DOFTargets), Vol(ActiveBall) * 0.45 * TargetSoundFactor
End Sub

Sub RandomSoundTargetHitWeak()
	PlaySoundAtLevelActiveBall SoundFX("Target_Hit_" & Int(Rnd * 4) + 1,DOFTargets), Vol(ActiveBall) * TargetSoundFactor
End Sub

Sub PlayTargetSound()
	Dim finalspeed
	finalspeed = Sqr(ActiveBall.velx * ActiveBall.velx + ActiveBall.vely * ActiveBall.vely)
	If finalspeed > 10 Then
		RandomSoundTargetHitStrong()
		RandomSoundBallBouncePlayfieldSoft ActiveBall
	Else
		RandomSoundTargetHitWeak()
	End If
End Sub

Sub Targets_Hit (idx)
	PlayTargetSound
End Sub

'/////////////////////////////  BALL BOUNCE SOUNDS  ////////////////////////////

Sub RandomSoundBallBouncePlayfieldSoft(aBall)
	Select Case Int(Rnd * 9) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_1"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_3"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.8, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_4"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.5, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Soft_5"), volz(aBall) * BallBouncePlayfieldSoftFactor, aBall
		Case 6
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_1"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 7
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_2"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 8
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_5"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.2, aBall
		Case 9
			PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_7"), volz(aBall) * BallBouncePlayfieldSoftFactor * 0.3, aBall
	End Select
End Sub

Sub RandomSoundBallBouncePlayfieldHard(aBall)
	PlaySoundAtLevelStatic ("Ball_Bounce_Playfield_Hard_" & Int(Rnd * 7) + 1), volz(aBall) * BallBouncePlayfieldHardFactor, aBall
End Sub

'/////////////////////////////  DELAYED DROP - TO PLAYFIELD - SOUND  ////////////////////////////

Sub RandomSoundDelayedBallDropOnPlayfield(aBall)
	Select Case Int(Rnd * 5) + 1
		Case 1
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_1_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 2
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_2_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 3
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_3_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 4
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_4_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
		Case 5
			PlaySoundAtLevelStatic ("Ball_Drop_Playfield_5_Delayed"), DelayedBallDropOnPlayfieldSoundLevel, aBall
	End Select
End Sub

'/////////////////////////////  BALL GATES AND BRACKET GATES SOUNDS  ////////////////////////////

Sub SoundPlayfieldGate()
	PlaySoundAtLevelStatic ("Gate_FastTrigger_" & Int(Rnd * 2) + 1), GateSoundLevel, ActiveBall
End Sub

Sub SoundHeavyGate()
	PlaySoundAtLevelStatic ("Gate_2"), GateSoundLevel, ActiveBall
End Sub

Sub Gates_hit(idx)
	SoundHeavyGate
End Sub

Sub GatesWire_hit(idx)
	SoundPlayfieldGate
End Sub

'/////////////////////////////  LEFT LANE ENTRANCE - SOUNDS  ////////////////////////////

Sub RandomSoundLeftArch()
	PlaySoundAtLevelActiveBall ("Arch_L" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub RandomSoundRightArch()
	PlaySoundAtLevelActiveBall ("Arch_R" & Int(Rnd * 4) + 1), Vol(ActiveBall) * ArchSoundFactor
End Sub

Sub Arch1_hit()
	If ActiveBall.velx > 1 Then SoundPlayfieldGate
	StopSound "Arch_L1"
	StopSound "Arch_L2"
	StopSound "Arch_L3"
	StopSound "Arch_L4"
End Sub

Sub Arch1_unhit()
	If ActiveBall.velx <  - 8 Then
		RandomSoundRightArch
	End If
End Sub

Sub Arch2_hit()
	If ActiveBall.velx < 1 Then SoundPlayfieldGate
	StopSound "Arch_R1"
	StopSound "Arch_R2"
	StopSound "Arch_R3"
	StopSound "Arch_R4"
End Sub

Sub Arch2_unhit()
	If ActiveBall.velx > 10 Then
		RandomSoundLeftArch
	End If
End Sub

'/////////////////////////////  SAUCERS (KICKER HOLES)  ////////////////////////////

Sub SoundSaucerLock()
	PlaySoundAtLevelStatic ("Saucer_Enter_" & Int(Rnd * 2) + 1), SaucerLockSoundLevel, ActiveBall
End Sub

Sub SoundSaucerKick(scenario, saucer)
	Select Case scenario
		Case 0
			PlaySoundAtLevelStatic SoundFX("Saucer_Empty", DOFContactors), SaucerKickSoundLevel, saucer
		Case 1
			PlaySoundAtLevelStatic SoundFX("Saucer_Kick", DOFContactors), SaucerKickSoundLevel, saucer
	End Select
End Sub

'/////////////////////////////  BALL COLLISION SOUND  ////////////////////////////

Sub OnBallBallCollision(ball1, ball2, velocity)

	FlipperCradleCollision ball1, ball2, velocity

	Dim snd
	Select Case Int(Rnd * 7) + 1
		Case 1
			snd = "Ball_Collide_1"
		Case 2
			snd = "Ball_Collide_2"
		Case 3
			snd = "Ball_Collide_3"
		Case 4
			snd = "Ball_Collide_4"
		Case 5
			snd = "Ball_Collide_5"
		Case 6
			snd = "Ball_Collide_6"
		Case 7
			snd = "Ball_Collide_7"
	End Select
	
	PlaySound (snd), 0, CSng(velocity) ^ 2 / 200 * BallWithBallCollisionSoundFactor * VolumeDial, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


'///////////////////////////  DROP TARGET HIT SOUNDS  ///////////////////////////

Sub RandomSoundDropTargetReset(obj)
	PlaySoundAtLevelStatic SoundFX("Drop_Target_Reset_" & Int(Rnd * 6) + 1,DOFContactors), 1, obj
End Sub

Sub SoundDropTargetDrop(obj)
	PlaySoundAtLevelStatic ("Drop_Target_Down_" & Int(Rnd * 6) + 1), 200, obj
End Sub

'/////////////////////////////  GI AND FLASHER RELAYS  ////////////////////////////

Const RelayFlashSoundLevel = 0.315  'volume level; range [0, 1];
Const RelayGISoundLevel = 1.05	  'volume level; range [0, 1];

Sub Sound_GI_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_GI_On"), 0.025 * RelayGISoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_GI_Off"), 0.025 * RelayGISoundLevel, obj
	End Select
End Sub

Sub Sound_Flash_Relay(toggle, obj)
	Select Case toggle
		Case 1
			PlaySoundAtLevelStatic ("Relay_Flash_On"), 0.025 * RelayFlashSoundLevel, obj
		Case 0
			PlaySoundAtLevelStatic ("Relay_Flash_Off"), 0.025 * RelayFlashSoundLevel, obj
	End Select
End Sub

'/////////////////////////////////////////////////////////////////
'					End Mechanical Sounds
'/////////////////////////////////////////////////////////////////

'******************************************************
'****  FLEEP MECHANICAL SOUNDS
'******************************************************


'******************************************************
' 	Z3DI:   3D INSERTS
'******************************************************
'
'
' Before you get started adding the inserts to your playfield in VPX, there are a few things you need to have done to prepare:
'	 1. Cut out all the inserts on the playfield image so there is alpha transparency where they should be.
'	  Make sure the playfield material has Opacity Active checkbox checked.
'	2. All the  insert text and/or images that lie over the insert plastic needs to be in its own file with
'	   alpha transparency. Many playfields may require finding the original font and remaking the insert text.
'
' To add the inserts:
'	1. Import all the textures (images) and materials from this file that start with the word "Insert" into your Table
'   2. Copy and past the two primitves that make up the insert you want to use. One primitive is for the on state, the other for the off state.
'   3. Align the primitives with the associated insert light. Name the on and off primitives correctly.
'   5. You will need to manually tweak the disable lighting value and material parameters to achielve the effect you want.
'
'
' Quick Reference:  Laying the Inserts ( Tutorial From Iaakki)
' - Each insert consists of two primitives. On and Off primitive. Suggested naming convention is to use lamp number in the name. For example
'   is lamp number is 57, the On primitive is "p57" and the Off primitive is "p57off". This makes it easier to work on script side.
' - When starting from a new table, I'd first select to make few inserts that look quite similar. Lets say there is total of 6 small triangle
'   inserts, 4 yellow and 2 blue ones.
' - Import the insert on/off images from the image manager and the vpx materials used from the sample project first, and those should appear
'   selected properly in the primitive settings when you paste your actual insert trays in your target table . Then open up your target project
'   at same time as the sample project and use copy&paste to copy desired inserts to target project.
' - There are quite many parameters in primitive that affect a lot how they will look. I wouldn't mess too much with them. Use Size options to
'   scale the insert properly into PF hole. Some insert primitives may have incorrect pivot point, which means that changing the depth, you may
'   also need to alter the Z-position too.
' - Once you have the first insert in place, wire it up in the script (detailed in part 3 below). Then set the light bulb's intensity to zero,
'   so it won't harass the adjustment.
' - Start up the game with F6 and visually see if the On-primitive blinks properly. If it is too dim, hit D and open editor. Write:
' - p57.BlendDisableLighting = 300 and hit enter
' - -> The insert should appear differently. Find good looking brightness level. Not too bright as the light bulb is still missing. Just generic good light.
'	 - If you cannot find proper light color or "mood", you can also fiddle with primitive material values. Provided material should be
'	   quite ok for most of the cases.
'	 - Now when you have found proper DL value (165), but that into script:
' - That one insert is now adjusted and you should be able to copy&paste rest of the triangle inserts in place and name them correctly. And add them
'   into script. And fine tune their brightness and color.
'
' Light bulbs and ball reflection:
'
' - This kind of lighted primitives are not giving you ball reflections. Also some more glow vould be needed to make the insert to bloom correctly.
' - Take the original lamp (l57), set the bulb mode enabled, set Halo Height to -3 (something that is inside the 2 insert primitives). I'd start with
'   falloff 100, falloff Power 2-2.5, Intensity 10, scale mesh 10, Transmit 5.
' - Start the game with F6, throw a ball on it and move the ball near the blinking insert. Visually see how the reflection looks.
' - Hit D once the reflection is the highest. Open light editor and start fine tuning the bulb values to achieve realistic look for the reflection.
' - Falloff Power value is the one that will affect reflection creatly. The higher the power value is, the brighter the reflection on the ball is.
'   This is the reason why falloff is rather large and falloff power is quite low. Change scale mesh if reflection is too small or large.
' - Transmit value can bring nice bloom for the insert, but it may also affect to other primitives nearby. Sometimes one need to set transmit to
'   zero to avoid affecting surrounding plastics. If you really need to have higher transmit value, you may set Disable Lighting From Below to 1
'   in surrounding primitive. This may remove the problem, but can make the primitive look worse too.

' ---- FOR EXAMPLE PUROSE ONLY. DELETE BELOW WHEN USING FOR REAL -----
' NOTE: The below timer is for flashing the inserts as a demonstration. Should be replaced by actual lamp states.
'	   In other words, delete this sub (InsertFlicker_timer) and associated timer if you are going to use with a ROM.
Dim flickerX, FlickerState
FlickerState = 0
Sub InsertFlicker_timer
	If FlickerState = 0 Then
		For flickerX = 0 To 19
			If flickerX < 6 Or flickerX > 9 Then
				AllLamps(flickerX).state = False
			ElseIf BonusX(flickerX - 6) = False Then
				AllLamps(flickerX).state = False
			End If
		Next
		FlickerState = 1
	Else
		For flickerX = 0 To 19
			If flickerX < 6 Or flickerX > 9 Then
				AllLamps(flickerX).state = True
			ElseIf BonusX(flickerX - 6) = False Then
				AllLamps(flickerX).state = True
			End If
		Next
		FlickerState = 0
	End If
End Sub



Sub l1_animate: p1.BlendDisableLighting = 200 * (l1.GetInPlayIntensity / l1.Intensity): End Sub
Sub l2_animate: p2.BlendDisableLighting = 200 * (l2.GetInPlayIntensity / l2.Intensity): End Sub
Sub l3_animate: p3.BlendDisableLighting = 200 * (l3.GetInPlayIntensity / l3.Intensity): End Sub
' Sub l4_animate: p4.BlendDisableLighting = 200 * (l4.GetInPlayIntensity / l4.Intensity): End Sub
' Sub l5_animate: p5.BlendDisableLighting = 200 * (l5.GetInPlayIntensity / l5.Intensity): End Sub
Sub l6_animate: p6.BlendDisableLighting = 200 * (l6.GetInPlayIntensity / l6.Intensity): End Sub
Sub l7_animate: p7.BlendDisableLighting = 200 * (l7.GetInPlayIntensity / l7.Intensity): End Sub
Sub l8_animate: p8.BlendDisableLighting = 200 * (l8.GetInPlayIntensity / l8.Intensity): End Sub
Sub l9_animate: p9.BlendDisableLighting = 200 * (l9.GetInPlayIntensity / l9.Intensity): End Sub

' Sub l10_animate: p10.BlendDisableLighting = 200 * (l10.GetInPlayIntensity / l10.Intensity): End Sub
Sub l11_animate: p11.BlendDisableLighting = 200 * (l11.GetInPlayIntensity / l11.Intensity): End Sub
Sub l12_animate: p12.BlendDisableLighting = 200 * (l12.GetInPlayIntensity / l12.Intensity): End Sub
Sub l13_animate: p13.BlendDisableLighting = 200 * (l13.GetInPlayIntensity / l13.Intensity): End Sub
Sub l14_animate: p14.BlendDisableLighting = 200 * (l14.GetInPlayIntensity / l14.Intensity): End Sub
Sub l15_animate: p15.BlendDisableLighting = 200 * (l15.GetInPlayIntensity / l15.Intensity): End Sub
Sub l16_animate: p16.BlendDisableLighting = 200 * (l16.GetInPlayIntensity / l16.Intensity): End Sub
Sub l17_animate: p17.BlendDisableLighting = 200 * (l17.GetInPlayIntensity / l17.Intensity): End Sub
Sub l18_animate: p18.BlendDisableLighting = 200 * (l18.GetInPlayIntensity / l18.Intensity): End Sub
Sub l19_animate: p19.BlendDisableLighting = 200 * (l19.GetInPlayIntensity / l19.Intensity): End Sub

Sub l20_animate: p20.BlendDisableLighting = 200 * (l20.GetInPlayIntensity / l20.Intensity): End Sub
Sub l21_animate: p21.BlendDisableLighting = 200 * (l21.GetInPlayIntensity / l21.Intensity): End Sub
Sub l22_animate: p22.BlendDisableLighting = 200 * (l22.GetInPlayIntensity / l22.Intensity): End Sub
Sub l23_animate: p23.BlendDisableLighting = 200 * (l23.GetInPlayIntensity / l23.Intensity): End Sub

'******************************************************
'*****   END 3D INSERTS
'******************************************************






'******************************************************
' 	ZFLD:  FLUPPER DOMES
'******************************************************
' Based on FlupperDoms2.2

' What you need in your table to use these flashers:
' Open this table and your table both in VPX
' Export all the materials domebasemat, Flashermaterial0 - 20 and import them in your table
' Export all textures (images) starting with the name "dome" and "ronddome" and import them into your table with the same names
' Export all textures (images) starting with the name "flasherbloom" and import them into your table with the same names
' Copy a set of 4 objects flasherbase, flasherlit, flasherlight and flasherflash from layer 7 to your table
' If you duplicate the four objects for a new flasher dome, be sure that they all end with the same number (in the 0-20 range)
' Copy the flasherbloom flashers from layer 10 to your table. you will need to make one per flasher dome that you plan to make
' Select the correct flasherbloom texture for each flasherbloom flasher, per flasher dome
' Copy the script below

' Place your flasher base primitive where you want the flasher located on your Table
' Then run InitFlasher in the script with the number of your flasher objects and the color of the flasher.  This will align the flasher object, light object, and
' flasher lit primitive.  It will also assign the appropriate flasher bloom images to the flasher bloom object.
'
' Example: InitFlasher 1, "green"
'
' Color Options: "blue", "green", "red", "purple", "yellow", "white", and "orange"

' You can use the RotateFlasher call to align the Rotz/ObjRotz of the flasher primitives with "handles".  Don't set those values in the editor,
' call the RotateFlasher sub instead (this call will likely crash VP if it's call for the flasher primitives without "handles")
'
' Example: RotateFlasher 1, 180		 'where 1 is the flasher number and 180 is the angle of Z rotation

' For flashing the flasher use in the script: "ObjLevel(1) = 1 : FlasherFlash1_Timer"
' This should also work for flashers with variable flash levels from the rom, just use ObjLevel(1) = xx from the rom (in the range 0-1)
'
' Notes (please read!!):
' - Setting TestFlashers = 1 (below in the ScriptsDirectory) will allow you to see how the flasher objects are aligned (need the targetflasher image imported to your table)
' - The rotation of the primitives with "handles" is done with a script command, not on the primitive itself (see RotateFlasher below)
' - Color of the objects are set in the script, not on the primitive itself
' - Screws are optional to copy and position manually
' - If your table is not named "Table1" then change the name below in the script
' - Every flasher uses its own material (Flashermaterialxx), do not use it for anything else
' - Lighting > Bloom Strength affects how the flashers look, do not set it too high
' - Change RotY and RotX of flasherbase only when having a flasher something other then parallel to the playfield
' - Leave RotX of the flasherflash object to -45; this makes sure that the flash effect is visible in FS and DT
' - If you want to resize a flasher, be sure to resize flasherbase, flasherlit and flasherflash with the same percentage
' - If you think that the flasher effects are too bright, change flasherlightintensity and/or flasherflareintensity below

' Some more notes for users of the v1 flashers and/or JP's fading lights routines:
' - Delete all textures/primitives/script/materials in your table from the v1 flashers and scripts before you start; they don't mix well with v2
' - Remove flupperflash(m) routines if you have them; they do not work with this new script
' - Do not try to mix this v2 script with the JP fading light routine (that is making it too complicated), just use the example script below

' example script for rom based tables (non modulated):

' SolCallback(25)="FlashRed"
'
' Sub FlashRed(flstate)
'	If Flstate Then
'		ObjTargetLevel(1) = 1
'	Else
'		ObjTargetLevel(1) = 0
'	End If
'   FlasherFlash1_Timer
' End Sub

' example script for rom based tables (modulated):

' SolModCallback(25)="FlashRed"
'
' Sub FlashRed(level)
'	ObjTargetLevel(1) = level/255 : FlasherFlash1_Timer
' End Sub

Sub Flash1(Enabled)
	If Enabled Then
		ObjTargetLevel(1) = 1
	Else
		ObjTargetLevel(1) = 0
	End If
	FlasherFlash1_Timer
	Sound_Flash_Relay enabled, Flasherbase1
End Sub

Sub Flash2(Enabled)
	If Enabled Then
		ObjTargetLevel(2) = 1
	Else
		ObjTargetLevel(2) = 0
	End If
	FlasherFlash2_Timer
	Sound_Flash_Relay enabled, Flasherbase2
End Sub

Sub Flash3(Enabled)
	If Enabled Then
		ObjTargetLevel(3) = 1
	Else
		ObjTargetLevel(3) = 0
	End If
	FlasherFlash3_Timer
	Sound_Flash_Relay enabled, Flasherbase3
End Sub

Sub Flash4(Enabled)
	If Enabled Then
		ObjTargetLevel(4) = 1
	Else
		ObjTargetLevel(4) = 0
	End If
	FlasherFlash4_Timer
	Sound_Flash_Relay enabled, Flasherbase1
End Sub

Dim TestFlashers, TableRef, FlasherLightIntensity, FlasherFlareIntensity, FlasherBloomIntensity, FlasherOffBrightness

' *********************************************************************
TestFlashers = 0				' *** set this to 1 to check position of flasher object			 ***
Set TableRef = Table1		   ' *** change this, if your table has another name				   ***
FlasherLightIntensity = 0.1	 ' *** lower this, if the VPX lights are too bright (i.e. 0.1)	   ***
FlasherFlareIntensity = 0.3	 ' *** lower this, if the flares are too bright (i.e. 0.1)		   ***
FlasherBloomIntensity = 0.2	 ' *** lower this, if the blooms are too bright (i.e. 0.1)		   ***
FlasherOffBrightness = 0.5	  ' *** brightness of the flasher dome when switched off (range 0-2)  ***
' *********************************************************************

Dim ObjLevel(20), objbase(20), objlit(20), objflasher(20), objbloom(20), objlight(20), ObjTargetLevel(20)
'Dim tablewidth, tableheight : tablewidth = TableRef.width : tableheight = TableRef.height

'initialise the flasher color, you can only choose from "green", "red", "purple", "blue", "white" and "yellow"
InitFlasher 1, "green"
InitFlasher 2, "red"
InitFlasher 3, "blue"
InitFlasher 4, "white"

' rotate the flasher with the command below (first argument = flasher nr, second argument = angle in degrees)
'   RotateFlasher 1,17
'   RotateFlasher 2,0
'   RotateFlasher 3,90
'   RotateFlasher 4,90

Sub InitFlasher(nr, col)
	' store all objects in an array for use in FlashFlasher subroutine
	Set objbase(nr) = Eval("Flasherbase" & nr)
	Set objlit(nr) = Eval("Flasherlit" & nr)
	Set objflasher(nr) = Eval("Flasherflash" & nr)
	Set objlight(nr) = Eval("Flasherlight" & nr)
	Set objbloom(nr) = Eval("Flasherbloom" & nr)
	
	' If the flasher is parallel to the playfield, rotate the VPX flasher object for POV and place it at the correct height
	If objbase(nr).RotY = 0 Then
		objbase(nr).ObjRotZ = Atn( (tablewidth / 2 - objbase(nr).x) / (objbase(nr).y - tableheight * 1.1)) * 180 / 3.14159
		objflasher(nr).RotZ = objbase(nr).ObjRotZ
		objflasher(nr).height = objbase(nr).z + 40
	End If
	
	' set all effects to invisible and move the lit primitive at the same position and rotation as the base primitive
	objlight(nr).IntensityScale = 0
	objlit(nr).visible = 0
	objlit(nr).material = "Flashermaterial" & nr
	objlit(nr).RotX = objbase(nr).RotX
	objlit(nr).RotY = objbase(nr).RotY
	objlit(nr).RotZ = objbase(nr).RotZ
	objlit(nr).ObjRotX = objbase(nr).ObjRotX
	objlit(nr).ObjRotY = objbase(nr).ObjRotY
	objlit(nr).ObjRotZ = objbase(nr).ObjRotZ
	objlit(nr).x = objbase(nr).x
	objlit(nr).y = objbase(nr).y
	objlit(nr).z = objbase(nr).z
	objbase(nr).BlendDisableLighting = FlasherOffBrightness
	
	'rothbauerw
	'Adjust the position of the flasher object to align with the flasher base.
	'Comment out these lines if you want to manually adjust the flasher object
	If objbase(nr).roty > 135 Then
		objflasher(nr).y = objbase(nr).y + 50
		objflasher(nr).height = objbase(nr).z + 20
	Else
		objflasher(nr).y = objbase(nr).y + 20
		objflasher(nr).height = objbase(nr).z + 50
	End If
	objflasher(nr).x = objbase(nr).x
	
	'rothbauerw
	'Adjust the position of the light object to align with the flasher base.
	'Comment out these lines if you want to manually adjust the flasher object
	objlight(nr).x = objbase(nr).x
	objlight(nr).y = objbase(nr).y
	objlight(nr).bulbhaloheight = objbase(nr).z - 10
	
	'rothbauerw
	'Assign the appropriate bloom image basked on the location of the flasher base
	'Comment out these lines if you want to manually assign the bloom images
	Dim xthird, ythird
	xthird = tablewidth / 3
	ythird = tableheight / 3
	If objbase(nr).x >= xthird And objbase(nr).x <= xthird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenter"
		objbloom(nr).imageB = "flasherbloomCenter"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird Then
		objbloom(nr).imageA = "flasherbloomUpperLeft"
		objbloom(nr).imageB = "flasherbloomUpperLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird Then
		objbloom(nr).imageA = "flasherbloomUpperRight"
		objbloom(nr).imageB = "flasherbloomUpperRight"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenterLeft"
		objbloom(nr).imageB = "flasherbloomCenterLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 2 Then
		objbloom(nr).imageA = "flasherbloomCenterRight"
		objbloom(nr).imageB = "flasherbloomCenterRight"
	ElseIf objbase(nr).x < xthird And objbase(nr).y < ythird * 3 Then
		objbloom(nr).imageA = "flasherbloomLowerLeft"
		objbloom(nr).imageB = "flasherbloomLowerLeft"
	ElseIf  objbase(nr).x > xthird * 2 And objbase(nr).y < ythird * 3 Then
		objbloom(nr).imageA = "flasherbloomLowerRight"
		objbloom(nr).imageB = "flasherbloomLowerRight"
	End If
	
	' set the texture and color of all objects
	Select Case objbase(nr).image
		Case "dome2basewhite"
			objbase(nr).image = "dome2base" & col
			objlit(nr).image = "dome2lit" & col
			
		Case "ronddomebasewhite"
			objbase(nr).image = "ronddomebase" & col
			objlit(nr).image = "ronddomelit" & col
			
		Case "domeearbasewhite"
			objbase(nr).image = "domeearbase" & col
			objlit(nr).image = "domeearlit" & col
	End Select
	If TestFlashers = 0 Then
		objflasher(nr).imageA = "domeflashwhite"
		objflasher(nr).visible = 0
	End If
	Select Case col
		Case "blue"
			objlight(nr).color = RGB(4,120,255)
			objflasher(nr).color = RGB(200,255,255)
			objbloom(nr).color = RGB(4,120,255)
			objlight(nr).intensity = 5000
			
		Case "green"
			objlight(nr).color = RGB(12,255,4)
			objflasher(nr).color = RGB(12,255,4)
			objbloom(nr).color = RGB(12,255,4)
			
		Case "red"
			objlight(nr).color = RGB(255,32,4)
			objflasher(nr).color = RGB(255,32,4)
			objbloom(nr).color = RGB(255,32,4)
			
		Case "purple"
			objlight(nr).color = RGB(230,49,255)
			objflasher(nr).color = RGB(255,64,255)
			objbloom(nr).color = RGB(230,49,255)
			
		Case "yellow"
			objlight(nr).color = RGB(200,173,25)
			objflasher(nr).color = RGB(255,200,50)
			objbloom(nr).color = RGB(200,173,25)
			
		Case "white"
			objlight(nr).color = RGB(255,240,150)
			objflasher(nr).color = RGB(100,86,59)
			objbloom(nr).color = RGB(255,240,150)
			
		Case "orange"
			objlight(nr).color = RGB(255,70,0)
			objflasher(nr).color = RGB(255,70,0)
			objbloom(nr).color = RGB(255,70,0)
	End Select
	objlight(nr).colorfull = objlight(nr).color
	If TableRef.ShowDT And ObjFlasher(nr).RotX =  - 45 Then
		objflasher(nr).height = objflasher(nr).height - 20 * ObjFlasher(nr).y / tableheight
		ObjFlasher(nr).y = ObjFlasher(nr).y + 10
	End If
End Sub

Sub RotateFlasher(nr, angle)
	angle = ((angle + 360 - objbase(nr).ObjRotZ) Mod 180) / 30
	objbase(nr).showframe(angle)
	objlit(nr).showframe(angle)
End Sub

Sub FlashFlasher(nr)
	If Not objflasher(nr).TimerEnabled Then
		objflasher(nr).TimerEnabled = True
		objflasher(nr).visible = 1
		objbloom(nr).visible = 1
		objlit(nr).visible = 1
	End If
	objflasher(nr).opacity = 1000 * FlasherFlareIntensity * ObjLevel(nr) ^ 2.5
	objbloom(nr).opacity = 100 * FlasherBloomIntensity * ObjLevel(nr) ^ 2.5
	objlight(nr).IntensityScale = 0.5 * FlasherLightIntensity * ObjLevel(nr) ^ 3
	objbase(nr).BlendDisableLighting = FlasherOffBrightness + 10 * ObjLevel(nr) ^ 3
	objlit(nr).BlendDisableLighting = 10 * ObjLevel(nr) ^ 2
	UpdateMaterial "Flashermaterial" & nr,0,0,0,0,0,0,ObjLevel(nr),RGB(255,255,255),0,0,False,True,0,0,0,0
	If Round(ObjTargetLevel(nr),1) > Round(ObjLevel(nr),1) Then
		ObjLevel(nr) = ObjLevel(nr) + 0.3
		If ObjLevel(nr) > 1 Then ObjLevel(nr) = 1
	ElseIf Round(ObjTargetLevel(nr),1) < Round(ObjLevel(nr),1) Then
		ObjLevel(nr) = ObjLevel(nr) * 0.85 - 0.01
		If ObjLevel(nr) < 0 Then ObjLevel(nr) = 0
	Else
		ObjLevel(nr) = Round(ObjTargetLevel(nr),1)
		objflasher(nr).TimerEnabled = False
	End If
	'   ObjLevel(nr) = ObjLevel(nr) * 0.9 - 0.01
	If ObjLevel(nr) < 0 Then
		objflasher(nr).TimerEnabled = False
		objflasher(nr).visible = 0
		objbloom(nr).visible = 0
		objlit(nr).visible = 0
	End If
End Sub

Sub FlasherFlash1_Timer()
	FlashFlasher(1)
End Sub
Sub FlasherFlash2_Timer()
	FlashFlasher(2)
End Sub
Sub FlasherFlash3_Timer()
	FlashFlasher(3)
End Sub
Sub FlasherFlash4_Timer()
	FlashFlasher(4)
End Sub

'******************************************************
'******  END FLUPPER DOMES
'******************************************************





'******************************************************
' 	ZFLB:  FLUPPER BUMPERS
'******************************************************
' Based on FlupperBumpers 0.145 final

' Explanation of how these bumpers work:
' There are 10 elements involved per bumper:
' - the shadow of the bumper ( a vpx flasher object)
' - the bumper skirt (primitive)
' - the bumperbase (primitive)
' - a vpx light which colors everything you can see through the bumpertop
' - the bulb (primitive)
' - another vpx light which lights up everything around the bumper
' - the bumpertop (primitive)
' - the VPX bumper object
' - the bumper screws (primitive)
' - the bulb highlight VPX flasher object
' All elements have a special name with the number of the bumper at the end, this is necessary for the fading routine and the initialisation.
' For the bulb and the bumpertop there is a unique material as well per bumpertop.
' To use these bumpers you have to first copy all 10 elements to your table.
' Also export the textures (images) with names that start with "Flbumper" and "Flhighlight" and materials with names that start with "bumper".
' Make sure that all the ten objects are aligned on center, if possible with the exact same x,y coordinates
' After that copy the script (below); also copy the BumperTimer vpx object to your table
' Every bumper needs to be initialised with the FlInitBumper command, see example below;
' Colors available are red, white, blue, orange, yellow, green, purple and blacklight.
' In a GI subroutine you can then call set the bumperlight intensity with the "FlBumperFadeTarget(nr) = value" command
' where nr is the number of the bumper, value is between 0 (off) and 1 (full on) (so you can also use 0.3 0.4 etc).

' Notes:
' - There is only one color for the disk; you can photoshop it to a different color
' - The bumpertops are angle independent up to a degree; my estimate is -45 to + 45 degrees horizontally, 0 (topview) to 70-80 degrees (frontview)
' - I built in correction for the day-night slider; this might not work perfectly, depending on your table lighting
' - These elements, textures and materials do NOT integrate with any of the lighting routines I have seen in use in many VPX tables
'   (just find the GI handling routine and insert the FlBumperFadeTarget statement)
' - If you want to use VPX native bumperdisks just copy my bumperdisk but make it invisible

' prepare some global vars to dim/brighten objects when using day-night slider
Dim DayNightAdjust , DNA30, DNA45, DNA90
If NightDay < 10 Then
	DNA30 = 0
	DNA45 = (NightDay - 10) / 20
	DNA90 = 0
	DayNightAdjust = 0.4
Else
	DNA30 = (NightDay - 10) / 30
	DNA45 = (NightDay - 10) / 45
	DNA90 = (NightDay - 10) / 90
	DayNightAdjust = NightDay / 25
End If

Dim FlBumperFadeActual(6), FlBumperFadeTarget(6), FlBumperColor(6), FlBumperTop(6), FlBumperSmallLight(6), Flbumperbiglight(6)
Dim FlBumperDisk(6), FlBumperBase(6), FlBumperBulb(6), FlBumperscrews(6), FlBumperActive(6), FlBumperHighlight(6)
Dim cnt
For cnt = 1 To 6
	FlBumperActive(cnt) = False
Next

' colors available are red, white, blue, orange, yellow, green, purple and blacklight
FlInitBumper 1, "red"
FlInitBumper 2, "white"
FlInitBumper 3, "blue"
FlInitBumper 4, "orange"
FlInitBumper 5, "yellow"

' ### uncomment the statement below to change the color for all bumpers ###
'   Dim ind
'   For ind = 1 To 5
'	   FlInitBumper ind, "green"
'   Next

Sub FlInitBumper(nr, col)
	FlBumperActive(nr) = True
	
	' store all objects in an array for use in FlFadeBumper subroutine
	FlBumperFadeActual(nr) = 1
	FlBumperFadeTarget(nr) = 1.1
	FlBumperColor(nr) = col
	Set FlBumperTop(nr) = Eval("bumpertop" & nr)
	FlBumperTop(nr).material = "bumpertopmat" & nr
	Set FlBumperSmallLight(nr) = Eval("bumpersmalllight" & nr)
	Set Flbumperbiglight(nr) = Eval("bumperbiglight" & nr)
	Set FlBumperDisk(nr) = Eval("bumperdisk" & nr)
	Set FlBumperBase(nr) = Eval("bumperbase" & nr)
	Set FlBumperBulb(nr) = Eval("bumperbulb" & nr)
	FlBumperBulb(nr).material = "bumperbulbmat" & nr
	Set FlBumperscrews(nr) = Eval("bumperscrews" & nr)
	FlBumperscrews(nr).material = "bumperscrew" & col
	Set FlBumperHighlight(nr) = Eval("bumperhighlight" & nr)
	
	' set the color for the two VPX lights
	Select Case col
		Case "red"
			FlBumperSmallLight(nr).color = RGB(255,4,0)
			FlBumperSmallLight(nr).colorfull = RGB(255,24,0)
			FlBumperBigLight(nr).color = RGB(255,32,0)
			FlBumperBigLight(nr).colorfull = RGB(255,32,0)
			FlBumperHighlight(nr).color = RGB(64,255,0)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.98
			FlBumperSmallLight(nr).TransmissionScale = 0
			
		Case "blue"
			FlBumperBigLight(nr).color = RGB(32,80,255)
			FlBumperBigLight(nr).colorfull = RGB(32,80,255)
			FlBumperSmallLight(nr).color = RGB(0,80,255)
			FlBumperSmallLight(nr).colorfull = RGB(0,80,255)
			FlBumperSmallLight(nr).TransmissionScale = 0
			MaterialColor "bumpertopmat" & nr, RGB(8,120,255)
			FlBumperHighlight(nr).color = RGB(255,16,8)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			
		Case "green"
			FlBumperSmallLight(nr).color = RGB(8,255,8)
			FlBumperSmallLight(nr).colorfull = RGB(8,255,8)
			FlBumperBigLight(nr).color = RGB(32,255,32)
			FlBumperBigLight(nr).colorfull = RGB(32,255,32)
			FlBumperHighlight(nr).color = RGB(255,32,255)
			MaterialColor "bumpertopmat" & nr, RGB(16,255,16)
			FlBumperSmallLight(nr).TransmissionScale = 0.005
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			
		Case "orange"
			FlBumperHighlight(nr).color = RGB(255,130,255)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).color = RGB(255,130,0)
			FlBumperSmallLight(nr).colorfull = RGB (255,90,0)
			FlBumperBigLight(nr).color = RGB(255,190,8)
			FlBumperBigLight(nr).colorfull = RGB(255,190,8)
			
		Case "white"
			FlBumperBigLight(nr).color = RGB(255,230,190)
			FlBumperBigLight(nr).colorfull = RGB(255,230,190)
			FlBumperHighlight(nr).color = RGB(255,180,100)
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 0.99
			
		Case "blacklight"
			FlBumperBigLight(nr).color = RGB(32,32,255)
			FlBumperBigLight(nr).colorfull = RGB(32,32,255)
			FlBumperHighlight(nr).color = RGB(48,8,255)
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			
		Case "yellow"
			FlBumperSmallLight(nr).color = RGB(255,230,4)
			FlBumperSmallLight(nr).colorfull = RGB(255,230,4)
			FlBumperBigLight(nr).color = RGB(255,240,50)
			FlBumperBigLight(nr).colorfull = RGB(255,240,50)
			FlBumperHighlight(nr).color = RGB(255,255,220)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
			FlBumperSmallLight(nr).TransmissionScale = 0
			
		Case "purple"
			FlBumperBigLight(nr).color = RGB(80,32,255)
			FlBumperBigLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).color = RGB(80,32,255)
			FlBumperSmallLight(nr).colorfull = RGB(80,32,255)
			FlBumperSmallLight(nr).TransmissionScale = 0
			FlBumperHighlight(nr).color = RGB(32,64,255)
			FlBumperSmallLight(nr).BulbModulateVsAdd = 1
	End Select
End Sub

Sub FlFadeBumper(nr, Z)
	FlBumperBase(nr).BlendDisableLighting = 0.5 * DayNightAdjust
	'   UpdateMaterial(string, float wrapLighting, float roughness, float glossyImageLerp, float thickness, float edge, float edgeAlpha, float opacity,
	'			   OLE_COLOR base, OLE_COLOR glossy, OLE_COLOR clearcoat, VARIANT_BOOL isMetal, VARIANT_BOOL opacityActive,
	'			   float elasticity, float elasticityFalloff, float friction, float scatterAngle) - updates all parameters of a material
	FlBumperDisk(nr).BlendDisableLighting = (0.5 - Z * 0.3 ) * DayNightAdjust
	
	Select Case FlBumperColor(nr)
		Case "blue"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(38 - 24 * Z,130 - 98 * Z,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 500 * Z / (0.5 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 5000 * (0.03 * Z + 0.97 * Z ^ 3)
			Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 10000 * (Z ^ 3) / (0.5 + DNA90)
			
		Case "green"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(16 + 16 * Sin(Z * 3.14),255,16 + 16 * Sin(Z * 3.14)), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 10 + 150 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 2 * DayNightAdjust + 20 * Z
			FlBumperBulb(nr).BlendDisableLighting = 7 * DayNightAdjust + 6000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 6000 * (Z ^ 3) / (1 + DNA90)
			
		Case "red"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 16 - 11 * Z + 16 * Sin(Z * 3.14),0), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 100 * Z / (1 + DNA30 ^ 2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 18 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 20 * DayNightAdjust + 9000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,20 + Z * 4,8 - Z * 8)
			
		Case "orange"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 100 - 22 * z + 16 * Sin(Z * 3.14),Z * 32), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 250 * Z / (1 + DNA30 ^ 2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 2500 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,100 + Z * 50, 0)
			
		Case "white"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255,230 - 100 * Z, 200 - 150 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 180 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 5 * DayNightAdjust + 30 * Z
			FlBumperBulb(nr).BlendDisableLighting = 18 * DayNightAdjust + 3000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 8 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255,255 - 20 * Z,255 - 65 * Z)
			FlBumperSmallLight(nr).colorfull = RGB(255,255 - 20 * Z,255 - 65 * Z)
			MaterialColor "bumpertopmat" & nr, RGB(255,235 - z * 36,220 - Z * 90)
			
		Case "blacklight"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 1, RGB(30 - 27 * Z ^ 0.03,30 - 28 * Z ^ 0.01, 255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 20 + 900 * Z / (1 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 60 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 30000 * Z ^ 3
			Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 2000 * (Z ^ 3) / (1 + DNA90)
			FlBumperSmallLight(nr).color = RGB(255 - 240 * (Z ^ 0.1),255 - 240 * (Z ^ 0.1),255)
			FlBumperSmallLight(nr).colorfull = RGB(255 - 200 * z,255 - 200 * Z,255)
			MaterialColor "bumpertopmat" & nr, RGB(255 - 190 * Z,235 - z * 180,220 + 35 * Z)
			
		Case "yellow"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(255, 180 + 40 * z, 48 * Z), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 17 + 200 * Z / (1 + DNA30 ^ 2)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 40 * Z / (1 + DNA90)
			FlBumperBulb(nr).BlendDisableLighting = 12 * DayNightAdjust + 2000 * (0.03 * Z + 0.97 * Z ^ 10)
			Flbumperbiglight(nr).intensity = 10 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 1000 * (Z ^ 3) / (1 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(255,200, 24 - 24 * z)
			
		Case "purple"
			UpdateMaterial "bumperbulbmat" & nr, 0, 0.75 , 0, 1 - Z, 1 - Z, 1 - Z, 0.9999, RGB(128 - 118 * Z - 32 * Sin(Z * 3.14), 32 - 26 * Z ,255), RGB(255,255,255), RGB(32,32,32), False, True, 0, 0, 0, 0
			FlBumperSmallLight(nr).intensity = 15 + 200 * Z / (0.5 + DNA30)
			FlBumperTop(nr).BlendDisableLighting = 3 * DayNightAdjust + 50 * Z
			FlBumperBulb(nr).BlendDisableLighting = 15 * DayNightAdjust + 10000 * (0.03 * Z + 0.97 * Z ^ 3)
			Flbumperbiglight(nr).intensity = 25 * Z / (1 + DNA45)
			FlBumperHighlight(nr).opacity = 4000 * (Z ^ 3) / (0.5 + DNA90)
			MaterialColor "bumpertopmat" & nr, RGB(128 - 60 * Z,32,255)
	End Select
End Sub

Sub BumperTimer_Timer
	Dim nr
	For nr = 1 To 6
		If FlBumperFadeActual(nr) < FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
			FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.8
			If FlBumperFadeActual(nr) > 0.99 Then FlBumperFadeActual(nr) = 1
			FlFadeBumper nr, FlBumperFadeActual(nr)
		End If
		If FlBumperFadeActual(nr) > FlBumperFadeTarget(nr) And FlBumperActive(nr)  Then
			FlBumperFadeActual(nr) = FlBumperFadeActual(nr) + (FlBumperFadeTarget(nr) - FlBumperFadeActual(nr)) * 0.4 / (FlBumperFadeActual(nr) + 0.1)
			If FlBumperFadeActual(nr) < 0.01 Then FlBumperFadeActual(nr) = 0
			FlFadeBumper nr, FlBumperFadeActual(nr)
		End If
	Next
End Sub

'******************************************************
'******  END FLUPPER BUMPERS
'******************************************************





'******************************************************
' 	ZTST:  Debug Shot Tester
'******************************************************

'****************************************************************
' Section; Debug Shot Tester v3.2
'
' 1.  Raise/Lower outlanes and drain posts by pressing 2 key
' 2.  Capture and Launch ball, Press and hold one of the buttons (W, E, R, Y, U, I, P, A) below to capture ball by flipper.  Release key to shoot ball
' 3.  To change the test shot angles, press and hold a key and use Flipper keys to adjust the shot angle.  Shot angles are saved into the User direction as cgamename.txt
' 4.  Set DebugShotMode = 0 to disable debug shot test code.
'
' HOW TO INSTALL: Copy all debug* objects from Layer 2 to table and adjust. Copy the Debug Shot Tester code section to the script.
'	Add "DebugShotTableKeyDownCheck keycode" to top of Table1_KeyDown sub and add "DebugShotTableKeyUpCheck keycode" to top of Table1_KeyUp sub
'****************************************************************
Const DebugShotMode = 1 'Set to 0 to disable.  1 to enable
Dim DebugKickerForce
DebugKickerForce = 55

' Enable Disable Outlane and Drain Blocker Wall for debug testing
Dim DebugBLState
debug_BLW1.IsDropped = 1
debug_BLP1.Visible = 0
debug_BLR1.Visible = 0
debug_BLW2.IsDropped = 1
debug_BLP2.Visible = 0
debug_BLR2.Visible = 0
debug_BLW3.IsDropped = 1
debug_BLP3.Visible = 0
debug_BLR3.Visible = 0

Sub BlockerWalls
	DebugBLState = (DebugBLState + 1) Mod 4
	'	debug.print "BlockerWalls"
	PlaySound ("Start_Button")
	
	Select Case DebugBLState
		Case 0
			debug_BLW1.IsDropped = 1
			debug_BLP1.Visible = 0
			debug_BLR1.Visible = 0
			debug_BLW2.IsDropped = 1
			debug_BLP2.Visible = 0
			debug_BLR2.Visible = 0
			debug_BLW3.IsDropped = 1
			debug_BLP3.Visible = 0
			debug_BLR3.Visible = 0
			
		Case 1
			debug_BLW1.IsDropped = 0
			debug_BLP1.Visible = 1
			debug_BLR1.Visible = 1
			debug_BLW2.IsDropped = 0
			debug_BLP2.Visible = 1
			debug_BLR2.Visible = 1
			debug_BLW3.IsDropped = 0
			debug_BLP3.Visible = 1
			debug_BLR3.Visible = 1
			
		Case 2
			debug_BLW1.IsDropped = 0
			debug_BLP1.Visible = 1
			debug_BLR1.Visible = 1
			debug_BLW2.IsDropped = 0
			debug_BLP2.Visible = 1
			debug_BLR2.Visible = 1
			debug_BLW3.IsDropped = 1
			debug_BLP3.Visible = 0
			debug_BLR3.Visible = 0
			
		Case 3
			debug_BLW1.IsDropped = 1
			debug_BLP1.Visible = 0
			debug_BLR1.Visible = 0
			debug_BLW2.IsDropped = 1
			debug_BLP2.Visible = 0
			debug_BLR2.Visible = 0
			debug_BLW3.IsDropped = 0
			debug_BLP3.Visible = 1
			debug_BLR3.Visible = 1
	End Select
End Sub

Sub DebugShotTableKeyDownCheck (Keycode)
	'Cycle through Outlane/Centerlane blocking posts
	'-----------------------------------------------
	If Keycode = 3 Then
		BlockerWalls
	End If
	
	If DebugShotMode = 1 Then
		'Capture and launch ball:
		'	Press and hold one of the buttons (W, E, R, T, Y, U, I, P) below to capture ball by flipper.  Release key to shoot ball
		'	To change the test shot angles, press and hold a key and use Flipper keys to adjust the shot angle.
		'--------------------------------------------------------------------------------------------
		If keycode = 17 Then 'W key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleW
		End If
		If keycode = 18 Then 'E key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleE
		End If
		If keycode = 19 Then 'R key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleR
		End If
		If keycode = 21 Then 'Y key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleY
		End If
		If keycode = 22 Then 'U key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleU
		End If
		If keycode = 23 Then 'I key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleI
		End If
		If keycode = 25 Then 'P key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleP
		End If
		If keycode = 30 Then 'A key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleA
		End If
		If keycode = 31 Then 'S key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleS
		End If
		If keycode = 33 Then 'F key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleF
		End If
		If keycode = 34 Then 'G key
			debugKicker.enabled = True
			TestKickerVar = TestKickAngleG
		End If
		
		If debugKicker.enabled = True Then		'Use Flippers to adjust angle while holding key
			If keycode = LeftFlipperKey Then
				debugKickAim.Visible = True
				TestKickerVar = TestKickerVar - 1
				Debug.print TestKickerVar
			ElseIf keycode = RightFlipperKey Then
				debugKickAim.Visible = True
				TestKickerVar = TestKickerVar + 1
				Debug.print TestKickerVar
			End If
			debugKickAim.ObjRotz = TestKickerVar
		End If
	End If
End Sub


Sub DebugShotTableKeyUpCheck (Keycode)
	' Capture and launch ball:
	' Release to shoot ball. Set up angle and force as needed for each shot.
	'--------------------------------------------------------------------------------------------
	If DebugShotMode = 1 Then
		If keycode = 17 Then 'W key
			TestKickAngleW = TestKickerVar
			debugKicker.kick TestKickAngleW, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 18 Then 'E key
			TestKickAngleE = TestKickerVar
			debugKicker.kick TestKickAngleE, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 19 Then 'R key
			TestKickAngleR = TestKickerVar
			debugKicker.kick TestKickAngleR, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 21 Then 'Y key
			TestKickAngleY = TestKickerVar
			debugKicker.kick TestKickAngleY, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 22 Then 'U key
			TestKickAngleU = TestKickerVar
			debugKicker.kick TestKickAngleU, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 23 Then 'I key
			TestKickAngleI = TestKickerVar
			debugKicker.kick TestKickAngleI, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 25 Then 'P key
			TestKickAngleP = TestKickerVar
			debugKicker.kick TestKickAngleP, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 30 Then 'A key
			TestKickAngleA = TestKickerVar
			debugKicker.kick TestKickAngleA, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 31 Then 'S key
			TestKickAngleS = TestKickerVar
			debugKicker.kick TestKickAngleS, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 33 Then 'F key
			TestKickAngleF = TestKickerVar
			debugKicker.kick TestKickAngleF, DebugKickerForce
			debugKicker.enabled = False
		End If
		If keycode = 34 Then 'G key
			TestKickAngleG = TestKickerVar
			debugKicker.kick TestKickAngleG, DebugKickerForce
			debugKicker.enabled = False
		End If
		
		'		EXAMPLE CODE to set up key to cycle through 3 predefined shots
		'		If keycode = 17 Then	 'Cycle through all left target shots
		'			If TestKickerAngle = -28 then
		'				TestKickerAngle = -24
		'			ElseIf TestKickerAngle = -24 Then
		'				TestKickerAngle = -19
		'			Else
		'				TestKickerAngle = -28
		'			End If
		'			debugKicker.kick TestKickerAngle, DebugKickerForce: debugKicker.enabled = false			 'W key
		'		End If
		
	End If
	
	If (debugKicker.enabled = False And debugKickAim.Visible = True) Then 'Save Angle changes
		debugKickAim.Visible = False
		SaveTestKickAngles
	End If
End Sub

Dim TestKickerAngle, TestKickerAngle2, TestKickerVar, TeskKickKey, TestKickForce
Dim TestKickAngleWDefault, TestKickAngleEDefault, TestKickAngleRDefault, TestKickAngleYDefault, TestKickAngleUDefault, TestKickAngleIDefault
Dim TestKickAnglePDefault, TestKickAngleADefault, TestKickAngleSDefault, TestKickAngleFDefault, TestKickAngleGDefault
Dim TestKickAngleW, TestKickAngleE, TestKickAngleR, TestKickAngleY, TestKickAngleU, TestKickAngleI
Dim TestKickAngleP, TestKickAngleA, TestKickAngleS, TestKickAngleF, TestKickAngleG
TestKickAngleWDefault =  - 27
TestKickAngleEDefault =  - 20
TestKickAngleRDefault =  - 14
TestKickAngleYDefault =  - 8
TestKickAngleUDefault =  - 3
TestKickAngleIDefault = 1
TestKickAnglePDefault = 5
TestKickAngleADefault = 11
TestKickAngleSDefault = 17
TestKickAngleFDefault = 19
TestKickAngleGDefault = 5
If DebugShotMode = 1 Then LoadTestKickAngles

Sub SaveTestKickAngles
	Dim FileObj, OutFile
	Set FileObj = CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) Then Exit Sub
	Set OutFile = FileObj.CreateTextFile(UserDirectory & cGameName & ".txt", True)
	
	OutFile.WriteLine TestKickAngleW
	OutFile.WriteLine TestKickAngleE
	OutFile.WriteLine TestKickAngleR
	OutFile.WriteLine TestKickAngleY
	OutFile.WriteLine TestKickAngleU
	OutFile.WriteLine TestKickAngleI
	OutFile.WriteLine TestKickAngleP
	OutFile.WriteLine TestKickAngleA
	OutFile.WriteLine TestKickAngleS
	OutFile.WriteLine TestKickAngleF
	OutFile.WriteLine TestKickAngleG
	OutFile.Close
	
	Set OutFile = Nothing
	Set FileObj = Nothing
End Sub

Sub LoadTestKickAngles
	Dim FileObj, OutFile, TextStr
	
	Set FileObj = CreateObject("Scripting.FileSystemObject")
	If Not FileObj.FolderExists(UserDirectory) Then
		MsgBox "User directory missing"
		Exit Sub
	End If
	
	If FileObj.FileExists(UserDirectory & cGameName & ".txt") Then
		Set OutFile = FileObj.GetFile(UserDirectory & cGameName & ".txt")
		Set TextStr = OutFile.OpenAsTextStream(1,0)
		If (TextStr.AtEndOfStream = True) Then
			Exit Sub
		End If
		
		TestKickAngleW = TextStr.ReadLine
		TestKickAngleE = TextStr.ReadLine
		TestKickAngleR = TextStr.ReadLine
		TestKickAngleY = TextStr.ReadLine
		TestKickAngleU = TextStr.ReadLine
		TestKickAngleI = TextStr.ReadLine
		TestKickAngleP = TextStr.ReadLine
		TestKickAngleA = TextStr.ReadLine
		TestKickAngleS = TextStr.ReadLine
		TestKickAngleF = TextStr.ReadLine
		TestKickAngleG = TextStr.ReadLine
		TextStr.Close
	Else
		'create file
		TestKickAngleW = TestKickAngleWDefault
		TestKickAngleE = TestKickAngleEDefault
		TestKickAngleR = TestKickAngleRDefault
		TestKickAngleY = TestKickAngleYDefault
		TestKickAngleU = TestKickAngleUDefault
		TestKickAngleI = TestKickAngleIDefault
		TestKickAngleP = TestKickAnglePDefault
		TestKickAngleA = TestKickAngleADefault
		TestKickAngleS = TestKickAngleSDefault
		TestKickAngleF = TestKickAngleFDefault
		TestKickAngleG = TestKickAngleGDefault
		SaveTestKickAngles
	End If
	
	Set OutFile = Nothing
	Set FileObj = Nothing
	
End Sub
'****************************************************************
' End of Section; Debug Shot Tester 3.2
'****************************************************************





'***************************************************************
' 	ZQUE: VPIN WORKSHOP ADVANCED QUEUING SYSTEM - 1.1.1
'***************************************************************
' WHAT IS IT?
' The VPin Workshop Advanced Queuing System allows table authors
' to put sub routine calls in a queue without creating a bunch
' of timers. There are many use cases for this: queuing sequences
' for light shows and DMD scenes, delaying solenoids until the
' DMD is finished playing all its sequences (such as holding a
' ball in a scoop), managing what actions take priority over
' others (e.g. an extra ball sequence is probably more important
' than a small jackpot), and many more.
'
' This system uses Scripting.Dictionary, a single timer, and the
' GameTime global to keep track of everything in the queue.
' This allows for better stability and a virtually unlimited
' number of items in the queue. It also allows for greater
' versatility, like pre-delays, queue delays, priorities, and
' even modifying items in the queue.
'
' The VPin Workshop Queuing System can replace vpmTimer as a
' proper queue system (each item depends on the previous)
' whereas vpmTimer is a collection of virtual timers that run
' in parallel. It also adds on other advanced functionality.
' However, this queue system does not have ROM support out of
' the box like vpmTimer does.
'
' I recommend reading all the comments before you implement the
' queuing system into your table.
'
' WHAT YOU NEED to use the queuing system:
' 1) Put this VBS file in your scripts folder.
' 2) Include this file via Scripting.FileSystemObject, and
'	ExecuteGlobal it.
' 3) Make one or more queues by constructing the vpwQueueManager:
'	Dim queue : Set queue = New vpwQueueManager
' 4) Create (or use) a timer that is always enabled and
'	preferably has an interval of 1 millisecond. Use a
'	higher number for less time precision but less resource
'	use. You only need one timer even if you
'	have multiple queues.
' 5) For each queue you created, call its Tick routine in
'	the timer's *_timer() routine:
'	queue.Tick
' 6) You're done! Refer to the routines in vpwQueueManager to
'	learn how to use the queuing system.
'***************************************************************

'===========================================
' vpwQueueManager
' This class manages a queue of
' vpwQueueItems and executes them.
'===========================================
Class vpwQueueManager
	Public qItems	   ' A dictionary of vpwQueueItems in the queue (do NOT use native Scripting.Dictionary.Add/Remove; use the vpwQueueManager's Add/Remove methods instead!)
	Public preQItems	' A dictionary of vpwQueueItems pending to be added to qItems
	Public debugOn	  ' Null = no debug. String = activate debug by using this unique label for the queue. REQUIRES baldgeek's error logs.
	
	'----------------------------------------------------------
	' vpwQueueManager.qCurrentItem
	' This contains a string of the key currently active / at
	' the top of the queue. An empty string means no items are
	' active right now.
	' This is an important property; it should be monitored
	' in another timer or routine whenever you Add a queue item
	' with a -1 (indefinite) preDelay or postDelay. Then, for
	' preDelay, ExecuteCurrentItem should be called to run the
	' queue item. And for postDelay, DoNextItem should be
	' called to move to the next item in the queue.
	'
	' For example, let's say you add a queue item with the
	' key "kickTheBall" and an indefinite preDelay. You want
	' to wait until another timer fires before this queue item
	' executes and kicks the ball out of a scoop. In the other
	' timer, you will monitor qCurrentItem. Once it equals
	' "kickTheBall", call ExecuteCurrentItem, which will run
	' the queue item and presumably kick out the ball.
	'
	' WARNING!: If you do not properly execute one of these
	' callback routines on an indefinite delayed item, then
	' the queue will effectively freeze / stop until you do.
	'---------------------------------------------------------
	Public qCurrentItem
	
	Public preDelayTime	 ' The GameTime the preDelay for the qCurrentItem was started
	Public postDelayTime	' The GameTime the postDelay for the qCurrentItem was started
	
	Private onQueueEmpty	' A string or object to be called every time the queue empties (use the QueueEmpty property to get/set this)
	Private queueWasEmpty   ' Boolean to determine if the queue was already empty when firing DoNextItem
	
	Private Sub Class_Initialize
		Set qItems = CreateObject("Scripting.Dictionary")
		Set preQItems = CreateObject("Scripting.Dictionary")
		qCurrentItem = ""
		onQueueEmpty = ""
		queueWasEmpty = True
		debugOn = Null
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Tick
	' This is where all the magic happens! Call this method in
	' your timer's _timer routine to check the queue and
	' execute the necessary methods. We do not iterate over
	' every item in the queue here, which allows for superior
	' performance even if you have hundreds of items in the
	' queue.
	'----------------------------------------------------------
	Public Sub Tick()
		Dim item
		If qItems.Count > 0 Then ' Don't waste precious resources if we have nothing in the queue
			' If no items are active, or the currently active item no longer exists, move to the next item in the queue.
			' (This is also a failsafe to ensure the queue continues to work even if an item gets manually deleted from the dictionary).
			If qCurrentItem = "" Or Not qItems.Exists(qCurrentItem) Then
				DoNextItem
			Else ' We are good; do stuff as normal
				Set item = qItems.item(qCurrentItem)
				
				If item.Executed Then
					' If the current item was executed and the post delay passed, go to the next item in the queue
					If item.postDelay >= 0 And GameTime >= (postDelayTime + item.postDelay) Then
						DebugLog qCurrentItem & " - postDelay of " & item.postDelay & " passed."
						DoNextItem
					End If
				Else
					' If the current item expires before it can be executed, go to the next item in the queue
					If item.timeToLive > 0 And GameTime >= (item.queuedOn + item.timeToLive) Then
						DebugLog qCurrentItem & " - expired (Time To live). Moving To the Next queue item."
						DoNextItem
					End If
					
					' If the current item was not executed yet and the pre delay passed, then execute it
					If item.preDelay >= 0 And GameTime >= (preDelayTime + item.preDelay) Then
						DebugLog qCurrentItem & " - preDelay of " & item.preDelay & " passed. Executing callback."
						item.Execute
						preDelayTime = 0
						postDelayTime = GameTime
					End If
				End If
			End If
		End If
		
		' Loop through each item in the pre-queue to find any that is ready to be added
		If preQItems.Count > 0 Then
			Dim k, key
			k = preQItems.Keys
			For Each key In k
				Set item = preQItems.Item(key)
				
				' If a queue item was pre-queued and is ready to be considered as actually in the queue, add it
				If GameTime >= (item.queuedOn + item.preQueueDelay) Then
					DebugLog key & " (preQueue) - preQueueDelay of " & item.preQueueDelay & " passed. Item added To the main queue."
					preQItems.Remove key
					item.preQueueDelay = 0
					item.queuedOn = GameTime
					qItems.Add key, item
					queueWasEmpty = False
				End If
			Next
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.DoNextItem
	' Goes to the next item in the queue and deletes the
	' currently active one.
	'----------------------------------------------------------
	Public Sub DoNextItem()
		If Not qCurrentItem = "" Then
			If qItems.Exists(qCurrentItem) Then qItems.Remove qCurrentItem ' Remove the current item from the queue if it still exists
			qCurrentItem = ""
		End If
		
		If qItems.Count > 0 Then
			Dim k, key
			Dim nextItem
			Dim nextItemPriority
			Dim nextItemTimeToLive
			Dim nextItemQueuedOn
			Dim item
			nextItemPriority = 0
			nextItem = ""
			
			' Find which item needs to run next based on priority first, queue order second (ignore items with an active preQueueDelay)
			k = qItems.Keys
			For Each key In k
				Set item = qItems.Item(key)
				
				If item.preQueueDelay <= 0 And item.priority > nextItemPriority Then
					nextItem = key
					nextItemPriority = item.priority
					nextItemTimeToLive = item.timeToLive
					nextItemQueuedOn = item.queuedOn
				End If
			Next
			
			If qItems.Exists(nextItem) Then
				DebugLog "DoNextItem - checking " & nextItem & " (priority " & nextItemPriority & ")"
				
				' Make sure the item is not expired. If it is, remove it and re-call doNextItem
				If nextItemTimeToLive > 0 And GameTime >= (nextItemQueuedOn + nextItemTimeToLive) Then
					DebugLog "DoNextItem - " & nextItem & " expired (Time To live). Removing And going To the Next item."
					qItems.Remove nextItem
					DoNextItem
					Exit Sub
				End If
				
				' Set item as current / active, and execute if it has no pre-delay (otherwise Tick will take care of pre-delay)
				qCurrentItem = nextItem
				Set item = qItems.Item(nextItem)
				If item.preDelay = 0 Then
					DebugLog "DoNextItem - " & nextItem & " Now active. It has no preDelay, so executing callback immediately."
					item.Execute
					preDelayTime = 0
					postDelayTime = GameTime
				Else
					DebugLog "DoNextItem - " & nextItem & " Now active. Waiting For a preDelay of " & item.preDelay & " before executing."
					preDelayTime = GameTime
					postDelayTime = 0
				End If
			End If
		ElseIf queueWasEmpty = False Then
			DebugLog "DoNextItem - Queue Is Now Empty; executing queueEmpty callback."
			CallQueueEmpty() ' Call QueueEmpty if this was the last item in the queue
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.ExecuteCurrentItem
	' Helper routine that can be used when the current item is
	' on an indefinite preDelay. Call this when you are ready
	' for that item to execute.
	'----------------------------------------------------------
	Public Sub ExecuteCurrentItem()
		If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
			DebugLog "ExecuteCurrentItem - Executing the callback For " & qCurrentItem & "."
			Dim item
			Set item = qItems.Item(qCurrentItem)
			item.Execute
			preDelayTime = 0
			postDelayTime = GameTime
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Add
	' REQUIRES Class vpwQueueItem
	'
	' Add an item to the queue.
	'
	' PARAMETERS:
	'
	' key (string) - Unique name for this queue item
	' (warning: Specifying a key that already exists will
	'  overwrite the item in the queue)
	'
	' qCallback (object|string) - An object to be called,
	' or string to be executed globally, when this queue item
	' runs. I highly recommend making sub routines for groups
	' of things that should be executed by the queue so that
	' your qCallback string does not get long, and you can
	' easily organize your callbacks. Also, use double
	' double-quotes when the call itself has quotes in it
	' (VBScript escaping).
	' Example: "playsound ""Plunger"""
	'
	' priority (number) - Items in the queue will be executed
	' in order from highest priority to lowest. Items with the
	' same priority will be executed in order according to
	' when they were added to the queue. Use any number
	' greater than 0. My recommendation is to make a plan for
	' your table on how you will prioritize various types of
	' queue items and what priority number each type should
	' have. Also, you should reserve priority 1 (lowest) to
	' items which should wait until everything else in the
	' queue is done (such as ejecting a ball from a scoop).
	'
	' preQueueDelay (number) - The number of
	' milliseconds before the queue actually considers this
	' item as "in the queue" (pretend you started a timer to
	' add this item into the queue after this delay; this
	' logically works in a similar way; the only difference is
	' timeToLive is still considered even when an item is
	' pre-queued.) Set to 0 to add to the queue immediately.
	' NOTE: this should be less than timeToLive.
	'
	' preDelay (number) - The number of milliseconds before
	' the qCallback executes once this item is active (top)
	' in the queue. Set this to 0 to immediately execute the
	' qCallback when this item becomes active.
	' Set this to -1 to have an indefinite delay until
	' vpwQueueManager.ExecuteCurrentItem is called (see the
	' comment for qCurrentItem for more information).
	' NOTE: this should be less than timeToLive. And, if
	' timeToLive runs out before preDelay runs out, the item
	' will be removed and will not execute.
	'
	' postDelay (number) - After the qCallback executes, the
	' number of milliseconds before moving on to the next item
	' in the queue. Set this to -1 to have an indefinite delay
	' until vpwQueueManager.DoNextItem is called (see the
	' comment for qCurrentItem for more information).
	'
	' timeToLive (number) - After this item is added to the
	' queue, the number of milliseconds before this queue item
	' expires / is removed if the qCallback is not executed by
	' then. Set to 0 to never expire. NOTE: If not 0, this should
	' be greater than preDelay + preQueueDelay or the item will
	' expire before the qCallback is executed.
	' Example use case: Maybe a player scored a jackpot, but
	' it would be awkward / irrelevant to play that jackpot
	' sequence if it hasn't played after a few seconds (e.g.
	' other items in the queue took priority).
	'
	' executeNow (boolean) - Specify true if this item
	' should interrupt the queue and run immediately. This
	' will only happen, however, if the currently active item
	' has a priority less than or equal to the item you are
	' adding. Note this does not bypass preQueueDelay nor
	' preDelay if set.
	' Example: If a player scores an extra ball, you might
	' want that to interrupt everything else going on as it
	' is an important milestone.
	'----------------------------------------------------------
	Public Sub Add(key, qCallback, priority, preQueueDelay, preDelay, postDelay, timeToLive, executeNow)
		DebugLog "Added " & key
		
		' Remove duplicate if it exists
		If preQueueDelay <= 0 And qItems.Exists(key) Then
			DebugLog key & " (Add) - Already exists In the queue. Replacing With the new one."
			qItems.Remove key
			If qCurrentItem = key Then qCurrentItem = "" ' Prevent infinite loops if this queue item has no preDelay and re-adds itself via the callback
		End If
		If preQueueDelay > 0 And preQItems.Exists(key) Then
			preQItems.Remove key
		End If
		
		' Construct the item class
		Dim newClass
		Set newClass = New vpwQueueItem
		With newClass
			.Callback = qCallback
			.priority = priority
			.preQueueDelay = preQueueDelay
			.preDelay = preDelay
			.postDelay = postDelay
			.timeToLive = timeToLive
		End With
		
		' Determine execution stuff if this item does not have a pre-queue delay
		If preQueueDelay <= 0 Then
			If executeNow = True Then
				' Make sure this item does not immediately execute if the current item has a higher priority
				If Not qCurrentItem = "" And qItems.Exists(qCurrentItem) Then
					Dim item
					Set item = qItems.Item(qCurrentItem)
					If item.priority <= priority Then
						DebugLog key & " (Add) - Execute Now was Set To True And this item's priority (" & priority & ") Is >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). Making it the current active queue item."
						qItems.Remove qCurrentItem ' TODO: Do we really want to remove an item if it has not executed yet (preDelay)?
						qCurrentItem = key
						If preDelay = 0 Then
							DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
							newClass.Execute
							preDelayTime = 0
							postDelayTime = GameTime
						Else
							DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
							preDelayTime = GameTime
							postDelayTime = 0
						End If
					Else
						DebugLog key & " (Add) - Execute Now was Set To True, but this item's priority (" & priority & ") Is Not >= the active item's priority (" & item.priority & " from " & qCurrentItem & "). This item will Not be executed Now And will be added To the queue normally."
					End If
				Else
					DebugLog key & " (Add) - Execute Now was Set To True And no item was active In the queue. Making it the current active queue item."
					qCurrentItem = key
					If preDelay = 0 Then
						DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
						newClass.Execute
						preDelayTime = 0
						postDelayTime = GameTime
					Else
						DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
						preDelayTime = GameTime
						postDelayTime = 0
					End If
				End If
			ElseIf qCurrentItem = key Then
				DebugLog key & " (Add) - An item With the same key Is currently the active queue item. Making this item the active one Now."
				If preDelay = 0 Then
					DebugLog key & " (Add) - No pre-delay. Executing the callback immediately."
					newClass.Execute
					preDelayTime = 0
					postDelayTime = GameTime
				Else
					DebugLog key & " (Add) - Waiting For a pre-delay of " & preDelay & " before executing the callback."
					preDelayTime = GameTime
					postDelayTime = 0
				End If
			Else
				DebugLog key & " (Add) - Execute Now was False. This item was added To the queue."
			End If
			qItems.Add key, newClass
			queueWasEmpty = False
		Else
			DebugLog key & " (Add) - Not actually added To the queue yet. It has a pre-queue delay of " & preQueueDelay
			preQItems.Add key, newClass
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.Remove
	'
	' Removes an item from the queue. It is better to use this
	' than to remove the item from qItems directly as this sub
	' will also call DoNextItem to advance the queue if
	' the item removed was the active item.
	' NOTE: This only removes items from qItems; to remove
	' an item from preQItems, use the standard
	' Scripting.Dictionary Remove method.
	'
	' PARAMETERS:
	'
	' key (string) - Unique name of the queue item to remove.
	'----------------------------------------------------------
	Public Sub Remove(key)
		If qItems.Exists(key) Then
			DebugLog key & " (Remove)"
			qItems.Remove key
			If qCurrentItem = key Or qCurrentItem = "" Then DoNextItem ' Ensure the queue does not get stuck
		End If
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueManager.RemoveAll
	'
	' Removes all items from the queue / clears the queue.
	' It is better to call this sub than to remove all items
	' from qItems directly because this sub cleans up the queue
	' to ensure it continues to work properly.
	'
	' PARAMETERS:
	'
	' preQueue (boolean) - Also clear the pre-queue.
	'----------------------------------------------------------
	Public Sub RemoveAll(preQueue)
		DebugLog "Queue was emptied via RemoveAll."
		
		' Loop through each item in the queue and remove it
		Dim k, key
		k = qItems.Keys
		For Each key In k
			qItems.Remove key
		Next
		qCurrentItem = ""
		
		If queueWasEmpty = False Then CallQueueEmpty() ' Queue is now empty, so call our callback if applicable
		
		If preQueue Then
			k = preQItems.Keys
			For Each key In k
				preQItems.Remove key
			Next
		End If
	End Sub
	
	'----------------------------------------------------------
	' Get vpwQueueManager.QueueEmpty
	' Get the current callback for when the queue is empty.
	'----------------------------------------------------------
	Public Property Get QueueEmpty()
		If IsObject(onQueueEmpty) Then
			Set QueueEmpty = onQueueEmpty
		Else
			QueueEmpty = onQueueEmpty
		End If
	End Property
	
	'----------------------------------------------------------
	' Let vpwQueueManager.QueueEmpty
	' Set the callback to call every time the queue empties.
	' This could be useful for setting a sub routine to be
	' called each time the queue empties for doing things such
	' as ejecting balls from scoops. Unlike using the Add
	' method, this callback is immune from getting removed by
	' higher priority items in the queue and will be called
	' every time the queue is emptied, not just once.
	'
	' PARAMETERS:
	'
	' callback (object|string) - The callback to call every
	' time the queue empties.
	'----------------------------------------------------------
	Public Property Let QueueEmpty(callback)
		If IsObject(callback) Then
			Set onQueueEmpty = callback
		ElseIf VarType(callback) = vbString Then
			onQueueEmpty = callback
		End If
	End Property
	
	'----------------------------------------------------------
	' Get vpwQueueManager.CallQueueEmpty
	' Private method that actually calls the QueueEmpty
	' callback.
	'----------------------------------------------------------
	Private Sub CallQueueEmpty()
		If queueWasEmpty = True Then Exit Sub
		
		If IsObject(onQueueEmpty) Then
			Call onQueueEmpty(0)
		ElseIf VarType(onQueueEmpty) = vbString Then
			If onQueueEmpty > "" Then ExecuteGlobal onQueueEmpty
		End If
		
		queueWasEmpty = True
	End Sub
	
	'----------------------------------------------------------
	' DebugLog
	' Log something if debugOn is not null.
	' REQUIRES / uses the WriteToLog sub from Baldgeek's
	' error log library.
	'----------------------------------------------------------
	Private Sub DebugLog(message)
		If Not IsNull(debugOn) Then
			WriteToLog "VPW Queue " & debugOn, message
		End If
	End Sub
End Class

'===========================================
' vpwQueueItem
' Represents a single item for the queue
' system. Do NOT use this class directly.
' Instead, use the vpwQueueManager.Add
' routine.

' You can, however, access an individual
' item in the queue via
' vpwQueueManager.qItems and then modify
' its properties while it is still in the
' queue.
'===========================================
Class vpwQueueItem
	Public priority		 ' The item's set priority
	Public timeToLive	   ' The item's set timeToLive milliseconds requested
	Public preQueueDelay	' The item's pre-queue milliseconds requested
	Public preDelay		 ' The item's pre delay milliseconds requested
	Public postDelay		' The item's post delay milliseconds requested
	Private qCallback	   ' The item's callback object or string (use the Callback property on the class to get/set it)
	
	Public executed		 ' Whether or not this item's qCallback was executed yet
	Public queuedOn		 ' The game time this item was added to the queue
	Public executedOn	   ' The game time this item was executed
	
	Private Sub Class_Initialize
		' Defaults
		priority = 0
		timeToLive = 0
		preQueueDelay = 0
		preDelay = 0
		postDelay = 0
		qCallback = ""
		queuedOn = GameTime
		executedOn = 0
	End Sub
	
	'----------------------------------------------------------
	' vpwQueueItem.Execute
	' Executes the qCallback on this item if it was not yet
	' already executed.
	'----------------------------------------------------------
	Public Sub Execute()
		If executed Then Exit Sub ' Do not allow an item's qCallback to ever Execute more than one time
		
		' Execute qCallback
		If IsObject(qCallback) Then
			Call qCallback(0)
		ElseIf VarType(qCallback) = vbString Then
			If qCallback > "" Then ExecuteGlobal qCallback
		End If
		
		executedOn = GameTime
		executed = True
	End Sub
	
	Public Property Get Callback()
		If IsObject(qCallback) Then
			Set Callback = qCallback
		Else
			Callback = qCallback
		End If
	End Property
	
	Public Property Let Callback(cb)
		If IsObject(cb) Then
			Set qCallback = cb
		ElseIf VarType(cb) = vbString Then
			qCallback = cb
		End If
	End Property
End Class


'*******************************************
' Routine called each time the queue is
' emptied.
'*******************************************
Sub QueueEmptyRoutine()
	ShowScene flexScenes(0), FlexDMD_RenderMode_DMD_RGB, 1 'Return FlexDMD back to score board
	
	' There is a ball waiting to be ejected from the kicker
	If ballInKicker1 = True Then
		' Prepare to eject the ball
		Kicker1.timerenabled = True
		
		' Warn the player with flashers the ball is about to come out
		Flash1 True
		vpmTimer.AddTimer 500,"Flash1 False'"
		Flash2 True
		vpmTimer.AddTimer 500,"Flash2 False'"
		
		' Reset the drop targets in this example
		SolDTBank123 1
		
		ballInKicker1 = False 'Set this to false now so that this does not get triggered a second time
	End If
End Sub

'*******************************************
' Bonus X example
'*******************************************

Sub sw6_Hit()
	l6.state = 1
	BonusX(0) = True
	CheckBonusX
End Sub

Sub sw7_Hit()
	l7.state = 1
	BonusX(1) = True
	CheckBonusX
End Sub

Sub sw8_Hit()
	l8.state = 1
	BonusX(2) = True
	CheckBonusX
End Sub

Sub sw9_Hit()
	l9.state = 1
	BonusX(3) = True
	CheckBonusX
End Sub

Sub CheckBonusX
	If BonusX(0) = True And BonusX(1) = True And BonusX(2) = True And BonusX(3) = True Then
		Dim i
		For i = 0 To 4
			BonusX(i) = False
			AllLamps(i + 6).state = 0
		Next
		queue.Add "flexBonusX", "ShowScene flexScenes(2), FlexDMD_RenderMode_DMD_RGB, 3", 2, 0, 0, 5000, 0, False
	End If
End Sub


'***************************************************************
'***  END VPIN WORKSHOP ADVANCED QUEUING SYSTEM
'***************************************************************







'*****************************************************************************************************************************************
'  	ZLOG: ERROR LOGS by baldgeek
'*****************************************************************************************************************************************

' Log File Usage:
'   WriteToLog "Label 1", "Message 1 "
'   WriteToLog "Label 2", "Message 2 "

Class DebugLogFile
	Private Filename
	Private TxtFileStream
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
		GetTimeStamp = _
		LZ(Year(CurrTime),   4) & "-" _
		& LZ(Month(CurrTime),  2) & "-" _
		& LZ(Day(CurrTime),	2) & " " _
		& LZ(Hour(CurrTime),   2) & ":" _
		& LZ(Minute(CurrTime), 2) & ":" _
		& LZ(Second(CurrTime), 2) & ":" _
		& LZ(MilliSecs, 4)
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message, code)
		Dim FormattedMsg, Timestamp
		'   Filename = UserDirectory + "\" + cGameName + "_debug_log.txt"
		Filename = cGameName + "_debug_log.txt"
		
		Set TxtFileStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(Filename, code, True)
		Timestamp = GetTimeStamp
		FormattedMsg = GetTimeStamp + " : " + label + " : " + message
		TxtFileStream.WriteLine FormattedMsg
		TxtFileStream.Close
		Debug.print label & " : " & message
	End Sub
End Class

Sub WriteToLog(label, message)
	If KeepLogs Then
		Dim LogFileObj
		Set LogFileObj = New DebugLogFile
		LogFileObj.WriteToLog label, message, 8
	End If
End Sub

Sub NewLog()
	If KeepLogs Then
		Dim LogFileObj
		Set LogFileObj = New DebugLogFile
		LogFileObj.WriteToLog "NEW Log", " ", 2
	End If
End Sub

'*****************************************************************************************************************************************
'***  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************






'**********************************************************************************************************
' 	ZCRD:  Instruction Card Zoom
'**********************************************************************************************************

Dim CardCounter, ScoreCard

Sub CardTimer_Timer
	If scorecard = 1 Then
		CardCounter = CardCounter + 2
		If CardCounter > 50 Then CardCounter = 50
	Else
		CardCounter = CardCounter - 4
		If CardCounter < 0 Then CardCounter = 0
	End If
	InstructionCard.transX = CardCounter * 6
	InstructionCard.transY = CardCounter * 6
	InstructionCard.transZ =  - cardcounter * 2
	'   InstructionCard.objRotX = -cardcounter/2
	InstructionCard.size_x = 1 + CardCounter / 25
	InstructionCard.size_y = 1 + CardCounter / 25
	If CardCounter = 0 Then
		CardTimer.Enabled = False
		InstructionCard.visible = 0
	Else
		InstructionCard.visible = 1
	End If
End Sub

'**********************************************************************************************************
'***  Instruction Card Zoom
'**********************************************************************************************************






'*******************************************
' 	ZVRR: VR Room / VR Cabinet
'*******************************************

Dim VRThings
If VRRoom <> 0 Then
	'	DMDPlayfield.dmd=False
	DMDPlayfield.visible = False
	'	DMDbackbox.DMD=True
	DMDbackbox.visible = True  ' flasher dmd
	If VRRoom = 1 Then
		For Each VRThings In VR_Cab
			VRThings.visible = 1
		Next
	End If
	If VRRoom = 2 Then
		For Each VRThings In VR_Cab
			VRThings.visible = 0
		Next
		PinCab_Backglass.visible = 1
	End If
Else
	If FlexONPlayfield Then
		'   DMDPlayfield.dmd=True
		DMDPlayfield.visible = True
	End If
	For Each VRThings In VR_Cab
		VRThings.visible = 0
	Next
End If
