'  Gottlieb's Royal Flush (1957) by steveman0  
'
'*********************************************************************************************************************************
' === TABLE OF CONTENTS  ===
'
' You can quickly jump to a section by searching the four letter tag (ZXXX)
'
'	ZOPT: User Options
'	ZCON: Constants and Global Variables
'	ZDMD: FlexDMD
'	ZTIM: Timers
'   ZGAM: Core gameplay logic and scoring
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

' Table script based on VPW Example table, Game Logic Framework (GLF), and related resources
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

Const BallSize = 50				  'Ball diameter in VPX units; must be 50
Const BallMass = 1				  'Ball mass must be 1
Const tnob = 5					  'Total number of balls the table can hold
Const lob = 0					  'Locked balls
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
	glf_ballsPerGame = 5
	
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

'*******************************************
'	ZGAM: Core gameplay logic and scoring
'*******************************************

Sub TriggerLaneL_Hit
	AddScore 100000
End Sub

Sub TriggerLaneL_Hit
	AddScore 100000
End Sub
' '******************************************************
' ' 	Z3DI:   3D INSERTS
' '******************************************************
' ' No 3d inserts on this table (yet)
' '
' ' Before you get started adding the inserts to your playfield in VPX, there are a few things you need to have done to prepare:
' '	 1. Cut out all the inserts on the playfield image so there is alpha transparency where they should be.
' '	  Make sure the playfield material has Opacity Active checkbox checked.
' '	2. All the  insert text and/or images that lie over the insert plastic needs to be in its own file with
' '	   alpha transparency. Many playfields may require finding the original font and remaking the insert text.
' '
' ' To add the inserts:
' '	1. Import all the textures (images) and materials from this file that start with the word "Insert" into your Table
' '   2. Copy and past the two primitves that make up the insert you want to use. One primitive is for the on state, the other for the off state.
' '   3. Align the primitives with the associated insert light. Name the on and off primitives correctly.
' '   5. You will need to manually tweak the disable lighting value and material parameters to achielve the effect you want.
' '
' '
' ' Quick Reference:  Laying the Inserts ( Tutorial From Iaakki)
' ' - Each insert consists of two primitives. On and Off primitive. Suggested naming convention is to use lamp number in the name. For example
' '   is lamp number is 57, the On primitive is "p57" and the Off primitive is "p57off". This makes it easier to work on script side.
' ' - When starting from a new table, I'd first select to make few inserts that look quite similar. Lets say there is total of 6 small triangle
' '   inserts, 4 yellow and 2 blue ones.
' ' - Import the insert on/off images from the image manager and the vpx materials used from the sample project first, and those should appear
' '   selected properly in the primitive settings when you paste your actual insert trays in your target table . Then open up your target project
' '   at same time as the sample project and use copy&paste to copy desired inserts to target project.
' ' - There are quite many parameters in primitive that affect a lot how they will look. I wouldn't mess too much with them. Use Size options to
' '   scale the insert properly into PF hole. Some insert primitives may have incorrect pivot point, which means that changing the depth, you may
' '   also need to alter the Z-position too.
' ' - Once you have the first insert in place, wire it up in the script (detailed in part 3 below). Then set the light bulb's intensity to zero,
' '   so it won't harass the adjustment.
' ' - Start up the game with F6 and visually see if the On-primitive blinks properly. If it is too dim, hit D and open editor. Write:
' ' - p57.BlendDisableLighting = 300 and hit enter
' ' - -> The insert should appear differently. Find good looking brightness level. Not too bright as the light bulb is still missing. Just generic good light.
' '	 - If you cannot find proper light color or "mood", you can also fiddle with primitive material values. Provided material should be
' '	   quite ok for most of the cases.
' '	 - Now when you have found proper DL value (165), but that into script:
' ' - That one insert is now adjusted and you should be able to copy&paste rest of the triangle inserts in place and name them correctly. And add them
' '   into script. And fine tune their brightness and color.
' '
' ' Light bulbs and ball reflection:
' '
' ' - This kind of lighted primitives are not giving you ball reflections. Also some more glow vould be needed to make the insert to bloom correctly.
' ' - Take the original lamp (l57), set the bulb mode enabled, set Halo Height to -3 (something that is inside the 2 insert primitives). I'd start with
' '   falloff 100, falloff Power 2-2.5, Intensity 10, scale mesh 10, Transmit 5.
' ' - Start the game with F6, throw a ball on it and move the ball near the blinking insert. Visually see how the reflection looks.
' ' - Hit D once the reflection is the highest. Open light editor and start fine tuning the bulb values to achieve realistic look for the reflection.
' ' - Falloff Power value is the one that will affect reflection creatly. The higher the power value is, the brighter the reflection on the ball is.
' '   This is the reason why falloff is rather large and falloff power is quite low. Change scale mesh if reflection is too small or large.
' ' - Transmit value can bring nice bloom for the insert, but it may also affect to other primitives nearby. Sometimes one need to set transmit to
' '   zero to avoid affecting surrounding plastics. If you really need to have higher transmit value, you may set Disable Lighting From Below to 1
' '   in surrounding primitive. This may remove the problem, but can make the primitive look worse too.

' ' ---- FOR EXAMPLE PUROSE ONLY. DELETE BELOW WHEN USING FOR REAL -----
' ' NOTE: The below timer is for flashing the inserts as a demonstration. Should be replaced by actual lamp states.
' '	   In other words, delete this sub (InsertFlicker_timer) and associated timer if you are going to use with a ROM.
' Dim flickerX, FlickerState
' FlickerState = 0
' Sub InsertFlicker_timer
	' If FlickerState = 0 Then
		' For flickerX = 0 To 19
			' If flickerX < 6 Or flickerX > 9 Then
				' AllLamps(flickerX).state = False
			' ElseIf BonusX(flickerX - 6) = False Then
				' AllLamps(flickerX).state = False
			' End If
		' Next
		' FlickerState = 1
	' Else
		' For flickerX = 0 To 19
			' If flickerX < 6 Or flickerX > 9 Then
				' AllLamps(flickerX).state = True
			' ElseIf BonusX(flickerX - 6) = False Then
				' AllLamps(flickerX).state = True
			' End If
		' Next
		' FlickerState = 0
	' End If
' End Sub



' Sub l1_animate: p1.BlendDisableLighting = 200 * (l1.GetInPlayIntensity / l1.Intensity): End Sub
' Sub l2_animate: p2.BlendDisableLighting = 200 * (l2.GetInPlayIntensity / l2.Intensity): End Sub
' Sub l3_animate: p3.BlendDisableLighting = 200 * (l3.GetInPlayIntensity / l3.Intensity): End Sub
' ' Sub l4_animate: p4.BlendDisableLighting = 200 * (l4.GetInPlayIntensity / l4.Intensity): End Sub
' ' Sub l5_animate: p5.BlendDisableLighting = 200 * (l5.GetInPlayIntensity / l5.Intensity): End Sub
' Sub l6_animate: p6.BlendDisableLighting = 200 * (l6.GetInPlayIntensity / l6.Intensity): End Sub
' Sub l7_animate: p7.BlendDisableLighting = 200 * (l7.GetInPlayIntensity / l7.Intensity): End Sub
' Sub l8_animate: p8.BlendDisableLighting = 200 * (l8.GetInPlayIntensity / l8.Intensity): End Sub
' Sub l9_animate: p9.BlendDisableLighting = 200 * (l9.GetInPlayIntensity / l9.Intensity): End Sub

' ' Sub l10_animate: p10.BlendDisableLighting = 200 * (l10.GetInPlayIntensity / l10.Intensity): End Sub
' Sub l11_animate: p11.BlendDisableLighting = 200 * (l11.GetInPlayIntensity / l11.Intensity): End Sub
' Sub l12_animate: p12.BlendDisableLighting = 200 * (l12.GetInPlayIntensity / l12.Intensity): End Sub
' Sub l13_animate: p13.BlendDisableLighting = 200 * (l13.GetInPlayIntensity / l13.Intensity): End Sub
' Sub l14_animate: p14.BlendDisableLighting = 200 * (l14.GetInPlayIntensity / l14.Intensity): End Sub
' Sub l15_animate: p15.BlendDisableLighting = 200 * (l15.GetInPlayIntensity / l15.Intensity): End Sub
' Sub l16_animate: p16.BlendDisableLighting = 200 * (l16.GetInPlayIntensity / l16.Intensity): End Sub
' Sub l17_animate: p17.BlendDisableLighting = 200 * (l17.GetInPlayIntensity / l17.Intensity): End Sub
' Sub l18_animate: p18.BlendDisableLighting = 200 * (l18.GetInPlayIntensity / l18.Intensity): End Sub
' Sub l19_animate: p19.BlendDisableLighting = 200 * (l19.GetInPlayIntensity / l19.Intensity): End Sub

' Sub l20_animate: p20.BlendDisableLighting = 200 * (l20.GetInPlayIntensity / l20.Intensity): End Sub
' Sub l21_animate: p21.BlendDisableLighting = 200 * (l21.GetInPlayIntensity / l21.Intensity): End Sub
' Sub l22_animate: p22.BlendDisableLighting = 200 * (l22.GetInPlayIntensity / l22.Intensity): End Sub
' Sub l23_animate: p23.BlendDisableLighting = 200 * (l23.GetInPlayIntensity / l23.Intensity): End Sub

' '******************************************************
' '*****   END 3D INSERTS
' '******************************************************

'******************************************************
'  ZANI: Misc Animations
'******************************************************

Sub LeftFlipper_Animate
	dim a: a = LeftFlipper.CurrentAngle
	FlipperLSh.RotZ = a
	' LFLogo.RotZ = a
	'Add any left flipper related animations here
End Sub

Sub RightFlipper_Animate
	dim a: a = RightFlipper.CurrentAngle
	FlipperRSh.RotZ = a
	' RFlogo.RotZ = a
	'Add any right flipper related animations here
End Sub

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
Dim objBallShadow(5)

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
	
	If Not IsArray(gBOT) Then Exit Sub
	
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

'*******************************************
'	ZBMP: Bumpers
'*******************************************

Sub Bumper1_Hit
'	Addscore 250
	RandomSoundBumperMiddle Bumper1
'	FlBumperFadeTarget(1) = 1   'Flupper bumper demo
'	Bumper1.timerenabled = True
End Sub

' Sub Bumper1_Timer
	' FlBumperFadeTarget(1) = 0
' End Sub

Sub Bumper2_Hit
	RandomSoundBumperMiddle Bumper2
End Sub

Sub Bumper3_Hit
	RandomSoundBumperMiddle Bumper3
End Sub

Sub BumperTL_Hit
	RandomSoundBumperTop BumperTL
End Sub

Sub BumperTR_Hit
	RandomSoundBumperTop BumperTR
End Sub

Sub BumperBL_Hit
	RandomSoundBumperBottom BumperBL
End Sub

Sub BumperBR_Hit
	RandomSoundBumperBottom BumperBR
End Sub
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

'*******************************************
'	ZDRN: Drain, Trough, and Ball Release
'*******************************************

' Original code retained for reference - largely replaced by GLF trough management

' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this,
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to
'	 the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 300 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


'TROUGH
' Sub swTrough1_Hit
	' UpdateTrough
' End Sub
' Sub swTrough1_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough2_Hit
	' UpdateTrough
' End Sub
' Sub swTrough2_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough3_Hit
	' UpdateTrough
' End Sub
' Sub swTrough3_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough4_Hit
	' UpdateTrough
' End Sub
' Sub swTrough4_UnHit
	' UpdateTrough
' End Sub
' Sub swTrough5_Hit
	' UpdateTrough
' End Sub
' Sub swTrough5_UnHit
	' UpdateTrough
' End Sub

Sub TroughTeleport_Hit
	Dim b : Set b = GetBallNear(TroughTeleport)
	If Not b Is Nothing Then 
		TroughTeleport.kick 90, 10
		b.X = LaneKicker.X
		b.Y = LaneKicker.Y
		KickBall b, -90, 2, 0, 0
	End If
End Sub

Function GetBallNear(obj)
	Dim b
	For b = 0 To UBound(gBOT)
		If abs(obj.X - gBOT(b).X) < 25 And abs(obj.Y - gBOT(b).Y) < 25 Then
			Set GetBallNear = gBOT(b)
		End If
	Next
End Function

' Sub UpdateTrough
	' UpdateTroughTimer.Interval = 300
	' UpdateTroughTimer.Enabled = 1
' End Sub

' Sub UpdateTroughTimer_Timer
	' If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	' If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	' If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	' If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	' Me.Enabled = 0
' End Sub

' ' DRAIN & RELEASE
' Sub Drain_Hit
	' BIP = BIP - 1
	' DMDBigText "DRAIN BLOCK",77,1
	' RandomSoundDrain Drain
	' vpmTimer.AddTimer 500, "Drain.kick 57, 20'"
' End Sub

' Sub SolRelease(enabled)
	' If enabled Then
		' BIP = BIP + 1
		' swTrough1.kick 90, 10
		' RandomSoundBallRelease swTrough1
	' End If
' End Sub

' '******************************************************
' ' 	ZRDT:  DROP TARGETS by Rothbauerw
' '******************************************************
' ' No Drop targets on this table.
' ' The Stand Up and Drop Target solutions improve the physics for targets to create more realistic behavior. It allows the ball
' ' to move through the target enabling the ability to score more than one target with a well placed shot.
' ' It also handles full target animation, switch handling and deflection on hit. For drop targets there is also a slight lift when
' ' the drop targets raise, bricking, and popping the ball up if it's over the drop target when it raises.
' '
' ' Add a Timers named DTAnim and STAnim to editor to handle drop & standup target animations, or run them off an always-on 10ms timer (GameTimer)
' ' DTAnim.interval = 10
' ' DTAnim.enabled = True

' ' Sub DTAnim_Timer
' ' 	DoDTAnim
' '	DoSTAnim
' ' End Sub

' ' For each drop target, we'll use two wall objects for physics calculations and one primitive for visuals and
' ' animation. We will not use target objects.  Place your drop target primitive the same as you would a VP drop target.
' ' The primitive should have it's pivot point centered on the x and y axis and at or just below the playfield
' ' level on the z axis. Orientation needs to be set using Rotz and bending deflection using Rotx. You'll find a hooded
' ' target mesh in this table's example. It uses the same texture map as the VP drop targets.
' '
' ' For each stand up target we'll use a vp target, a laid back collidable primitive, and one primitive for visuals and animation.
' ' The visual primitive should should have it's pivot point centered on the x and y axis and the z should be at or just below the playfield.
' ' The target should animate backwards using transy.
' '
' ' To create visual target primitives that work with the stand up and drop target code, follow the below instructions:
' ' (Other methods will work as well, but this is easy for even non-blender users to do)
' ' 1) Open a new blank table. Delete everything off the table in editor.
' ' 2) Copy and paste the VP target from your table into this blank table.
' ' 3) Place the target at x = 0, y = 0  (upper left hand corner) with an orientation of 0 (target facing the front of the table)
' ' 4) Under the file menu, select Export "OBJ Mesh"
' ' 5) Go to "https://threejs.org/editor/". Here you can modify the exported obj file. When you export, it exports your target and also 
' '    the playfield mesh. You need to delete the playfield mesh here. Under the file menu, chose import, and select the obj you exported
' '    from VPX. In the right hand panel, find the Playfield object and click on it and delete. Then use the file menu to Export OBJ.
' ' 6) In VPX, you can add a primitive and use "Import Mesh" to import the exported obj from the previous step. X,Y,Z scale should be 1.
' '    The primitive will use the same target texture as the VP target object. 
' '
' ' * Note, each target must have a unique switch number. If they share a same number, add 100 to additional target with that number.
' ' For example, three targets with switch 32 would use 32, 132, 232 for their switch numbers.
' ' The 100 and 200 will be removed when setting the switch value for the target.

' '******************************************************
' '  DROP TARGETS INITIALIZATION
' '******************************************************

' Class DropTarget
  ' Private m_primary, m_secondary, m_prim, m_sw, m_animate, m_isDropped

  ' Public Property Get Primary(): Set Primary = m_primary: End Property
  ' Public Property Let Primary(input): Set m_primary = input: End Property

  ' Public Property Get Secondary(): Set Secondary = m_secondary: End Property
  ' Public Property Let Secondary(input): Set m_secondary = input: End Property

  ' Public Property Get Prim(): Set Prim = m_prim: End Property
  ' Public Property Let Prim(input): Set m_prim = input: End Property

  ' Public Property Get Sw(): Sw = m_sw: End Property
  ' Public Property Let Sw(input): m_sw = input: End Property

  ' Public Property Get Animate(): Animate = m_animate: End Property
  ' Public Property Let Animate(input): m_animate = input: End Property

  ' Public Property Get IsDropped(): IsDropped = m_isDropped: End Property
  ' Public Property Let IsDropped(input): m_isDropped = input: End Property

  ' Public default Function init(primary, secondary, prim, sw, animate, isDropped)
    ' Set m_primary = primary
    ' Set m_secondary = secondary
    ' Set m_prim = prim
    ' m_sw = sw
    ' m_animate = animate
    ' m_isDropped = isDropped

    ' Set Init = Me
  ' End Function
' End Class

' 'Define a variable for each drop target
' Dim DT1, DT2, DT3

' 'Set array with drop target objects
' '
' 'DropTargetvar = Array(primary, secondary, prim, swtich, animate)
' '   primary:	primary target wall to determine drop
' '   secondary:  wall used to simulate the ball striking a bent or offset target after the initial Hit
' '   prim:	   primitive target used for visuals and animation
' '				   IMPORTANT!!!
' '				   rotz must be used for orientation
' '				   rotx to bend the target back
' '				   transz to move it up and down
' '				   the pivot point should be in the center of the target on the x, y and at or below the playfield (0) on z
' '   switch:	 ROM switch number
' '   animate:	Array slot for handling the animation instrucitons, set to 0
' '				   Values for animate: 1 - bend target (hit to primary), 2 - drop target (hit to secondary), 3 - brick target (high velocity hit to secondary), -1 - raise target
' '   isDropped:  Boolean which determines whether a drop target is dropped. Set to false if they are initially raised, true if initially dropped.
' '					Use the function DTDropped(switchid) to check a target's drop status.

' Set DT1 = (new DropTarget)(sw1, sw1a, sw1p, 1, 0, False)
' Set DT2 = (new DropTarget)(sw2, sw2a, sw2p, 2, 0, False)
' Set DT3 = (new DropTarget)(sw3, sw3a, sw3p, 3, 0, False)

' Dim DTArray
' DTArray = Array(DT1, DT2, DT3)

' 'Configure the behavior of Drop Targets.
' Const DTDropSpeed = 90 'in milliseconds
' Const DTDropUpSpeed = 40 'in milliseconds
' Const DTDropUnits = 44 'VP units primitive drops so top of at or below the playfield
' Const DTDropUpUnits = 10 'VP units primitive raises above the up position on drops up
' Const DTMaxBend = 8 'max degrees primitive rotates when hit
' Const DTDropDelay = 20 'time in milliseconds before target drops (due to friction/impact of the ball)
' Const DTRaiseDelay = 40 'time in milliseconds before target drops back to normal up position after the solenoid fires to raise the target
' Const DTBrickVel = 30 'velocity at which the target will brick, set to '0' to disable brick
' Const DTEnableBrick = 0 'Set to 0 to disable bricking, 1 to enable bricking
' Const DTMass = 0.2 'Mass of the Drop Target (between 0 and 1), higher values provide more resistance

' '******************************************************
' '  DROP TARGETS FUNCTIONS
' '******************************************************

' Sub DTHit(switch)
	' Dim i
	' i = DTArrayID(switch)
	
	' PlayTargetSound
	' DTArray(i).animate = DTCheckBrick(ActiveBall,DTArray(i).prim)
	' If DTArray(i).animate = 1 Or DTArray(i).animate = 3 Or DTArray(i).animate = 4 Then
		' DTBallPhysics ActiveBall, DTArray(i).prim.rotz, DTMass
	' End If
	' DoDTAnim
' End Sub

' Sub DTRaise(switch)
	' Dim i
	' i = DTArrayID(switch)
	
	' DTArray(i).animate =  - 1
	' DoDTAnim
' End Sub

' Sub DTDrop(switch)
	' Dim i
	' i = DTArrayID(switch)
	
	' DTArray(i).animate = 1
	' DoDTAnim
' End Sub

' Function DTArrayID(switch)
	' Dim i
	' For i = 0 To UBound(DTArray)
		' If DTArray(i).sw = switch Then
			' DTArrayID = i
			' Exit Function
		' End If
	' Next
' End Function

' Sub DTBallPhysics(aBall, angle, mass)
	' Dim rangle,bangle,calc1, calc2, calc3
	' rangle = (angle - 90) * 3.1416 / 180
	' bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	
	' calc1 = cor.BallVel(aball.id) * Cos(bangle - rangle) * (aball.mass - mass) / (aball.mass + mass)
	' calc2 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Cos(rangle + 4 * Atn(1) / 2)
	' calc3 = cor.BallVel(aball.id) * Sin(bangle - rangle) * Sin(rangle + 4 * Atn(1) / 2)
	
	' aBall.velx = calc1 * Cos(rangle) + calc2
	' aBall.vely = calc1 * Sin(rangle) + calc3
' End Sub

' 'Check if target is hit on it's face or sides and whether a 'brick' occurred
' Function DTCheckBrick(aBall, dtprim)
	' Dim bangle, bangleafter, rangle, rangle2, Xintersect, Yintersect, cdist, perpvel, perpvelafter, paravel, paravelafter
	' rangle = (dtprim.rotz - 90) * 3.1416 / 180
	' rangle2 = dtprim.rotz * 3.1416 / 180
	' bangle = atn2(cor.ballvely(aball.id),cor.ballvelx(aball.id))
	' bangleafter = Atn2(aBall.vely,aball.velx)
	
	' Xintersect = (aBall.y - dtprim.y - Tan(bangle) * aball.x + Tan(rangle2) * dtprim.x) / (Tan(rangle2) - Tan(bangle))
	' Yintersect = Tan(rangle2) * Xintersect + (dtprim.y - Tan(rangle2) * dtprim.x)
	
	' cdist = Distance(dtprim.x, dtprim.y, Xintersect, Yintersect)
	
	' perpvel = cor.BallVel(aball.id) * Cos(bangle - rangle)
	' paravel = cor.BallVel(aball.id) * Sin(bangle - rangle)
	
	' perpvelafter = BallSpeed(aBall) * Cos(bangleafter - rangle)
	' paravelafter = BallSpeed(aBall) * Sin(bangleafter - rangle)
	
	' If perpvel > 0 And  perpvelafter <= 0 Then
		' If DTEnableBrick = 1 And  perpvel > DTBrickVel And DTBrickVel <> 0 And cdist < 8 Then
			' DTCheckBrick = 3
		' Else
			' DTCheckBrick = 1
		' End If
	' ElseIf perpvel > 0 And ((paravel > 0 And paravelafter > 0) Or (paravel < 0 And paravelafter < 0)) Then
		' DTCheckBrick = 4
	' Else
		' DTCheckBrick = 0
	' End If
' End Function

' Sub DoDTAnim()
	' Dim i
	' For i = 0 To UBound(DTArray)
		' DTArray(i).animate = DTAnimate(DTArray(i).primary,DTArray(i).secondary,DTArray(i).prim,DTArray(i).sw,DTArray(i).animate)
	' Next
' End Sub

' Function DTAnimate(primary, secondary, prim, switch, animate)
	' Dim transz, switchid
	' Dim animtime, rangle
	
	' switchid = switch
	
	' Dim ind
	' ind = DTArrayID(switchid)
	
	' rangle = prim.rotz * PI / 180
	
	' DTAnimate = animate
	
	' If animate = 0 Then
		' primary.uservalue = 0
		' DTAnimate = 0
		' Exit Function
	' ElseIf primary.uservalue = 0 Then
		' primary.uservalue = GameTime
	' End If
	
	' animtime = GameTime - primary.uservalue
	
	' If (animate = 1 Or animate = 4) And animtime < DTDropDelay Then
		' primary.collidable = 0
		' If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0
		' prim.rotx = DTMaxBend * Cos(rangle)
		' prim.roty = DTMaxBend * Sin(rangle)
		' DTAnimate = animate
		' Exit Function
	' ElseIf (animate = 1 Or animate = 4) And animtime > DTDropDelay Then
		' primary.collidable = 0
		' If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 1 'If animate = 1 Then secondary.collidable = 1 Else secondary.collidable = 0 'updated by rothbauerw to account for edge case
		' prim.rotx = DTMaxBend * Cos(rangle)
		' prim.roty = DTMaxBend * Sin(rangle)
		' animate = 2
		' SoundDropTargetDrop prim
	' End If
	
	' If animate = 2 Then
		' transz = (animtime - DTDropDelay) / DTDropSpeed * DTDropUnits *  - 1
		' If prim.transz >  - DTDropUnits  Then
			' prim.transz = transz
		' End If
		
		' prim.rotx = DTMaxBend * Cos(rangle) / 2
		' prim.roty = DTMaxBend * Sin(rangle) / 2
		
		' If prim.transz <= - DTDropUnits Then
			' prim.transz =  - DTDropUnits
			' secondary.collidable = 0
			' DTArray(ind).isDropped = True 'Mark target as dropped
			' If UsingROM Then
				' controller.Switch(Switchid mod 100) = 1
			' Else
				' DTAction switchid
			' End If
			' primary.uservalue = 0
			' DTAnimate = 0
			' Exit Function
		' Else
			' DTAnimate = 2
			' Exit Function
		' End If
	' End If
	
	' If animate = 3 And animtime < DTDropDelay Then
		' primary.collidable = 0
		' secondary.collidable = 1
		' prim.rotx = DTMaxBend * Cos(rangle)
		' prim.roty = DTMaxBend * Sin(rangle)
	' ElseIf animate = 3 And animtime > DTDropDelay Then
		' primary.collidable = 1
		' secondary.collidable = 0
		' prim.rotx = 0
		' prim.roty = 0
		' primary.uservalue = 0
		' DTAnimate = 0
		' Exit Function
	' End If
	
	' If animate =  - 1 Then
		' transz = (1 - (animtime) / DTDropUpSpeed) * DTDropUnits *  - 1
		
		' If prim.transz =  - DTDropUnits Then
			' Dim b
			' 'Dim gBOT
			' 'gBOT = GetBalls
			
			' For b = 0 To UBound(gBOT)
				' If InRotRect(gBOT(b).x,gBOT(b).y,prim.x, prim.y, prim.rotz, - 25, - 10,25, - 10,25,25, - 25,25) And gBOT(b).z < prim.z + DTDropUnits + 25 Then
					' gBOT(b).velz = 20
				' End If
			' Next
		' End If
		
		' If prim.transz < 0 Then
			' prim.transz = transz
		' ElseIf transz > 0 Then
			' prim.transz = transz
		' End If
		
		' If prim.transz > DTDropUpUnits Then
			' DTAnimate =  - 2
			' prim.transz = DTDropUpUnits
			' prim.rotx = 0
			' prim.roty = 0
			' primary.uservalue = GameTime
		' End If
		' primary.collidable = 0
		' secondary.collidable = 1
		' DTArray(ind).isDropped = False 'Mark target as not dropped
		' If UsingROM Then controller.Switch(Switchid mod 100) = 0
	' End If
	
	' If animate =  - 2 And animtime > DTRaiseDelay Then
		' prim.transz = (animtime - DTRaiseDelay) / DTDropSpeed * DTDropUnits *  - 1 + DTDropUpUnits
		' If prim.transz < 0 Then
			' prim.transz = 0
			' primary.uservalue = 0
			' DTAnimate = 0
			
			' primary.collidable = 1
			' secondary.collidable = 0
		' End If
	' End If
' End Function

' Function DTDropped(switchid)
	' Dim ind
	' ind = DTArrayID(switchid)
	
	' DTDropped = DTArray(ind).isDropped
' End Function

' Sub DTAction(switchid)
	' Select Case switchid
		' Case 1
			' Addscore 1000
			' ShadowDT(0).visible = False
			
		' Case 2
			' Addscore 1000
			' ShadowDT(1).visible = False
			
		' Case 3
			' Addscore 1000
			' ShadowDT(2).visible = False
	' End Select
' End Sub


' '******************************************************
' '****  END DROP TARGETS
' '******************************************************

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
' FlInitBumper 1, "red"
' FlInitBumper 2, "white"
' FlInitBumper 3, "blue"
' FlInitBumper 4, "orange"
' FlInitBumper 5, "yellow"

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
' InitFlasher 1, "green"
' InitFlasher 2, "red"
' InitFlasher 3, "blue"
' InitFlasher 4, "white"

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
'	ZPHY:  GENERAL ADVICE ON PHYSICS
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


'*******************************************
' Late 70's to early 80's

Sub InitPolarity()
   dim x, a : a = Array(LF, RF)
	for each x in a
		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
		x.enabled = True
		x.TimeDelay = 80
		x.DebugOn=False ' prints some info in debugger


        x.AddPt "Polarity", 0, 0, 0
        x.AddPt "Polarity", 1, 0.05, - 2.7
        x.AddPt "Polarity", 2, 0.16, - 2.7
        x.AddPt "Polarity", 3, 0.22, - 0
        x.AddPt "Polarity", 4, 0.25, - 0
        x.AddPt "Polarity", 5, 0.3, - 1
        x.AddPt "Polarity", 6, 0.4, - 2
        x.AddPt "Polarity", 7, 0.5, - 2.7
        x.AddPt "Polarity", 8, 0.65, - 1.8
        x.AddPt "Polarity", 9, 0.75, - 0.5
        x.AddPt "Polarity", 10, 0.81, - 0.5
        x.AddPt "Polarity", 11, 0.88, 0
        x.AddPt "Polarity", 12, 1.3, 0

		x.AddPt "Velocity", 0, 0, 0.85
		x.AddPt "Velocity", 1, 0.15, 0.85
		x.AddPt "Velocity", 2, 0.2, 0.9
		x.AddPt "Velocity", 3, 0.23, 0.95
		x.AddPt "Velocity", 4, 0.41, 0.95
		x.AddPt "Velocity", 5, 0.53, 0.95 '0.982
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

'Sub InitPolarity()
'	Dim x, a
'	a = Array(LF, RF)
'	For Each x In a
'		x.AddPt "Ycoef", 0, RightFlipper.Y-65, 1 'disabled
'		x.AddPt "Ycoef", 1, RightFlipper.Y-11, 1
'		x.enabled = True
'		x.TimeDelay = 60
'		x.DebugOn=False ' prints some info in debugger
'
'		x.AddPt "Polarity", 0, 0, 0
'		x.AddPt "Polarity", 1, 0.05, - 5.5
'		x.AddPt "Polarity", 2, 0.16, - 5.5
'		x.AddPt "Polarity", 3, 0.20, - 0.75
'		x.AddPt "Polarity", 4, 0.25, - 1.25
'		x.AddPt "Polarity", 5, 0.3, - 1.75
'		x.AddPt "Polarity", 6, 0.4, - 3.5
'		x.AddPt "Polarity", 7, 0.5, - 5.25
'		x.AddPt "Polarity", 8, 0.7, - 4.0
'		x.AddPt "Polarity", 9, 0.75, - 3.5
'		x.AddPt "Polarity", 10, 0.8, - 3.0
'		x.AddPt "Polarity", 11, 0.85, - 2.5
'		x.AddPt "Polarity", 12, 0.9, - 2.0
'		x.AddPt "Polarity", 13, 0.95, - 1.5
'		x.AddPt "Polarity", 14, 1, - 1.0
'		x.AddPt "Polarity", 15, 1.05, -0.5
'		x.AddPt "Polarity", 16, 1.1, 0
'		x.AddPt "Polarity", 17, 1.3, 0
'
'		x.AddPt "Velocity", 0, 0, 0.85
'		x.AddPt "Velocity", 1, 0.23, 0.85
'		x.AddPt "Velocity", 2, 0.27, 1
'		x.AddPt "Velocity", 3, 0.3, 1
'		x.AddPt "Velocity", 4, 0.35, 1
'		x.AddPt "Velocity", 5, 0.6, 1 '0.982
'		x.AddPt "Velocity", 6, 0.62, 1.0
'		x.AddPt "Velocity", 7, 0.702, 0.968
'		x.AddPt "Velocity", 8, 0.95,  0.968
'		x.AddPt "Velocity", 9, 1.03,  0.945
'		x.AddPt "Velocity", 10, 1.5,  0.945
'
'	Next
'	
'	' SetObjects arguments: 1: name of object 2: flipper object: 3: Trigger object around flipper
'	LF.SetObjects "LF", LeftFlipper, TriggerLF
'	RF.SetObjects "RF", RightFlipper, TriggerRF
'End Sub

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
Const EOSTnew = 1.5 'EM's to late 80's - new recommendation by rothbauerw (previously 1)
'Const EOSTnew = 1.2 '90's and later - new recommendation by rothbauerw (previously 0.8)
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
Const EOSReturn = 0.055  'EM's
'   Const EOSReturn = 0.045  'late 70's to mid 80's
'   Const EOSReturn = 0.035  'mid 80's to early 90's
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

'*******************************************
'	ZSCR: Scoring
'*******************************************

Sub Addscore (value)
	PlayerScore(CurrentPlayer) = PlayerScore(CurrentPlayer) + value
	If value > 99999 Then DMDBGFlash = 15
	If value > 499999 Then DMDFire = Flexframe + 50
	
	' Add chimes based on score amount here
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

' InitSlingCorrection

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

'****************************************************************
'	ZSLG: Slingshots
'****************************************************************

' RStep and LStep are the variables that increment the animation
Dim RStep, LStep

Sub RightSlingShot_Slingshot(args)
	Dim enabled, ball : enabled = args(0)
	If enabled then
		If Not IsNull(args(1)) Then
			RS.VelocityCorrect(args(1))
		End If
		Addscore 10000
		RSling1.Visible = 1
		Sling1.TransY =  - 20   'Sling Metal Bracket
		RStep = 0
		RightSlingShot_Timer
		RightSlingShot.TimerEnabled = 1
		RightSlingShot.TimerInterval = 17
		RandomSoundSlingshotRight Sling1
		'DOF 104, DOFPulse
	End If
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

Sub LeftSlingShot_Slingshot(args)
	Dim enabled, ball : enabled = args(0)
	If enabled then
		If Not IsNull(args(1)) Then
			LS.VelocityCorrect(args(1))
		End If
		Addscore 10000
		LSling1.Visible = 1
		Sling2.TransY =  - 20   'Sling Metal Bracket		
		LStep = 0
		LeftSlingShot_Timer
		LeftSlingShot.TimerEnabled = 1
		LeftSlingShot.TimerInterval = 17
		RandomSoundSlingshotLeft Sling2
		'DOF 103, DOFPulse
	End If
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
' Dim ST11, ST12, ST13

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


' Set ST11 = (new StandupTarget)(sw11, psw11,11, 0)
' Set ST12 = (new StandupTarget)(sw12, psw12,12, 0)
' Set ST13 = (new StandupTarget)(sw13, psw13,13, 0)

'Add all the Stand-up Target Arrays to Stand-up Target Animation Array
'   STAnimationArray = Array(ST1, ST2, ....)
Dim STArray
' STArray = Array(ST11, ST12, ST13)
STArray = Array()

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
'	DoDTAnim				'handle drop target animations
	DoSTAnim				'handle stand up target animations
	queue.Tick	      		'handle the queue system
	BSUpdate
End Sub

'The CorTimer interval should be 10. It's sole purpose is to update the Cor calculations
CorTimer.Interval = 10
Sub CorTimer_Timer(): Cor.Update: End Sub

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

'*******************************************
'  ZOPT: User Options
'*******************************************


'----- DMD Options -----
Const UseFlexDMD = 0				'0 = no FlexDMD, 1 = enable FlexDMD
Const FlexONPlayfield = False	   'False = off, True=DMD on playfield ( vrroom overrides this )

'----- VR Room -----
Const VRRoomChoice = 0			  ' 1 - Minimal Room, 2 - Ultra Minimal Room

Dim LightLevel : LightLevel = 0.25				' Level of room lighting (0 to 1), where 0 is dark and 100 is brightest
Dim ColorLUT : ColorLUT = 1						' Color desaturation LUTs: 1 to 11, where 1 is normal and 11 is black'n'white
Dim VolumeDial : VolumeDial = 0.8				' Overall Mechanical sound effect volume. Recommended values should be no greater than 1.
Dim BallRollVolume : BallRollVolume = 0.5		' Level of ball rolling volume. Value between 0 and 1
Dim RampRollVolume : RampRollVolume = 0.5		' Level of ramp rolling volume. Value between 0 and 1
Dim StagedFlippers : StagedFlippers = 0			' Staged Flippers. 0 = Disabled, 1 = Enabled
Dim RotoDofMode : RotoDofMode = 0				' Shaker/Gear motor DOF preference. 1 = Both, 2 = Gear Only, 3 = Shaker Only, 4 = None
Dim TiltWarnings : TiltWarnings = 2				' Number of tilt warnings before full tilt.


' Called when options are tweaked by the player. 
' - 0: game has started, good time to load options and adjust accordingly
' - 1: an option has changed
' - 2: options have been reseted
' - 3: player closed the tweak UI, good time to update staticly prerendered parts
' Table1.Option arguments are: 
' - option name, minimum value, maximum value, step between valid values, default value, unit (0=None, 1=Percent), an optional arry of literal strings
Dim dspTriggered : dspTriggered = False
Sub Table1_OptionEvent(ByVal eventId)
	If eventId = 1 And Not dspTriggered Then dspTriggered = True : DisableStaticPreRendering = True : End If

	' DOF
	RotoDofMode = Table1.Option("Roto DOF Mode", 1, 4, 1, 0, 0, Array("Both", "Gear Motor Only", "Shaker Motor Only", "None"))
	
	' Difficulty
	TiltWarnings = Table1.Option("Tilt Warnings", 0, 5, 1, 2, 0)
	
	' Color Saturation
    ColorLUT = Table1.Option("Color Saturation", 1, 11, 1, 1, 0, _
		Array("Normal", "Desaturated 10%", "Desaturated 20%", "Desaturated 30%", "Desaturated 40%", "Desaturated 50%", _
        "Desaturated 60%", "Desaturated 70%", "Desaturated 80%", "Desaturated 90%", "Black 'n White"))
	if ColorLUT = 1 Then Table1.ColorGradeImage = ""
	if ColorLUT = 2 Then Table1.ColorGradeImage = "colorgradelut256x16-10"
	if ColorLUT = 3 Then Table1.ColorGradeImage = "colorgradelut256x16-20"
	if ColorLUT = 4 Then Table1.ColorGradeImage = "colorgradelut256x16-30"
	if ColorLUT = 5 Then Table1.ColorGradeImage = "colorgradelut256x16-40"
	if ColorLUT = 6 Then Table1.ColorGradeImage = "colorgradelut256x16-50"
	if ColorLUT = 7 Then Table1.ColorGradeImage = "colorgradelut256x16-60"
	if ColorLUT = 8 Then Table1.ColorGradeImage = "colorgradelut256x16-70"
	if ColorLUT = 9 Then Table1.ColorGradeImage = "colorgradelut256x16-80"
	if ColorLUT = 10 Then Table1.ColorGradeImage = "colorgradelut256x16-90"
	if ColorLUT = 11 Then Table1.ColorGradeImage = "colorgradelut256x16-100"

    ' Sound volumes
    VolumeDial = Table1.Option("Mech Volume", 0, 1, 0.01, 0.8, 1)
    BallRollVolume = Table1.Option("Ball Roll Volume", 0, 1, 0.01, 0.5, 1)
	RampRollVolume = Table1.Option("Ramp Roll Volume", 0, 1, 0.01, 0.5, 1)

	' Room brightness
'	LightLevel = Table1.Option("Table Brightness (Ambient Light Level)", 0, 1, 0.01, .5, 1)
	LightLevel = NightDay/100
'	SetRoomBrightness LightLevel   'Uncomment this line for lightmapped tables.

    ' Staged Flippers
    StagedFlippers = Table1.Option("Staged Flippers", 0, 1, 1, 0, 0, Array("Disabled", "Enabled"))

	If eventId = 3 And dspTriggered Then dspTriggered = False : DisableStaticPreRendering = False : End If

	Glf_Options(eventId)
End Sub

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

'*******************************************
' 	ZVRR: VR Room / VR Cabinet
'*******************************************

' Disabled as non-functional without unknown "VR_Cab" object

' Dim VRThings
' If VRRoom <> 0 Then
	' '	DMDPlayfield.dmd=False
	' DMDPlayfield.visible = False
	' '	DMDbackbox.DMD=True
	' DMDbackbox.visible = True  ' flasher dmd
	' If VRRoom = 1 Then
		' For Each VRThings In VR_Cab
			' VRThings.visible = 1
		' Next
	' End If
	' If VRRoom = 2 Then
		' For Each VRThings In VR_Cab
			' VRThings.visible = 0
		' Next
		' PinCab_Backglass.visible = 1
	' End If
' Else
	' If FlexONPlayfield Then
		' '   DMDPlayfield.dmd=True
		' DMDPlayfield.visible = True
	' End If
	' For Each VRThings In VR_Cab
		' VRThings.visible = 0
	' Next
' End If


'******************************************************
'	ZGCF:  GLF Configurations
'******************************************************

Sub ConfigureGlfDevices

    
    ' Plunger
    With CreateGlfBallDevice("plunger")
        .BallSwitches = Array("s_Plunger")
        .EjectTimeout = 200
        .MechanicalEject = True
        .DefaultDevice = True
        .EjectCallback = "PlungerEjectCallback"
    End With
	
	' Flippers
    With CreateGlfFlipper("left")
        .Switch = "s_left_flipper"
        .ActionCallback = "LeftFlipperAction"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started", "enable_flippers")
    End With

    With CreateGlfFlipper("right")
        .Switch = "s_right_flipper"
        .ActionCallback = "RightFlipperAction"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started", "enable_flippers")
    End With
	
	' Slingshots
    With CreateGlfAutoFireDevice("left_sling")
        .Switch = "LeftSlingShot"
        .ActionCallback = "LeftSlingshotAction"
        .DisabledCallback = "LeftSlingshotDisabled"
        .EnabledCallback = "LeftSlingshotEnabled"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started","enable_flippers")
    End With

    With CreateGlfAutoFireDevice("right_sling")
        .Switch = "RightSlingShot"
        .ActionCallback = "RightSlingshotAction"
        .DisabledCallback = "RightSlingshotDisabled"
        .EnabledCallback = "RightSlingshotEnabled"
        .DisableEvents = Array("kill_flippers")
        .EnableEvents = Array("ball_started","enable_flippers")
    End With

    CreateBaseMode()
	CreateGIMode()
	CreateScoreMode()

End Sub

Sub LeftFlipperAction(Enabled)
	If Enabled Then
		LeftFlipper.RotateToEnd
	Else
		LeftFlipper.RotateToStart
	End If
End Sub

Sub RightFlipperAction(Enabled)
	If Enabled Then
		RightFlipper.RotateToEnd
	Else
		RightFlipper.RotateToStart
	End If
End Sub



Sub CreateBaseMode()

    With CreateGlfMode("base", 100)
        .StartEvents = Array("ball_started")
        .StopEvents = Array("ball_ended")   
    End With

End Sub

Public Sub CreateGIMode()

    With CreateGlfMode("gi_control", 1000)
        .StartEvents = Array("game_started")
        .StopEvents = Array("game_ended") 
        With .LightPlayer()
            With .EventName("mode_gi_control_started")
                With .Lights("GI")
                    .Color = "ffffff"
                End With
            End With
        End With
    End With

End Sub

Public Sub CreateScoreMode()

    With CreateGlfMode("score", 2000)
        .StartEvents = Array("game_started")
        .StopEvents = Array("game_ended")

        With .VariablePlayer()
            With .EventName("score_10k") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "10000"
                End With
            End With
            With .EventName("score_100k") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "100000"
                End With
            End With
            With .EventName("score_500k") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "500000"
                End With
            End With
            With .EventName("score_1m") 
                With .Variable("score")
                    .Action = "add"
                    .Int = "1000000"
                End With
            End With

		End With

        
    End With

End Sub

Public Sub CreateLitModes()

	' Left bumper and slignshots lit for 100k
	With CreateGlfMode("left_lit", 500)
		.StartEvents = Array("light_left")
		.StopEvents = Array("light_right")

		With .EventPlayer()
			.Add "Bumper1_hit", Array("score_100k")
			.Add "LeftSlingShot_hit", Array("score_100k")
			.Add "Bumper3_hit", Array("score_10k")
			.Add "RightSlingShot_hit", Array("score_10k")
			.Add "RubberSWx_hit", Array("light_right") ' TODO: hook up individual rubber switches 
		End With
		
		With .LightPlayer()
			With .EventName("light_left")
				With .Lights("GI_LEFT_BUMP")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_left")
				With .Lights("GI_LEFT_SLING")
					.Color = "ffffff"
				End With
			End With
			
			' Turn out lights for other side
			With .EventName("light_left")
				With .Lights("GI_RIGHT_BUMP")
					.Color = "000000"
				End With
			End With
			With .EventName("light_left")
				With .Lights("GI_RIGHT_SLING")
					.Color = "000000"
				End With
			End With
		End With
        
	End With
	
	' Right bumper and slignshots lit for 100k
	With CreateGlfMode("right_lit", 500)
		.StartEvents = Array("light_right")
		.StopEvents = Array("light_left")

		With .EventPlayer()
			.Add "Bumper1_hit", Array("score_10k")
			.Add "LeftSlingShot_hit", Array("score_10k")
			.Add "Bumper3_hit", Array("score_100k")
			.Add "RightSlingShot_hit", Array("score_100k")
			.Add "RubberSWx_hit", Array("light_left") ' TODO: hook up individual rubber switches 
		End With
		
		With .LightPlayer()
			With .EventName("light_right")
				With .Lights("GI_RIGHT_BUMP")
					.Color = "ffffff"
				End With
			End With
			With .EventName("light_right")
				With .Lights("GI_RIGHT_SLING")
					.Color = "ffffff"
				End With
			End With
			
			' Turn out lights for other side
			With .EventName("light_right")
				With .Lights("GI_LEFT_BUMP")
					.Color = "000000"
				End With
			End With
			With .EventName("light_right")
				With .Lights("GI_LEFT_SLING")
					.Color = "000000"
				End With
			End With
		End With
        
	End With

End Sub
'VPX Game Logic Framework (https://mpcarr.github.io/vpx-glf/)

'
Dim glf_currentPlayer : glf_currentPlayer = Null
Dim glf_canAddPlayers : glf_canAddPlayers = True
Dim glf_PI : glf_PI = 4 * Atn(1)
Dim glf_plunger, glf_ballsearch
glf_ballsearch = Null
Dim glf_ballsearch_enabled : glf_ballsearch_enabled = False
Dim glf_gameStarted : glf_gameStarted = False
Dim glf_gameTilted : glf_gameTilted = False
Dim glf_gameEnding : glf_gameEnding = False
Dim glf_last_switch_hit_time : glf_last_switch_hit_time = 0
Dim glf_last_ballsearch_reset_time : glf_last_ballsearch_reset_time = 0
Dim glf_last_switch_hit : glf_last_switch_hit = ""
Dim glf_current_virtual_tilt : glf_current_virtual_tilt = 0
Dim glf_tilt_sensitivity : glf_tilt_sensitivity = 7
Dim glf_flex_alphadmd : Set glf_flex_alphadmd = Nothing
Dim glf_flex_alphadmd_enabled : glf_flex_alphadmd_enabled = False
Dim glf_flex_alphadmd_segments(31)
Dim glf_pinEvents : Set glf_pinEvents = CreateObject("Scripting.Dictionary")
Dim glf_pinEventsOrder : Set glf_pinEventsOrder = CreateObject("Scripting.Dictionary")
Dim glf_playerEvents : Set glf_playerEvents = CreateObject("Scripting.Dictionary")
Dim Glf_EventBlocks : Set Glf_EventBlocks = CreateObject("Scripting.Dictionary")
Dim Glf_ShotProfiles : Set Glf_ShotProfiles = CreateObject("Scripting.Dictionary")
Dim Glf_ShowStartQueue : Set Glf_ShowStartQueue = CreateObject("Scripting.Dictionary")
Dim glf_playerEventsOrder : Set glf_playerEventsOrder = CreateObject("Scripting.Dictionary")
Dim glf_playerState : Set glf_playerState = CreateObject("Scripting.Dictionary")
Dim glf_running_shows : Set glf_running_shows = CreateObject("Scripting.Dictionary")
Dim glf_cached_shows : Set glf_cached_shows = CreateObject("Scripting.Dictionary")
Dim glf_cached_rgb_fades : Set glf_cached_rgb_fades = CreateObject("Scripting.Dictionary")
Dim glf_lightPriority : Set glf_lightPriority = CreateObject("Scripting.Dictionary")
Dim glf_lightColorLookup : Set glf_lightColorLookup = CreateObject("Scripting.Dictionary")
Dim glf_lightMaps : Set glf_lightMaps = CreateObject("Scripting.Dictionary")
Dim glf_lightStacks : Set glf_lightStacks = CreateObject("Scripting.Dictionary")
Dim glf_lightTags : Set glf_lightTags = CreateObject("Scripting.Dictionary")
Dim glf_lightNames : Set glf_lightNames = CreateObject("Scripting.Dictionary")
Dim glf_modes : Set glf_modes = CreateObject("Scripting.Dictionary")
Dim glf_timers : Set glf_timers = CreateObject("Scripting.Dictionary")

Dim glf_state_machines : Set glf_state_machines = CreateObject("Scripting.Dictionary")
Dim glf_ball_devices : Set glf_ball_devices = CreateObject("Scripting.Dictionary")
Dim glf_diverters : Set glf_diverters = CreateObject("Scripting.Dictionary")
Dim glf_flippers : Set glf_flippers = CreateObject("Scripting.Dictionary")
Dim glf_autofiredevices : Set glf_autofiredevices = CreateObject("Scripting.Dictionary")
Dim glf_ball_holds : Set glf_ball_holds = CreateObject("Scripting.Dictionary")
Dim glf_magnets : Set glf_magnets = CreateObject("Scripting.Dictionary")
Dim glf_segment_displays : Set glf_segment_displays = CreateObject("Scripting.Dictionary")
Dim glf_droptargets : Set glf_droptargets = CreateObject("Scripting.Dictionary")
Dim glf_multiball_locks : Set glf_multiball_locks = CreateObject("Scripting.Dictionary")
Dim glf_multiballs : Set glf_multiballs = CreateObject("Scripting.Dictionary")
Dim glf_shows : Set glf_shows = CreateObject("Scripting.Dictionary")
Dim glf_initialVars : Set glf_initialVars = CreateObject("Scripting.Dictionary")
Dim glf_dispatch_await : Set glf_dispatch_await = CreateObject("Scripting.Dictionary")
Dim glf_dispatch_handlers_await : Set glf_dispatch_handlers_await = CreateObject("Scripting.Dictionary")
Dim glf_dispatch_current_kwargs
Dim glf_dispatch_lightmaps_await : Set glf_dispatch_lightmaps_await = CreateObject("Scripting.Dictionary")
Dim glf_machine_vars : Set glf_machine_vars = CreateObject("Scripting.Dictionary")
Dim glf_achievements : Set glf_achievements = CreateObject("Scripting.Dictionary")
Dim glf_sound_buses : Set glf_sound_buses = CreateObject("Scripting.Dictionary")
Dim glf_sounds : Set glf_sounds = CreateObject("Scripting.Dictionary")
Dim glf_combo_switches : Set glf_combo_switches = CreateObject("Scripting.Dictionary")

Dim bcpController : bcpController = Null
Dim glf_debugBcpController : glf_debugBcpController = Null
Dim glf_monitor_player_state : glf_monitor_player_state = ""
Dim glf_monitor_modes : glf_monitor_modes = ""
Dim glf_monitor_event_stream : glf_monitor_event_stream = ""
Dim glf_running_modes : glf_running_modes = ""

Dim useGlfBCPMonitor : useGlfBCPMonitor = False
Dim useBCP : useBCP = False
Dim bcpPort : bcpPort = 5050
Dim bcpExeName : bcpExeName = ""
Dim glf_BIP : glf_BIP = 0
Dim glf_FuncCount : glf_FuncCount = 0
Dim glf_SeqCount : glf_SeqCount = 0
Dim glf_max_dispatch : glf_max_dispatch = 25
Dim glf_max_lightmap_sync : glf_max_lightmap_sync = -1
Dim glf_max_lightmap_sync_enabled : glf_max_lightmap_sync_enabled = False
Dim glf_max_lights_test : glf_max_lights_test = 0

Dim glf_master_volume : glf_master_volume = 0.8

Dim glf_ballsPerGame : glf_ballsPerGame = 3
Dim glf_troughSize : glf_troughSize = tnob
Dim glf_lastTroughSw : glf_lastTroughSw = Null

Dim glf_debugLog : Set glf_debugLog = (new GlfDebugLogFile)()
Dim glf_debugEnabled : glf_debugEnabled = False
Dim glf_debug_level : glf_debug_level = "Info"

Glf_RegisterLights()
Dim glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8	

Public Sub Glf_ConnectToBCPMediaController(args)
    Set bcpController = (new GlfVpxBcpController)(bcpPort, bcpExeName)
End Sub

Public Sub Glf_ConnectToDebugBCPMediaController(args)
    Set glf_debugBcpController = (new GlfMonitorBcpController)(5051, "glf_monitor.exe")
End Sub

Public Sub Glf_WriteDebugLog(name, message)
	If glf_debug_level = "Debug" Then
		glf_debugLog.WriteToLog name, message
		'Glf_MonitorEventStream name, message
	End If
End Sub

Public Function SwitchHandler(handler, args)
	SwitchHandler = False
	Select Case handler
		Case "BaseModeDeviceEventHandler"
			BaseModeDeviceEventHandler args
			SwitchHandler = True
	End Select

End Function

Public Sub Glf_Init()
	Glf_Options Null 'Force Options Check


	If glf_troughSize > 0 Then : swTrough1.DestroyBall : Set glf_ball1 = swTrough1.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1) : Set glf_lastTroughSw = swTrough1 : End If
	If glf_troughSize > 1 Then : swTrough2.DestroyBall : Set glf_ball2 = swTrough2.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2) : Set glf_lastTroughSw = swTrough2 : End If
	If glf_troughSize > 2 Then : swTrough3.DestroyBall : Set glf_ball3 = swTrough3.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3) : Set glf_lastTroughSw = swTrough3 : End If
	If glf_troughSize > 3 Then : swTrough4.DestroyBall : Set glf_ball4 = swTrough4.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4) : Set glf_lastTroughSw = swTrough4 : End If
	If glf_troughSize > 4 Then : swTrough5.DestroyBall : Set glf_ball5 = swTrough5.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5) : Set glf_lastTroughSw = swTrough5 : End If
	If glf_troughSize > 5 Then : swTrough6.DestroyBall : Set glf_ball6 = swTrough6.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6) : Set glf_lastTroughSw = swTrough6 : End If
	If glf_troughSize > 6 Then : swTrough7.DestroyBall : Set glf_ball7 = swTrough7.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7) : Set glf_lastTroughSw = swTrough7 : End If
	If glf_troughSize > 7 Then : Drain.DestroyBall : Set glf_ball8 = Drain.CreateSizedballWithMass(Ballsize / 2,Ballmass) : gBot = Array(glf_ball1, glf_ball2, glf_ball3, glf_ball4, glf_ball5, glf_ball6, glf_ball7, glf_ball8) : End If
	
	Dim switch, switchHitSubs
	switchHitSubs = ""
	For Each switch in Glf_Switches
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_Hit() : If Not glf_gameTilted Then : DispatchPinEvent """ & switch.Name & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& switch.Name &""": End If : End Sub" & vbCrLf
		switchHitSubs = switchHitSubs & "Sub " & switch.Name & "_UnHit() : If Not glf_gameTilted Then : DispatchPinEvent """ & switch.Name & "_inactive"", ActiveBall : End If  : End Sub" & vbCrLf
	Next
	
	ExecuteGlobal switchHitSubs

	Dim slingshot, slingshotHitSubs
	slingshotHitSubs = ""
	For Each slingshot in Glf_Slingshots
		slingshotHitSubs = slingshotHitSubs & "Sub " & slingshot.Name & "_Slingshot() : If Not glf_gameTilted Then : DispatchPinEvent """ & slingshot.Name & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& slingshot.Name &""": End If  : End Sub" & vbCrLf
	Next
	ExecuteGlobal slingshotHitSubs

	Dim spinner, spinnerHitSubs
	spinnerHitSubs = ""
	For Each spinner in Glf_Spinners
		spinnerHitSubs = spinnerHitSubs & "Sub " & spinner.Name & "_Spin() : If Not glf_gameTilted Then : DispatchPinEvent """ & spinner.Name & "_active"", ActiveBall : glf_last_switch_hit_time = gametime : glf_last_switch_hit = """& spinner.Name &""": End If  : End Sub" & vbCrLf
	Next
	ExecuteGlobal spinnerHitSubs

	If glf_debugEnabled = True Then

		' Calculate the scale factor
		Dim scaleFactor
		scaleFactor = 1080 / tableheight

		Dim light
		Dim switchNumber : switchNumber = 0
		Dim lightsNumber : lightsNumber = 0
		Dim coilsNumber : coilsNumber = 0
		Dim switchesYaml : switchesYaml = "#config_version=6" & vbCrLf & vbCrLf
		Dim coilsYaml : coilsYaml = "#config_version=6" & vbCrLf & vbCrLf
		coilsYaml = coilsYaml + "coils:" & vbCrLf
		Dim ballDevicesYaml : ballDevicesYaml = "#config_version=6" & vbCrLf & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "ball_devices:" & vbCrLf
		Dim configYaml : configYaml = "#config_version=6" & vbCrLf & vbCrLf
		configYaml = configYaml + "config:" & vbCrLf
		configYaml = configYaml + "  - lights.yaml" & vbCrLf
		configYaml = configYaml + "  - ball_devices.yaml" & vbCrLf
		configYaml = configYaml + "  - coils.yaml" & vbCrLf
		configYaml = configYaml + "  - switches.yaml" & vbCrLf
		configYaml = configYaml + vbCrLf
		configYaml = configYaml + "playfields:" & vbCrLf
		configYaml = configYaml + "  playfield:" & vbCrLf
		configYaml = configYaml + "    tags: default" & vbCrLf
		configYaml = configYaml + "    default_source_device: balldevice_plunger" & vbCrLf

		
		
		Dim lightsYaml : lightsYaml = "#config_version=6" & vbCrLf & vbCrLf
		lightsYaml = lightsYaml + "lights:" & vbCrLf
		Dim monitorYaml : monitorYaml = "light:" & vbCrLf
		Dim godotLightScene : godotLightScene = ""
		For Each light in glf_lights
			monitorYaml = monitorYaml + "  " & light.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    size: 0.04" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& light.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& light.y/tableheight & vbCrLf

			lightsYaml = lightsYaml + "  " & light.name & ":"&vbCrLf
			lightsYaml = lightsYaml + "    number: " & lightsNumber & vbCrLf
			lightsYaml = lightsYaml + "    subtype: led" & vbCrLf
			lightsYaml = lightsYaml + "    type: rgb" & vbCrLf
			lightsYaml = lightsYaml + "    tags: " & light.BlinkPattern & vbCrLf
			lightsNumber = lightsNumber + 1

			godotLightScene = godotLightScene + "[node name="""&light.name&""" type=""Sprite2D"" parent=""lights""]" & vbCrLf
			godotLightScene = godotLightScene + "position = Vector2("&light.x*scaleFactor&", "&light.y*scaleFactor&")" & vbCrLf
			godotLightScene = godotLightScene + "script = ExtResource(""3_qb2nn"")" & vbCrLf

			Dim splitTagArray : splitTagArray = Split(light.BlinkPattern, ",")
			Dim outputTagString : outputTagString = ""
			Dim i
			For i = LBound(splitTagArray) To UBound(splitTagArray)
				outputTagString = outputTagString & """" & Trim(splitTagArray(i)) & """"
				If i < UBound(splitTagArray) Then
					outputTagString = outputTagString & ", "
				End If
			Next

			godotLightScene = godotLightScene + "tags = ["&outputTagString&"]" & vbCrLf
			godotLightScene = godotLightScene + vbCrLf
		Next

		monitorYaml = monitorYaml + vbCrLf
		monitorYaml = monitorYaml + "switch:" & vbCrLf
		switchesYaml = switchesYaml + "switches:" & vbCrLf

		For Each switch in glf_switches
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& switch.y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  " & switch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		For Each switch in glf_spinners
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& switch.y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  " & switch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		For Each switch in glf_slingshots
			monitorYaml = monitorYaml + "  " & switch.name & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& switch.BlendDisableLighting & vbCrLf
			monitorYaml = monitorYaml + "    y: "& 1-switch.BlendDisableLightingFromBelow & vbCrLf
			switchesYaml = switchesYaml + "  " & switch.name & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
		Next
		Dim troughCount, troughSwitchesArr()
		ReDim troughSwitchesArr(tnob)
		configYaml = configYaml + vbCrLf & "virtual_platform_start_active_switches:" & vbCrLf
		For troughCount=1 to tnob
			monitorYaml = monitorYaml + "  s_trough" & troughCount & ":"&vbCrLf
			monitorYaml = monitorYaml + "    shape: RECTANGLE" & vbCrLf
			monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
			monitorYaml = monitorYaml + "    x: "& Eval("swTrough"&troughCount).x/tablewidth & vbCrLf
			monitorYaml = monitorYaml + "    y: "& Eval("swTrough"&troughCount).y/tableheight & vbCrLf
			switchesYaml = switchesYaml + "  s_trough" & troughCount & ":"&vbCrLf
			switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
			switchesYaml = switchesYaml + "    tags: " & vbCrLf
			switchNumber = switchNumber + 1
			troughSwitchesArr(troughCount-1) = "s_trough" & troughCount
			configYaml = configYaml + "  - s_trough" & troughCount & vbCrLf
		Next

		

		switchesYaml = switchesYaml + "  s_trough_jam" & ":"&vbCrLf
		switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
		switchesYaml = switchesYaml + "    tags: " & vbCrLf
		switchNumber = switchNumber + 1

		monitorYaml = monitorYaml + "  s_start:"&vbCrLf
		monitorYaml = monitorYaml + "    size: 0.06" & vbCrLf
		monitorYaml = monitorYaml + "    x: 0.95" & vbCrLf
		monitorYaml = monitorYaml + "    y: 0.95" & vbCrLf
		switchesYaml = switchesYaml + "  s_start:"&vbCrLf
		switchesYaml = switchesYaml + "    number: " & switchNumber & vbCrLf
		switchesYaml = switchesYaml + "    tags: start" & vbCrLf
		switchNumber = switchNumber + 1

		dim device

		ballDevicesYaml = ballDevicesYaml + "  bd_trough:" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    ball_switches: "&Join(troughSwitchesArr, ",")&" s_trough_jam" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    eject_coil: c_trough_eject" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    tags: trough, home, drain" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    jam_switch: s_trough_jam" & vbCrLf
		ballDevicesYaml = ballDevicesYaml + "    eject_targets: balldevice_plunger" & vbCrLf
		

		coilsYaml = coilsYaml + "  c_trough_eject:" & vbCrLf
		coilsYaml = coilsYaml + "    number: " & coilsNumber & vbCrLf 
		coilsNumber = coilsNumber + 1

		For Each device in glf_ball_devices.Items()
			ballDevicesYaml = ballDevicesYaml + device.ToYaml()
			coilsYaml = coilsYaml + "  c_" & device.Name & "_eject:" & vbCrLf
			coilsYaml = coilsYaml + "    number: " & coilsNumber & vbCrLf 
			coilsNumber = coilsNumber + 1
		Next

		Dim fso, modesFolder, TxtFileStream, monitorFolder, configFolder
		Set fso = CreateObject("Scripting.FileSystemObject")
		monitorFolder = "glf_mpf\monitor\"
		configFolder = "glf_mpf\config\"
		If Not fso.FolderExists("glf_mpf") Then
			fso.CreateFolder "glf_mpf"
		End If
		If Not fso.FolderExists("glf_mpf\monitor") Then
			fso.CreateFolder "glf_mpf\monitor"
		End If
		If Not fso.FolderExists("glf_mpf\config") Then
			fso.CreateFolder "glf_mpf\config"
		End If
		Set TxtFileStream = fso.OpenTextFile(monitorFolder & "\monitor.yaml", 2, True)
		TxtFileStream.WriteLine monitorYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\config.yaml", 2, True)
		TxtFileStream.WriteLine configYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\ball_devices.yaml", 2, True)
		TxtFileStream.WriteLine ballDevicesYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\coils.yaml", 2, True)
		TxtFileStream.WriteLine coilsYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\switches.yaml", 2, True)
		TxtFileStream.WriteLine switchesYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(configFolder & "\lights.yaml", 2, True)
		TxtFileStream.WriteLine lightsYaml
		TxtFileStream.Close
		Set TxtFileStream = fso.OpenTextFile(monitorFolder & "\gotdotlights.txt", 2, True)
		TxtFileStream.WriteLine godotLightScene
		TxtFileStream.Close
		
	End If

	'Cache Shows
	Dim mode, show_count, shot_count, cached_show
	show_count = 0
	shot_count = 0
	For Each mode in glf_modes.Items()
		Glf_WriteDebugLog "Init", mode.Name
		If Not IsNull(mode.ShowPlayer) Then
			With mode.ShowPlayer()
				Dim show_settings
				For Each show_settings in .EventShows()
					If Not IsNull(show_settings.Show) And show_settings.Action = "play" Then
						show_settings.InternalCacheId = CStr(show_count)
						show_count = show_count + 1
						cached_show = Glf_ConvertShow(show_settings.Show, show_settings.Tokens)
						glf_cached_shows.Add "show_player_" & mode.name & "_" & show_settings.Key & "__" & show_settings.InternalCacheId, cached_show
					End If 
				Next
			End With
		End If

		If Not IsNull(mode.LightPlayer) Then
			With mode.LightPlayer()
				.ReloadLights()
			End With
		End If
		Dim x,state
		If UBound(mode.ModeShots) > -1 Then
			Dim mode_shot
			For Each mode_shot in mode.ModeShots
				Dim shot_profile : Set shot_profile = Glf_ShotProfiles(mode_shot.Profile)
				
				If mode_shot.InternalCacheId = -1 Then
					mode_shot.InternalCacheId = shot_count
					shot_count = shot_count + 1
				End If
				For x=0 to shot_profile.StatesCount
					Set state = shot_profile.StateForIndex(x)
					If state.InternalCacheId = -1 Then
						state.InternalCacheId = CStr(show_count)
						show_count = show_count + 1
					End If

					Dim key
					Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
					If Not IsNull(state.Tokens) Then
						For Each key In state.Tokens.Keys()
							mergedTokens.Add key, state.Tokens()(key)
						Next
					End If
					Dim tokens
					If Not IsNull(mode_shot.Tokens) Then
						Set tokens = mode_shot.Tokens
						For Each key In tokens.Keys
							If mergedTokens.Exists(key) Then
								mergedTokens(key) = tokens(key)
							Else
								mergedTokens.Add key, tokens(key)
							End If
						Next
					End If
					cached_show = Glf_ConvertShow(state.Show, mergedTokens)
					glf_cached_shows.Add mode.name & "_" & x & "_" & mode_shot.Name & "_" & state.Key & "_" & mode_shot.InternalCacheId & "_" & state.InternalCacheId, cached_show
				Next
			Next
		End If

		If UBound(mode.ModeStateMachines) > -1 Then
			Dim mode_state_machine,state_count
			state_count = 0
			For Each mode_state_machine in mode.ModeStateMachines
				
				For x=0 to UBound(mode_state_machine.StateItems)
					Set state = mode_state_machine.StateItems()(x)
					If state.InternalCacheId = -1 Then
						state.InternalCacheId = CStr(state_count)
						state_count = state_count + 1
					End If
					If Not IsNull(state.ShowWhenActive().Show) Then
						If state.ShowWhenActive().Action = "play" Then
							state.ShowWhenActive().InternalCacheId = CStr(show_count)
							show_count = show_count + 1
							cached_show = Glf_ConvertShow(state.ShowWhenActive().Show, state.ShowWhenActive().Tokens)
							glf_cached_shows.Add mode.name & "_" & mode_state_machine.Name & "_" & state.Name & "_" & state.ShowWhenActive().Key & "_" & state.InternalCacheId & "_" & state.ShowWhenActive().InternalCacheId, cached_show
						End If
					End If
				Next
			Next
		End If
	Next

	With CreateMachineVar("player1_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With
	With CreateMachineVar("player2_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With
	With CreateMachineVar("player3_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With
	With CreateMachineVar("player4_score")
        .InitialValue = 0
        .ValueType = "int"
        .Persist = True
    End With

	Glf_ReadMachineVars()

	Glf_Reset()
End Sub

Sub Glf_Reset()

	DispatchQueuePinEvent "reset_complete", Null
End Sub

AddPinEventListener "reset_complete", "initial_segment_displays", "Glf_SegmentInit", 100, Null
AddPinEventListener "reset_virtual_segment_lights", "reset_segment_displays", "Glf_SegmentInit", 100, Null
Sub Glf_SegmentInit(args)
	Dim segment_display
	For Each segment_display in glf_segment_displays.Items()	
		segment_display.SetVirtualDMDLights Not glf_flex_alphadmd_enabled
	Next
End Sub

Sub Glf_ReadMachineVars()
    Dim objFSO, objFile, arrLines, line, inSection
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If Not objFSO.FileExists(CGameName & "_glf.ini") Then Exit Sub
    
    Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
    arrLines = Split(objFile.ReadAll, vbCrLf)
    objFile.Close
    
    inSection = False
    For Each line In arrLines
        line = Trim(line)
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase("MachineVars"))
        ElseIf inSection And InStr(line, "=") > 0 Then
			Dim key : key = Trim(Split(line, "=")(0))
			If glf_machine_vars.Exists(key) Then
				If glf_machine_vars(key).Persist = True Then
	            	glf_machine_vars(key).Value = Trim(Split(line, "=")(1))
				End If
			End If
        End If
    Next
End Sub

Sub Glf_DisableVirtualSegmentDmd()
	If Not glf_flex_alphadmd is Nothing Then
		glf_flex_alphadmd.Show = False
		glf_flex_alphadmd.Run = False
		Set glf_flex_alphadmd = Nothing
	End If
	glf_flex_alphadmd_enabled = False
	DispatchPinEvent "reset_virtual_segment_lights", Null
End Sub

Sub Glf_EnableVirtualSegmentDmd()
	Glf_DisableVirtualSegmentDmd()
	Dim i
	Set glf_flex_alphadmd = CreateObject("FlexDMD.FlexDMD")
	With glf_flex_alphadmd
		.TableFile = Table1.Filename & ".vpx"
		.Color = RGB(255, 88, 32)
		.Width = 128
		.Height = 32
		.Clear = True
		.Run = True
		.GameName = cGameName
		.RenderMode = 3
	End With
	For i = 0 To 31
		glf_flex_alphadmd_segments(i) = 0
	Next
	glf_flex_alphadmd.Segments = glf_flex_alphadmd_segments
	glf_flex_alphadmd_enabled = True
	DispatchPinEvent "reset_virtual_segment_lights", Null
End Sub

Sub Glf_WriteMachineVars()
    Dim objFSO, objFile, arrLines, line, inSection, foundSection
    Dim outputLines, key
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(CGameName & "_glf.ini") Then
        Set objFile = objFSO.OpenTextFile(CGameName & "_glf.ini", 1)
        arrLines = Split(objFile.ReadAll, vbCrLf)
        objFile.Close
    Else
        arrLines = Array()
    End If
    
    outputLines = ""
    inSection = False
    foundSection = False
    
    For Each line In arrLines
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            inSection = (LCase(Mid(line, 2, Len(line) - 2)) = LCase("MachineVars"))
            foundSection = foundSection Or inSection
        End If
        
        If inSection And InStr(line, "=") > 0 Then
            key = Trim(Split(line, "=")(0))
            If glf_machine_vars.Exists(key) Then
				If glf_machine_vars(key).Persist = True Then
                	line = key & "=" & glf_machine_vars(key).Value
                	glf_machine_vars.Remove key
				End If
            End If
        End If

		If line = "" And inSection Then
			' Add remaining keys in the section
			For Each key In glf_machine_vars.Keys
				If glf_machine_vars(key).Persist = True Then
					outputLines = outputLines & key & "=" & glf_machine_vars(key).Value & vbCrLf
				End If
			Next
			glf_machine_vars.RemoveAll
		End If
        If line <> "" Then
			outputLines = outputLines & line & vbCrLf
		End If
    Next
    
    If Not foundSection Then
        outputLines = outputLines & "[MachineVars]" & vbCrLf
        For Each key In glf_machine_vars.Keys
			If glf_machine_vars(key).Persist = True Then
            	outputLines = outputLines & key & "=" & glf_machine_vars(key).Value & vbCrLf
			End If
        Next
    End If
    
    Set objFile = objFSO.CreateTextFile(CGameName & "_glf.ini", True)
    objFile.Write outputLines
    objFile.Close
End Sub

Sub Glf_Options(ByVal eventId)
	Dim ballsPerGame : ballsPerGame = Table1.Option("Balls Per Game", 1, 2, 1, 2, 0, Array("3 Balls", "5 Balls"))
	If ballsPerGame = 1 Then
		glf_ballsPerGame = 3
	Else
		glf_ballsPerGame = 5
	End If

	Dim tilt_sensitivity : tilt_sensitivity = Table1.Option("Tilt Sensitivity (digital nudge)", 1, 10, 1, 5, 0, Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10"))
	glf_tilt_sensitivity = tilt_sensitivity

	Dim glfDebug : glfDebug = Table1.Option("Glf Debug Log", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfDebug = 1 Then
		glf_debugEnabled = True
		glf_debugLog.EnableLogs
	Else
		glf_debugEnabled = False
		glf_debugLog.DisableLogs
	End If

	Dim glfDebugLevel : glfDebugLevel = Table1.Option("Glf Debug Log Level", 0, 1, 1, 0, 0, Array("Info", "Debug"))
	If glfDebugLevel = 1 Then
		glf_debug_level = "Debug"
	Else
		glf_debug_level = "Info"
	End If

	Dim glfMaxDispatch : glfMaxDispatch = Table1.Option("Glf Frame Dispatch", 1, 10, 1, 1, 0, Array("5", "10", "15", "20", "25", "30", "35", "40", "45", "50"))
	glf_max_dispatch = glfMaxDispatch*5

	Dim glfuseBCP : glfuseBCP = Table1.Option("Glf Backbox Control Protocol", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseBCP = 1 Then
		If IsNull(bcpController) Then
			SetDelay "start_glf_bcp", "Glf_ConnectToBCPMediaController", Null, 500
		End If
	Else
		useBCP = False
		If Not IsNull(bcpController) Then
			bcpController.Disconnect
			bcpController = Null
		End If
	End If

	Dim glfuseDebugBCP : glfuseDebugBCP = Table1.Option("Glf Monitor", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseDebugBCP = 1 And useGlfBCPMonitor = False Then
		useGlfBCPMonitor = True
		If IsNull(glf_debugBcpController) Then
			SetDelay "start_glf_monitor", "Glf_ConnectToDebugBCPMediaController", Null, 500
		End If
	ElseIf glfuseDebugBCP = 0 And useGlfBCPMonitor = True Then
		useGlfBCPMonitor = False
		If Not IsNull(glf_debugBcpController) Then
			glf_debugBcpController.Disconnect
			glf_debugBcpController = Null
		End If
	End If

	Dim glfuseVirtualSegmentDMD : glfuseVirtualSegmentDMD = Table1.Option("Glf Virtual Segment DMD", 0, 1, 1, 0, 0, Array("Off", "On"))
	If glfuseVirtualSegmentDMD = 1 And glf_flex_alphadmd_enabled = False Then
		Glf_EnableVirtualSegmentDmd()
	ElseIf glfuseVirtualSegmentDMD = 0 And  glf_flex_alphadmd_enabled = True Then
		Glf_DisableVirtualSegmentDmd()
	End If

	Dim min_lightmap_update_rate : min_lightmap_update_rate = Table1.Option("Glf Min Lightmap Update Rate", 1, 6, 1, 1, 0, Array("Disabled", "30 Hz", "60 Hz", "120 Hz", "144 Hz", "165 Hz"))
    Select Case min_lightmap_update_rate
		Case 1: glf_max_lightmap_sync_enabled = False
		Case 2: glf_max_lightmap_sync = 30 : glf_max_lightmap_sync_enabled = True
        Case 3: glf_max_lightmap_sync = 15 : glf_max_lightmap_sync_enabled = True
        Case 4: glf_max_lightmap_sync = 8 : glf_max_lightmap_sync_enabled = True
        Case 5: glf_max_lightmap_sync = 7 : glf_max_lightmap_sync_enabled = True
        Case 6: glf_max_lightmap_sync = 6 : glf_max_lightmap_sync_enabled = True
    End Select
	
End Sub

Public Sub Glf_Exit()
	If Not IsNull(bcpController) Then
		bcpController.Disconnect
		bcpController = Null
	End If
	If Not IsNull(glf_debugBcpController) Then
		glf_debugBcpController.Disconnect
		glf_debugBcpController = Null
	End If
	If glf_debugEnabled = True Then
		glf_debugLog.WriteToLog "Max Lights", glf_max_lights_test
		glf_debugLog.DisableLogs
	End If
	Glf_DisableVirtualSegmentDmd()
	Glf_WriteMachineVars()
End Sub

Public Sub Glf_KeyDown(ByVal keycode)
    If glf_gameStarted = True Then
		If keycode = StartGameKey Then
			If glf_canAddPlayers = True Then
				Glf_AddPlayer()
			End If
		End If
	Else
		If keycode = StartGameKey Then
			DispatchRelayPinEvent "request_to_start_game", True
		End If
	End If

	If keycode = LeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_flipper_active", Null
	End If
	
	If keycode = RightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_flipper_active", Null
	End If
	
	If keycode = LockbarKey Then
		RunAutoFireDispatchPinEvent "s_lockbar_key_active", Null
	End If

	If KeyCode = PlungerKey Then
		RunAutoFireDispatchPinEvent "s_plunger_key_active", Null
	End If

	If KeyCode = LeftMagnaSave Then
		RunAutoFireDispatchPinEvent "s_left_magna_key_active", Null
	End If

	If KeyCode = RightMagnaSave Then
		RunAutoFireDispatchPinEvent "s_right_magna_key_active", Null
	End If

	If KeyCode = StagedRightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_staged_flipper_key_active", Null
	End If

	If KeyCode = StagedLeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_staged_flipper_key_active", Null
	End If
	
	If keycode = MechanicalTilt Then 
		RunAutoFireDispatchPinEvent "s_tilt_warning_active", Null
    End If

	If keycode = LeftTiltKey Then 
		Nudge 90, 2
		Glf_CheckTilt
	End If
    If keycode = RightTiltKey Then
		Nudge 270, 2
		Glf_CheckTilt
	End If
    If keycode = CenterTiltKey Then 
		Nudge 0, 3
		Glf_CheckTilt
	End If

	If KeyCode = AddCreditKey Then
		RunAutoFireDispatchPinEvent "s_add_credit_key_active", Null
	End If

	If KeyCode = AddCreditKey2 Then
		RunAutoFireDispatchPinEvent "s_add_credit_key2_active", Null
	End If
End Sub

Public Sub Glf_KeyUp(ByVal keycode)
	
	If KeyCode = PlungerKey Then
		RunAutoFireDispatchPinEvent "s_plunger_key_inactive", Null
	End If

	If keycode = LeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_flipper_inactive", Null
	End If
	
	If keycode = RightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_flipper_inactive", Null
	End If

	If keycode = LockbarKey Then
		RunAutoFireDispatchPinEvent "s_lockbar_key_inactive", Null
	End If

	If KeyCode = LeftMagnaSave Then
		RunAutoFireDispatchPinEvent "s_left_magna_key_inactive", Null
	End If

	If KeyCode = RightMagnaSave Then
		RunAutoFireDispatchPinEvent "s_right_magna_key_inactive", Null
	End If

	If KeyCode = StagedRightFlipperKey Then
		RunAutoFireDispatchPinEvent "s_right_staged_flipper_key_inactive", Null
	End If

	If KeyCode = StagedLeftFlipperKey Then
		RunAutoFireDispatchPinEvent "s_left_staged_flipper_key_inactive", Null
	End If		

	If KeyCode = AddCreditKey Then
		RunAutoFireDispatchPinEvent "s_add_credit_key_inactive", Null
	End If

	If KeyCode = AddCreditKey2 Then
		RunAutoFireDispatchPinEvent "s_add_credit_key2_inactive", Null
	End If
End Sub

Dim glf_lastEventExecutionTime, glf_lastBcpExecutionTime, glf_lastLightUpdateExecutionTime, glf_lastTiltUpdateExecutionTime
glf_lastEventExecutionTime = 0
glf_lastBcpExecutionTime = 0
glf_lastLightUpdateExecutionTime = 0
glf_lastTiltUpdateExecutionTime = 0

Public Sub Glf_GameTimer_Timer()

	If (gametime - glf_lastEventExecutionTime) > 25 Then
		'debug.print "Slow GLF Frame: " & gametime - glf_lastEventExecutionTime & ". Dispatch Count: " & glf_frame_dispatch_count & ". Handler Count: " & glf_frame_handler_count
	End If
	glf_frame_dispatch_count = 0
	glf_frame_handler_count = 0
	'glf_temp1 = 0

	Dim i, key, keys, lightMap
	i = 0
	keys = glf_dispatch_handlers_await.Keys()
	i = Glf_RunHandlers(i)
	If i<glf_max_dispatch Then
		keys = glf_dispatch_await.Keys()
		For Each key in keys
			RunDispatchPinEvent key, glf_dispatch_await(key)
			glf_dispatch_await.Remove key
			If UBound(glf_dispatch_handlers_await.Keys())>-1 Then
				'Handlers were added, process those first.
				i = Glf_RunHandlers(i)
			End If
			i = i + 1
			If i>=glf_max_dispatch Then
				Exit For
			End If
		Next
	End If

	DelayTick

	If glf_max_lightmap_sync_enabled = True Then
		keys = glf_dispatch_lightmaps_await.Keys()
		'debug.print(ubound(keys))
		If glf_max_lights_test < Ubound(keys) Then
			glf_max_lights_test = Ubound(keys)
		End If
		For Each key in keys
			For Each lightMap in glf_lightMaps(key)
				lightMap.Color = glf_lightNames(key).Color
			Next
			glf_dispatch_lightmaps_await.Remove key
			If (gametime - glf_lastEventExecutionTime) > glf_max_lightmap_sync Then
				'debug.print("Exiting")
				Exit For
			End If
		Next
	End If
	If (gametime - glf_lastTiltUpdateExecutionTime) >=50 And glf_current_virtual_tilt > 0 Then
		glf_current_virtual_tilt = glf_current_virtual_tilt - 0.1
		glf_lastTiltUpdateExecutionTime = gametime
		'Debug.print("Tilt Cooldown: " & glf_current_virtual_tilt) 
	End If

	If (gametime - glf_lastBcpExecutionTime) >= 300 Then
        glf_lastBcpExecutionTime = gametime
		Glf_BcpUpdate
		Glf_MonitorPlayerStateUpdate "GLF BIP", glf_BIP
		Glf_MonitorBcpUpdate
    End If

	If glf_last_switch_hit_time > 0 And (gametime - glf_last_ballsearch_reset_time) > 2000 Then
		Glf_ResetBallSearch
		glf_last_switch_hit_time = 0
		glf_last_ballsearch_reset_time = gametime
	End If
	glf_lastEventExecutionTime = gametime
End Sub

Sub Glf_CheckTilt()
	glf_current_virtual_tilt = glf_current_virtual_tilt + glf_tilt_sensitivity
	If (glf_current_virtual_tilt > 10) Then 
		RunAutoFireDispatchPinEvent "s_tilt_warning_active", Null
		glf_current_virtual_tilt = glf_tilt_sensitivity
	End If
End Sub

Sub Glf_ResetBallSearch()
	If glf_ballsearch_enabled = True Then
		glf_ballsearch.Reset()
	End If
End Sub

Public Function Glf_RunHandlers(i)
	Dim key, keys, args
	keys = glf_dispatch_handlers_await.Keys()
	If UBound(keys) = -1 Then
		Glf_RunHandlers = i
		Exit Function
	End If
	For Each key in keys
		args = glf_dispatch_handlers_await(key)
		Dim wait_for : wait_for = DispatchPinHandlers(key, args)
		glf_dispatch_handlers_await.Remove key
		If Not IsEmpty(wait_for) Then
			Dim remaining_handlers_keys : remaining_handlers_keys = glf_dispatch_handlers_await.Keys
			Dim remaining_handlers_items : remaining_handlers_items = glf_dispatch_handlers_await.Items
			AddPinEventListener wait_for, key & "_wait_for", "ContinueDispatchQueuePinEvent", 1000, Array(remaining_handlers_keys, remaining_handlers_items)
			glf_dispatch_handlers_await.RemoveAll
			Exit For
		End If
		i = i + 1
		If i=glf_max_dispatch Then
			Exit For
		End If
	Next
	If Ubound(glf_dispatch_handlers_await.Keys())=-1 Then
		'Finished processing Handlers for current event.
		'Remove any blocks for this event.
		Glf_EventBlocks(args(2)).RemoveAll
	End If
	Glf_RunHandlers = i
End Function

Public Function Glf_RegisterLights()

	Dim light, tags, tag
	For Each light In Glf_Lights
		tags = Split(light.BlinkPattern, ",")
		For Each tag in tags
			
			tag = "T_" & Trim(tag)
			If Not glf_lightTags.Exists(tag) Then
				Set glf_lightTags(tag) = CreateObject("Scripting.Dictionary")
			End If
			If Not glf_lightTags(tag).Exists(light.Name) Then
				glf_lightTags(tag).Add light.Name, True
			End If
		Next
		glf_lightPriority.Add light.Name, 0
		
		Dim e, lmStr: lmStr = "lmArr = Array("    
		For Each e in GetElements()
			On Error Resume Next
			If InStr(LCase(e.Name), LCase("_" & light.Name & "_")) Then
				lmStr = lmStr & e.Name & ","
			End If
			For Each tag in tags
				tag = "T_" & Trim(tag)
				If InStr(LCase(e.Name), LCase("_" & tag & "_")) Then
					lmStr = lmStr & e.Name & ","
				End If
			Next
			If Err Then Log "Error: " & Err
		Next
		lmStr = lmStr & "Null)"
		lmStr = Replace(lmStr, ",Null)", ")")
		lmStr = Replace(lmStr, "Null)", ")")
		ExecuteGlobal "Dim lmArr : "&lmStr
		glf_lightMaps.Add light.Name, lmArr
		glf_lightNames.Add light.Name, light
		Dim lightStack : Set lightStack = (new GlfLightStack)()
		glf_lightStacks.Add light.Name, lightStack
		light.State = 1
		Glf_SetLight light.Name, "000000"
	Next
End Function

Public Function Glf_SetLight(light, color)

	Dim rgbColor
	If glf_lightColorLookup.Exists(color) Then
		rgbColor = glf_lightColorLookup(color)
	Else
		glf_lightColorLookup.Add color, RGB( CInt("&H" & (Left(color, 2))), CInt("&H" & (Mid(color, 3, 2))), CInt("&H" & (Right(color, 2)) ))
		rgbColor = glf_lightColorLookup(color)
	End If
	
	glf_lightNames(light).Color = rgbColor

	If glf_max_lightmap_sync_enabled = True Then
		If Not glf_dispatch_lightmaps_await.Exists(light) Then
			glf_dispatch_lightmaps_await.Add light, True
		End If
	Else
		dim lightMap
		For Each lightMap in glf_lightMaps(light)
			lightMap.Color = glf_lightNames(light).Color
		Next
	End If
End Function

Public Function Glf_ParseInput(value)
	Dim templateCode : templateCode = ""
	Dim tmp: tmp = value
	Dim isVariable, parts
    Select Case VarType(value)
        Case 8 ' vbString
			tmp = Glf_ReplaceCurrentPlayerAttributes(tmp)
			tmp = Glf_ReplaceAnyPlayerAttributes(tmp)
			tmp = Glf_ReplaceDeviceAttributes(tmp)
			tmp = Glf_ReplaceMachineAttributes(tmp)
			tmp = Glf_ReplaceModeAttributes(tmp)
			tmp = Glf_ReplaceGameAttributes(tmp)
			tmp = Glf_ReplaceKwargsAttributes(tmp)
			'msgbox tmp
			If InStr(tmp, " if ") Then
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & Glf_ConvertIf(tmp, "Glf_" & glf_FuncCount) & vbCrLf
				templateCode = templateCode & "End Function"
			Else
				isVariable = Glf_IsCondition(tmp)
				If Not IsNull(isVariable) Then
					'The input needs formatting
					parts = Split(isVariable, ":")
					If UBound(parts) = 1 Then
						tmp = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
					End If
				End If
				templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
				templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
				templateCode = templateCode & "End Function"
			End IF
        Case Else
			templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf			
			isVariable = Glf_IsCondition(tmp)
			If Not IsNull(isVariable) Then
				'The input needs formatting
				parts = Split(isVariable, ":")
				If UBound(parts) = 1 Then
					tmp = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
				End If
			End If
			templateCode = templateCode & vbTab & "Glf_" & glf_FuncCount & " = " & tmp & vbCrLf
			templateCode = templateCode & "End Function"
    End Select
	'msgbox templateCode
	ExecuteGlobal templateCode
	Dim funcRef : funcRef = "Glf_" & glf_FuncCount
	glf_FuncCount = glf_FuncCount + 1
	Glf_ParseInput = Array(funcRef, value, True)
End Function

Public Function Glf_ParseEventInput(value)
	Dim templateCode : templateCode = ""
	Dim parts : parts = Split(value, ":")
	Dim event_delay : event_delay = 0
	Dim priority : priority = 0
	If UBound(parts) = 1 Then
		value = parts(0)
		event_delay= parts(1)
	End If

	Dim condition : condition = Glf_IsCondition(value)
	If IsNull(condition) Then
		parts = Split(value, ".")
		If UBound(parts) = 1 Then
			value = parts(0)
			priority= parts(1)
		End If
		Glf_ParseEventInput = Array(value, value, Null, event_delay, priority)
	Else
		dim conditionReplaced : conditionReplaced = Glf_ReplaceCurrentPlayerAttributes(condition)
		conditionReplaced = Glf_ReplaceAnyPlayerAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceDeviceAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceMachineAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceModeAttributes(conditionReplaced)
		conditionReplaced = Glf_ReplaceGameAttributes(conditionReplaced)

		conditionReplaced = Glf_ReplaceKwargsAttributes(conditionReplaced)
		templateCode = "Function Glf_" & glf_FuncCount & "()" & vbCrLf
		templateCode = templateCode & vbTab & "On Error Resume Next" & vbCrLf
		templateCode = templateCode & vbTab & Glf_ConvertCondition(conditionReplaced, "Glf_" & glf_FuncCount) & vbCrLf
		templateCode = templateCode & vbTab & "If Err Then Glf_" & glf_FuncCount & " = False" & vbCrLf
		templateCode = templateCode & "End Function"
		ExecuteGlobal templateCode
		Dim funcRef : funcRef = "Glf_" & glf_FuncCount
		glf_FuncCount = glf_FuncCount + 1

		value = Replace(value, "{"&condition&"}", "")
		parts = Split(value, ".")
		If UBound(parts) = 1 Then
			value = parts(0)
			priority= parts(1)
		End If
		Glf_ParseEventInput = Array(value & funcRef ,value, funcRef, event_delay, priority)
	End If
End Function

Function Glf_ReplaceCurrentPlayerAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "current_player\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    replacement = "GetPlayerState(""$1"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceCurrentPlayerAttributes = outputString
End Function

Function Glf_ReplaceAnyPlayerAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "players\[([0-3]+)\]\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    replacement = "GetPlayerStateForPlayer($1, ""$2"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceAnyPlayerAttributes = outputString
End Function

Function Glf_ReplaceDeviceAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "devices\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
	replacement = "glf_$1(""$2"").GetValue(""$3"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceDeviceAttributes = outputString
End Function

Function Glf_ReplaceMachineAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "machine\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
	replacement = "glf_machine_vars(""$1"").GetValue()"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceMachineAttributes = outputString
End Function

Function Glf_ReplaceGameAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "game\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
	replacement = "Glf_GameVariable(""$1"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceGameAttributes = outputString
End Function

Function Glf_ReplaceModeAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "modes\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
	replacement = "glf_modes(""$1"").GetValue(""$2"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceModeAttributes = outputString
End Function

Function Glf_ReplaceKwargsAttributes(inputString)
    Dim pattern, replacement, regex, outputString
    pattern = "kwargs\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = True
    replacement = "glf_dispatch_current_kwargs(""$1"")"
    outputString = regex.Replace(inputString, replacement)
    Set regex = Nothing
    Glf_ReplaceKwargsAttributes = outputString
End Function

Function Glf_GameVariable(value)
	Glf_GameVariable = False
	Select Case value
		Case "tilted"
			Glf_GameVariable = glf_gameTilted
		Case "balls_per_game"
			Glf_GameVariable = glf_ballsPerGame
	End Select
End Function

Function Glf_CheckForGetPlayerState(inputString)
    Dim pattern, regex, matches, match, hasGetPlayerState, attribute, playerNumber
	inputString = Glf_ReplaceCurrentPlayerAttributes(inputString)
	inputString = Glf_ReplaceAnyPlayerAttributes(inputString)
    pattern = "GetPlayerState\(""([a-zA-Z0-9_]+)""\)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False
    
    Set matches = regex.Execute(inputString)
    If matches.Count > 0 Then
        hasGetPlayerState = True
		playerNumber = -1 'Current Player
        attribute = matches(0).SubMatches(0)
    Else
        hasGetPlayerState = False
        attribute = ""
		playerNumber = Null
    End If


	pattern = "GetPlayerStateForPlayer\(([0-3]), ""([a-zA-Z0-9_]+)""\)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False
    
    Set matches = regex.Execute(inputString)
    If matches.Count > 0 Then
        hasGetPlayerState = True
		playerNumber = Int(matches(0).SubMatches(0))
        attribute = matches(0).SubMatches(1)
    Else
        hasGetPlayerState = False
        attribute = ""
		playerNumber = Null
    End If

    Set regex = Nothing
    Set matches = Nothing
    
    Glf_CheckForGetPlayerState = Array(hasGetPlayerState, attribute, playerNumber)
End Function

Function Glf_CheckForDeviceState(inputString)
    Dim pattern, regex, matches, match, hasDeviceState, attribute, deviceType, deviceName
    pattern = "devices\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)\.([a-zA-Z0-9_]+)"
    Set regex = New RegExp
    regex.Pattern = pattern
    regex.IgnoreCase = True
    regex.Global = False
    Set matches = regex.Execute(inputString)
    If matches.Count > 0 Then
        hasDeviceState = True
		deviceType = matches(0).SubMatches(0)
		deviceType = Left(deviceType, Len(deviceType)-1)
        deviceName = matches(0).SubMatches(1)
		attribute = matches(0).SubMatches(2)
		attribute = Left(attribute, Len(attribute)-1)
    Else
        hasDeviceState = False
        deviceType = Empty
		deviceName = Empty
		attribute = Empty
    End If

    Set regex = Nothing
    Set matches = Nothing
    
    Glf_CheckForDeviceState = Array(hasDeviceState, deviceType, deviceName, attribute)
End Function

Function Glf_ConvertIf(value, retName)
    Dim parts, condition, truePart, falsePart, isVariable
    parts = Split(value, " if ")
    truePart = Trim(parts(0))
    Dim conditionAndFalsePart
    conditionAndFalsePart = Split(parts(1), " else ")
    condition = Trim(conditionAndFalsePart(0))
    falsePart = Trim(conditionAndFalsePart(1))
	isVariable = Glf_IsCondition(truePart)
	If Not IsNull(isVariable) Then
		'The input needs formatting
		parts = Split(isVariable, ":")
		If UBound(parts) = 1 Then
			truePart = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
		End If
	End If
	isVariable = Glf_IsCondition(falsePart)
	If Not IsNull(isVariable) Then
		'The input needs formatting
		parts = Split(isVariable, ":")
		If UBound(parts) = 1 Then
			falsePart = "Glf_FormatValue(" & parts(0) & ", """ & parts(1) & """)"
		End If
	End If

    Dim vbscriptIfStatement
    vbscriptIfStatement = "If " & condition & " Then" & vbCrLf & _
                          "    "&retName&" = " & truePart & vbCrLf & _
                          "Else" & vbCrLf & _
                          "    "&retName&" = " & falsePart & vbCrLf & _
                          "End If"
	Glf_ConvertIf = vbscriptIfStatement
End Function

Function Glf_ConvertCondition(value, retName)
	value = Replace(value, "==", "=")
	value = Replace(value, "!=", "<>")
	value = Replace(value, "&&", "And")
	Glf_ConvertCondition = "    "&retName&" = " & value
End Function

Function Glf_FormatValue(value, formatString)
    Dim padChar, width, result, align, hasCommas
	
	If CStr(value) = "False" Then
		Glf_FormatValue = ""
		Exit Function
	End If

    ' Default values
    padChar = " " ' Default padding character is space
    align = ">"   ' Default alignment is right
    width = 0     ' Default width is 0 (no padding)
    hasCommas = False ' Default: No thousand separators

    ' Check for :, in the format string
    If InStr(formatString, ",") > 0 Then
        hasCommas = True
        formatString = Replace(formatString, ",", "") ' Remove , from format string
    End If

    ' Parse the remaining format string
    If Len(formatString) >= 2 Then
        padChar = Mid(formatString, 1, 1)
        align = Mid(formatString, 2, 1)
        width = CInt(Mid(formatString, 3))
    End If

    ' Format the value
    If hasCommas And IsNumeric(value) Then
        ' Add commas as thousand separators
        Dim numStr, decimalPart
        numStr = CStr(value)
        If InStr(numStr, ".") > 0 Then
            decimalPart = Mid(numStr, InStr(numStr, "."))
            numStr = Left(numStr, InStr(numStr, ".") - 1)
        Else
            decimalPart = ""
        End If

        Dim i, formattedNum
        formattedNum = ""
        For i = Len(numStr) To 1 Step -1
            formattedNum = Mid(numStr, i, 1) & formattedNum
            If ((Len(numStr) - i) Mod 3 = 2) And (i > 1) Then
                formattedNum = "," & formattedNum
            End If
        Next
        value = formattedNum & decimalPart
    End If

    ' Apply alignment and padding
    Select Case align
        Case ">" ' Right-align with padding
            If Len(value) < width Then
                result = String(width - Len(value), padChar) & value
            Else
                result = value
            End If
        Case "<" ' Left-align with padding
            If Len(value) < width Then
                result = value & String(width - Len(value), padChar)
            Else
                result = value
            End If
        Case "^" ' Center-align with padding
            Dim leftPad, rightPad
            If Len(value) < width Then
                leftPad = (width - Len(value)) \ 2
                rightPad = width - Len(value) - leftPad
                result = String(leftPad, padChar) & value & String(rightPad, padChar)
            Else
                result = value
            End If
        Case Else ' Default: Return value as is
            result = value
    End Select

    Glf_FormatValue = result
End Function

Public Sub Glf_SetInitialPlayerVar(variable_name, initial_value)
	glf_initialVars.Add variable_name, initial_value
End Sub

Function glf_ReadShowYAMLFiles()
    Dim fso, folder, file, yamlFiles, fileContent
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Check if the directory exists
    If Not fso.FolderExists(directoryPath) Then
        WScript.Echo "Directory does not exist: " & directoryPath
        Exit Function
    End If
    
    ' Initialize the array to store file contents
    ReDim yamlFiles(-1)
    
    ' Get the folder object
    Set folder = fso.GetFolder(directoryPath)
    
    ' Iterate through the files in the directory
    For Each file In folder.Files
        ' Check if the file has a .yaml extension
        If LCase(fso.GetExtensionName(file.Name)) = "yaml" Then
            ' Read the file content
            Set fileContent = fso.OpenTextFile(file.Path, 1) ' 1 = ForReading
            ReDim Preserve yamlFiles(UBound(yamlFiles) + 1)
            yamlFiles(UBound(yamlFiles)) = fileContent.ReadAll
            fileContent.Close
        End If
    Next
    
    ' Return the array of YAML file contents
    ReadYAMLFiles = yamlFiles
End Function

Sub glf_ConvertYamlShowToGlfShow(yamlFilePath)
    Dim fso, file, content, lines, line, output, i, stepLights
    Dim glf_ShowName, stepTime, lightsDict, key, lightName, color, intensity, lightParts
    
    ' Initialize FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Read the YAML file
    Set file = fso.OpenTextFile(yamlFilePath, 1)
    content = file.ReadAll
    file.Close
    
    ' Split the content into lines
    lines = Split(content, vbLf)
    
    ' Initialize variables
    glf_ShowName = fso.GetBaseName(yamlFilePath)
    output = "Dim glf_Show" & glf_ShowName & " : Set glf_Show" & glf_ShowName & " = (new GlfShow)(""" & glf_ShowName & """)" & vbCrLf
    output = output & "With glf_Show" & glf_ShowName & vbCrLf
    
    ' Iterate through lines to extract steps and lights
	stepLights = ""
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        
        ' Close the step when a new time step or end of file is reached
        If InStr(line, "- time:") = 1 And stepLights <> "" Then
			'output = output & vbTab & vbTab & ".Lights = Array("& Left(stepLights, Len(stepLights) - 1)&")" & vbCrLf
			output = output & vbTab & vbTab & ".Lights = Array(" & _ 
            SplitStringWithUnderscore(Left(stepLights, Len(stepLights) - 1), 1500) & ")" & vbCrLf
            output = output & vbTab & "End With" & vbCrLf
			stepLights = ""
		End If

        ' Identify time steps
        If InStr(line, "- time:") = 1 Then
            stepTime = Trim(Split(line, ":")(1))
            output = output & vbTab & "With .AddStep(" & stepTime & ", Null, Null)" & vbCrLf
        
		ElseIf InStr(line, "lights:") = 1 Then
			'Do Nothing
        ' Identify lights and colors
        ElseIf InStr(line, ":") > 0 Then
            lightParts = Split(line, ":")
			key = lightParts(0)
            lightName = Trim(key)
			
            color = lightParts(1)
			color = Trim(color)
			color = Replace(color, """", "")
            'msgbox key & "<>" & color
            ' Default intensity to 100 if color is not "000000"
            intensity = 100
            'If color = "000000" Then
            '    intensity = 0
            'End If
            
            ' Add lights to output
            'If intensity > 0 Then
                
            'End If
			stepLights = stepLights + """" & lightName & "|" & intensity & "|" & color & ""","

        End If
    Next
    'msgbox Len(stepLights)
    ' Close the final step and the show
	If Len(stepLights) = 0 Then
		output = output & vbTab & vbTab & ".Lights = Array()" & vbCrLf
	Else
		output = output & vbTab & vbTab & ".Lights = Array(" & _ 
		SplitStringWithUnderscore(Left(stepLights, Len(stepLights) - 1), 1500) & ")" & vbCrLf
	End If
	
    output = output & vbTab & "End With" & vbCrLf
    output = output & "End With" & vbCrLf
    
    ' Write the output to a VBScript file
    Dim outputFilePath
    outputFilePath = fso.GetParentFolderName(yamlFilePath) & "\" & glf_ShowName & ".vbs"
    Set file = fso.CreateTextFile(outputFilePath, True)
    file.Write output
    file.Close
    
    ' Clean up
    Set fso = Nothing
    Set file = Nothing
End Sub

Function SplitStringWithUnderscore(str, maxLength)
    Dim result, i, strLen
    strLen = Len(str)
    result = ""
    
    If strLen <= maxLength Then
        result = str
    Else
        For i = 1 To strLen Step maxLength
            If i + maxLength - 1 < strLen Then
                result = result & Mid(str, i, maxLength) & "_" & vbCrLf & vbTab & vbTab & vbTab
            Else
                result = result & Mid(str, i, maxLength)
            End If
        Next
    End If
    
    SplitStringWithUnderscore = result
End Function

Function GlfShotProfiles(name)
	If Glf_ShotProfiles.Exists(name) Then
		Set GlfShotProfiles = Glf_ShotProfiles(name)
	Else
		Dim new_shotprofile : Set new_shotprofile = (new GlfShotProfile)(name)
		Glf_ShotProfiles.Add name, new_shotprofile
		Set GlfShotProfiles = new_shotprofile
	End If
End Function

Function CreateGlfMode(name, priority)
	If Not glf_modes.Exists(name) Then 
		Dim mode : Set mode = (new Mode)(name, priority)
		glf_modes.Add name, mode
		Set CreateGlfMode = mode
	End If
End Function

Function GlfModes(name)
	If glf_modes.Exists(name) Then 
		Set GlfModes = glf_modes(name)
	Else
		GlfModes = Null
	End If
End Function

Function GlfKwargs()
	Set GlfKwargs = CreateObject("Scripting.Dictionary")
End Function

Function Glf_ConvertShow(show, tokens)

	Dim showStep, light, lightsCount, x,tagLight, tagLights, lightParts, token, stepIdx
	Dim newShow, lightsInShow
	Set lightsInShow = CreateObject("Scripting.Dictionary")

	ReDim newShow(UBound(show.Steps().Keys()))
	stepIdx = 0
	For Each showStep in show.Steps().Items()

		If UBound(showStep.ShowsInStep().Keys()) > -1 Then
			Dim show_item
       	
			Dim show_items : show_items = showStep.ShowsInStep().Items()
        	For Each show_item in show_items
				If Not glf_cached_shows.Exists(show_item.Key & "__" & show_item.InternalCacheId) Then
					Dim cached_show
					cached_show = Glf_ConvertShow(show_item.Show, show_item.Tokens)
					glf_cached_shows.Add show_item.Key & "__" & show_item.InternalCacheId, cached_show
				End If
			Next 
		End If

		lightsCount = 0 
		For Each light in showStep.Lights
			lightParts = Split(light, "|")
			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And Not glf_lightNames.Exists(lightParts(0)) Then
					tagLights = glf_lightTags("T_"&lightParts(0)).Keys()
					lightsCount = UBound(tagLights)+1
				Else
					If IsNull(token) Then
						lightsCount = lightsCount + 1
					Else
						'resolve token lights
						If Not glf_lightNames.Exists(tokens(token)) Then
							'token is a tag
							tagLights = glf_lightTags("T_"&tokens(token)).Keys()
							lightsCount = UBound(tagLights)+1
						Else
							lightsCount = lightsCount + 1
						End If
					End If
				End If
			End If
		Next
	
		Dim seqArray
		ReDim seqArray(lightsCount-1)
		x=0
		For Each light in showStep.Lights
			lightParts = Split(light, "|")
			Dim lightColor : lightColor = ""
			Dim fadeMs : fadeMs = ""
			Dim fade_duration : fade_duration = -1
			Dim intensity
			Dim step_number : step_number = -1
			Dim localLightsSet : Set localLightsSet = CreateObject("Scripting.Dictionary")
			If IsNull(Glf_IsToken(lightParts(1))) Then
				intensity = lightParts(1)
			Else
				intensity = tokens(Glf_IsToken(lightParts(1)))
			End If
			If Ubound(lightParts) >= 2 Then

				If IsNull(Glf_IsToken(lightParts(2))) Then
					lightColor = lightParts(2)
				Else
					lightColor = tokens(Glf_IsToken(lightParts(2)))
				End If
				If UBound(lightParts) = 3 Then
					If IsNull(Glf_IsToken(lightParts(3))) Then
						fade_duration = lightParts(3)
					Else
						fade_duration = tokens(Glf_IsToken(lightParts(3)))
					End If
					step_number = fade_duration * 0.01
					step_number = Round(step_number, 0) + 2
					If step_number<4 Then
						step_number = 4
					End If
					fadeMs = "|" & step_number
				End If
			End If

			Dim resolved_light_name

			If IsArray(lightParts) Then
				token = Glf_IsToken(lightParts(0))
				If IsNull(token) And Not glf_lightNames.Exists(lightParts(0)) Then
					tagLights = glf_lightTags("T_"&lightParts(0)).Keys()
					resolved_light_name = "T_"&lightParts(0)
					For Each tagLight in tagLights
						If UBound(lightParts) >=1 Then
							seqArray(x) = tagLight & "|"&intensity&"|" & AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						Else
							seqArray(x) = tagLight & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						End If
						If Not localLightsSet.Exists(tagLight) Then
							localLightsSet.Add tagLight, True
						End If
						If Not lightsInShow.Exists(tagLight) Then
							lightsInShow.Add tagLight, True
						End If
						x=x+1
					Next
				Else
					If IsNull(token) Then
						resolved_light_name = lightParts(0)
						If UBound(lightParts) >= 1 Then
							seqArray(x) = lightParts(0) & "|"&intensity&"|"&AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						Else
							seqArray(x) = lightParts(0) & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
						End If
						If Not localLightsSet.Exists(lightParts(0)) Then
							localLightsSet.Add lightParts(0), True
						End If
						If Not lightsInShow.Exists(lightParts(0)) Then
							lightsInShow.Add lightParts(0), True
						End If
						x=x+1
					Else
						'resolve token lights
						If Not glf_lightNames.Exists(tokens(token)) Then
							'token is a tag
							tagLights = glf_lightTags("T_"&tokens(token)).Keys()
							resolved_light_name = "T_"&tokens(token)
							For Each tagLight in tagLights
								If UBound(lightParts) >=1 Then
									seqArray(x) = tagLight & "|"&intensity&"|"&AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
								Else
									seqArray(x) = tagLight & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
								End If
								If Not localLightsSet.Exists(tagLight) Then
									localLightsSet.Add tagLight, True
								End If
								If Not lightsInShow.Exists(tagLight) Then
									lightsInShow.Add tagLight, True
								End If
								x=x+1
							Next
						Else
							resolved_light_name = tokens(token)
							If UBound(lightParts) >= 1 Then
								seqArray(x) = tokens(token) & "|"&intensity&"|"&AdjustHexColor(lightColor, intensity) & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
							Else
								seqArray(x) = tokens(token) & "|"&intensity & "|000000" & "|fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number & fadeMs
							End If
							If Not localLightsSet.Exists(tokens(token)) Then
								localLightsSet.Add tokens(token), True
							End If
							If Not lightsInShow.Exists(tokens(token)) Then
								lightsInShow.Add tokens(token), True
							End If
							x=x+1
						End If
					End If
				End If

				'Generate a fake show for the fade in the format light from ? color to lightColor over x steps
				If fadeMs <> "" Then
					Dim fade_seq, i, step_duration,cached_rgb_seq
					 
					Dim cache_name : cache_name = "fade_" & resolved_light_name & "_?_" & AdjustHexColor(lightColor, intensity) & "_steps_" & step_number
					If Not glf_cached_shows.Exists(cache_name & "__-1") Then
						'MsgBox cache_name
						Dim fade_show : Set fade_show = CreateGlfShow(cache_name)
						
					
						step_duration = (fade_duration / step_number)/1000
						For i=1 to step_number

							Dim lightsArr
							ReDim lightsArr(UBound(localLightsSet.Keys))
							Dim localLightItem, k
							k=0
							For Each localLightItem in localLightsSet.Keys()
								cached_rgb_seq = Glf_FadeRGB("FF0000", AdjustHexColor(lightColor, intensity), step_number)
								ReDim fade_seq(step_number - 1)
								Dim j
								For j = 0 To UBound(fade_seq)
									fade_seq(j) = localLightItem & "|100|" & cached_rgb_seq(j)
								Next
								lightsArr(k) = fade_seq(i-1)
								k=k+1
							Next
							With fade_show
								With .AddStep(Null, Null, step_duration)
									.Lights = lightsArr
								End With
							End With
						Next
						cached_show = Glf_ConvertShow(fade_show, Null)
						'msgbox "Converted show: " & cache_name & ", steps: " & ubound(cached_show(0)) & ". Fade Replacements: " & ubound(cached_rgb_seq)
						glf_cached_shows.Add cache_name & "__-1", cached_show
					End If
				End If

			End If
		Next
		'Glf_WriteDebugLog "Convert Show", Join(seqArray)
		newShow(stepIdx) = seqArray
		stepIdx = stepIdx + 1
	Next
	Glf_ConvertShow = Array(newShow, lightsInShow)
End Function

Private Function Glf_IsToken(mainString)
	' Check if the string contains an opening parenthesis and ends with a closing parenthesis
	If InStr(mainString, "(") > 0 And Right(mainString, 1) = ")" Then
		' Extract the substring within the parentheses
		Dim startPos, subString
		startPos = InStr(mainString, "(")
		subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
		Glf_IsToken = subString
	Else
		Glf_IsToken = Null
	End If
End Function

Private Function Glf_IsCondition(mainString)
	' Check if the string contains an opening { and ends with a closing }
	If InStr(mainString, "{") > 0 And Right(mainString, 1) = "}" Then
		Dim startPos, subString
		startPos = InStr(mainString, "{")
		subString = Mid(mainString, startPos + 1, Len(mainString) - startPos - 1)
		Glf_IsCondition = subString
	Else
		Glf_IsCondition = Null
	End If
End Function

Function Glf_RotateArray(arr, direction)
    Dim n, rotatedArray, i
    ReDim rotatedArray(UBound(arr))
 
    If LCase(direction) = "l" Then
        For i = 0 To UBound(arr) - 1
            rotatedArray(i) = arr(i + 1)
        Next
        rotatedArray(UBound(arr)) = arr(0)
    ElseIf LCase(direction) = "r" Then
        For i = UBound(arr) To 1 Step -1
            rotatedArray(i) = arr(i - 1)
        Next
        rotatedArray(0) = arr(UBound(arr))
    Else
        ' Invalid direction
        Glf_RotateArray = arr
        Exit Function
    End If
    
    ' Return the rotated array
    Glf_RotateArray = rotatedArray
End Function

Function Glf_CopyArray(arr)
    Dim newArr, i
    ReDim newArr(UBound(arr))
    For i = 0 To UBound(arr)
        newArr(i) = arr(i)
    Next
    Glf_CopyArray = newArr
End Function

Function Glf_IsInArray(value, arr)
    Dim i
    Glf_IsInArray = False

    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            Glf_IsInArray = True
            Exit Function
        End If
    Next
End Function

Function CreateGlfInput(value)
	Set CreateGlfInput = (new GlfInput)(value)
End Function

Class GlfInput
	Private m_raw, m_value, m_isGetRef, m_isPlayerState, m_playerStateValue, m_playerStatePlayer, m_isDeviceState, m_deviceStateDeviceType, m_deviceStateDeviceName, m_deviceStateDeviceAttr
  
    Public Property Get Value() 
		If m_isGetRef = True Then
			Value = GetRef(m_value)()
		Else
			Value = m_value
		End If
	End Property

    Public Property Get Raw() : Raw = m_raw : End Property

	Public Property Get IsPlayerState() : IsPlayerState = m_isPlayerState : End Property
	Public Property Get PlayerStateValue() : PlayerStateValue = m_playerStateValue : End Property		
	Public Property Get PlayerStatePlayer() : PlayerStatePlayer = m_playerStatePlayer : End Property		

	Public Property Get IsDeviceState() : IsDeviceState = m_isDeviceState : End Property
	Public Property Get DeviceStateEvent() : DeviceStateEvent = m_deviceStateDeviceType & "_" & m_deviceStateDeviceName & "_" & m_deviceStateDeviceAttr : End Property
		
	Public default Function init(input)
        m_raw = input
        Dim parsedInput : parsedInput = Glf_ParseInput(input)
		Dim playerState : playerState = Glf_CheckForGetPlayerState(input)
		m_isPlayerState = playerState(0)
		m_playerStateValue = playerState(1)
		m_playerStatePlayer = playerState(2)
		Dim deviceState : deviceState = Glf_CheckForDeviceState(input)
		m_isDeviceState = deviceState(0)
		m_deviceStateDeviceType = deviceState(1)
		m_deviceStateDeviceName = deviceState(2)
		m_deviceStateDeviceAttr = deviceState(3)
		
        m_value = parsedInput(0)
        m_isGetRef = parsedInput(2)
	    Set Init = Me
	End Function

End Class

Function Glf_FadeRGB(color1, color2, steps)
	If glf_cached_rgb_fades.Exists(color1&"_"&color2&"_"&CStr(steps)) Then
		Glf_FadeRGB = glf_cached_rgb_fades(color1&"_"&color2&"_"&CStr(steps))
		Exit Function
	End If

	Dim cached_rgb_seq : cached_rgb_seq = Array()
	Dim r1, g1, b1, r2, g2, b2, c1,c2
	Dim i
	Dim r, g, b
	c1 = clng( RGB( Glf_HexToInt(Left(color1, 2)), Glf_HexToInt(Mid(color1, 3, 2)), Glf_HexToInt(Right(color1, 2)))  )
	c2 = clng( RGB( Glf_HexToInt(Left(color2, 2)), Glf_HexToInt(Mid(color2, 3, 2)), Glf_HexToInt(Right(color2, 2)))  )
	
	r1 = c1 Mod 256
	g1 = (c1 \ 256) Mod 256
	b1 = (c1 \ (256 * 256)) Mod 256

	r2 = c2 Mod 256
	g2 = (c2 \ 256) Mod 256
	b2 = (c2 \ (256 * 256)) Mod 256

	ReDim cached_rgb_seq(steps - 1)
	Dim rgb_color
	For i = 0 To steps - 1
		r = r1 + (r2 - r1) * i / (steps - 1)
		g = g1 + (g2 - g1) * i / (steps - 1)
		b = b1 + (b2 - b1) * i / (steps - 1)
		rgb_color = Glf_RGBToHex(CInt(r), CInt(g), CInt(b))
		cached_rgb_seq(i) = rgb_color
	Next
	glf_cached_rgb_fades.Add color1&"_"&color2&"_"&CStr(steps), cached_rgb_seq	
	Glf_FadeRGB = cached_rgb_seq
End Function

Function Glf_RGBToHex(r, g, b)
	Glf_RGBToHex = Right("0" & Hex(r), 2) & _
	Right("0" & Hex(g), 2) & _
	Right("0" & Hex(b), 2)
End Function

Private Function Glf_HexToInt(hex)
	Glf_HexToInt = CInt("&H" & hex)
End Function

'******************************************************
'*****   GLF Shows 		                           ****
'******************************************************

Function CreateGlfShow(name)
	Dim show : Set show = (new GlfShow)(name)
	'msgbox name
	glf_shows.Add name, show
	Set CreateGlfShow = show
End Function

With CreateGlfShow("on")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100")
	End With
End With

With CreateGlfShow("off")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

With CreateGlfShow("flash")
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|000000")
	End With
End With



With CreateGlfShow("flash_color")
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|(color)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|000000")
	End With
End With

With CreateGlfShow("led_color")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|(color)")
	End With
End With

With CreateGlfShow("fade_led_color")
	With .AddStep(Null, Null, -1)
		.Lights = Array("(lights)|100|(color)|(fade)")
	End With
End With

With CreateGlfShow("fade_rgb_test")
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|ff0000|(fade)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|00ff00|(fade)")
	End With
	With .AddStep(Null, Null, 1)
		.Lights = Array("(lights)|100|0000ff|(fade)")
	End With
End With

With GlfShotProfiles("default")
	With .States("on")
		.Show = "flash"
	End With
	With .States("off")
		.Show = "off"
	End With
End With

With GlfShotProfiles("flash_color")
	With .States("off")
		.Show = "off"
	End With
	With .States("on")
		.Show = "flash_color"
	End With	
End With

Function AdjustHexColor(hexColor, percentage)
    ' Ensure percentage is between 0 and 100
    If percentage < 0 Then percentage = 0
    If percentage > 100 Then percentage = 100

    ' Parse the R, G, B components
    Dim r, g, b
    r = CLng("&H" & Mid(hexColor, 1, 2))
    g = CLng("&H" & Mid(hexColor, 3, 2))
    b = CLng("&H" & Mid(hexColor, 5, 2))

    ' Adjust the RGB components by the percentage
    r = Int(r * (percentage / 100))
    g = Int(g * (percentage / 100))
    b = Int(b * (percentage / 100))

    ' Ensure the values are within the valid range (0 to 255)
    r = FixRange(r)
    g = FixRange(g)
    b = FixRange(b)

    ' Convert back to hex and return the adjusted color
    AdjustHexColor = PadHex(Hex(r)) & PadHex(Hex(g)) & PadHex(Hex(b))
End Function

' Helper function to ensure values stay within range
Function FixRange(value)
    If value < 0 Then value = 0
    If value > 255 Then value = 255
    FixRange = value
End Function

' Helper function to pad single digit hex values with a leading zero
Function PadHex(hexValue)
    If Len(hexValue) < 2 Then
        PadHex = "0" & hexValue
    Else
        PadHex = hexValue
    End If
End Function


'******************************************************
'*****   GLF Pin Events                            ****
'******************************************************

Const GLF_GAME_START = "game_start"
Const GLF_GAME_STARTED = "game_started"
Const GLF_GAME_OVER = "game_ended"
Const GLF_BALL_WILL_END = "ball_will_end"
Const GLF_BALL_ENDING = "ball_ending"
Const GLF_BALL_ENDED = "ball_ended"
Const GLF_NEXT_PLAYER = "next_player"
Const GLF_BALL_DRAIN = "ball_drain"
Const GLF_BALL_STARTED = "ball_started"

'******************************************************
'*****   GLF Player State                          ****
'******************************************************

Const GLF_SCORE = "score"
Const GLF_CURRENT_BALL = "ball"
Const GLF_INITIALS = "initials"


Function EnableGlfBallSearch()
	Dim ball_search : Set ball_search = (new GlfBallSearch)()
    With CreateGlfMode("glf_ball_search", 100)
        .StartEvents = Array("reset_complete")

        With .TimedSwitches("flipper_cradle")
            .Switches = Array("s_left_flipper", "s_right_flipper")
            .Time = 3000
            .EventsWhenActive = Array("flipper_cradle")
            .EventsWhenReleased = Array("flipper_release")
        End With
    End With
    glf_ballsearch_enabled = True
	Set EnableGlfBallSearch = ball_search
End Function

Class GlfBallSearch

    Private m_debug
    Private m_timeout
    Private m_search_interval
    Private m_ball_search_wait_after_iteration
    Private m_phase
    Private m_devices
    Private m_current_device_type

    Public Property Get GetValue(value)
        'Select Case value
            'Case   
        'End Select
        GetValue = True
    End Property

    Public Property Let Timeout(value): Set m_timeout = CreateGlfInput(value): End Property
    Public Property Get Timeout(): Timeout = m_timeout.Value(): End Property

    Public Property Let SearchInterval(value): Set m_search_interval = CreateGlfInput(value): End Property
    Public Property Get SearchInterval(): SearchInterval = m_search_interval.Value(): End Property

    Public Property Let BallSearchWaitAfterIteration(value): Set m_ball_search_wait_after_iteration = CreateGlfInput(value): End Property
    Public Property Get BallSearchWaitAfterIteration(): BallSearchWaitAfterIteration = m_ball_search_wait_after_iteration.Value(): End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init()
        Set m_timeout = CreateGlfInput(15000)
        Set m_search_interval = CreateGlfInput(150)
        Set m_ball_search_wait_after_iteration = CreateGlfInput(10000)
        m_phase = 0
        m_devices = Array()
        m_current_device_type = Empty
        Set glf_ballsearch = Me
        SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
        AddPinEventListener "flipper_cradle", "ball_search_flipper_cradle", "BallSearchHandler", 30, Array("stop", Me)
        AddPinEventListener "flipper_release", "ball_search_flipper_cradle", "BallSearchHandler", 30, Array("reset", Me)
        Set Init = Me
    End Function

    Public Sub Start(phase)
        Dim ball_hold, held_balls
        held_balls = 0
        For Each ball_hold in glf_ball_holds.Items()
            held_balls = held_balls + ball_hold.GetValue("balls_held")
        Next
        If glf_gameStarted = True And glf_BIP > 0 And (glf_BIP-held_balls)>0 And glf_plunger.HasBall() = False Then
            m_phase = phase
            glf_last_switch_hit_time = 0
            'Fire all auto fire devices, slings, pops.
            m_devices = glf_autofiredevices.Items()
            m_current_device_type = "autofire"
            If UBound(m_devices) > -1 Then
                m_devices(0).BallSearch(m_phase)
                SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
            End If
        Else
            SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
        End If
    End Sub

    Public Sub NextDevice(device_index)
        If UBound(m_devices) > device_index Then
            m_devices(device_index+1).BallSearch(m_phase)
            SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, device_index+1), Null), m_search_interval.Value
        Else
            If m_current_device_type = "autofire" Then
                m_devices = glf_ball_devices.Items()
                m_current_device_type = "balldevices"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            ElseIf m_current_device_type = "balldevices" Then
                m_devices = glf_droptargets.Items()
                m_current_device_type = "droptargets"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            ElseIf m_current_device_type = "droptargets" Then
                m_devices = glf_diverters.Items()
                m_current_device_type = "diverters"
                If UBound(m_devices) > -1 Then
                    m_devices(0).BallSearch(m_phase)
                    SetDelay "ball_search_next_device" , "BallSearchHandler", Array(Array("next_device", Me, 0), Null), m_search_interval.Value
                End If
            Else
                m_current_device_type = Empty
                If m_phase < 3 Then
                    Start m_phase+1
                Else
                    m_phase = 0
                    SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
                End If
            End If
        End If
    End Sub

    Public Sub Reset()
        RemoveDelay "ball_search_next_device"
        m_phase = 0
        SetDelay "ball_search" , "BallSearchHandler", Array(Array("start", Me), Null), m_timeout.Value
    End Sub

    Public Sub StopBallSearch()
        RemoveDelay "ball_search_next_device"
        m_phase = 0
        RemoveDelay "ball_search"
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function BallSearchHandler(args)
    Dim ownProps, kwargs
    ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ball_search : Set ball_search = ownProps(1)
    Select Case evt
        Case "start"
            ball_search.Start 1
        Case "reset":
            ball_search.Reset
        Case "stop":
            ball_search.StopBallSearch
        Case "next_device"
            ball_search.NextDevice ownProps(2)
    End Select
End Function



'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************

Class GlfVpxBcpController

    Private m_bcpController, m_connected, m_mode_list

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_controller.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        useBcp = True
        m_mode_list = ""
        If Err Then MsgBox("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

	Public Sub Send(commandMessage)
		If m_connected = True Then
            m_bcpController.Send commandMessage
        End If
	End Sub

    Public Function GetMessages
		If m_connected Then
            GetMessages = m_bcpController.GetMessages
        End If
	End Function

    Public Sub Reset()
		If m_connected Then
            m_bcpController.Send "reset"
        End If
	End Sub
    
    Public Sub PlaySlide(slide, context, calling_context, priorty)
		If m_connected Then
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""" & slide & """: {""action"": ""play"", ""expire"": 0}}, ""context"": """ & context & """, ""calling_context"": """ & calling_context & """, ""priority"": " & priorty & "}"
        End If
	End Sub

    Public Sub ModeList()
        If m_connected Then
            If m_mode_list <> glf_running_modes Then
                m_bcpController.Send "mode_list?json={""running_modes"": ["&glf_running_modes&"]}"
                m_mode_list = glf_running_modes
            End If
        End If
    End Sub

    Public Sub SendPlayerVariable(name, value, prevValue)
		If m_connected Then
            m_bcpController.Send "player_variable?name=" & name & "&value=" & EncodeVariable(value) & "&prev_value=" & EncodeVariable(prevValue) & "&change=" & EncodeVariable(VariableVariance(value, prevValue)) & "&player_num=int:" & Getglf_currentPlayerNumber
            '06:34:34.644 : VERBOSE : BCP : Received BCP command: ball_start?player_num=int:1&ball=int:1
        End If
	End Sub

    Private Function EncodeVariable(value)
        Dim retValue
        Select Case VarType(value)
            Case vbInteger, vbLong
                retValue = "int:" & value
            Case vbSingle, vbDouble
                retValue = "float:" & value
            Case vbString
                retValue = "string:" & value
            Case vbBoolean
                retValue = "bool:" & CStr(value)
            Case Else
                retValue = "NoneType:"
        End Select
        EncodeVariable = retValue
    End Function
    
    Private Function VariableVariance(v1, v2)
        Dim retValue
        Select Case VarType(v1)
            Case vbInteger, vbLong, vbSingle, vbDouble
                retValue = Abs(v1 - v2)
            Case Else
                retValue = True 
        End Select
        VariableVariance = retValue
    End Function

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
            useBcp = False
        End If
    End Sub
End Class

Sub Glf_BcpSendPlayerVar(args)
    If useBcp=False Then
        Exit Sub
    End If
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim player_var : player_var = kwargs(0)
    Dim value : value = kwargs(1)
    Dim prevValue : prevValue = kwargs(2)
    bcpController.SendPlayerVariable player_var, value, prevValue
End Sub

Sub Glf_BcpAddPlayer(playerNum)
    If useBcp Then
        bcpController.Send("player_added?player_num=int:"&playerNum)
    End If
End Sub

Sub Glf_BcpUpdate()
    If useBcp=False Then
        Exit Sub
    End If
    Dim messages : messages = bcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter, eventName
        For Each message in messages
            'debug.print(message.Command)
            Select Case message.Command
                case "hello"
                    bcpController.Reset
                case "monitor_start"
                    Dim category : category = message.GetValue("category")
                    If category = "player_vars" Then
                        AddPlayerStateEventListener "score", "bcp_player_var_score_0", 0, "Glf_BcpSendPlayerVar", 1000, Null
                        AddPlayerStateEventListener "current_ball", "bcp_player_var_ball_0", 0, "Glf_BcpSendPlayerVar", 1000, Null
                    End If
                case "register_trigger"
                    eventName = message.GetValue("event")
            End Select
        Next
    End If
    bcpController.ModeList()
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************


'*****************************************************************************************************************************************
'  Glf Monitor Bcp Controller
'*****************************************************************************************************************************************

Class GlfMonitorBcpController

    Private m_bcpController, m_connected, m_isInMonitor

    Public default Function init(port, backboxCommand)
        On Error Resume Next
        Set m_bcpController = CreateObject("vpx_bcp_controller.VpxBcpController")
        m_bcpController.Connect port, backboxCommand
        m_connected = True
        If Err Then MsgBox("Can not start VPX BCP Controller") : m_connected = False
        Set Init = Me
	End Function

    Public Function IsInMonitior()
        If m_connected = True And m_isInMonitor = True Then
            IsInMonitior=True
        Else
            IsInMonitior=False
        End If
    End Function

	Public Sub Send(commandMessage)
		If m_connected = True Then
            m_bcpController.Send commandMessage
        End If
	End Sub

    Public Function GetMessages
		If m_connected Then
            GetMessages = m_bcpController.GetMessages
        End If
	End Function

    Public Sub Reset()
		If m_connected Then
            m_bcpController.Send "reset"
            m_bcpController.Send "trigger?json={""name"": ""slides_play"", ""settings"": {""monitor"": {""action"": ""play"", ""expire"": 0}}, ""context"": """", ""priority"": 1}"
            
            Dim mode
            For Each mode in glf_modes.Items()
                Glf_MonitorModeUpdate mode
            Next
            m_isInMonitor = True
        End If
	End Sub

    Public Sub Disconnect()
        If m_connected Then
            m_bcpController.Disconnect()
            m_connected = False
        End If
    End Sub
End Class

Sub Glf_MonitorModeUpdate(mode)
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If
    glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """&mode.Status&""", ""debug"": " & mode.IsDebug & "},"
    Dim config_item
    For Each config_item in mode.BallSavesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.CountersItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.TimersItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.MultiballLocksItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.MultiballsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ModeShots()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ShotGroupsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.BallHoldsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.SequenceShotsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ExtraBallsItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ComboSwitchesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.TimedSwitchesItems()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    For Each config_item in mode.ModeStateMachines()
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & config_item.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & config_item.Name & """},"
    Next
    If Not IsNull(mode.LightPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.LightPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.LightPlayer.Name & """},"
    End If
    If Not IsNull(mode.EventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.EventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.EventPlayer.Name & """},"
    End If
    If Not IsNull(mode.TiltConfig) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.TiltConfig.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.TiltConfig.Name & """},"
    End If
    If Not IsNull(mode.QueueEventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.QueueEventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.QueueEventPlayer.Name & """},"
    End If
    If Not IsNull(mode.QueueRelayPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.QueueRelayPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.QueueRelayPlayer.Name & """},"
    End If
    If Not IsNull(mode.RandomEventPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.RandomEventPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.RandomEventPlayer.Name & """},"
    End If
    If Not IsNull(mode.ShowPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.ShowPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.ShowPlayer.Name & """},"
    End If
    If Not IsNull(mode.SegmentDisplayPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.SegmentDisplayPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.SegmentDisplayPlayer.Name & """},"
    End If
    If Not IsNull(mode.VariablePlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.VariablePlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.VariablePlayer.Name & """},"
    End If
    If Not IsNull(mode.DOFPlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.DOFPlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.DOFPlayer.Name & """},"
    End If
    If Not IsNull(mode.SlidePlayer) Then
        glf_monitor_modes = glf_monitor_modes & "{""mode"": """&mode.Name&""", ""value"": """", ""debug"": " & mode.SlidePlayer.IsDebug & ", ""mode_device"": 1, ""mode_device_name"": """ & mode.SlidePlayer.Name & """},"
    End If
End Sub

Sub Glf_MonitorPlayerStateUpdate(key, value)
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If    
    glf_monitor_player_state = glf_monitor_player_state & "{""key"": """&key&""", ""value"": """&value&"""},"
End Sub

Sub Glf_MonitorEventStream(label, message)
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If
    glf_monitor_event_stream = glf_monitor_event_stream & "{""label"": """&label&""", ""message"": """&message&"""},"
End Sub


Sub Glf_MonitorBcpUpdate()
    If IsNull(glf_debugBcpController) Then
        Exit Sub
    End If

    'Send Updates
    If glf_debugBcpController.IsInMonitior Then
        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_player_state"",""changes"": ["&glf_monitor_player_state&"]}"
        glf_monitor_player_state = ""

        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_monitor_modes"",""changes"": ["&glf_monitor_modes&"]}"
        glf_monitor_modes = ""

        glf_debugBcpController.Send "glf_monitor?json={""name"": ""glf_event_stream"",""changes"": ["&glf_monitor_event_stream&"]}"
        glf_monitor_event_stream = ""
    End If
    

    Dim messages : messages = glf_debugBcpController.GetMessages()
    If IsEmpty(messages) Then
        Exit Sub
    End If
    If IsArray(messages) and UBound(messages)>-1 Then
        Dim message, parameters, parameter, eventName
        For Each message in messages
            debug.print(message.Command)
            Select Case message.Command
                case "hello"
                    glf_debugBcpController.Reset
                case "trigger"
                    eventName = message.GetValue("name")
                    Dim mode_name, device_name
                    debug.print eventName
                    If eventName = "glf_monitor_debug_mode" Then
                        mode_name = message.GetValue("mode")
                        debug.print mode_name
                        If Not IsNull(GlfModes(mode_name)) Then
                            debug.print("got mode")
                            If GlfModes(mode_name).IsDebug = 1 Then
                                debug.print("Turning off debug")
                                GlfModes(mode_name).Debug = False
                            Else
                                debug.print("Turning on debug")
                                GlfModes(mode_name).Debug = True
                            End If
                        End If
                    End If
                    If eventName = "glf_monitor_debug_mode_device" Then
                        mode_name = message.GetValue("mode")
                        device_name = message.GetValue("mode_device")
                        device_name = Replace(device_name, mode_name & "_", "")
                        debug.print mode_name
                        debug.print device_name
                        If Not IsNull(GlfModes(mode_name)) Then
                            debug.print("got mode")
                            Dim config_item, mode, is_debug
                            is_debug = message.GetValue("debug")
                            debug.print is_debug
                            If is_debug = "bool:true" Then
                                is_debug = True
                            Else
                                is_debug = False
                            End If
                            Set mode = GlfModes(mode_name)
                            For Each config_item in mode.BallSavesItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.CountersItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.TimersItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.MultiballLocksItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.MultiballsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ModeShots()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ShotGroupsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.BallHoldsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.SequenceShotsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ExtraBallsItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ComboSwitchesItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.TimedSwitchesItems()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            For Each config_item in mode.ModeStateMachines()
                                If config_item.Name = device_name Then : config_item.Debug = is_debug : End If
                            Next
                            If Not IsNull(mode.LightPlayer) Then
                                If mode.LightPlayer.Name = device_name Then : mode.LightPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.EventPlayer) Then
                                If mode.EventPlayer.Name = device_name Then : mode.EventPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.TiltConfig) Then
                                If mode.TiltConfig.Name = device_name Then : mode.TiltConfig.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.RandomEventPlayer) Then
                                If mode.RandomEventPlayer.Name = device_name Then : mode.RandomEventPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.ShowPlayer) Then
                                If mode.ShowPlayer.Name = device_name Then : mode.ShowPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.SegmentDisplayPlayer) Then
                                If mode.SegmentDisplayPlayer.Name = device_name Then : mode.SegmentDisplayPlayer.Debug = is_debug : End If
                            End If
                            If Not IsNull(mode.VariablePlayer) Then
                                If mode.VariablePlayer.Name = device_name Then : mode.VariablePlayer.Debug = is_debug : End If
                            End If
                        End If
                    End If
            End Select
        Next
    End If
End Sub

'*****************************************************************************************************************************************
'  Vpx Glf Bcp Controller
'*****************************************************************************************************************************************a
Class GlfAchievements

    Private m_name
    Private m_priority
    Private m_complete_events
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
        End Select
    End Property

    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let CompleteEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_complete_events.Add x, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property

    Public default Function Init(name, mode)
        m_name = "achievements_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_complete_events = CreateObject("Scripting.Dictionary")

        Set m_base_device = (new GlfBaseModeDevice)(mode, "achievement", Me)
        glf_achievements.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim key
        For Each key in m_complete_events.Keys
            AddPinEventListener m_complete_events(key).EventName, m_name & "_complete_event_" & key, "AchievementsEventHandler", m_priority+m_complete_events(key).Priority, Array("complete", Me)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim key
        For Each key in m_complete_events.Keys
            RemovePinEventListener m_complete_events(key).EventName, m_name & "_complete_event_" & key
        Next
    End Sub

    Public Sub Complete()
        'TODO: Implement Complete Events
    End Sub

    Private Sub Log(message)
        If m_debug Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function AchievementsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1)
    End If

    Dim evt : evt = ownProps(0)
    Dim achievement : Set achievement = ownProps(1)

    Select Case evt
        Case "complete"
            achievement.Complete
    End Select

    If IsObject(args(1)) Then
        Set AchievementsHandler = kwargs
    Else
        AchievementsHandler = kwargs
    End If
End Function



Class GlfBallHold

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_enabled
    Private m_balls_to_hold
    Private m_hold_devices
    Private m_balls_held
    Private m_hold_queue
    Private m_release_all_events
    Private m_release_one_events
    Private m_release_one_if_full_events

    Public Property Get Name() : Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "balls_held":
                GetValue = m_balls_held
        End Select
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    
    Public Property Get BallsToHold() : BallsToHold = m_balls_to_hold : End Property
    Public Property Let BallsToHold(value) : m_balls_to_hold = value : End Property

    Public Property Get HoldDevices() : HoldDevices = m_hold_devices : End Property
    Public Property Let HoldDevices(value) : m_hold_devices = value : End Property

    Public Property Get ReleaseAllEvents(): Set ReleaseAllEvents = m_release_all_events: End Property
    Public Property Let ReleaseAllEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_all_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get ReleaseOneEvents(): Set ReleaseOneEvents = m_release_one_events: End Property
    Public Property Let ReleaseOneEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_one_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get ReleaseOneIfFullEvents(): Set ReleaseOneIfFullEvents = m_release_one_if_full_events: End Property
    Public Property Let ReleaseOneIfFullEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_release_one_if_full_events.Add newEvent.Name, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = "ball_hold_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_balls_to_hold = 0
        m_balls_held = 0
        m_hold_devices = Array()
        Set m_hold_queue = CreateObject("Scripting.Dictionary")
        Set m_release_all_events = CreateObject("Scripting.Dictionary")
        Set m_release_one_events = CreateObject("Scripting.Dictionary")
        Set m_release_one_if_full_events = CreateObject("Scripting.Dictionary")
        ReleaseAllEvents = Array("tilt")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "ball_hold", Me)
        glf_ball_holds.Add name, Me
        Set Init = Me
	End Function

    Public Sub Activate()
        If UBound(m_base_device.EnableEvents.Keys()) = -1 Then
            Enable()
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        m_enabled = True
        'Add Event Listeners
        Dim device
        For Each device in m_hold_devices
            AddPinEventListener "balldevice_" & device & "_ball_enter", m_mode & "_" & name & "_hold", "BallHoldsEventHandler", m_priority, Array("hold", me, device)
        Next
        Dim evt
        For Each evt in m_release_all_events.Keys
            AddPinEventListener m_release_all_events(evt).EventName, m_mode & "_" & name & "_release_all", "BallHoldsEventHandler", m_priority+m_release_all_events(evt).Priority, Array("release_all", me, m_release_all_events(evt))
        Next
        For Each evt in m_release_one_events.Keys
            AddPinEventListener m_release_one_events(evt).EventName, m_mode & "_" & name & "_release_one", "BallHoldsEventHandler", m_priority+m_release_one_events(evt).Priority, Array("release_one", me, m_release_one_events(evt))
        Next
        For Each evt in m_release_one_if_full_events.Keys
            AddPinEventListener m_release_one_if_full_events(evt).EventName, m_mode & "_" & name & "_release_one_if_full", "BallHoldsEventHandler", m_priority+m_release_one_if_full_events(evt).Priority, Array("release_one_if_full", me, m_release_one_if_full_events(evt))
        Next
    End Sub

    Public Sub Disable()
        m_enabled = False
        ReleaseAll()
        Dim device
        For Each device in m_hold_devices
            RemovePinEventListener "balldevice_" & device & "_ball_enter", m_mode & "_" & name & "_hold"
        Next
        Dim evt
        For Each evt in m_release_all_events.Keys
            RemovePinEventListener m_release_all_events(evt).EventName, m_mode & "_" & name & "_release_all"
        Next
        For Each evt in m_release_one_events.Keys
            RemovePinEventListener m_release_one_events(evt).EventName, m_mode & "_" & name & "_release_one"
        Next
        For Each evt in m_release_one_if_full_events.Keys
            RemovePinEventListener m_release_one_if_full_events(evt).EventName, m_mode & "_" & name & "_release_one_if_full"
        Next
    End Sub

    Public Function IsFull()
        'Return true if hold is full
        If RemainingSpaceInHold() = 0 Then
            IsFull = True
        Else
            IsFull = False
        End If
    End Function

    Public Function RemainingSpaceInHold()
        'Return the remaining capacity of the hold.
        Dim balls
        balls = m_balls_to_hold - m_balls_held
        If balls < 0 Then
            balls = 0
        End If
        RemainingSpaceInHold = balls
    End Function

    Public Function HoldBall(device, unclaimed_balls)        
        ' Handle result of _ball_enter event of hold_devices.
        If IsFull() Then
            Log "Cannot hold balls. Hold is full."
            HoldBall = unclaimed_balls
            Exit Function
        End If

        If unclaimed_balls <= 0 Then
            HoldBalls = unclaimed_balls
            Exit Function
        End If

        Dim capacity : capacity = RemainingSpaceInHold()
        Dim balls_to_hold
        If unclaimed_balls > capacity Then
            balls_to_hold = capacity
        Else
            balls_to_hold = unclaimed_balls
        End If
        m_balls_held = m_balls_held + balls_to_hold
        Log "Held " & balls_to_hold & " balls"

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls_held", balls_to_hold
            .Add "total_balls_held", m_balls_held
        End With
        DispatchPinEvent m_name & "_held_ball", kwargs

        'check if we are full now and post event if yes
        If IsFull() Then
            Set kwargs = GlfKwargs()
            With kwargs
                .Add "balls", m_balls_held
            End With
            DispatchPinEvent m_name & "_full", kwargs
        End If

        m_hold_queue.Add device, unclaimed_balls

        HoldBall = unclaimed_balls - balls_to_hold
    End Function

    Public Function ReleaseAll()
        'Release all balls in hold.
        ReleaseAll = ReleaseBalls(m_balls_held)
    End Function

    Public Function ReleaseBalls(balls_to_release)
        'Release all balls and return the actual amount of balls released.
        '
        'Args:
        '----
        '    balls_to_release: number of ball to release from hold
        
        If Ubound(m_hold_queue.Keys()) = -1 Then
            ReleaseBalls = 0
            Exit Function
        End If

        Dim remaining_balls_to_release : remaining_balls_to_release = balls_to_release

        Log "Releasing up to " & balls_to_release & " balls from hold"
        Dim balls_released : balls_released = 0
        Do While Ubound(m_hold_queue.Keys()) > -1
            Dim keys : keys = m_hold_queue.Keys()
            Dim device, balls_held
            device = keys(0)
            balls_held = m_hold_queue(device)
            m_hold_queue.Remove device

            Dim deviceControl : Set deviceControl = glf_ball_devices(device)
            
            Dim balls : balls = balls_held
            Dim balls_in_device : balls_in_device = deviceControl.Balls
            If balls > balls_in_device Then
                balls = balls_in_device
            End If

            If balls > remaining_balls_to_release Then
                m_hold_queue.Add device, balls_held - remaining_balls_to_release
                balls = remaining_balls_to_release
            End If

            deviceControl.EjectBalls balls
            balls_released = balls_released + balls
            remaining_balls_to_release = remaining_balls_to_release - balls
            If remaining_balls_to_release <= 0 Then
               Exit Do
            End If
        Loop

        If balls_released > 0 Then
            Dim kwargs : Set kwargs = GlfKwargs()
            With kwargs
                .Add "balls_released", balls_released
            End With
            DispatchPinEvent m_name & "_balls_released", kwargs
        End If

        m_balls_held = m_balls_held - balls_released
        ReleaseBalls = balls_released
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt, x
        yaml = "  " & Replace(m_name, "ball_hold_", "") & ":" & vbCrLf
        If UBound(m_base_device.EnableEvents().Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents().keys
                yaml = yaml & m_base_device.EnableEvents()(key).Raw
                If x <> UBound(m_base_device.EnableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_base_device.DisableEvents().Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents().keys
                yaml = yaml & m_base_device.DisableEvents()(key).Raw
                If x <> UBound(m_base_device.DisableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        yaml = yaml & "    hold_devices: " & Join(m_hold_devices, ",") & vbCrLf
        If m_balls_to_hold > 0 Then
            yaml = yaml & "    balls_to_hold: " & m_balls_to_hold & vbCrLf
        End If
        If UBound(m_release_all_events.Keys) > -1 Then
            yaml = yaml & "    release_all_events: "
            x=0
            For Each key in m_release_all_events.keys
                yaml = yaml & m_release_all_events(key).Raw
                If x <> UBound(m_release_all_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_release_one_events.Keys) > -1 Then
            yaml = yaml & "    release_one_events: "
            x=0
            For Each key in m_release_one_events.keys
                yaml = yaml & m_release_one_events(key).Raw
                If x <> UBound(m_release_one_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_release_one_if_full_events.Keys) > -1 Then
            yaml = yaml & "    release_one_if_full_events: "
            x=0
            For Each key in m_release_one_if_full_events.keys
                yaml = yaml & m_release_one_if_full_events(key).Raw
                If x <> UBound(m_release_one_if_full_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function BallHoldsEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ball_hold : Set ball_hold = ownProps(1)
    Dim glfEvent
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If

    Select Case evt
        Case "hold"
            kwargs = ball_hold.HoldBall(ownProps(2), kwargs)
        Case "release_all"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            ball_hold.ReleaseAll
        Case "release_one"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            ball_hold.ReleaseBalls 1
        Case "release_one_if_full"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            If ball_hold.IsFull Then
                ball_hold.ReleaseBalls 1
            End If
    End Select

    If IsObject(args(1)) Then
        Set BallHoldsEventHandler = kwargs
    Else
        BallHoldsEventHandler = kwargs
    End If
End Function
Class BallSave

    Private m_name
    Private m_mode
    Private m_priority
    Private m_active_time
    Private m_grace_period
    Private m_enable_events
    Private m_timer_start_events
    Private m_auto_launch
    Private m_balls_to_save
    Private m_saving_balls
    Private m_enabled
    Private m_timer_started
    Private m_tick
    Private m_in_grace
    Private m_in_hurry_up
    Private m_hurry_up_time
    private m_base_device
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AutoLaunch(): AutoLaunch = m_auto_launch: End Property
    Public Property Let ActiveTime(value) : m_active_time = Glf_ParseInput(value) : End Property
    Public Property Let GracePeriod(value) : m_grace_period = Glf_ParseInput(value) : End Property
    Public Property Let HurryUpTime(value) : m_hurry_up_time = Glf_ParseInput(value) : End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property

    Public Property Let TimerStartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_timer_start_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let AutoLaunch(value) : m_auto_launch = value : End Property
    Public Property Let BallsToSave(value) : m_balls_to_save = value : End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property

	Public default Function init(name, mode)
        m_name = "ball_save_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_active_time = Null
	    m_grace_period = Null
        m_hurry_up_time = Null
        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_timer_start_events = CreateObject("Scripting.Dictionary")
	    m_auto_launch = False
	    m_balls_to_save = 1
        m_enabled = False
        m_timer_started = False
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "ball_save", Me)
	    Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_timer_start_events.Keys
            AddPinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start", "BallSaveEventHandler", m_priority+m_timer_start_events(evt).Priority, Array("timer_start", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_timer_start_events.Keys
            RemovePinEventListener m_timer_start_events(evt).EventName, m_name & "_timer_start"
        Next
    End Sub

    Public Sub Enable()
        If m_enabled = True Then
            Exit Sub
        End If
        m_enabled = True
        m_saving_balls = m_balls_to_save
        Log "Enabling. Auto launch: "&m_auto_launch&", Balls to save: "&m_balls_to_save
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "BallSaveEventHandler", 1000, Array("drain", Me)
        DispatchPinEvent m_name&"_enabled", Null
        If UBound(m_timer_start_events.Keys) = -1 Then
            Log "Timer Starting as no timer start events are set"
            TimerStart()
        End If
    End Sub

    Public Sub Disable
        'Disable ball save
        If m_enabled = False Then
            Exit Sub
        End If
        m_enabled = False
        m_saving_balls = m_balls_to_save
        m_timer_started = False
        Log "Disabling..."
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        RemoveDelay "_ball_save_"&m_name&"_disable"
        RemoveDelay m_name&"_grace_period"
        RemoveDelay m_name&"_hurry_up_time"
            
    End Sub

    Sub Drain(ballsToSave)
        If m_enabled = True And ballsToSave > 0 Then
            If m_saving_balls > 0 Then
                m_saving_balls = m_saving_balls -1
            End If
            Log "Ball(s) drained while active. Requesting new one(s). Auto launch: "& m_auto_launch
            DispatchPinEvent m_name&"_saving_ball", Null
            SetDelay m_name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", Me),Null), 1000
            If m_saving_balls = 0 Then
                Disable()
            End If
        End If
    End Sub

    Public Sub TimerStart
        'Start the timer.
        'This is usually called after the ball was ejected while the ball save may have been enabled earlier.
        If m_timer_started=True Or m_enabled=False Then
            Exit Sub
        End If
        m_timer_started=True
        DispatchPinEvent m_name&"_timer_start", Null
        If Not IsNull(m_active_time) Then
            Dim active_time : active_time = GetRef(m_active_time(0))()
            Dim grace_period, hurry_up_time
            If Not IsNull(m_grace_period) Then
                grace_period = GetRef(m_grace_period(0))()
            Else
                grace_period = 0
            End If
            If Not IsNull(m_hurry_up_time) Then
                hurry_up_time = GetRef(m_hurry_up_time(0))()
            Else
                hurry_up_time = 0
            End If
            Log "Starting ball save timer: " & active_time
            Log "gametime: "& gametime & ". disabled at: " & gametime+active_time+grace_period
            SetDelay m_name&"_disable", "BallSaveEventHandler" , Array(Array("disable", Me),Null), active_time+grace_period
            SetDelay m_name&"_grace_period", "BallSaveEventHandler", Array(Array("grace_period", Me),Null), active_time
            SetDelay m_name&"_hurry_up_time", "BallSaveEventHandler", Array(Array("hurry_up_time", Me), Null), active_time-hurry_up_time
        End If
    End Sub

    Public Sub EnterGracePeriod
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Sub EnterHurryUpTime
        DispatchPinEvent m_name & "_hurry_up", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & Replace(m_name, "ball_save_", "") & ":" & vbCrLf
        yaml = yaml & "    active_time: " & m_active_time(1) & "s" & vbCrLf
        yaml = yaml & "    grace_period: " & m_grace_period(1) & "s" & vbCrLf
        yaml = yaml & "    hurry_up_time: " & m_hurry_up_time(1) & "s" & vbCrLf
        yaml = yaml & "    enable_events: "
        Dim evt,x : x = 0
        For Each evt in m_enable_events.Keys
            yaml = yaml & m_enable_events(evt).Raw
            If x <> UBound(m_enable_events.Keys) Then
                yaml = yaml & ", "
            End If
            x = x +1
        Next
        yaml = yaml & vbCrLf
        yaml = yaml & "    timer_start_events: "
        x=0
        For Each evt in m_timer_start_events.Keys
            yaml = yaml & m_timer_start_events(evt).Raw
            If x <> UBound(m_timer_start_events.Keys) Then
                yaml = yaml & ", "
            End If
            x = x +1
        Next
        yaml = yaml & vbCrLf
        yaml = yaml & "    auto_launch: " & LCase(m_auto_launch) & vbCrLf
        yaml = yaml & "    balls_to_save: " & m_balls_to_save & vbCrLf
        ToYaml = yaml
    End Function
End Class

Function BallSaveEventHandler(args)
    Dim ownProps, ballsToSave : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    ballsToSave = args(1) 
    Dim ballSave : Set ballSave = ownProps(1)
    Select Case evt
        Case "activate"
            ballSave.Activate
        Case "deactivate"
            ballSave.Deactivate
        Case "enable"
            ballSave.Enable ownProps(2)
        Case "disable"
            ballSave.Disable
        Case "grace_period"
            ballSave.EnterGracePeriod
        Case "hurry_up_time"
            ballSave.EnterHurryUpTime
        Case "drain"
            If ballsToSave > 0 Then
                ballSave.Drain ballsToSave
                ballsToSave = ballsToSave - 1
            End If
        Case "timer_start"
            ballSave.TimerStart
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True  And glf_plunger.IncomingBalls = 0  Then
                Glf_ReleaseBall(Null)
                If ballSave.AutoLaunch = True Then
                    SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave),Null), 500
                End If
            Else
                SetDelay ballSave.Name&"_queued_release", "BallSaveEventHandler" , Array(Array("queue_release", ballSave), Null), 1000
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay ballSave.Name&"_auto_launch", "BallSaveEventHandler" , Array(Array("auto_launch", ballSave), Null), 500
            End If
    End Select
    BallSaveEventHandler = ballsToSave
End Function


Class GlfComboSwitches

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_switch_1
    Private m_switch_2
    Private m_events_when_both
    Private m_events_when_inactive
    Private m_events_when_one
    Private m_events_when_switch_1
    Private m_events_when_switch_2
    Private m_hold_time
    Private m_max_offset_time
    Private m_release_time

    Private m_switch_1_active
    Private m_switch_2_active

    Private m_switch_state

    Public Property Get Name() : Name = m_name : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get GetValue(value)
        Select Case value
            'Case ""
            '    GetValue = m_ticks
        End Select
    End Property

    Public Property Let Switch1(value): m_switch_1 = value: End Property
    Public Property Let Switch2(value): m_switch_2 = value: End Property
    Public Property Let HoldTime(value): Set m_hold_time = CreateGlfInput(value): End Property
    Public Property Let MaxOffsetTime(value): Set m_max_offset_time = CreateGlfInput(value): End Property
    Public Property Let ReleaseTime(value): Set m_release_time = CreateGlfInput(value): End Property
    Public Property Let EventsWhenBoth(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_both.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenInactive(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_inactive.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenOne(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_one.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenSwitch1(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_switch_1.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenSwitch2(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_switch_2.Add newEvent.Raw, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = "combo_switch_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
    
        m_switch_1 = Empty
        m_switch_2 = Empty
        Set m_events_when_both = CreateObject("Scripting.Dictionary")
        Set m_events_when_inactive = CreateObject("Scripting.Dictionary")
        Set m_events_when_one = CreateObject("Scripting.Dictionary")
        Set m_events_when_switch_1 = CreateObject("Scripting.Dictionary")
        Set m_events_when_switch_2 = CreateObject("Scripting.Dictionary")
        Set m_hold_time = CreateGlfInput(0)
        Set m_max_offset_time = CreateGlfInput(-1)
        Set m_release_time = CreateGlfInput(0)

        m_switch_1_active = 0
        m_switch_2_active = 0

        m_switch_state = Empty
        Set m_base_device = (new GlfBaseModeDevice)(mode, "combo_switch", Me)

        glf_combo_switches.Add name, Me

        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating Combo Switch"
        AddPinEventListener m_switch_1 & "_active" , m_name & "_switch1_active" , "ComboSwitchEventHandler", m_priority, Array("switch1_active", Me, m_switch_1)
        AddPinEventListener m_switch_1 & "_inactive" , m_name & "_switch1_inactive" , "ComboSwitchEventHandler", m_priority, Array("switch1_inactive", Me, m_switch_1)
        AddPinEventListener m_switch_2 & "_active" , m_name & "_switch2_active" , "ComboSwitchEventHandler", m_priority, Array("switch2_active", Me, m_switch_2)
        AddPinEventListener m_switch_2 & "_inactive" , m_name & "_switch2_inactive" , "ComboSwitchEventHandler", m_priority, Array("switch2_inactive", Me, m_switch_2)
    End Sub

    Public Sub Deactivate()
        Log "Deactivating Combo Switch"
        RemovePinEventListener m_switch_1 & "_active" , m_name & "_switch1_active"
        RemovePinEventListener m_switch_1 & "_inactive" , m_name & "_switch1_inactive"
        RemovePinEventListener m_switch_2 & "_active" , m_name & "_switch2_active"
        RemovePinEventListener m_switch_2 & "_inactive" , m_name & "_switch2_inactive"
        RemoveDelay m_name & "_" & "switch_1_inactive"
        RemoveDelay m_name & "_" & "switch_2_inactive"
        RemoveDelay m_name & "_" & "switch_1_active"
        RemoveDelay m_name & "_" & "switch_1_active"
        RemoveDelay m_name & "_" & "switch_2_only"
        RemoveDelay m_name & "_" & "switch_1_only"
    End Sub

    Public Sub Switch1WentActive(switch_name)
        Log "switch_1 just went active"
        RemoveDelay m_name & "_" & "switch_1_inactive"

        If m_switch_1_active > 0 Then
            Exit Sub
        End If

        If m_hold_time.Value() = 0 Then
            ActivateSwitches1 switch_name
        Else
            SetDelay m_name & "_switch_1_active", "ComboSwitchEventHandler", Array(Array("activate_switch1", Me, switch_name),Null), m_hold_time.Value()
        End If
    End Sub

    Public Sub Switch2WentActive(switch_name)
        Log "switch_2 just went active"
        RemoveDelay m_name & "_" & "switch_2_inactive"

        If m_switch_2_active > 0 Then
            Exit Sub
        End If

        If m_hold_time.Value() = 0 Then
            ActivateSwitches2 switch_name
        Else
            SetDelay m_name & "_switch_2_active", "ComboSwitchEventHandler", Array(Array("activate_switch2", Me, switch_name),Null), m_hold_time.Value()
        End If
    End Sub

    Public Sub Switch1WentInactive(switch_name)
        Log "switch_1 just went inactive"
        RemoveDelay m_name & "_" & "switch_1_active"

        If m_release_time.Value() = 0 Then
            ReleaseSwitch1 switch_name
        Else
            SetDelay m_name & "_switch_1_inactive", "ComboSwitchEventHandler", Array(Array("release_switch1", Me, switch_name),Null), m_release_time.Value()
        End If
    End Sub

    Public Sub Switch2WentInactive(switch_name)
        Log "switch_2 just went inactive"
        RemoveDelay m_name & "_" & "switch_2_active"

        If m_release_time.Value() = 0 Then
            ReleaseSwitch2 switch_name
        Else
            SetDelay m_name & "_switch_2_inactive", "ComboSwitchEventHandler", Array(Array("release_switch2", Me, switch_name),Null), m_release_time.Value()
        End If
    End Sub

    Public Sub ActivateSwitches1(switch_name)
        Log "Switch_1 has passed the hold time and is now active"
        m_switch_1_active = gametime
        RemoveDelay m_name & "_" & "switch_2_only"

        If m_switch_2_active > 0 Then
            If m_max_offset_time.Value() >= 0 And (m_switch_1_active - m_switch_2_active) > m_max_offset_time.Value() Then
                Log "Switches_2 is active, but the max_offset_time=" & m_max_offset_time.Value() & " which is larger than when a Switches_1 switch was first activated, so the state will not switch to both"
                Exit Sub
            End If
            PostSwitchStateEvents "both"
        ElseIf m_max_offset_time.Value()>=0 Then
            SetDelay m_name & "_switch_1_only", "ComboSwitchEventHandler", Array(Array("switch_1_only", Me, switch_name),Null), max_offset_time.Value()
        End If
    End Sub

    Public Sub ActivateSwitches2(switch_name)
        Log "Switch_2 has passed the hold time and is now active"
        m_switch_2_active = gametime
        RemoveDelay m_name & "_" & "switch_1_only"

        If m_switch_1_active > 0 Then
            If m_max_offset_time.Value() >= 0 And (m_switch_2_active - m_switch_1_active) > m_max_offset_time.Value() Then
                Log "Switches_1 is active, but the max_offset_time=" & m_max_offset_time.Value() & " which is larger than when a Switches_2 switch was first activated, so the state will not switch to both"
                Exit Sub
            End If
            PostSwitchStateEvents "both"
        ElseIf m_max_offset_time.Value()>=0 Then
            SetDelay m_name & "_switch_2_only", "ComboSwitchEventHandler", Array(Array("switch_2_only", Me, switch_name),Null), max_offset_time.Value()
        End If
    End Sub

    Public Sub ReleaseSwitch1(switch_name)
        Log "Switches_1 has passed the release time and is now released"
        m_switch_1_active = 0

        If m_switch_2_active > 0 And m_switch_state = "both" Then
            PostSwitchStateEvents "one"
        ElseIf m_switch_state = "one" Then
            PostSwitchStateEvents "inactive"
        End If
    End Sub

    Public Sub ReleaseSwitch2(switch_name)
        Log "Switches_2 has passed the release time and is now released"
        m_switch_2_active = 0

        If m_switch_1_active > 0 And m_switch_state = "both" Then
            PostSwitchStateEvents "one"
        ElseIf m_switch_state = "one" Then
            PostSwitchStateEvents "inactive"
        End If
    End Sub

    Public Sub PostSwitchStateEvents(state)
        If m_switch_state = state Then
            Exit Sub
        End If
        m_switch_state = state
        Log "New State " & state

        Dim evt
        Select Case state
            Case "both"
                For Each evt in m_events_when_both.Keys
                    If m_events_when_both(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_both(evt).EventName, Null
                    End If
                Next
            Case "one"
                For Each evt in m_events_when_one.Keys
                    If m_events_when_one(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_one(evt).EventName, Null
                    End If
                Next
            Case "inactive"
                For Each evt in m_events_when_inactive.Keys
                    If m_events_when_inactive(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_inactive(evt).EventName, Null
                    End If
                Next
            Case "switches_1"
                For Each evt in m_events_when_switch_1.Keys
                    If m_events_when_switch_1(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_switch_1(evt).EventName, Null
                    End If
                Next
            Case "switches_2"
                For Each evt in m_events_when_switch_2.Keys
                    If m_events_when_switch_2(evt).Evaluate() Then
                        DispatchPinEvent m_events_when_switch_2(evt).EventName, Null
                    End If
                Next
        End Select

    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function ComboSwitchEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim combo_switch : Set combo_switch = ownProps(1)
    Dim switch_name : switch_name = ownProps(2)
    Select Case evt
        Case "switch1_active"
            combo_switch.Switch1WentActive switch_name
        Case "switch2_active"
            combo_switch.Switch2WentActive switch_name
        Case "switch1_inactive"
            combo_switch.Switch1WentInactive switch_name
        Case "switch2_inactive"
            combo_switch.Switch2WentInactive switch_name
        Case "activate_switch1"
            combo_switch.ActivateSwitches1 switch_name
        Case "activate_switch2"
            combo_switch.ActivateSwitches2 switch_name
        Case "release_switch1"
            combo_switch.ReleaseSwitch1 switch_name
        Case "release_switch2"
            combo_switch.ReleaseSwitch2 switch_name
        Case "switch_1_only"
            combo_switch.PostSwitchStateEvents "switches_1"
        Case "switch_2_only"
            combo_switch.PostSwitchStateEvents "switches_2"
    End Select
    If IsObject(args(1)) Then
        Set ComboSwitchEventHandler = kwargs
    Else
        ComboSwitchEventHandler = kwargs
    End If
End Function


Class GlfCounter

    Private m_name
    Private m_priority
    Private m_mode
    Private m_enable_events
    Private m_count_events
    Private m_count_complete_value
    Private m_disable_on_complete
    Private m_reset_on_complete
    Private m_events_when_complete
    Private m_persist_state
    Private m_debug

    Private m_count

    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CountEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_count_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CountCompleteValue(value) : m_count_complete_value = value : End Property
    Public Property Let DisableOnComplete(value) : m_disable_on_complete = value : End Property
    Public Property Let ResetOnComplete(value) : m_reset_on_complete = value : End Property
    Public Property Let EventsWhenComplete(value) : m_events_when_complete = value : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(name, mode)
        m_name = "counter_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_count = -1
        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_count_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode & "_starting", m_name & "_activate", "CounterEventHandler", m_priority, Array("activate", Me)
        AddPinEventListener m_mode & "_stopping", m_name & "_deactivate", "CounterEventHandler", m_priority, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub SetValue(value)
        If value = "" Then
            value = 0
        End If
        m_count = value
        If m_persist_state Then
            SetPlayerState m_name & "_state", m_count
        End If
    End Sub

    Public Sub Activate()
        If m_persist_state And m_count > -1 Then
            If GetPlayerState(m_name & "_state")=False Then
                SetValue 0
            Else
                SetValue GetPlayerState(m_name & "_state")
            End If
        Else
            SetValue 0
        End If
        Dim evt
        For Each evt in m_enable_events.Keys
            AddPinEventListener m_enable_events(evt).EventName, m_name & "_enable", "CounterEventHandler", m_priority+m_enable_events(evt).Priority, Array("enable", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Disable()
        If Not m_persist_state Then
            SetValue -1
        End If
        Dim evt
        For Each evt in m_enable_events.Keys
            RemovePinEventListener m_enable_events(evt).EventName, m_name & "_enable"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_count_events.Keys
            AddPinEventListener m_count_events(evt).EventName, m_name & "_count", "CounterEventHandler", m_priority+m_count_events(evt).Priority, Array("count", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_count_events.Keys
            RemovePinEventListener m_count_events(evt).EventName, m_name & "_count"
        Next
    End Sub

    Public Sub Count()
        Log "counting: old value: "& m_count & ", new Value: " & m_count+1 & ", target: "& m_count_complete_value
        SetValue m_count + 1
        If m_count = m_count_complete_value Then
            Dim evt
            For Each evt in m_events_when_complete
                DispatchPinEvent evt, Null
            Next
            If m_disable_on_complete Then
                Disable()
            End If
            If m_reset_on_complete Then
                SetValue 0
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function CounterEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0) : kwargs = args(1) 
    Dim evt : evt = ownProps(0)
    Dim counter : Set counter = ownProps(1)
    Select Case evt
        Case "activate"
            counter.Activate
        Case "deactivate"
            counter.Deactivate
        Case "enable"
            counter.Enable
        Case "count"
            counter.Count
    End Select
    CounterEventHandler = kwargs
End Function


Class GlfDofPlayer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "dof_player" : End Property


    Public Property Get EventDOF() : EventDOF = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_dof : Set new_dof = (new GlfDofPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_dof
        Set EventName = new_dof
        
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "dof_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "dof_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_dof_player_play", "DofPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_dof_player_play"
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            Log "Firing DOF Event: " & m_eventValues(evt).DOFEvent & " State: " & m_eventValues(evt).Action
            DOF m_eventValues(evt).DOFEvent, m_eventValues(evt).Action  
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_dof_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function DofPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim dof_player : Set dof_player = ownProps(1)
    Select Case evt
        Case "activate"
            dof_player.Activate
        Case "deactivate"
            dof_player.Deactivate
        Case "play"
            dof_player.Play(ownProps(2))
    End Select
    If IsObject(args(1)) Then
        Set DofPlayerEventHandler = kwargs
    Else
        DofPlayerEventHandler = kwargs
    End If
End Function

Class GlfDofPlayerItem
	Private m_dof_event, m_action
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        Select Case input
            Case "DOF_OFF"
                m_action = 0
            Case "DOF_ON"
                m_action = 1
            Case "DOF_PULSE"
                m_action = 2
        End Select
    End Property

    Public Property Get DOFEvent(): DOFEvent = m_dof_event: End Property
    Public Property Let DOFEvent(input): m_dof_event = CInt(input): End Property

	Public default Function init()
        m_action = Empty
        m_dof_event = Empty
        Set Init = Me
	End Function

End Class



Class GlfEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "event_player" : End Property

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Name, newEvent
        'msgbox newEvent.Name
        m_eventValues.Add newEvent.Name, value  
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            Log "Adding Event Listener for: " & m_events(evt).EventName
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_event_player_play", "EventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        Log "Dispatching Event: " & evt
        If Not IsNull(m_events(evt).Condition) Then
            'msgbox m_events(evt).Condition
            If GetRef(m_events(evt).Condition)() = False Then
                Exit Sub
            End If
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            Log "Dispatching Event: " & evtValue
            DispatchPinEvent evtValue, Null
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_event_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & m_events(key).Raw & ": " & vbCrLf
                For Each evt in m_eventValues(key)
                    yaml = yaml & "    - " & evt & vbCrLf
                Next
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function EventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            eventPlayer.FireEvent ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set EventPlayerEventHandler = kwargs
    Else
        EventPlayerEventHandler = kwargs
    End If
End Function

Class GlfExtraBall

    Private m_name
    private m_command_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_award_events
    Private m_max_per_game

    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
    Public Property Get GetValue(value)
        Select Case value
            'Case "":
            '    GetValue = 
        End Select
    End Property

    Public Property Let AwardEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_award_events.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Let MaxPerGame(value) : Set m_max_per_game = CreateGlfInput(value) : End Property

	Public default Function init(name, mode)
        m_name = "extra_ball_" & name
        m_command_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        Set m_award_events = CreateObject("Scripting.Dictionary")
        Set m_max_per_game = CreateGlfInput(0)

        Glf_SetInitialPlayerVar m_name & "_awarded", 0

        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "extra_ball", Me)
        
        Set Init = Me
	End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_award_events.Keys
            AddPinEventListener m_award_events(evt).EventName, m_name & "_" & evt & "_award", "ExtraBallsHandler", m_priority+m_award_events(evt).Priority, Array("award", Me, m_award_events(evt))
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_award_events.Keys
            RemovePinEventListener m_award_events(evt).EventName, m_name & "_" & evt & "_award"
        Next
    End Sub

    Public Sub Award(evt)
        If evt.Evaluate() Then
            If GetPlayerState(m_name & "_awarded") < m_max_per_game.Value() Then
                SetPlayerState "extra_balls", GetPlayerState("extra_balls") + 1
                SetPlayerState m_name & "_awarded", GetPlayerState(m_name & "_awarded") + 1
                DispatchPinEvent m_name & "_awarded", Null
                DispatchPinEvent "extra_ball_awarded", Null
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function ExtraBallsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim extra_ball : Set extra_ball = ownProps(1)
    Select Case evt
        Case "award"
            extra_ball.Award ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set ExtraBallsHandler = kwargs
    Else
        ExtraBallsHandler = kwargs
    End If
End Function

Class GlfLightPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = m_name : End Property
    
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get EventName(name)
        If m_events.Exists(name) Then
            Set EventName = m_events(name)
        Else
            Dim new_event : Set new_event = (new GlfLightPlayerEventItem)()
            m_events.Add name, new_event
            Set EventName = new_event
        End If
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "light_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_events = CreateObject("Scripting.Dictionary")
        
        Set m_base_device = (new GlfBaseModeDevice)(mode, "light_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            Log "Adding Event Listener for event: " & evt
            AddPinEventListener evt, m_mode & "_light_player_play", "LightPlayerEventHandler", m_priority, Array("play", Me, m_events(evt), evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener evt, m_mode & "_light_player_play"
            PlayOff evt, m_events(evt)
        Next
    End Sub

    Public Sub ReloadLights()
        Log "Reloading Lights"
        Dim evt
        For Each evt in m_events.Keys()
            Dim lightName, light
            'First get light counts
            For Each lightName in m_events(evt).LightNames
                Set light = m_events(evt).Lights(lightName)
                Dim lightsCount, x,tagLight, tagLights
                lightsCount = 0
                If Not glf_lightNames.Exists(lightName) Then
                    tagLights = glf_lightTags("T_"&lightName).Keys()
                    Log "Tag Lights: " & Join(tagLights)
                    For Each tagLight in tagLights
                        lightsCount = lightsCount + 1
                    Next
                Else
                    lightsCount = lightsCount + 1
                End If
            Next
            Log "Adding " & lightsCount & " lights for event: " & evt 
            Dim seqArray
            ReDim seqArray(lightsCount-1)
            x=0
            'Build Seq
            For Each lightName in m_events(evt).LightNames
                Set light = m_events(evt).Lights(lightName)

                If Not glf_lightNames.Exists(lightName) Then
                    tagLights = glf_lightTags("T_"&lightName).Keys()
                    For Each tagLight in tagLights
                        seqArray(x) = tagLight & "|100|" & light.Color & "|" & light.Fade
                        x=x+1
                    Next
                Else
                    seqArray(x) = lightName & "|100|" & light.Color & "|" & light.Fade
                    x=x+1
                End If
            Next
            Log "Light List: " & Join(seqArray)
            m_events(evt).LightSeq = seqArray
        Next   
    End Sub

    Public Sub Play(evt, lights)
        LightPlayerCallbackHandler evt, Array(lights.LightSeq), m_name, m_priority, True, 1, Empty
    End Sub

    Public Sub PlayOff(evt, lights)
        LightPlayerCallbackHandler evt, Array(lights.LightSeq), m_name, m_priority, False, 1, Empty
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function
End Class

Class GlfLightPlayerEventItem
	Private m_lights, m_lightSeq
  
    Public Property Get LightNames() : LightNames = m_lights.Keys() : End Property
    Public Property Get Lights(name)
        If m_lights.Exists(name) Then
            Set Lights = m_lights(name)
        Else
            Dim new_event : Set new_event = (new GlfLightPlayerItem)()
            m_lights.Add name, new_event
            Set Lights = new_event
        End If    
    End Property

    Public Property Get LightSeq() : LightSeq = m_lightSeq : End Property
    Public Property Let LightSeq(input) : m_lightSeq = input : End Property

    Public Function ToYaml()
        Dim yaml
        If UBound(m_lights.Keys) > -1 Then
            For Each key in m_lights.keys
                yaml = yaml & "    " & key & ": " & vbCrLf
                yaml = yaml & m_lights(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        Set m_lights = CreateObject("Scripting.Dictionary")
        m_lightSeq = Array()
        Set Init = Me
	End Function

End Class

Class GlfLightPlayerItem
	Private m_light, m_color, m_fade, m_priority
  
    Public Property Get Light(): Light = m_light: End Property
    Public Property Let Light(input): m_light = input: End Property

    Public Property Get Color(): Color = m_color: End Property
    Public Property Let Color(input): m_color = input: End Property

    Public Property Get Fade(): Fade = m_fade: End Property
    Public Property Let Fade(input): m_fade = input: End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input): m_priority = input: End Property      
  
    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "      color: " & m_color & vbCrLf
        If Not IsEmpty(m_fade) Then
            yaml = yaml & "      fade: " & m_fade & vbCrLf
        End If
        If m_priority <> 0 Then
            yaml = yaml & "      priority: " & m_priority & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        m_color = "ffffff"
        m_fade = Empty
        m_priority = 0
        Set Init = Me
	End Function

End Class

Function LightPlayerCallbackHandler(key, lights, mode, priority, play, speed, color_replacement)
    Dim shows_added
    Dim lightStack
    Dim lightParts, light
    If play = False Then
        For Each light in lights(0)
            lightParts = Split(light,"|")
            Set lightStack = glf_lightStacks(lightParts(0))
            If Not lightStack.IsEmpty() Then
                'glf_debugLog.WriteToLog "LightPlayer", "Removing Light " & lightParts(0)
                lightStack.PopByKey(mode & "_" & key)
                Dim show_key
                For Each show_key in glf_running_shows.Keys
                    If Left(show_key, Len("fade_" & mode & "_" & key & "_" & lightParts(0))) = "fade_" & mode & "_" & key & "_" & lightParts(0) Then
                        glf_running_shows(show_key).StopRunningShow()
                    End If
                Next
                If Not lightStack.IsEmpty() Then
                    ' Set the light to the next color on the stack
                    Dim nextColor
                    Set nextColor = lightStack.Peek()
                    Glf_SetLight lightParts(0), nextColor("Color")
                Else
                    ' Turn off the light since there's nothing on the stack
                    Glf_SetLight lightParts(0), "000000"
                End If
            End If
        Next
        Exit Function
        'glf_debugLog.WriteToLog "LightPlayer", "Removing Light Seq" & mode & "_" & key
    Else
        If UBound(lights) = -1 Then
            Exit Function
        End If
        If IsArray(lights) Then
            'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key    
        Else
            'glf_debugLog.WriteToLog "LightPlayer", "Lights not an array!?"
        End If
        'glf_debugLog.WriteToLog "LightPlayer", "Adding Light Seq" & Join(lights) & ". Key:" & mode & "_" & key
        Set shows_added = CreateObject("Scripting.Dictionary")
        For Each light in lights(0)
            lightParts = Split(light,"|")
            
            Set lightStack = glf_lightStacks(lightParts(0))
            Dim oldColor : oldColor = Empty
            Dim newColor : newColor = lightParts(2)
            If Not IsEmpty(color_replacement) Then
                newColor = color_replacement
            End If

            If lightStack.IsEmpty() Then
                oldColor = "000000"
                ' If stack is empty, push the color onto the stack and set the light color
                lightStack.Push mode & "_" & key, newColor, priority
                Glf_SetLight lightParts(0), newColor
            Else
                Dim current
                Set current = lightStack.Peek()                
                If priority >= current("Priority") Then
                    oldColor = current("Color")
                    ' If the new priority is higher, push it onto the stack and change the light color
                    lightStack.Push mode & "_" & key, newColor, priority
                    Glf_SetLight lightParts(0), newColor
                Else
                    ' Otherwise, just push it onto the stack without changing the light color
                    lightStack.Push mode & "_" & key, newColor, priority
                End If
            End If
            
            If Not IsEmpty(oldColor) And Ubound(lightParts)=4 Then
                If lightParts(4) <> "" Then
                    'FadeMs
                    Dim cache_name, new_running_show,cached_show,show_settings, fade_seq
                    cache_name = lightParts(3)
                    fade_seq = Glf_FadeRGB(oldColor, newColor, lightParts(4))
                    'MsgBox cache_name
                    shows_added.Add cache_name, True
                    
                    If glf_cached_shows.Exists(cache_name & "__-1") Then
                        'msgbox ubound(glf_cached_shows(cache_name & "__-1")(0))
                        'msgbox ubound(fade_seq)
                        'msgbox "Converted show: " & cache_name & ", steps: " & ubound(glf_cached_shows(cache_name & "__-1")(0)) & ". Fade Replacements: " & ubound(fade_seq) & ". Extracted Step Number: " & lightParts(4) 
                        Set show_settings = (new GlfShowPlayerItem)()
                        show_settings.Show = cache_name
                        show_settings.Loops = 1
                        show_settings.Speed = speed
                        show_settings.ColorLookup = fade_seq
                        Set new_running_show = (new GlfRunningShow)(cache_name, show_settings.Key, show_settings, priority+1, Null, Null)
                    End If
                End If
            End If
        Next
    End If
    Set LightPlayerCallbackHandler = shows_added
End Function

Function LightPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim LightPlayer : Set LightPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            LightPlayer.Activate
        Case "deactivate"
            LightPlayer.Deactivate
        Case "play"
            LightPlayer.Play ownProps(3), ownProps(2)
    End Select
    LightPlayerEventHandler = Null
End Function

Class GlfLightStack
    Private stack

    Public default Function Init()
        ReDim stack(-1)  ' Initialize an empty array
        Set Init = Me
    End Function

    Public Sub Push(key, color, priority)
        Dim found : found = False
        Dim i

        ' Check if the key already exists in the stack and update it
        For i = LBound(stack) To UBound(stack)
            If stack(i)("Key") = key Then
                ' Replace the existing item if the key matches
                Set stack(i) = CreateColorPriorityObject(key, color, priority)
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            ' Insert the new item into the array maintaining priority order
            ReDim Preserve stack(UBound(stack) + 1)
            Set stack(UBound(stack)) = CreateColorPriorityObject(key, color, priority)
            SortStackByPriority
        End If
    End Sub

    Public Function PopByKey(key)
        Dim i, removedItem, found
        found = False
        Set removedItem = Nothing
    
        ' Loop through the stack to find the item with the matching key
        For i = LBound(stack) To UBound(stack)
            If stack(i)("Key") = key Then
                ' Store the item to be removed
                Set removedItem = stack(i)
                found = True
    
                ' Shift all elements after the removed item to the left
                Dim j
                For j = i To UBound(stack) - 1
                    Set stack(j) = stack(j + 1)
                Next
    
                ' Resize the array to remove the last element
                ReDim Preserve stack(UBound(stack) - 1)
                Exit For
            End If
        Next
    
        ' Return the removed item (or Nothing if not found)
        If found Then
            Set PopByKey = removedItem
        Else
            Set PopByKey = Nothing
        End If
    End Function
    

    ' Get the current top color without popping it
    Public Function Peek()
        If UBound(stack) >= 0 Then
            Set Peek = stack(LBound(stack))
        Else
            Set Peek = Nothing
        End If
    End Function

    ' Check if the stack is empty
    Public Function IsEmpty()
        IsEmpty = (UBound(stack) < 0)
    End Function

    ' Create a color-priority object
    Private Function CreateColorPriorityObject(key, color, priority)
        Dim colorPriorityObject
        Set colorPriorityObject = CreateObject("Scripting.Dictionary")
        colorPriorityObject.Add "Key", key
        colorPriorityObject.Add "Color", color
        colorPriorityObject.Add "Priority", priority
        Set CreateColorPriorityObject = colorPriorityObject
    End Function

    ' Sort the stack by priority (descending)
    Private Sub SortStackByPriority()
        Dim i, j
        Dim temp
        For i = LBound(stack) To UBound(stack) - 1
            For j = i + 1 To UBound(stack)
                If stack(i)("Priority") < stack(j)("Priority") Then
                    ' Swap the elements
                    Set temp = stack(i)
                    Set stack(i) = stack(j)
                    Set stack(j) = temp
                End If
            Next
        Next
    End Sub

    Public Sub PrintStackOrder()
        Dim i
        Debug.Print "Stack Order:" 
        For i = LBound(stack) To UBound(stack)
            Debug.Print "Key: " & stack(i)("Key") & ", Color: " & stack(i)("Color") & ", Priority: " & stack(i)("Priority")
        Next
    End Sub
End Class

Class GlfBaseModeDevice

    Private m_mode
    Private m_priority
    Private m_enable_events
    Private m_disable_events
    Private m_device
    Private m_parent
    Private m_debug

    Public Property Get EnableEvents(): Set EnableEvents = m_enable_events: End Property
    Public Property Let EnableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Get DisableEvents(): Set DisableEvents = m_disable_events: End Property
    Public Property Let DisableEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get Mode(): Set Mode = m_mode: End Property

    Public Property Let Debug(value) : m_debug = value : End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode, device, parent)
        Set m_mode = mode
        m_priority = mode.Priority
        m_device = device
        Set m_parent = parent
        If mode.IsDebug = 1 Then
            m_debug = True
        End If

        Set m_enable_events = CreateObject("Scripting.Dictionary")
        Set m_disable_events = CreateObject("Scripting.Dictionary")

        AddPinEventListener m_mode.Name & "_starting", m_device & "_" & m_parent.Name & "_activate", "BaseModeDeviceEventHandler", m_priority+1, Array("activate", Me)
        AddPinEventListener m_mode.Name & "_stopping", m_device & "_" & m_parent.Name & "_deactivate", "BaseModeDeviceEventHandler", m_priority-1, Array("deactivate", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_enable_events.Keys()
            AddPinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable", "BaseModeDeviceEventHandler", m_priority+m_enable_events(evt).Priority, Array("enable", m_parent, m_enable_events(evt))
        Next
        For Each evt In m_disable_events.Keys()
            AddPinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable", "BaseModeDeviceEventHandler", m_priority+m_disable_events(evt).Priority, Array("disable", m_parent, m_disable_events(evt))
        Next
        m_parent.Activate
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        Dim evt
        For Each evt In m_enable_events.Keys()
            RemovePinEventListener m_enable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_enable"
        Next
        For Each evt In m_disable_events.Keys()
            RemovePinEventListener m_disable_events(evt).EventName, m_mode.Name & m_device & "_" & m_parent.Name & "_disable"
        Next
        m_parent.Deactivate
    End Sub

    Public Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode.Name & "_" & m_device & "_" & m_parent.Name, message
        End If
    End Sub
End Class


Function BaseModeDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    Dim glfEvent
    Select Case evt
        Case "activate"
            device.Activate
        Case "deactivate"
            device.Deactivate
        Case "enable"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            device.Enable
        Case "disable"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            device.Disable
    End Select
    If IsObject(args(1)) Then
        Set BaseModeDeviceEventHandler = kwargs
    Else
        BaseModeDeviceEventHandler = kwargs
    End If
End Function

Class Mode

    Private m_name
    Private m_modename 
    Private m_start_events
    Private m_stop_events
    private m_priority
    Private m_debug
    Private m_started
    Private m_ballsaves
    Private m_counters
    Private m_multiball_locks
    Private m_multiballs
    Private m_shots
    Private m_shot_groups
    Private m_ballholds
    Private m_timers
    Private m_lightplayer
    Private m_segment_display_player
    Private m_showplayer
    Private m_variableplayer
    Private m_eventplayer
    Private m_queueEventplayer
    Private m_queueRelayPlayer
    Private m_random_event_player
    Private m_sound_player
    Private m_dof_player
    Private m_slide_player
    Private m_shot_profiles
    Private m_sequence_shots
    Private m_state_machines
    Private m_extra_balls
    Private m_combo_switches
    Private m_timed_switches
    Private m_tilt
    Private m_use_wait_queue

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "active":
                If m_started Then
                    GetValue = True
                Else
                    GetValue = False
                End If
        End Select
    End Property
    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Get Status()
        If m_started Then
            Status = "started"
        Else
            Status = "stopped"
        End If
    End Property
    Public Property Get LightPlayer()
        If IsNull(m_lightplayer) Then
            Set m_lightplayer = (new GlfLightPlayer)(Me)
        End If
        Set LightPlayer = m_lightplayer
    End Property
    Public Property Get ShowPlayer()
        If IsNull(m_showplayer) Then
            Set m_showplayer = (new GlfShowPlayer)(Me)
        End If
        Set ShowPlayer = m_showplayer
    End Property
    Public Property Get SegmentDisplayPlayer()
        If IsNull(m_segment_display_player) Then
            Set m_segment_display_player = (new GlfSegmentDisplayPlayer)(Me)
        End If
        Set SegmentDisplayPlayer = m_segment_display_player
    End Property
    Public Property Get EventPlayer() : Set EventPlayer = m_eventplayer: End Property
    Public Property Get QueueEventPlayer() : Set QueueEventPlayer = m_queueEventplayer: End Property
    Public Property Get QueueRelayPlayer() : Set QueueRelayPlayer = m_queueRelayPlayer: End Property
    Public Property Get RandomEventPlayer() : Set RandomEventPlayer = m_random_event_player : End Property
    Public Property Get VariablePlayer(): Set VariablePlayer = m_variableplayer: End Property
    Public Property Get SoundPlayer() : Set SoundPlayer = m_sound_player : End Property
    Public Property Get DOFPlayer() : Set DOFPlayer = m_dof_player : End Property
    Public Property Get SlidePlayer() : Set SlidePlayer = m_slide_player : End Property

    Public Property Get ShotProfiles(name)
        If m_shot_profiles.Exists(name) Then
            Set ShotProfiles = m_shot_profiles(name)
        Else
            Dim new_shotprofile : Set new_shotprofile = (new GlfShotProfile)(name)
            m_shot_profiles.Add name, new_shotprofile
            Glf_ShotProfiles.Add name, new_shotprofile
            Set ShotProfiles = new_shotprofile
        End If
    End Property

    Public Property Get BallSavesItems() : BallSavesItems = m_ballsaves.Items() : End Property
    Public Property Get BallSaves(name)
        If m_ballsaves.Exists(name) Then
            Set BallSaves = m_ballsaves(name)
        Else
            Dim new_ballsave : Set new_ballsave = (new BallSave)(name, Me)
            m_ballsaves.Add name, new_ballsave
            Set BallSaves = new_ballsave
        End If
    End Property

    Public Property Get TimersItems() : TimersItems = m_timers.Items() : End Property
    Public Property Get Timers(name)
        If m_timers.Exists(name) Then
            Set Timers = m_timers(name)
        Else
            Dim new_timer : Set new_timer = (new GlfTimer)(name, Me)
            m_timers.Add name, new_timer
            Set Timers = new_timer
        End If
    End Property

    Public Property Get CountersItems() : CountersItems = m_counters.Items() : End Property
    Public Property Get Counters(name)
        If m_counters.Exists(name) Then
            Set Counters = m_counters(name)
        Else
            Dim new_counter : Set new_counter = (new GlfCounter)(name, Me)
            m_counters.Add name, new_counter
            Set Counters = new_counter
        End If
    End Property

    Public Property Get MultiballLocksItems() : MultiballLocksItems = m_multiball_locks.Items() : End Property
    Public Property Get MultiballLocks(name)
        If m_multiball_locks.Exists(name) Then
            Set MultiballLocks = m_multiball_locks(name)
        Else
            Dim new_multiball_lock : Set new_multiball_lock = (new GlfMultiballLocks)(name, Me)
            m_multiball_locks.Add name, new_multiball_lock
            Set MultiballLocks = new_multiball_lock
        End If
    End Property

    Public Property Get MultiballsItems() : MultiballsItems = m_multiballs.Items() : End Property
    Public Property Get Multiballs(name)
        If m_multiballs.Exists(name) Then
            Set Multiballs = m_multiballs(name)
        Else
            Dim new_multiball : Set new_multiball = (new GlfMultiballs)(name, Me)
            m_multiballs.Add name, new_multiball
            Set Multiballs = new_multiball
        End If
    End Property

    Public Property Get SequenceShotsItems() : SequenceShotsItems = m_sequence_shots.Items() : End Property
    Public Property Get SequenceShots(name)
        If m_sequence_shots.Exists(name) Then
            Set SequenceShots = m_sequence_shots(name)
        Else
            Dim new_sequence_shot : Set new_sequence_shot = (new GlfSequenceShots)(name, Me)
            m_sequence_shots.Add name, new_sequence_shot
            Set SequenceShots = new_sequence_shot
        End If
    End Property

    Public Property Get ExtraBallsItems() : ExtraBallsItems = m_extra_balls.Items() : End Property
    Public Property Get ExtraBalls(name)
        If m_extra_balls.Exists(name) Then
            Set ExtraBalls = m_extra_balls(name)
        Else
            Dim new_extra_ball : Set new_extra_ball = (new GlfExtraBall)(name, Me)
            m_extra_balls.Add name, new_extra_ball
            Set ExtraBalls = new_extra_ball
        End If
    End Property
    Public Property Get ComboSwitchesItems() : ComboSwitchesItems = m_combo_switches.Items() : End Property
    Public Property Get ComboSwitches(name)
        If m_combo_switches.Exists(name) Then
            Set ComboSwitches = m_combo_switches(name)
        Else
            Dim new_combo_switch : Set new_combo_switch = (new GlfComboSwitches)(name, Me)
            m_combo_switches.Add name, new_combo_switch
            Set ComboSwitches = new_combo_switch
        End If
    End Property
    Public Property Get TimedSwitchesItems() : TimedSwitchesItems = m_timed_switches.Items() : End Property
        Public Property Get TimedSwitches(name)
            If m_timed_switches.Exists(name) Then
                Set TimedSwitches = m_timed_switches(name)
            Else
                Dim new_timed_switch : Set new_timed_switch = (new GlfTimedSwitches)(name, Me)
                m_timed_switches.Add name, new_timed_switch
                Set TimedSwitches = new_timed_switch
            End If
        End Property
    Public Property Get Tilt()
        If Not IsNull(m_tilt) Then
            Set Tilt = m_tilt
        Else
            Set m_tilt = (new GlfTilt)(Me)
            Set Tilt = m_tilt
        End If
    End Property
    Public Property Get TiltConfig()
        If Not IsNull(m_tilt) Then
            Set TiltConfig = m_tilt
        Else
            TiltConfig = Null
        End If
    End Property

    Public Property Get StateMachines(name)
        If m_state_machines.Exists(name) Then
            Set StateMachines = m_state_machines(name)
        Else
            Dim new_state_machine : Set new_state_machine = (new GlfStateMachine)(name, Me)
            m_state_machines.Add name, new_state_machine
            Set StateMachines = new_state_machine
        End If
    End Property
    Public Property Get ModeStateMachines(): ModeStateMachines = m_state_machines.Items(): End Property

    Public Property Get ModeShots(): ModeShots = m_shots.Items(): End Property
    Public Property Get Shots(name)
        If m_shots.Exists(name) Then
            Set Shots = m_shots(name)
        Else
            Dim new_shot : Set new_shot = (new GlfShot)(name, Me)
            m_shots.Add name, new_shot
            Set Shots = new_shot
        End If
    End Property

    Public Property Get ShotGroupsItems() : ShotGroupsItems = m_shot_groups.Items() : End Property
    Public Property Get ShotGroups(name)
        If m_shot_groups.Exists(name) Then
            Set ShotGroups = m_shot_groups(name)
        Else
            Dim new_shot_group : Set new_shot_group = (new GlfShotGroup)(name, Me)
            m_shot_groups.Add name, new_shot_group
            Set ShotGroups = new_shot_group
        End If
    End Property

    Public Property Get BallHoldsItems() : BallHoldsItems = m_ballholds.Items() : End Property
    Public Property Get BallHolds(name)
        If m_ballholds.Exists(name) Then
            Set BallHolds = m_shots(name)
        Else
            Dim new_ballhold : Set new_ballhold = (new GlfBallHold)(name, Me)
            m_ballholds.Add name, new_ballhold
            Set BallHolds = new_ballhold
        End If
    End Property

    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Name, newEvent
            AddPinEventListener newEvent.EventName, m_name & "_start", "ModeEventHandler", m_priority+newEvent.Priority, Array("start", Me, newEvent)
        Next
    End Property
    
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Name, newEvent
            AddPinEventListener newEvent.EventName, m_name & "_stop", "ModeEventHandler", m_priority+newEvent.Priority+1, Array("stop", Me, newEvent)
        Next
    End Property

    Public Property Get UseWaitQueue(): UseWaitQueue = m_use_wait_queue: End Property
    Public Property Let UseWaitQueue(input): m_use_wait_queue = input: End Property

    Public Property Get IsDebug()
        If m_debug = True Then
            IsDebug = 1
        Else
            IsDebug = 0
        End If
    End Property
    Public Property Let Debug(value)
        m_debug = value
        Dim config_item
        For Each config_item in m_ballsaves.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_counters.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_timers.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_multiball_locks.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_multiballs.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_shots.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_shot_groups.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_ballholds.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_sequence_shots.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_state_machines.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_extra_balls.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_combo_switches.Items()
            config_item.Debug = value
        Next
        For Each config_item in m_timed_switches.Items()
            config_item.Debug = value
        Next
        If Not IsNull(m_tilt) Then
            m_tilt.Debug = value
        End If
        If Not IsNull(m_lightplayer) Then
            m_lightplayer.Debug = value
        End If
        If Not IsNull(m_eventplayer) Then
            m_eventplayer.Debug = value
        End If
        If Not IsNull(m_queueEventplayer) Then
            m_queueEventplayer.Debug = value
        End If
        If Not IsNull(m_queueRelayPlayer) Then
            m_queueRelayPlayer.Debug = value
        End If
        If Not IsNull(m_random_event_player) Then
            m_random_event_player.Debug = value
        End If
        If Not IsNull(m_sound_player) Then
            m_sound_player.Debug = value
        End If
        If Not IsNull(m_dof_player) Then
            m_dof_player.Debug = value
        End If
        If Not IsNull(m_slide_player) Then
            m_slide_player.Debug = value
        End If
        If Not IsNull(m_showplayer) Then
            m_showplayer.Debug = value
        End If
        If Not IsNull(m_segment_display_player) Then
            m_segment_display_player.Debug = value
        End If
        If Not IsNull(m_variableplayer) Then
            m_variableplayer.Debug = value
        End If
        Glf_MonitorModeUpdate Me
    End Property

	Public default Function init(name, priority)
        m_name = "mode_"&name
        m_modename = name
        m_priority = priority
        m_started = False
        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")
        Set m_ballsaves = CreateObject("Scripting.Dictionary")
        Set m_counters = CreateObject("Scripting.Dictionary")
        Set m_timers = CreateObject("Scripting.Dictionary")
        Set m_multiball_locks = CreateObject("Scripting.Dictionary")
        Set m_multiballs = CreateObject("Scripting.Dictionary")
        Set m_shots = CreateObject("Scripting.Dictionary")
        Set m_shot_groups = CreateObject("Scripting.Dictionary")
        Set m_ballholds = CreateObject("Scripting.Dictionary")
        Set m_shot_profiles = CreateObject("Scripting.Dictionary")
        Set m_sequence_shots = CreateObject("Scripting.Dictionary")
        Set m_state_machines = CreateObject("Scripting.Dictionary")
        Set m_extra_balls = CreateObject("Scripting.Dictionary")
        Set m_combo_switches = CreateObject("Scripting.Dictionary")
        Set m_timed_switches = CreateObject("Scripting.Dictionary")

        m_use_wait_queue = False
        m_lightplayer = Null
        m_tilt = Null
        m_showplayer = Null
        m_segment_display_player = Null
        Set m_eventplayer = (new GlfEventPlayer)(Me)
        Set m_queueEventplayer = (new GlfQueueEventPlayer)(Me)
        Set m_queueRelayPlayer = (new GlfQueueRelayPlayer)(Me)
        Set m_random_event_player = (new GlfRandomEventPlayer)(Me)
        Set m_sound_player = (new GlfSoundPlayer)(Me)
        Set m_dof_player = (new GlfDofPlayer)(Me)
        Set m_slide_player = (new GlfSlidePlayer)(Me)
        Set m_variableplayer = (new GlfVariablePlayer)(Me)
        Glf_MonitorModeUpdate Me
        AddPinEventListener m_name & "_starting", m_name & "_starting_end", "ModeEventHandler", -99, Array("started", Me, "")
        AddPinEventListener m_name & "_stopping", m_name & "_stopping_end", "ModeEventHandler", -99, Array("stopped", Me, "")
        Set Init = Me
	End Function

    Public Sub StartMode()
        Log "Starting"
        m_started=True
        DispatchQueuePinEvent m_name & "_starting", Null
    End Sub

    Public Sub StopMode()
        If m_started = True Then
            m_started = False
            Log "Stopping"
            DispatchQueuePinEvent m_name & "_stopping", Null
        End If
    End Sub

    Public Sub Started()
        DispatchPinEvent m_name & "_started", Null
        Glf_MonitorModeUpdate Me
        glf_running_modes = glf_running_modes & "["""&m_modename&""", " & m_priority & "],"
        Log "Started"
    End Sub

    Public Sub Stopped()
        'MsgBox m_name & "Stopped"
        DispatchPinEvent m_name & "_stopped", Null
        Glf_MonitorModeUpdate Me
        glf_running_modes = Replace(glf_running_modes, "["""&m_modename&""", " & m_priority & "],", "")
        Log "Stopped"
    End Sub

    Private Sub Log(message) 
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        dim yaml, child
        yaml = "#config_version=6" & vbCrLf & vbCrLf

        yaml = yaml & "mode:" & vbCrLf

        If UBound(m_start_events.Keys) > -1 Then
            yaml = yaml & "  start_events: "
            x=0
            For Each key in m_start_events.keys
                yaml = yaml & m_start_events(key).Raw
                If x <> UBound(m_start_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
        If UBound(m_stop_events.Keys) > -1 Then
            yaml = yaml & "  stop_events: "
            x=0
            For Each key in m_stop_events.keys
                yaml = yaml & m_stop_events(key).Raw
                If x <> UBound(m_stop_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        yaml = yaml & "  priority: " & m_priority & vbCrLf
        
        If UBound(m_ballsaves.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "ball_saves: " & vbCrLf
            For Each child in m_ballsaves.Keys
                yaml = yaml & m_ballsaves(child).ToYaml
            Next
        End If
        
        If UBound(m_shot_profiles.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "shot_profiles: " & vbCrLf
            For Each child in m_shot_profiles.Keys
                yaml = yaml & m_shot_profiles(child).ToYaml
            Next
        End If
        
        If UBound(m_shots.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "shots: " & vbCrLf
            For Each child in m_shots.Keys
                yaml = yaml & m_shots(child).ToYaml
            Next
        End If
        
        If UBound(m_shot_groups.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "shot_groups: " & vbCrLf
            For Each child in m_shot_groups.Keys
                yaml = yaml & m_shot_groups(child).ToYaml
            Next
        End If
        
        If UBound(m_eventplayer.Events.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "event_player: " & vbCrLf
            yaml = yaml & m_eventplayer.ToYaml()
        End If
        
        If Not IsNull(m_showPlayer) Then
            If UBound(m_showplayer.EventShows)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "show_player: " & vbCrLf
                yaml = yaml & m_showplayer.ToYaml()
            End If
        End If
        
        If Not IsNull(m_lightplayer) Then
            If UBound(m_lightplayer.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "light_player: " & vbCrLf
                For Each child in m_lightplayer.EventNames
                    yaml = yaml & m_lightplayer.ToYaml()
                Next
            End If
        End If

        If Not IsNull(m_segment_display_player) Then
            If UBound(m_segment_display_player.EventNames)>-1 Then
                yaml = yaml & vbCrLf
                yaml = yaml & "segment_display_player: " & vbCrLf
                For Each child in m_segment_display_player.EventNames
                    yaml = yaml & m_segment_display_player.ToYaml()
                Next
            End If
        End If
        
        If UBound(m_ballholds.Keys)>-1 Then
            yaml = yaml & vbCrLf
            yaml = yaml & "ball_holds: " & vbCrLf
            For Each child in m_ballholds.Keys
                yaml = yaml & m_ballholds(child).ToYaml
            Next
        End If
        yaml = yaml & vbCrLf
        
        Dim fso, modesFolder, TxtFileStream
        Set fso = CreateObject("Scripting.FileSystemObject")
        modesFolder = "glf_mpf\modes\" & Replace(m_name, "mode_", "") & "\config"

        If Not fso.FolderExists("glf_mpf") Then
            fso.CreateFolder "glf_mpf"
        End If

        Dim currentFolder
        Dim folderParts
        Dim i
    
        ' Split the path into parts
        folderParts = Split(modesFolder, "\")
        
        ' Initialize the current folder as the root
        currentFolder = folderParts(0)
    
        ' Iterate over each part of the path and create folders as needed
        For i = 1 To UBound(folderParts)
            currentFolder = currentFolder & "\" & folderParts(i)
            If Not fso.FolderExists(currentFolder) Then
                fso.CreateFolder(currentFolder)
            End If
        Next


        
        Set TxtFileStream = fso.OpenTextFile(modesFolder & "\" & Replace(m_name, "mode_", "") & ".yaml", 2, True)
        TxtFileStream.WriteLine yaml
        TxtFileStream.Close

        ToYaml = yaml
    End Function
End Class

Function ModeEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim mode : Set mode = ownProps(1)
    Dim glfEvent
    Select Case evt
        Case "start"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            mode.StartMode
            If mode.UseWaitQueue = True Then
                kwargs.Add "wait_for", mode.Name & "_stopped"
            End If
        Case "stop"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            mode.StopMode
        Case "started"
            mode.Started
        Case "stopped"
            mode.Stopped
    End Select
    If IsObject(args(1)) Then
        Set ModeEventHandler = kwargs
    Else
        ModeEventHandler = kwargs
    End If
End Function

Class GlfMultiballLocks

    Private m_name
    Private m_lock_device
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_enable_events
    Private m_disable_events
    Private m_balls_to_lock
    Private m_balls_locked
    Private m_balls_to_replace
    Private m_lock_events
    Private m_reset_events
    Private m_enabled
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "locked_balls":
                GetValue = m_balls_locked
        End Select
    End Property
    Public Property Get LockDevice() : LockDevice = m_lock_device : End Property
    Public Property Let LockDevice(value) : m_lock_device = value : End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let BallsToLock(value) : m_balls_to_lock = value : End Property
    Public Property Let LockEvents(value) : m_lock_events = value : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let BallsToReplace(value) : m_balls_to_replace = value : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(name, mode)
        m_name = "multiball_lock_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_lock_events = Array()
        m_reset_events = Array()
        m_lock_device = Empty
        m_balls_to_lock = 0
        m_balls_to_replace = -1
        m_enabled = False
        m_balls_locked = 0
        Set m_base_device = (new GlfBaseModeDevice)(mode, "multiball_lock", Me)
        glf_multiball_locks.Add name, Me
        Set Init = Me
	End Function

    Public Sub Activate()
        If UBound(m_base_device.EnableEvents.Keys()) = -1 Then
            Enable()
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        If Not IsEmpty(m_lock_device) Then
            AddPinEventListener "balldevice_" & m_lock_device & "_ball_enter", m_mode & "_" & name & "_lock", "MultiballLocksHandler", m_priority, Array("lock", me, m_lock_device)
        End If
        Dim evt
        For Each evt in m_lock_events
            AddPinEventListener evt, m_name & "_ball_locked", "MultiballLocksHandler", m_priority, Array("virtual_lock", Me, Null)
        Next
        For Each evt in m_reset_events
            AddPinEventListener evt, m_name & "_reset", "MultiballLocksHandler", m_priority, Array("reset", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        If Not IsEmpty(m_lock_device) Then
            RemovePinEventListener "balldevice_" & m_lock_device & "_ball_enter", m_mode & "_" & name & "_lock"
        End If
        Dim evt
        For Each evt in m_lock_events
            RemovePinEventListener evt, m_name & "_ball_locked"
        Next
        For Each evt in m_reset_events
            RemovePinEventListener evt, m_name & "_reset"
        Next
    End Sub

    Public Function Lock(device, unclaimed_balls)
        
        If unclaimed_balls <= 0 Then
            Lock = unclaimed_balls
            Exit Function
        End If
        
        Dim balls_locked
        If GetPlayerState(m_name & "_balls_locked") = False Then
            balls_locked = 1
        Else
            balls_locked = GetPlayerState(m_name & "_balls_locked") + 1
        End If
        If balls_locked > m_balls_to_lock Then
            Log "Cannot lock balls. Lock is full."
            Lock = unclaimed_balls
            Exit Function
        End If

        SetPlayerState m_name & "_balls_locked", balls_locked
        

        If Not IsNull(device) Then
            
            If glf_ball_devices(device).Balls() > balls_locked Then
                glf_ball_devices(device).Eject()
            Else
                If m_balls_to_replace = -1 Or balls_locked <= m_balls_to_replace Then
                    ' glf_BIP = glf_BIP - 1
                    SetDelay m_name & "_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", Me),Null), 1000
                End If
            End If
        End If

        DispatchPinEvent m_name & "_locked_ball", balls_locked
        
        If balls_locked = m_balls_to_lock Then
            DispatchPinEvent m_name & "_full", balls_locked
        End If

        Lock = unclaimed_balls - 1
    End Function

    Public Sub Reset
        Log "Resetting multiball lock count"
        SetPlayerState m_name & "_balls_locked", 0
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function MultiballLocksHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    Select Case evt
        Case "activate"
            multiball.Activate
        Case "deactivate"
            multiball.Deactivate
        Case "enable"
            multiball.Enable
        Case "disable"
            multiball.Disable
        Case "lock"
            kwargs = multiball.Lock(ownProps(2), kwargs)
        Case "virtual_lock"
            multiball.Lock Null, 1
        Case "reset"
            multiball.Reset
        Case "queue_release"
            If glf_plunger.HasBall = False And ballInReleasePostion = True And glf_plunger.IncomingBalls = 0  Then
                Glf_ReleaseBall(Null)
                SetDelay multiball.Name&"_auto_launch", "MultiballLocksHandler" , Array(Array("auto_launch", multiball),Null), 500
            Else
                SetDelay multiball.Name&"_queued_release", "MultiballLocksHandler" , Array(Array("queue_release", multiball), Null), 1000
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay multiball.Name&"_auto_launch", "MultiballLocksHandler" , Array(Array("auto_launch", multiball), Null), 500
            End If
    End Select
    If IsObject(args(1)) Then
        Set MultiballLocksHandler = kwargs
    Else
        MultiballLocksHandler = kwargs
    End If
End Function
Class GlfMultiballs

    Private m_name
    Private m_configname
    Private m_mode
    Private m_priority
    Private m_base_device
    Private m_ball_count
    Private m_ball_lock
    Private m_add_a_ball_events
    Private m_add_a_ball_grace_period
    Private m_add_a_ball_hurry_up_time
    Private m_add_a_ball_shoot_again
    Private m_ball_count_type
    Private m_disable_events
    Private m_enable_events
    Private m_grace_period
    Private m_grace_period_enabled
    Private m_hurry_up
    Private m_hurry_up_enabled
    Private m_replace_balls_in_play
    Private m_reset_events
    Private m_shoot_again
    Private m_source_playfield
    Private m_start_events
    Private m_stop_events
    Private m_balls_added_live
    Private m_balls_live_target
    Private m_enabled
    Private m_shoot_again_enabled
    Private m_queued_balls
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
        End Select
    End Property

    Public Property Let BallCount(value): Set m_ball_count = CreateGlfInput(value): End Property
    Public Property Let AddABallEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_add_a_ball_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let AddABallGracePeriod(value): Set m_add_a_ball_grace_period = CreateGlfInput(value): End Property
    Public Property Let AddABallHurryUpTime(value): Set m_add_a_ball_hurry_up_time = CreateGlfInput(value): End Property
    Public Property Let AddABallShootAgain(value): Set m_add_a_ball_shoot_again = CreateGlfInput(value): End Property
    Public Property Let BallCountType(value): m_ball_count_type = value: End Property
    Public Property Let BallLock(value): m_ball_lock = value: End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Let GracePeriod(value): Set m_grace_period = CreateGlfInput(value): End Property
    Public Property Let HurryUp(value): Set m_hurry_up = CreateGlfInput(value): End Property
    Public Property Let ReplaceBallsInPlay(value): m_replace_balls_in_play = value: End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ShootAgain(value): Set m_shoot_again = CreateGlfInput(value): End Property
    Public Property Let StartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_start_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let StopEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_stop_events.Add newEvent.Raw, newEvent
        Next
    End Property
        
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(name, mode)
        m_name = "multiball_" & name
        m_configname = name
        m_mode = mode.Name
        m_priority = mode.Priority
        Set m_ball_count = CreateGlfInput(0)
        Set m_add_a_ball_events = CreateObject("Scripting.Dictionary")
        Set m_add_a_ball_grace_period = CreateGlfInput(0)
        Set m_add_a_ball_hurry_up_time = CreateGlfInput(0)
        Set m_add_a_ball_shoot_again = CreateGlfInput(5000)
        m_ball_count_type = "total"
        m_ball_lock = Empty
        Set m_grace_period = CreateGlfInput(0)
        Set m_hurry_up = CreateGlfInput(0)
        m_replace_balls_in_play = False
        Set m_shoot_again = CreateGlfInput(10000)
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_start_events = CreateObject("Scripting.Dictionary")
        Set m_stop_events = CreateObject("Scripting.Dictionary")
        m_replace_balls_in_play = False
        m_balls_added_live = 0
        m_balls_live_target = 0
        m_queued_balls = 0
        m_enabled = False
        m_shoot_again_enabled = False
        m_grace_period_enabled = False
        m_hurry_up_enabled = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "multiball", Me)
        glf_multiballs.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        If UBound(m_base_device.EnableEvents.Keys()) = -1 Then
            Enable()
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling " & m_name
        m_enabled = True
        Dim evt
        For Each evt in m_start_events.Keys
            AddPinEventListener m_start_events(evt).EventName, m_name & "_" & evt & "_start", "MultiballsHandler", m_priority+m_start_events(evt).Priority, Array("start", Me, m_start_events(evt))
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset", "MultiballsHandler", m_priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_add_a_ball_events.Keys
            AddPinEventListener m_add_a_ball_events(evt).EventName, m_name & "_" & evt & "_add_a_ball", "MultiballsHandler", m_priority, Array("add_a_ball", Me, m_add_a_ball_events(evt))
        Next
        For Each evt in m_stop_events.Keys
            AddPinEventListener m_stop_events(evt).EventName, m_name & "_" & evt & "_stop", "MultiballsHandler", m_priority+m_stop_events(evt).Priority, Array("stop", Me, m_stop_events(evt))
        Next
    End Sub
    
    Public Sub Disable()
        Log "Disabling " & m_name
        m_enabled = False
        m_balls_added_live = 0
        m_balls_live_target = 0
        m_shoot_again_enabled = False
        StopMultiball()
        Dim evt
        For Each evt in m_start_events.Keys
            RemovePinEventListener m_start_events(evt).EventName, m_name & "_" & evt & "_start"
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_add_a_ball_events.Keys
            RemovePinEventListener m_add_a_ball_events(evt).EventName, m_name & "_" & evt & "_add_a_ball"
        Next
        For Each evt in m_stop_events.Keys
            RemovePinEventListener m_stop_events(evt).EventName, m_name & "_" & evt & "_stop"
        Next
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        'RemoveDelay m_name & "_queued_release"
    End Sub
    
    Private Sub HandleBallsInPlayAndBallsLive()
        'Dim balls_to_replace
        'If m_replace_balls_in_play = True Then
        '    balls_to_replace = glf_BIP
        'Else
        '    balls_to_replace = 0
        'End If
        'Log("Going to add an additional " & balls_to_replace & " balls for replace_balls_in_play")
        m_balls_added_live = 0 
        Dim ball_count_value : ball_count_value = m_ball_count.Value
        If m_ball_count_type = "total" Then
            Log "glf_BIP: " & glf_BIP
            If ball_count_value > glf_BIP Then
                m_balls_added_live = ball_count_value - glf_BIP
                'glf_BIP = m_ball_count
            End If
            m_balls_live_target = ball_count_value
        Else
            m_balls_added_live = ball_count_value
            'glf_BIP = glf_BIP + m_balls_added_live
            m_balls_live_target = glf_BIP + m_balls_added_live
        End If

    End Sub

    Public Function BallsDrained(balls)
        If m_shoot_again_enabled Then
            balls = BallDrainShootAgain(balls)
        Else
            BallDrainCountBalls(balls)
        End If
        BallsDrained = balls
    End Function

    Public Sub Start()
        ' Start multiball.
        If not m_enabled Then
            Exit Sub
        End If

        If m_balls_live_target > 0 Then
            Log("Cannot start MB because " & m_balls_live_target & " are still in play")
            Exit Sub
        End If

        m_shoot_again_enabled = True

        HandleBallsInPlayAndBallsLive()
        Log("Starting multiball with " & m_balls_live_target & " balls (added " & m_balls_added_live & ")")
        'msgbox("Starting multiball with " & m_balls_live_target & " balls (added " & m_balls_added_live & ")")    
        Dim balls_added : balls_added = 0

        'eject balls from locks
        If Not IsEmpty(m_ball_lock) Then
            Dim available_balls : available_balls = glf_ball_devices(m_ball_lock).Balls()
            If available_balls > 0 Then
                glf_ball_devices(m_ball_lock).EjectAll()
            End If
            balls_added = available_balls
        End If

        glf_BIP = m_balls_live_target

        'request remaining balls
        m_queued_balls = (m_balls_added_live - balls_added)
        If m_queued_balls > 0 Then
            SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me),Null), 1000
        End If

        If m_shoot_again.Value = 0 Then
            'No shoot again. Just stop multiball right away
            StopMultiball()
        else
            'Enable shoot again
            TimerStart()
        End If
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "MultiballsHandler", m_priority, Array("drain", Me)

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls", m_balls_live_target
        End With
        DispatchPinEvent m_name & "_started", kwargs
    End Sub

    Sub TimerStart()
        DispatchPinEvent "ball_save_" & m_configname & "_timer_start", Null 'desc: The multiball ball save called (name) has just start its countdown timer.
        StartShootAgain m_shoot_again.Value, m_grace_period.Value, m_hurry_up.Value
    End Sub

    Sub StartShootAgain(shoot_again_ms, grace_period_ms, hurry_up_time_ms)
        'Set callbacks for shoot again, grace period, and hurry up, if values above 0 are provided.
        'This is started for both beginning multiball ball save and add a ball ball save
        If shoot_again_ms > 0 Then
            Log("Starting ball save timer: " & shoot_again_ms)
            SetDelay m_name&"_disable_shoot_again", "MultiballsHandler" , Array(Array("stop", Me),Null), shoot_again_ms+grace_period_ms
        End If
        If grace_period_ms > 0 Then
            m_grace_period_enabled = True
            SetDelay m_name&"_grace_period", "MultiballsHandler" , Array(Array("grace_period", Me),Null), shoot_again_ms
        End If
        If hurry_up_time_ms > 0 Then
            m_hurry_up_enabled = True
            SetDelay m_name&"_hurry_up", "MultiballsHandler" , Array(Array("hurry_up", Me),Null), shoot_again_ms - hurry_up_time_ms
        End If
    End Sub

    Sub RunHurryUp()
        Log("Starting Hurry Up")
        m_hurry_up_enabled = False
        DispatchPinEvent m_name & "_hurry_up", Null
    End Sub

    Sub RunGracePeriod()
        Log("Starting Grace Period")
        m_grace_period_enabled = False
        DispatchPinEvent m_name & "_grace_period", Null
    End Sub

    Public Function BallDrainShootAgain(balls):
        Dim balls_to_save, kwargs

        If balls = 0 Then
            BallDrainShootAgain = balls
            Exit Function
        End If

        balls_to_save = m_balls_live_target - balls

        Log "Balls to save: " & balls_to_save & ". Balls live target: " & m_balls_live_target & ". Balls in Play: " & glf_BIP & ". Balls Drained: " & balls
        
        If balls_to_save <= 0 Then
            BallDrainShootAgain = balls
        End If

        If balls_to_save > balls Then
            balls_to_save = balls
        End If

        Set kwargs = GlfKwargs()
        With kwargs
            .Add "balls", balls_to_save
        End With
        DispatchPinEvent m_name & "_shoot_again", kwargs
        
        Log("Ball drained during MB. Requesting a new one")
        m_queued_balls = m_queued_balls + 1
        SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me, m_queued_balls),Null), 1000

        BallDrainShootAgain = balls - balls_to_save
    End Function

    Function BallDrainCountBalls(balls):
        DispatchPinEvent m_name & "_ball_lost", Null
        If not glf_gameStarted or (glf_BIP - balls) = 1 Then
            m_balls_added_live = 0
            m_balls_live_target = 0
            DispatchPinEvent m_name & "_ended", Null
            RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
            Log("Ball drained. MB ended.")
        End If
        BallDrainCountBalls = balls
    End Function

    Public Sub Reset()
        Log "Resetting multiball: " & m_name
        DispatchPinEvent m_name & "_reset_event", Null

        Disable()
        m_shoot_again_enabled = False
        m_balls_added_live = 0
        m_balls_live_target = 0
    End Sub

    Public Sub AddABall()
        If m_balls_live_target > 0 Then
            Log "Adding a ball to multiball: " & m_name
            m_balls_live_target = m_balls_live_target + 1
            m_balls_added_live = m_balls_added_live + 1
            m_queued_balls = m_queued_balls + 1
            glf_BIP = glf_BIP + 1
            SetDelay m_name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", Me, m_queued_balls),Null), 1000
        End If
    End Sub

    Public Sub AddAballTimerStart()
        'Start the timer for add a ball ball save.
        'This is started when multiball add a ball is triggered if configured,
        'and the default timer is not still running.
        If m_shoot_again_enabled = True Then
            Exit Sub
        End If

        m_shoot_again_enabled = True

        Dim shoot_again_ms : shoot_again_ms = m_add_a_ball_shoot_again.Value()
        if shoot_again_ms = 0 Then
            'No shoot again. Just stop multiball right away
            StopMultiball()
            Exit Sub
        End If

        DispatchPinEvent "ball_save_" & m_configname & "_add_a_ball_timer_start", Null

        Dim grace_period_ms : grace_period_ms = m_add_a_ball_grace_period.Value()
        Dim hurry_up_time_ms : hurry_up_time_ms = m_add_a_ball_hurry_up_time.Value()
        StartShootAgain shoot_again_ms, grace_period_ms, hurry_up_time_ms
    End Sub

    Public Sub StopMultiball()
        '"""Stop shoot again."""
        Log("Stopping shoot again of multiball")
        m_shoot_again_enabled = False

        '# disable shoot again
        RemoveDelay m_name&"_disable_shoot_again"

        If m_grace_period_enabled Then
            RemoveDelay m_name&"_grace_period"
            RunGracePeriod()
        End If
        If m_hurry_up_enabled Then
            RemoveDelay m_name&"_hurry_up"
            RunHurryUp()
        End If
        Log "Stop Shoot Again, Queued Balls: " & QueuedBalls()
        'RemoveDelay m_name & "_queued_release"

        DispatchPinEvent m_name & "_shoot_again_ended", Null
    End Sub

    Public Function QueuedBalls()
        QueuedBalls = m_queued_balls
    End Function

    Public Function ReleaseQueuedBalls()
        m_queued_balls = m_queued_balls - 1
        Log "Queued Balls: " & m_queued_balls
        ReleaseQueuedBalls = m_queued_balls
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function MultiballsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim multiball : Set multiball = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set MultiballsHandler = kwargs
                Else
                    MultiballsHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "start"
            multiball.Start
        Case "reset"
            multiball.Reset
        Case "add_a_ball"
            multiball.AddABall
        Case "stop"
            multiball.StopMultiball
        Case "grace_period"
            multiball.RunGracePeriod
        Case "hurry_up"
            multiball.RunHurryUp
        Case "drain"
            kwargs = multiball.BallsDrained(kwargs)
        Case "queue_release"
            If multiball.QueuedBalls() > 0 Then
                If glf_plunger.HasBall = False And ballInReleasePostion = True And glf_plunger.IncomingBalls = 0 Then
                    Glf_ReleaseBall(Null)
                    SetDelay multiball.Name&"_auto_launch", "MultiballsHandler" , Array(Array("auto_launch", multiball),Null), 500
                    If multiball.ReleaseQueuedBalls() > 0 Then
                        SetDelay multiball.Name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", multiball), Null), 1000    
                    End If
                Else
                    SetDelay multiball.Name&"_queued_release", "MultiballsHandler" , Array(Array("queue_release", multiball), Null), 1000
                End If
            End If
        Case "auto_launch"
            If glf_plunger.HasBall = True Then
                glf_plunger.Eject
            Else
                SetDelay multiball.Name&"_auto_launch", "MultiballsHandler" , Array(Array("auto_launch", multiball), Null), 500
            End If
    End Select

    If IsObject(args(1)) Then
        Set MultiballsHandler = kwargs
    Else
        MultiballsHandler = kwargs
    End If
End Function



Class GlfQueueEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "queue_event_player" : End Property

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "queue_event_player", Me)
        Set Init = Me
	End Function

    Public Sub Add(key, value)
        Dim newEvent : Set newEvent = (new GlfEvent)(key)
        m_events.Add newEvent.Name, newEvent
        'msgbox newEvent.Name
        m_eventValues.Add newEvent.Name, value  
    End Sub

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_queue_event_player_play", "QueueEventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_events(evt).Name & "_queue_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        If m_events(evt).Evaluate() = False Then
            Exit Sub
        End If
        Dim evtValue
        For Each evtValue In m_eventValues(evt)
            Log "Dispatching Event: " & evtValue
            DispatchQueuePinEvent evtValue, Null
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_queue_event_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & m_events(key).Raw & ": " & vbCrLf
                For Each evt in m_eventValues(key)
                    yaml = yaml & "    - " & evt & vbCrLf
                Next
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function QueueEventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            eventPlayer.FireEvent ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set QueueEventPlayerEventHandler = kwargs
    Else
        QueueEventPlayerEventHandler = kwargs
    End If
End Function


Class GlfQueueRelayPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "queue_relay_player" : End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "queue_relay_player", Me)
        Set Init = Me
	End Function

    Public Property Get Events() : Set Events = m_events : End Property
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    Public Property Get EventName(name)
        If m_events.Exists(name) Then
            Set EventName = m_eventValues(name)
        Else
            Dim new_event : Set new_event = (new GlfEvent)(name)
            m_events.Add new_event.Raw, new_event
            Dim new_event_value : Set new_event_value = (new GlfQueueRelayEvent)()
            m_eventValues.Add new_event.Raw, new_event_value
            Set EventName = new_event_value
        End If
    End Property

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_queue_relay_player_play", "QueueRelayPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_queue_relay_player_play"
        Next
    End Sub

    Public Function FireEvent(evt)
        FireEvent=Empty
        If m_events(evt).Evaluate() Then
            'post a new event, and wait for the release
            DispatchPinEvent m_eventValues(evt).Post, Null
            FireEvent = m_eventValues(evt).WaitFor
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_queue_relay_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & m_events(key).Raw & ": " & vbCrLf
                For Each evt in m_eventValues(key)
                    yaml = yaml & "    - " & evt & vbCrLf
                Next
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function QueueRelayPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            Dim wait_for : wait_for = eventPlayer.FireEvent(ownProps(2))
            If Not IsEmpty(wait_for) Then
                kwargs.Add "wait_for", wait_for
            End If
    End Select
    If IsObject(args(1)) Then
        Set QueueRelayPlayerEventHandler = kwargs
    Else
        QueueRelayPlayerEventHandler = kwargs
    End If
End Function

Class GlfQueueRelayEvent

	Private m_wait_for, m_post
  
    Public Property Get WaitFor() : WaitFor = m_wait_for : End Property
    Public Property Let WaitFor(input) : m_wait_for = input : End Property
    Public Property Get Post() : Post = m_post : End Property
    Public Property Let Post(input) : m_post = input : End Property
        
	Public default Function init()
        m_wait_for = Empty
        m_post = Empty
	    Set Init = Me
	End Function

End Class


Class GlfRandomEventPlayer

    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "random_event_player" : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get EventName(value)
        
        Dim newEvent : Set newEvent = (new GlfEvent)(value)
        m_events.Add newEvent.Raw, newEvent
        Dim newRandomEvent : Set newRandomEvent = (new GlfRandomEvent)(value, m_mode, UBound(m_events.Keys))
        m_eventValues.Add newEvent.Raw, newRandomEvent
        
        Set EventName = newRandomEvent
    End Property

	Public default Function init(mode)
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "random_event_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        Log "Activating"
        For Each evt In m_events.Keys()
            Log "Adding: " & m_events(evt).EventName & ". For Key: " & m_mode & "_" & evt & "_random_event_player_play"
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_random_event_player_play", "RandomEventPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_random_event_player_play"
        Next
    End Sub

    Public Sub FireEvent(evt)
        Log "Firing Random Event:  " & evt
        If m_events(evt).Evaluate() Then
            Dim event_to_fire
            event_to_fire = m_eventValues(evt).GetNextRandomEvent()
            If Not IsEmpty(event_to_fire) Then
                Log "Dispatching Event: " & event_to_fire
                DispatchPinEvent event_to_fire, Null
            Else
                Log "No event available to fire"
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_random_event_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml : yaml = ""
        ToYaml = yaml
    End Function

End Class

Function RandomEventPlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim eventPlayer : Set eventPlayer = ownProps(1)
    Select Case evt
        Case "play"
            eventPlayer.FireEvent ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set RandomEventPlayerEventHandler = kwargs
    Else
        RandomEventPlayerEventHandler = kwargs
    End If
End Function

Class GlfSegmentDisplayPlayer

    Private m_priority
    Private m_mode
    Private m_name
    Private m_debug
    Private m_events
    private m_base_device

    Public Property Get Name() : Name = "segment_player" : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
    Public Property Get EventNames() : EventNames = m_events.Keys() : End Property    
    
    Public Property Get EventName(value)
        Dim newEvent : Set newEvent = (new GlfEvent)(value)

        If m_events.Exists(newEvent.Raw) Then
            Set EventName = m_events(newEvent.Raw)
        Else
            Dim new_segment_event : Set new_segment_event = (new GlfSegmentDisplayPlayerEvent)(newEvent)
            m_events.Add newEvent.Raw, new_segment_event
            Set EventName = new_segment_event
        End If
    End Property

	Public default Function init(mode)
        m_name = "segment_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "segment_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).GlfEvent.EventName, m_mode & "_" & evt & "_segment_player_play", "SegmentPlayerEventHandler", m_priority+m_events(evt).GlfEvent.Priority, Array("play", Me, m_events(evt), m_events(evt).GlfEvent.EventName)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt, display
        Dim displays_to_update : Set displays_to_update = CreateObject("Scripting.Dictionary")
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).GlfEvent.EventName, m_mode & "_" & evt & "_segment_player_play"
            Set displays_to_update = PlayOff(m_events(evt).GlfEvent.EventName, m_events(evt), displays_to_update)
        Next
        
        For Each display in displays_to_update.Keys()
            glf_segment_displays(display).UpdateStack()
        Next
    End Sub

    Public Sub Play(evt, segment_event)
        Dim i
        For i=0 to UBound(segment_event.Displays())
            SegmentPlayerCallbackHandler evt, segment_event.Displays()(i), m_mode, m_priority
        Next
    End Sub

    Public Function PlayOff(evt, segment_event, displays_to_update)
        Dim i, segment_item
        For i=0 to UBound(segment_event.Displays())
            Set segment_item = segment_event.Displays()(i)
            Dim key
            key = m_mode & "." & "segment_player_player." & segment_item.Display
            If Not IsEmpty(segment_item.Key) Then
                key = key & segment_item.Key
            End If
            Dim display : Set display = glf_segment_displays(segment_item.Display)
            RemoveDelay key
            display.RemoveTextByKeyNoUpdate key
            If Not displays_to_update.Exists(segment_item.Display) Then
                displays_to_update.Add segment_item.Display, True
            End If 
        Next
        Set PlayOff = displays_to_update
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml()
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfSegmentDisplayPlayerEvent

    Private m_items
    Private m_event

    Public Property Get Displays() : Displays = m_items.Items() : End Property

    Public Property Get GlfEvent() : Set GlfEvent = m_event : End Property

    Public Property Get Display(value)
        If m_items.Exists(value) Then
            Set Display = m_items(value)
        Else
            Dim new_item : Set new_item = (new GlfSegmentPlayerEventItem)()
            new_item.Display = value
            m_items.add value, new_item
            Set Display = new_item
        End If
    End Property

    Public default Function init(evt)
        Set m_items = CreateObject("Scripting.Dictionary")
        Set m_event = evt
        Set Init = Me
	End Function

End Class

Class GlfSegmentPlayerEventItem
	
    private m_display
    private m_text
    private m_priority
    private m_action
    private m_expire
    private m_flash_mask
    private m_flashing
    private m_key
    private m_transition
    private m_transition_out
    private m_color

    Public Property Get Display() : Display = m_display : End Property
    Public Property Let Display(input) : m_display = input : End Property
    
    Public Property Get Text()
        If Not IsNull(m_text) Then
            Set Text = m_text
        Else
            Text = Null
        End If
    End Property
    Public Property Let Text(input) 
        Set m_text = (new GlfInput)(input)
    End Property

    Public Property Get Priority() : Priority = m_priority : End Property
    Public Property Let Priority(input) : m_priority = input : End Property
        
    Public Property Get Action() : Action = m_action : End Property
    Public Property Let Action(input) : m_action = input : End Property
                
    Public Property Get Expire() : Expire = m_expire : End Property
    Public Property Let Expire(input) : m_expire = input : End Property

    Public Property Get FlashMask() : FlashMask = m_flash_mask : End Property
    Public Property Let FlashMask(input) : m_flash_mask = input : End Property
                        
    Public Property Get Flashing() : Flashing = m_flashing : End Property
    Public Property Let Flashing(input) : m_flashing = input : End Property
                            
    Public Property Get Key() : Key = m_key : End Property
    Public Property Let Key(input) : m_key = input : End Property

    Public Property Get Color() : Color = m_color : End Property
    Public Property Let Color(input) : m_color = input : End Property

    Public Property Get HasTransition() : HasTransition = Not IsNull(m_transition) : End Property    
    Public Property Get HasTransitionOut() : HasTransitionOut = Not IsNull(m_transition_out) : End Property

    Public Property Get Transition()
        If IsNull(m_transition) Then
            Set m_transition = (new GlfSegmentPlayerTransition)("bee")
            Set Transition = m_transition   
        Else
            Set Transition = m_transition
        End If
    End Property

    Public Property Get TransitionOut()
        If IsNull(m_transition_out) Then
            Set m_transition_out = (new GlfSegmentPlayerTransition)()
            Set TransitionOut = m_transition_out   
        Else
            Set TransitionOut = m_transition_out
        End If
    End Property
                                
	Public default Function init()
        m_display = Empty
        m_text = Null
        m_priority = 0
        m_action = "add"
        m_expire = 0
        m_flash_mask = Empty
        m_flashing = "no_flash"
        m_key = Empty
        m_transition = Null
        m_transition_out = Null
        m_color = Rgb(255,255,255)
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        If Not IsEmpty(m_display) Then
            yaml = yaml & "    " & m_display & ": " & vbCrLf
        End If
        If Not IsNull(m_text) Then
            yaml = yaml & "    " & m_text.Raw() & ": " & vbCrLf
        End If
        If m_priority > 0 Then
            yaml = yaml & "    " & m_priority & ": " & vbCrLf
        End If
        If m_action <> "add" Then
            yaml = yaml & "    " & m_action & ": " & vbCrLf
        End If
        If m_expire > 0 Then
            yaml = yaml & "    " & m_expire & ": " & vbCrLf
        End If
        If Not IsEmpty(m_flash_mask) Then
            yaml = yaml & "    " & m_flash_mask & ": " & vbCrLf
        End If
        If m_flashing <> "not_set" Then
            yaml = yaml & "    " & m_flashing & ": " & vbCrLf
        End If
        If Not IsEmpty(m_key) Then
            yaml = yaml & "    " & m_key & ": " & vbCrLf
        End If
        If Not IsEmpty(m_color) Then
            yaml = yaml & "    " & m_color & ": " & vbCrLf
        End If
        If Not IsNull(m_transition) Then
            yaml = yaml & m_transition.ToYaml()
        End If
        If Not IsNull(m_transition_out) Then
            yaml = yaml & m_transition_out.ToYaml()
        End If
        ToYaml = yaml
    End Function

End Class

Class GlfSegmentPlayerTransition
	
    private m_type
    private m_text
    private m_direction

    Public Property Get TransitionType() : TransitionType = m_type : End Property
    Public Property Let TransitionType(input) : m_type = input : End Property
    
    Public Property Get Text()
        Text = m_text
    End Property
    Public Property Let Text(input)
        m_text = input
    End Property

    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let Direction(input) : m_direction = input : End Property                          

	Public default Function Init(loo)
        m_type = "push"
        m_text = Empty
        m_direction = "right"
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    transition:" & vbCrLf
        yaml = yaml & "      " & m_type & ": " & vbCrLf
        yaml = yaml & "      " & m_direction & ": " & vbCrLf
        yaml = yaml & "      " & m_text & ": " & vbCrLf
        ToYaml = yaml
    End Function

End Class

Function SegmentPlayerEventHandler(args)
    Dim ownProps : ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim SegmentPlayer : Set SegmentPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            SegmentPlayer.Activate
        Case "deactivate"
            SegmentPlayer.Deactivate
        Case "play"
            If ownProps(2).GlfEvent.Evaluate() Then
                SegmentPlayer.Play ownProps(3), ownProps(2)
            End If
        Case "remove"
            RemoveDelay ownProps(2)
            SegmentPlayer.RemoveTextByKey ownProps(2)
    End Select
    SegmentPlayerEventHandler = Null
End Function


Function SegmentPlayerCallbackHandler(evt, segment_item, mode, priority)

    If IsObject(segment_item) Then
        'Shot Text on a display
        Dim key
        key = mode & "." & "segment_player_player." & segment_item.Display
        If Not IsEmpty(segment_item.Key) Then
            key = key & segment_item.Key
        End If

        Dim display : Set display = glf_segment_displays(segment_item.Display)
        
        If segment_item.Action = "add" Then
            RemoveDelay key
            Dim transition, transition_out : transition = Null : transition_out = Null
            If segment_item.HasTransition() Then
                Set transition = segment_item.Transition
            End If
            If segment_item.HasTransitionOut() Then
                Set transition_out = segment_item.TransitionOut
            End If
            display.AddTextEntry segment_item.Text, segment_item.Color, segment_item.Flashing, segment_item.FlashMask, transition, transition_out, priority + segment_item.Priority, key
                                
            If segment_item.Expire > 0 Then
                SetDelay key & "_expire", "SegmentPlayerEventHandler",  Array(Array("remove", display, key)), segment_item.Expire
            End If

        ElseIf segment_item.Action = "remove" Then
            RemoveDelay key
            display.RemoveTextByKey key        
        ElseIf segment_item.Action = "flash" Then
            display.SetFlashing "all"
        ElseIf segment_item.Action = "flash_match" Then
            display.SetFlashing "match"
        ElseIf segment_item.Action = "flash_mask" Then
            display.SetFlashingMask segment_item.FlashMask
        ElseIf segment_item.Action = "no_flash" Then
            display.SetFlashing "no_flash"
        ElseIf segment_item.Action = "set_color" Then
            If Not IsNull(segment_item.Color) Then
                display.SetColor segment_item.Color
            End If
        End If
    End If

End Function

Class GlfSequenceShots

    Private m_name
    Private m_command_name
    Private m_lock_device
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_cancel_events
    Private m_cancel_switches
    Private m_delay_event_list
    Private m_delay_switch_list
    Private m_event_sequence
    Private m_sequence_timeout
    Private m_switch_sequence
    Private m_start_event
    Private m_sequence_count
    Private m_active_delays
    Private m_active_sequences
    Private m_sequence_events
    Private m_start_time

    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
    Public Property Get GetValue(value)
        Select Case value
            'Case "":
            '    GetValue = 
        End Select
    End Property

    Public Property Let CancelEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_cancel_events.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let CancelSwitches(value): m_cancel_switches = value: End Property
    Public Property Let DelayEventList(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_delay_event_list.Add newEvent.Name, newEvent
        Next
    End Property
    Public Property Let DelaySwitchList(value): m_delay_switch_list = value: End Property
    Public Property Let EventSequence(value)
        m_event_sequence = value
        If m_sequence_count = 0 Then
            m_sequence_events = value
        Else
            Redim Preserve m_sequence_events(m_sequence_count+UBound(value))
            Dim i
            For i = 0 To UBound(value)
                m_sequence_events(m_sequence_count + i) = value(i)
            Next
        End If
        m_start_event = value(0)
        m_sequence_count = m_sequence_count + (UBound(m_sequence_events)+1)
    End Property
    Public Property Let SequenceTimeout(value): Set m_sequence_timeout = CreateGlfInput(value): End Property
    Public Property Let SwitchSequence(value)
        m_switch_sequence = value
        If m_sequence_count = 0 Then
            m_start_event = value(0) & "_active"
        End If
        Redim Preserve m_sequence_events(m_sequence_count+UBound(value))
        Dim i
        For i = 0 To UBound(value)
            m_sequence_events(m_sequence_count + i) = value(i) & "_active"
        Next
        m_sequence_count = m_sequence_count + (UBound(m_sequence_events)+1)
    End Property

	Public default Function init(name, mode)
        m_name = "sequence_shot_" & name
        m_command_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        
        Set m_cancel_events = CreateObject("Scripting.Dictionary")
        Set m_delay_event_list = CreateObject("Scripting.Dictionary")
        Set m_active_sequences = CreateObject("Scripting.Dictionary")
        Set m_active_delays = CreateObject("Scripting.Dictionary")
        
        m_sequence_events = Array()
        m_cancel_switches = Array()
        m_start_time = 0
        m_event_sequence = Array()
        m_switch_sequence = Array()
        Set m_sequence_timeout = CreateGlfInput(0)
        m_sequence_count = 0
        m_start_event = Empty
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "sequence_shot_", Me)
        
        Set Init = Me
	End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_event_sequence
            AddPinEventListener evt, m_name & "_" & evt & "_advance", "SequenceShotsHandler", m_priority, Array("advance", Me, evt)
        Next
        For Each evt in m_switch_sequence
            AddPinEventListener evt & "_active", m_name & "_" & evt & "_advance", "SequenceShotsHandler", m_priority, Array("advance", Me, evt & "_active")
        Next
        For Each evt in m_cancel_events.Keys
            AddPinEventListener m_cancel_events(evt).EventName, m_name & "_" & evt & "_cancel", "SequenceShotsHandler", m_priority+m_cancel_events(evt).Priority, Array("cancel_event", Me, m_cancel_events(evt))
        Next
        For Each evt in m_delay_event_list.Keys
            AddPinEventListener m_delay_event_list(evt).EventName, m_name & "_" & evt & "_delay", "SequenceShotsHandler", m_priority+m_delay_event_list(evt).Priority, Array("delay_event", Me, m_delay_event_list(evt))
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_event_sequence
            RemovePinEventListener evt, m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_switch_sequence
            RemovePinEventListener evt & "_active", m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_cancel_events.Keys
            RemovePinEventListener m_cancel_events(evt).EventName, m_name & "_" & evt & "_cancel"
        Next
        For Each evt in  m_delay_event_list.Keys
            RemovePinEventListener m_delay_event_list(evt).EventName, m_name & "_" & evt & "_delay"
        Next
    End Sub

    Sub SequenceAdvance(event_name)
        ' Since we can track multiple simultaneous sequences (e.g. two balls
        ' going into an orbit in a row), we first have to see whether this
        ' switch is starting a new sequence or continuing an existing one

        Log "Sequence advance: " & event_name

        If event_name = m_start_event Then
            If m_sequence_count > 1 Then
                ' start a new sequence
                StartNewSequence()
            ElseIf UBound(m_active_delays.Keys) = -1 Then
                ' if it only has one step it will finish right away
                Completed()
            End If
        Else
            ' Get the seq_id of the first sequence this switch is next for.
            ' This is not a loop because we only want to advance 1 sequence
            Dim k, seq
            seq = Null
            For Each k In m_active_sequences.Keys
                Log m_active_sequences(k).NextEvent
                If m_active_sequences(k).NextEvent = event_name Then
                    Set seq = m_active_sequences(k)
                    Exit For
                End If
            Next

            If Not IsNull(seq) Then
                ' advance this sequence
                AdvanceSequence(seq)
            End If
        End If
    End Sub

    Public Sub StartNewSequence()
        ' If the sequence hasn't started, make sure we're not within the
        ' delay_switch hit window

        If UBound(m_active_delays.Keys)>-1 Then
            Log "There's a delay timer in effect. Sequence will not be started."
            Exit Sub
        End If

        'record start time
        m_start_time = gametime

        ' create a new sequence
        Dim seq_id : seq_id = "seq_" & glf_SeqCount
        glf_SeqCount = glf_SeqCount + 1

        Dim next_event : next_event = m_sequence_events(1)

        Log "Setting up a new sequence. Next: " & next_event

        m_active_sequences.Add seq_id, (new GlfActiveSequence)(seq_id, 0, next_event)

        ' if this sequence has a time limit, set that up
        If m_sequence_timeout.Value > 0 Then
            Log "Setting up a sequence timer for " & m_sequence_timeout.Value
            SetDelay seq_id, "SequenceShotsHandler" , Array(Array("seq_timeout", Me, seq_id),Null), m_sequence_timeout.Value
        End If
    End Sub

    Public Sub AdvanceSequence(sequence)
        ' Remove this sequence from the list
        If sequence.CurrentPositionIndex = (m_sequence_count - 2) Then  ' complete
            Log "Sequence complete!"
            RemoveDelay sequence.SeqId
            m_active_sequences.Remove sequence.SeqId
            Completed()
        Else
            Dim current_position_index : current_position_index = sequence.CurrentPositionIndex + 1
            Dim next_event : next_event = m_sequence_events(current_position_index + 1)
            Log "Advancing the sequence. Next: " & next_event
            sequence.CurrentPositionIndex = current_position_index
            sequence.NextEvent = next_event
        End If
    End Sub

    Public Sub Completed()
        'measure the elapsed time between start and completion of the sequence
        Dim elapsed
        If m_start_time > 0 Then
            elapsed = gametime - m_start_time
        Else
            elapsed = 0
        End If

        'Post sequence complete event including its elapsed time to complete.
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "elapsed", elapsed
        End With
        DispatchPinEvent m_command_name & "_hit", kwargs
    End Sub

    Public Sub ResetAllSequences()
        'Reset all sequences."""
        Dim k
        For Each k in m_active_sequences.Keys
            RemoveDelay m_active_sequences(k).SeqId
        Next

        m_active_sequences.RemoveAll()
    End Sub

    Public Sub DelayEvent(delay, name)
        Log "Delaying sequence by " & delay
        SetDelay name & "_delay_timer", "SequenceShotsHandler" , Array(Array("release_delay", Me, name),Null), delay
        m_active_delays.Add name, True
    End Sub

    Public Sub ReleaseDelay(name)
        m_active_delays.Remove name
    End Sub

    Public Sub FireSequenceTimeout(seq_id)
        Log "Sequence " & seq_id & " timeouted"
        m_active_sequences.Remove seq_id
        DispatchPinEvent m_name & "_timeout", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Class GlfActiveSequence

    Private m_next_event, m_seq_id, m_idx

    Public Property Get SeqId(): SeqId = m_seq_id: End Property
    Public Property Get NextEvent(): NextEvent = m_next_event: End Property
    Public Property Let NextEvent(value): m_next_event = value: End Property
    Public Property Get CurrentPositionIndex(): CurrentPositionIndex = m_idx: End Property
    Public Property Let CurrentPositionIndex(value): m_idx = value: End Property

    Public default Function init(seq_id, idx, next_event)
        m_seq_id = seq_id
        m_idx = idx
        m_next_event = next_event
        Set Init = Me
    End Function

End Class

Function SequenceShotsHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim sequence_shot : Set sequence_shot = ownProps(1)
    Select Case evt
        Case "advance"
            sequence_shot.SequenceAdvance ownProps(2)
        Case "cancel"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            sequence_shot.ResetAllSequences
        Case "delay"
            Set glfEvent = ownProps(2)
            If glfEvent.Evaluate() = False Then
                Exit Function
            End If
            sequence_shot.DelayEvent glfEvent.Delay, glfEvent.EventName
        Case "seq_timeout"
            sequence_shot.FireSequenceTimeout ownProps(2)
        Case "release_delay"
            sequence_shot.ReleaseDelay ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set SequenceShotsHandler = kwargs
    Else
        SequenceShotsHandler = kwargs
    End If
End Function
Class GlfShotGroup
    Private m_name
    Private m_mode
    Private m_priority
    private m_base_device
    private m_debug
    private m_shots
    private m_common_state
    Private m_enable_rotation_events
    Private m_disable_rotation_events
    Private m_restart_events
    Private m_reset_events
    Private m_rotate_events
    Private m_rotate_left_events
    Private m_rotate_right_events
    Private rotation_enabled
    Private m_temp_shots
    Private m_rotation_pattern
    Private m_rotation_enabled
    Private m_isRotating
 
    Public Property Get Name(): Name = m_name: End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public Property Get CommonState()
        Dim state : state = m_base_device.Mode.Shots(m_shots(0)).State
        Dim shot
        For Each shot in m_shots
            If state <> m_base_device.Mode.Shots(shot).State Then
                CommonState = Empty
                Exit Property
            End If
        Next
        CommonState = state
    End Property
 
    Public Property Let Shots(value)
        m_shots = value
        m_rotation_pattern = Glf_CopyArray(Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).RotationPattern)
    End Property
 
    Public Property Let EnableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_enable_rotation_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let DisableRotationEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_disable_rotation_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RotateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RotateLeftEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_left_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RotateRightEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_rotate_right_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
 
	Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_common_state = Empty
        m_rotation_enabled = True
        m_rotation_pattern = Empty
        m_isRotating = False
 
        Set m_enable_rotation_events = CreateObject("Scripting.Dictionary")
        Set m_disable_rotation_events = CreateObject("Scripting.Dictionary")
        Set m_restart_events = CreateObject("Scripting.Dictionary")
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_left_events = CreateObject("Scripting.Dictionary")
        Set m_rotate_right_events = CreateObject("Scripting.Dictionary")
        Set m_temp_shots = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot_group", Me)
 
        Set Init = Me
	End Function
 
    Public Sub Activate
        Dim evt
        For Each evt in m_enable_rotation_events.Keys()
            m_rotation_enabled = False
            AddPinEventListener m_enable_rotation_events(evt).EventName, m_name & "_" & evt & "_enable_rotation", "ShotGroupEventHandler", m_priority+m_enable_rotation_events(evt).Priority, Array("enable_rotation", Me, m_enable_rotation_events(evt))
        Next
        For Each evt in m_disable_rotation_events.Keys()
            AddPinEventListener m_disable_rotation_events(evt).EventName, m_name & "_" & evt & "_disable_rotation", "ShotGroupEventHandler", m_priority+m_disable_rotation_events(evt).Priority, Array("disable_rotation", Me, m_disable_rotation_events(evt))
        Next
        For Each evt in m_restart_events.Keys()
            AddPinEventListener m_restart_events(evt).EventName, m_name & "_" & evt & "_restart", "ShotGroupEventHandler", m_priority+m_restart_events(evt).Priority, Array("restart", Me, m_restart_events(evt))
        Next
        For Each evt in m_reset_events.Keys()
            AddPinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset", "ShotGroupEventHandler", m_priority+m_reset_events(evt).Priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_rotate_events.Keys()
            AddPinEventListener m_rotate_events(evt).EventName, m_name & "_" & evt & "_rotate", "ShotGroupEventHandler", m_priority+m_rotate_events(evt).Priority, Array("rotate", Me, m_rotate_events(evt))
        Next
        For Each evt in m_rotate_left_events.Keys()
            AddPinEventListener m_rotate_left_events(evt).EventName, m_name & "_" & evt & "_rotate_left", "ShotGroupEventHandler", m_priority+m_rotate_left_events(evt).Priority, Array("rotate_left", Me, m_rotate_left_events(evt))
        Next
        For Each evt in m_rotate_right_events.Keys()
            AddPinEventListener m_rotate_right_events(evt).EventName, m_name & "_" & evt & "_rotate_right", "ShotGroupEventHandler", m_priority+m_rotate_right_events(evt).Priority, Array("rotate_right", Me, m_rotate_right_events(evt))
        Next
        Dim shot_name
        For Each shot_name in m_shots
            AddPinEventListener shot_name & "_hit", m_name & "_" & m_mode & "_hit", "ShotGroupEventHandler", m_priority, Array("hit", Me, Null)
            AddPlayerStateEventListener "shot_" & shot_name, m_name & "_" & m_mode & "_complete", -1, "ShotGroupEventHandler", m_priority, Array("complete", Me, Null)
        Next
    End Sub
 
    Public Sub Deactivate
        Dim evt
        m_rotation_enabled = True
        For Each evt in m_enable_rotation_events.Keys()
            RemovePinEventListener m_enable_rotation_events(evt).EventName, m_name & "_" & evt & "_enable_rotation"
        Next
        For Each evt in m_disable_rotation_events.Keys()
            RemovePinEventListener m_disable_rotation_events(evt).EventName, m_name & "_" & evt & "_disable_rotation"
        Next
        For Each evt in m_restart_events.Keys()
            RemovePinEventListener m_restart_events(evt).EventName, m_name & "_" & evt & "_restart"
        Next
        For Each evt in m_reset_events.Keys()
            RemovePinEventListener m_reset_events(evt).EventName, m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_rotate_events.Keys()
            RemovePinEventListener m_rotate_events(evt).EventName, m_name & "_" & evt & "_rotate"
        Next
        For Each evt in m_rotate_left_events.Keys()
            RemovePinEventListener m_rotate_left_events(evt).EventName, m_name & "_" & evt & "_rotate_left"
        Next
        For Each evt in m_rotate_right_events.Keys()
            RemovePinEventListener m_rotate_right_events(evt).EventName, m_name & "_" & evt & "_rotate_right"
        Next
        Dim shot_name
        For Each shot_name in m_shots
            RemovePinEventListener shot_name & "_hit", m_name & "_" & m_mode & "_hit"
            RemovePlayerStateEventListener "shot_" & shot_name, m_name & "_" & m_mode & "_complete"
        Next
    End Sub
 
    Public Function CheckForComplete()
        If m_isRotating Then
            Exit Function
        End If
        Dim state : state = CommonState()
        If state = m_common_state Then
            Exit Function
        End If
 
        m_common_state = state
 
        If state = Empty Then
            Exit Function
        End If
 
        Dim state_name : state_name = Glf_ShotProfiles(m_base_device.Mode.Shots(m_shots(0)).Profile).StateName(m_common_state)
 
        Log "Shot group is complete with state: " & state_name
        Dim kwargs : Set kwargs = GlfKwargs()
		With kwargs
            .Add "state", state_name
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        DispatchPinEvent m_name & "_" & state_name & "_complete", Null
 
    End Function
 
    Public Sub Enable()
        Dim shot
        Log "Enabling"
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Enable()
        Next
        Dim evt
    End Sub
 
    Public Sub Disable()
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Disable()
        Next
    End Sub
 
    Public Sub EnableRotation
        Log "Enabling Rotation"
        m_rotation_enabled = True
    End Sub
 
    Public Sub DisableRotation
        Log "Disabling Rotation"
        m_rotation_enabled = False
    End Sub
 
    Public Sub Restart
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Restart()
        Next
    End Sub
 
    Public Sub Reset
        Dim shot
        For Each shot in m_shots
            m_base_device.Mode.Shots(shot).Reset()
        Next
    End Sub
 
    Public Sub Rotate(direction)
 
        If m_rotation_enabled = False Then
            Exit Sub
        End If
        Dim shots_to_rotate : shots_to_rotate = Array()
 
        m_temp_shots.RemoveAll
        Dim shot
        For Each shot in m_shots
            If m_base_device.Mode.Shots(shot).CanRotate Then
                m_temp_shots.Add shot, m_base_device.Mode.Shots(shot)
            End If
        Next
 
        Dim shot_states, x
        x=0
        ReDim shot_states(UBound(m_temp_shots.Keys))
        For Each shot in m_temp_shots.Keys
            shot_states(x) = m_temp_shots(shot).State
            x=x+1
        Next 
 
        If direction = Empty Then
            direction = m_rotation_pattern(0)
            Glf_RotateArray m_rotation_pattern, "l"
        End If
 
        shot_states = Glf_RotateArray(shot_states, direction)
        x=0
        m_isRotating = True
        For Each shot in m_temp_shots.Keys
            Log "Rotating Shot:" & shot
            m_temp_shots(shot).Jump shot_states(x), True, False
            x=x+1
        Next 
        m_isRotating = False
        CheckForComplete()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
 
    Public Function ToYaml
        Dim yaml
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml & "    shots: " & Join(m_shots, ",") & vbCrLf
 
        If UBound(m_enable_rotation_events.Keys) > -1 Then
            yaml = yaml & "    enable_rotation_events: "
            x=0
            For Each key in m_enable_rotation_events.keys
                yaml = yaml & m_enable_rotation_events(key).Raw
                If x <> UBound(m_enable_rotation_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_disable_rotation_events.Keys) > -1 Then
            yaml = yaml & "    disable_rotation_events: "
            x=0
            For Each key in m_disable_rotation_events.keys
                yaml = yaml & m_disable_rotation_events(key).Raw
                If x <> UBound(m_disable_rotation_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_restart_events.Keys) > -1 Then
            yaml = yaml & "    restart_events: "
            x=0
            For Each key in m_restart_events.keys
                yaml = yaml & m_restart_events(key).Raw
                If x <> UBound(m_restart_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_reset_events.Keys) > -1 Then
            yaml = yaml & "    reset_events: "
            x=0
            For Each key in m_reset_events.keys
                yaml = yaml & m_reset_events(key).Raw
                If x <> UBound(m_reset_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_events.Keys) > -1 Then
            yaml = yaml & "    rotate_events: "
            x=0
            For Each key in m_rotate_events.keys
                yaml = yaml & m_rotate_events(key).Raw
                If x <> UBound(m_rotate_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_left_events.Keys) > -1 Then
            yaml = yaml & "    rotate_left_events: "
            x=0
            For Each key in m_rotate_left_events.keys
                yaml = yaml & m_rotate_left_events(key).Raw
                If x <> UBound(m_rotate_left_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_rotate_right_events.Keys) > -1 Then
            yaml = yaml & "    rotate_right_events: "
            x=0
            For Each key in m_rotate_right_events.keys
                yaml = yaml & m_rotate_right_events(key).Raw
                If x <> UBound(m_rotate_right_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_base_device.EnableEvents.Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents.keys
                yaml = yaml & m_base_device.EnableEvents(key).Raw
                If x <> UBound(m_base_device.EnableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
        If UBound(m_base_device.DisableEvents.Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents.keys
                yaml = yaml & m_base_device.DisableEvents(key).Raw
                If x <> UBound(m_base_device.DisableEvents.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If
 
 
        ToYaml = yaml
        End Function
End Class
 
Function ShotGroupEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Dim device : Set device = ownProps(1)
    If Not IsNull(ownProps(2)) Then
        Dim glf_event : Set glf_event = ownProps(2)
        If glf_event.Evaluate() = False Then
            If IsObject(args(1)) Then
                Set ShotGroupEventHandler = kwargs
            Else
                ShotGroupEventHandler = kwargs
            End If
            Exit Function
        End If
    End If
    Select Case evt
        Case "hit"
            DispatchPinEvent device.Name & "_hit", Null
            DispatchPinEvent device.Name & "_" & kwargs("state") & "_hit", Null
        Case "complete"
            device.CheckForComplete
        Case "enable_rotation"
            device.EnableRotation
        Case "disable_rotation"
            device.DisableRotation
        Case "restart"
            device.Restart
        Case "reset"
            device.Reset
        Case "rotate"
            device.Rotate Empty
        Case "rotate_left"
            device.Rotate "l"
        Case "rotate_right"
            device.Rotate "r"
    End Select
    If IsObject(args(1)) Then
        Set ShotGroupEventHandler = kwargs
    Else
        ShotGroupEventHandler = kwargs
    End If
 
End Function

Class GlfShotProfile

    Private m_name
    Private m_advance_on_hit
    Private m_block
    Private m_loop
    Private m_rotation_pattern
    Private m_states
    Private m_states_not_to_rotate
    Private m_states_to_rotate

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get AdvanceOnHit(): AdvanceOnHit = m_advance_on_hit: End Property
    Public Property Get Block(): Block = m_block: End Property
    Public Property Let Block(input): m_block = input: End Property
    Public Property Get ProfileLoop(): ProfileLoop = m_loop: End Property
    Public Property Get RotationPattern(): RotationPattern = m_rotation_pattern: End Property
    Public Property Get States(name)
        If m_states.Exists(name) Then
            Set States = m_states(name)
        Else
            Dim new_state : Set new_state = (new GlfShowPlayerItem)()
            m_states.Add name, new_state
            Set States = new_state
        End If
    End Property
    Public Property Get StateForIndex(index)
        Dim stateItems : stateItems = m_states.Items()
        If UBound(stateItems) >= index Then
            Set StateForIndex = stateItems(index)
        Else
            StateForIndex = Null
        End If
    End Property
    Public Property Get StateName(index)
        Dim stateKeys : stateKeys = m_states.Keys()
        If UBound(stateKeys) >= index Then
            StateName = stateKeys(index)
        Else
            StateName = Empty
        End If
    End Property
    Public Property Get StatesCount()
        StatesCount = UBound(m_states.Keys())
    End Property

    Public Property Get StateNamesToRotate(): StateNamesToRotate = m_states_to_rotate: End Property
    Public Property Let StateNamesToRotate(input): m_states_to_rotate = input: End Property
    Public Property Get StateNamesNotToRotate(): StateNamesNotToRotate = m_states_not_to_rotate: End Property
    Public Property Let StateNamesNotToRotate(input): m_states_not_to_rotate = input: End Property
    
	Public default Function init(name)
        m_name = "shotprofile_" & name
        m_advance_on_hit = True
        m_block = False
        m_loop = False
        m_rotation_pattern = Array("r")
        m_states_to_rotate = Array()
        m_states_not_to_rotate = Array()
        Set m_states = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "  " & Replace(m_name, "shotprofile_", "") & ":" & vbCrLf
        yaml = yaml & "    states: " & vbCrLf
        Dim token,evt,state,x : x = 0
        For Each evt in m_states.Keys
            Set state = StateForIndex(x)
            yaml = yaml & "     - name: " & StateName(x) & vbCrLf
            yaml = yaml & "       show: " & state.Show.Name & vbCrLf
            yaml = yaml & "       loops: " & m_states(evt).Loops & vbCrLf
            yaml = yaml & "       speed: " & m_states(evt).Speed & vbCrLf
            yaml = yaml & "       sync_ms: " & m_states(evt).SyncMs & vbCrLf

            If Ubound(state.Tokens().Keys)>-1 Then
                yaml = yaml & "       show_tokens: " & vbCrLf
                For Each token in state.Tokens().Keys()
                    yaml = yaml & "         " & token & ": " & state.Tokens()(token) & vbCrLf
                Next
            End If

            'yaml = yaml & "     block: " & m_block & vbCrLf
            'yaml = yaml & "     advance_on_hit: " & m_advance_on_hit & vbCrLf
            'yaml = yaml & "     loop: " & m_loop & vbCrLf
            'yaml = yaml & "     rotation_pattern: " & m_rotation_pattern & vbCrLf
            'yaml = yaml & "     state_names_to_not_rotate: " & m_states_not_to_rotate & vbCrLf
            'yaml = yaml & "     state_names_to_rotate: " & m_states_to_rotate & vbCrLf
            x = x +1
        Next
        ToYaml = yaml
    End Function

End Class

Class GlfShot

    Private m_name
    Private m_mode
    Private m_priority
    Private m_base_device
    Private m_debug
    Private m_profile
    Private m_control_events
    Private m_advance_events
    Private m_reset_events
    Private m_restart_events
    Private m_switches
    Private m_tokens
    Private m_hit_events
    Private m_start_enabled
    Private m_show_cache
    Private m_state
    Private m_enabled
    Private m_player_var_name
    Private m_persist
    Private m_internal_cache_id

    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    Public Property Get Profile(): Profile = m_profile: End Property
    Public Property Get ShotKey(): ShotKey = m_name & "_" & m_profile: End Property
    Public Property Get State(): State = m_state: End Property
    Public Property Get Tokens() : Set Tokens = m_tokens : End Property
    Public Property Get CanRotate()
        If Glf_IsInArray(Glf_ShotProfiles(m_profile).StateName(m_state), Glf_ShotProfiles(m_profile).StateNamesNotToRotate) Then
            CanRotate = False
        Else
            CanRotate = True
        End If
    End Property
    
    Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property

    Public Property Let EnableEvents(value) : m_base_device.EnableEvents = value : End Property
    Public Property Let DisableEvents(value) : m_base_device.DisableEvents = value : End Property
    Public Property Get ControlEvents()
            Dim control_event_count : control_event_count = UBound(m_control_events.Keys)
            Dim newEvent : Set newEvent = (new GlfShotControlEvent)()
            m_control_events.Add CStr(control_event_count+1), newEvent
            Set ControlEvents = newEvent
    End Property
    Public Property Let AdvanceEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_advance_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ResetEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let RestartEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_restart_events.Add newEvent.Raw, newEvent
        Next
    End Property   
    Public Property Let Persist(value) : m_persist = value : End Property
    Public Property Let Profile(value) : m_profile = value : End Property
    Public Property Let Switch(value) : m_switches = Array(value) : End Property
    Public Property Let Switches(value) : m_switches = value : End Property
    Public Property Let StartEnabled(value) : m_start_enabled = value : End Property
    Public Property Let HitEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_hit_events.Add newEvent.Raw, newEvent
        Next
    End Property

	Public default Function init(name, mode)
        m_name = name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_internal_cache_id = -1
        m_enabled = False
        m_persist = True
        Set m_base_device = (new GlfBaseModeDevice)(mode, "shot", Me)
        m_debug = False
        m_profile = "default"
        m_player_var_name = "shot_" & m_name
        m_state = -1
        m_switches = Array()
        m_start_enabled = Empty
        Set m_hit_events = CreateObject("Scripting.Dictionary")
        Set m_tokens = CreateObject("Scripting.Dictionary")
        Set m_show_cache = CreateObject("Scripting.Dictionary")
        Set m_advance_events = CreateObject("Scripting.Dictionary")
        Set m_control_events = CreateObject("Scripting.Dictionary")
        Set m_reset_events = CreateObject("Scripting.Dictionary")
        Set m_restart_events = CreateObject("Scripting.Dictionary")

        Set Init = Me
	End Function

    Public Sub Activate()
        If GetPlayerState(m_player_var_name) = False Then
            m_state = 0
            If m_persist Then
                SetPlayerState m_player_var_name, 0
            End If
        Else
            m_state = GetPlayerState(m_player_var_name)
        End If
        If m_start_enabled = True Then
            Enable()
        Else
            If IsEmpty(m_start_enabled) And UBound(m_base_device.EnableEvents.Keys) = -1 Then
                Enable()
            End If
        End If
    End Sub

    Public Sub Deactivate()
        Disable()
        Dim evt
        For Each evt in m_switches
            RemovePinEventListener evt, m_mode & "_" & m_name & "_hit"
        Next
        For Each evt in m_hit_events.Keys
            RemovePinEventListener m_hit_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_hit"
        Next
        For Each evt in m_advance_events.Keys
            RemovePinEventListener m_advance_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_advance"
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events().Keys
                RemovePinEventListener m_control_events(evt).Events()(cEvt).EventName, m_mode & "_" & m_name & "_control_" & cEvt
            Next
        Next
        For Each evt in m_reset_events.Keys
            RemovePinEventListener m_reset_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_reset"
        Next
        For Each evt in m_restart_events.Keys
            RemovePinEventListener m_restart_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_restart"
        Next
    End Sub

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_switches
            AddPinEventListener evt & "_active", m_mode & "_" & m_name & "_hit", "ShotEventHandler", m_priority, Array("hit", Me, Null)
        Next
        For Each evt in m_hit_events.Keys
            AddPinEventListener m_hit_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_hit", "ShotEventHandler", m_priority, Array("hit", Me, m_hit_events(evt))
        Next
        For Each evt in m_advance_events.Keys
            AddPinEventListener m_advance_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_advance", "ShotEventHandler", m_priority, Array("advance", Me, m_advance_events(evt))
        Next
        For Each evt in m_control_events.Keys
            Dim cEvt
            For Each cEvt in m_control_events(evt).Events().Keys
                AddPinEventListener m_control_events(evt).Events()(cEvt).EventName, m_mode & "_" & m_name & "_control_" & cEvt, "ShotEventHandler", m_priority+m_control_events(evt).Events()(cEvt).Priority, Array("control", Me, m_control_events(evt).Events()(cEvt), m_control_events(evt))
            Next
        Next
        For Each evt in m_reset_events.Keys
            AddPinEventListener m_reset_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_reset", "ShotEventHandler", m_priority, Array("reset", Me, m_reset_events(evt))
        Next
        For Each evt in m_restart_events.Keys
            AddPinEventListener m_restart_events(evt).EventName, m_mode & "_" & m_name & "_" & evt & "_restart", "ShotEventHandler", m_priority, Array("restart", Me, m_restart_events(evt))
        Next
        'Play the show for the active state
        PlayShowForState(m_state)
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_hit_events.Keys
            RemovePinEventListener m_hit_events(evt).EventName, m_mode & "_" & m_name & "_" & evt  & "_hit"
        Next
        Dim x
        For x=0 to Glf_ShotProfiles(m_profile).StatesCount()
            StopShowForState(x)
        Next
    End Sub

    Private Sub StopShowForState(state)
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        Log "Removing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        If glf_running_shows.Exists(m_mode & "_" & CStr(state) & "_" & m_name & "_" & profileState.Key) Then 
            glf_running_shows(m_mode & "_" & CStr(state) & "_" & m_name & "_" & profileState.Key).StopRunningShow()
        End If
    End Sub

    Private Sub PlayShowForState(state)
        If m_enabled = False Then
            Exit Sub
        End If
        Dim profileState : Set profileState = Glf_ShotProfiles(m_profile).StateForIndex(state)
        Log "Playing Shot Show: " & m_mode & "_" & m_name & ". Key: " & profileState.Key
        If IsObject(profileState) Then
            If Not IsNull(profileState.Show) Then
                Dim new_running_show
                Set new_running_show = (new GlfRunningShow)(m_mode & "_" & CStr(m_state) & "_" & m_name & "_" & profileState.Key, profileState.Key, profileState, m_priority + profileState.Priority, m_tokens, m_internal_cache_id)
            End If
        End If
    End Sub

    Public Sub Hit(evt)
        If m_enabled = False Then
            Exit Sub
        End If

        Dim profile : Set profile = Glf_ShotProfiles(m_profile)
        Dim old_state : old_state = m_state
        Log "Hit! Profile: "&m_profile&", State: " & profile.StateName(m_state)

        Dim advancing
        If profile.AdvanceOnHit Then
            Log "Advancing shot because advance_on_hit is True."
            advancing = Advance(False)
        Else
            Log "Not advancing shot"
            advancing = False
        End If

    
        If profile.Block Then
            Glf_EventBlocks(evt).Add Name, True
        Else
            Glf_EventBlocks(evt).Add ShotKey, True
        End If
        Dim kwargs : Set kwargs = GlfKwargs()
		With kwargs
            .Add "profile", m_profile
            .Add "state", profile.StateName(old_state)
            .Add "advancing", advancing
        End With

        DispatchPinEvent m_name & "_hit", kwargs
        DispatchPinEvent m_name & "_" & m_profile & "_hit", kwargs
        DispatchPinEvent m_name & "_" & m_profile & "_" & profile.StateName(old_state) & "_hit", kwargs
        DispatchPinEvent m_name & "_" & profile.StateName(old_state) & "_hit", kwargs
        
    End Sub

    Public Function Advance(force)

        If m_enabled = False And force = False Then
            Advance = False
            Exit Function
        End If
        Dim profile : Set profile = Glf_ShotProfiles(m_profile)

        Log "Advancing 1 step. Profile: "&m_profile&", Current State: " & profile.StateName(m_state)

        If profile.StatesCount() = m_state Then
            If profile.ProfileLoop Then
                StopShowForState(m_state)
                m_state = 0
                If m_persist Then
                    SetPlayerState m_player_var_name, 0
                End If
                PlayShowForState(m_state)
            Else
                Advance = False
                Exit Function
            End If
        Else
            StopShowForState(m_state)
            m_state = m_state + 1
            If m_persist Then
                SetPlayerState m_player_var_name, m_state
            End If
            PlayShowForState(m_state)
        End If

        Advance = True
        
    End Function

    Public Sub Reset()
        Jump 0, True, False
    End Sub

    Public Sub Jump(state, force, force_show)
        Log "Received jump request. State: " & state & ", Force: "& force

        If Not m_enabled And Not force Then
            Log "Profile is disabled and force is False. Not jumping"
            Exit Sub
        End If
        If state = m_state And Not force_show Then
            Log "Shot is already in the jump destination state"
            Exit Sub
        End If
        Log "Jumping to profile state " & state

        StopShowForState(m_state)
        m_state = state
        If m_persist Then
            SetPlayerState m_player_var_name, m_state
        End If
        PlayShowForState(m_state)
    End Sub

    Public Sub Restart()
        Reset()
        Enable()
    End Sub

    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & Replace(m_name, "shot_", "") & ":" & vbCrLf
        If UBound(m_switches) = 0 Then
            yaml = yaml & "    switch: " & m_switches(0) & vbCrLf
        ElseIf UBound(m_switches) > 0 Then
            yaml = yaml & "    switches: " & Join(m_switches, ",") & vbCrLf
        End If
        yaml = yaml & "    show_tokens: " & vbCrLf
        dim key
        For Each key in m_tokens.keys
            If IsArray(m_tokens(key)) Then
                yaml = yaml & "      " & key & ": " & Join(m_tokens(key), ",") & vbCrLf
            Else  
                yaml = yaml & "      " & key & ": " & m_tokens(key) & vbCrLf
            End If
        Next

        If UBound(m_base_device.EnableEvents().Keys) > -1 Then
            yaml = yaml & "    enable_events: "
            x=0
            For Each key in m_base_device.EnableEvents().keys
                yaml = yaml & m_base_device.EnableEvents()(key).Raw
                If x <> UBound(m_base_device.EnableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_base_device.DisableEvents().Keys) > -1 Then
            yaml = yaml & "    disable_events: "
            x=0
            For Each key in m_base_device.DisableEvents().keys
                yaml = yaml & m_base_device.DisableEvents()(key).Raw
                If x <> UBound(m_base_device.DisableEvents().Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_advance_events.Keys) > -1 Then
            yaml = yaml & "    advance_events: "
            x=0
            For Each key in m_advance_events.keys
                yaml = yaml & m_advance_events(key).Raw
                If x <> UBound(m_advance_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_hit_events.Keys) > -1 Then
            yaml = yaml & "    hit_events: "
            x=0
            For Each key in m_hit_events.keys
                yaml = yaml & m_hit_events(key).Raw
                If x <> UBound(m_hit_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        yaml = yaml & "    profile: " & m_profile & vbCrLf
        If Not IsEmpty(m_start_enabled) Then
            yaml = yaml & "    start_enabled: " & m_start_enabled & vbCrLf
        End If
        If UBound(m_restart_events.Keys) > -1 Then
            yaml = yaml & "    restart_events: "
            x=0
            For Each key in m_restart_events.keys
                yaml = yaml & m_restart_events(key).Raw
                If x <> UBound(m_restart_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_reset_events.Keys) > -1 Then
            yaml = yaml & "    reset_events: "
            x=0
            For Each key in m_reset_events.keys
                yaml = yaml & m_reset_events(key).Raw
                If x <> UBound(m_reset_events.Keys) Then
                    yaml = yaml & ", "
                End If
                x = x + 1
            Next
            yaml = yaml & vbCrLf
        End If

        If UBound(m_control_events.Keys) > -1 Then
            yaml = yaml & "    control_events: " & vbCrLf
            For Each key in m_control_events.keys
                yaml = yaml & "      - events: "
                Dim cEvt
                x=0
                For Each cEvt in m_control_events(key).Events
                    yaml = yaml & cEvt
                    If x <> UBound(m_control_events(key).Events) Then
                        yaml = yaml & ", "
                    End If
                    x = x + 1
                Next
                yaml = yaml & vbCrLf
                yaml = yaml & "        state: " & m_control_events(key).State & vbCrLf
            Next
        End If
        
        ToYaml = yaml
    End Function
End Class

Function ShotEventHandler(args)
    Dim ownProps, kwargs, e
    ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    e = args(2)
    Dim evt : evt = ownProps(0)
    Dim shot : Set shot = ownProps(1)
    If Not IsNull(ownProps(2)) Then
        Dim glf_event : Set glf_event = ownProps(2)
        If glf_event.Evaluate() = False Then
            If IsObject(args(1)) Then
                Set ShotEventHandler = kwargs
            Else
                ShotEventHandler = kwargs
            End If
            Exit Function
        End If
    End If
    Select Case evt
        Case "activate"
            shot.Activate
        Case "deactivate"
            shot.Deactivate
        Case "enable"
            shot.Enable
        Case "hit"
            If Not Glf_EventBlocks(e).Exists(shot.Name) And Not Glf_EventBlocks(e).Exists(shot.ShotKey) Then
                shot.Hit e
            End If
        Case "advance"
            shot.Advance False
        Case "control"
            shot.Jump ownProps(3).State, ownProps(3).Force, ownProps(3).ForceShow
        Case "reset"
            shot.Reset
        Case "restart"
            shot.Restart
            
    End Select
    If IsObject(args(1)) Then
        Set ShotEventHandler = kwargs
    Else
        ShotEventHandler = kwargs
    End If
End Function

Class GlfShotControlEvent
	Private m_events, m_state, m_force, m_force_show
  
	Public Property Get Events(): Set Events = m_events: End Property
    Public Property Let Events(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events.Add newEvent.Name, newEvent
        Next
    End Property

    Public Property Get State(): State = m_state End Property
    Public Property Let State(input): m_state = input End Property

    Public Property Get Force(): Force = m_force: End Property
	Public Property Let Force(input): m_force = input: End Property
  
	Public Property Get ForceShow(): ForceShow = m_force_show: End Property
	Public Property Let ForceShow(input): m_force_show = input: End Property   

	Public default Function init()
        Set m_events = CreateObject("Scripting.Dictionary")
        m_state = 0
        m_force = True
        m_force_show = False
	    Set Init = Me
	End Function

End Class

Class GlfShowPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_eventValues
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "show_player" : End Property
    Public Property Get EventShows() : EventShows = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_show : Set new_show = (new GlfShowPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_show
        Set EventName = new_show
        
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "show_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "show_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_show_player_play", "ShowPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_show_player_play"
            PlayOff m_eventValues(evt).Key
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            If m_eventValues(evt).Action = "stop" Then
                PlayOff m_eventValues(evt).Key
            Else
                Dim new_running_show
                Set new_running_show = (new GlfRunningShow)(m_name & "_" & m_eventValues(evt).Key, m_eventValues(evt).Key, m_eventValues(evt), m_priority, Null, Null)
                If m_eventValues(evt).BlockQueue = True Then
                    Play = m_name & "_" & m_eventValues(evt).Key & "_" & m_eventValues(evt).Key  & "_unblock_queue"
                End If
            End If
        End If
    End Function

    Public Sub PlayOff(key)
        If glf_running_shows.Exists(m_name & "_" & key) Then 
            glf_running_shows(m_name & "_" & key).StopRunningShow()
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function ShowPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim ShowPlayer : Set ShowPlayer = ownProps(1)
    Select Case evt
        Case "activate"
            ShowPlayer.Activate
        Case "deactivate"
            ShowPlayer.Deactivate
        Case "play"
            Dim block_queue
            block_queue = ShowPlayer.Play(ownProps(2))
            If Not IsEmpty(block_queue) Then
                kwargs.Add "wait_for", block_queue
            End If
    End Select
    If IsObject(args(1)) Then
        Set ShowPlayerEventHandler = kwargs
    Else
        ShowPlayerEventHandler = kwargs
    End If
End Function

Class GlfShowPlayerItem
	Private m_key, m_show, m_loops, m_speed, m_tokens, m_action, m_syncms, m_duration, m_priority, m_internal_cache_id
    Private m_block_queue
    Private m_events_when_completed
    Private m_color_lookup
  
	Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Key(): Key = m_key End Property
    Public Property Let Key(input): m_key = input End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property

    Public Property Get ColorLookup(): ColorLookup = m_color_lookup End Property
    Public Property Let ColorLookup(input): m_color_lookup = input End Property

    Public Property Get Show()
        If IsNull(m_show) Then
            Show = Null
        Else
            Set Show = m_show
        End If
    End Property
	Public Property Let Show(input)
        'msgbox "input:" & input
        If glf_shows.Exists(input) Then
            Set m_show = glf_shows(input)
        End If
    End Property
  
	Public Property Get Loops(): Loops = m_loops: End Property
	Public Property Let Loops(input): m_loops = input: End Property
  
	Public Property Get Speed(): Speed = m_speed: End Property
	Public Property Let Speed(input): m_speed = input: End Property

    Public Property Get SyncMs(): SyncMs = m_syncms: End Property
    Public Property Let SyncMs(input): m_syncms = input: End Property      
        
    Public Property Get BlockQueue(): BlockQueue = m_block_queue : End Property
    Public Property Let BlockQueue(input): m_block_queue = input : End Property
    
    Public Property Get EventsWhenCompleted(): EventsWhenCompleted = m_events_when_completed : End Property
    Public Property Let EventsWhenCompleted(input): m_events_when_completed = input: End Property
            
    Public Property Get Tokens()
        Set Tokens = m_tokens
    End Property        
  
    Public Function ToYaml()
        Dim yaml
        yaml = yaml & "    " & m_show.Name &": " & vbCrLf
        If m_action <> "play" Then
            yaml = yaml & "      action: " & m_action & vbCrLf
        End If
        If m_key <> "" Then
            yaml = yaml & "      key: " & m_key & vbCrLf
        End If
        If m_priority <> 0 Then
            yaml = yaml & "      priority: " & m_priority & vbCrLf
        End If
        If m_loops > -1 Then
            yaml = yaml & "      loops: " & m_loops & vbCrLf
        End If
        If m_speed <> 1 Then
            yaml = yaml & "      speed: " & m_speed & vbCrLf
        End If
        If UBound(m_tokens.Keys) > -1 Then
            yaml = yaml & "      show_tokens: " & vbCrLf
            Dim key
            For Each key in m_tokens.Keys
                yaml = yaml & "        " & key & ": " & m_tokens(key) & vbCrLf
            Next
        End If
        If m_syncms > 0 Then
            yaml = yaml & "      sync_ms: " & m_syncms & vbCrLf
        End If
        ToYaml = yaml
    End Function

	Public default Function init()
        m_action = "play"
        m_key = ""
        m_priority = 0
        m_loops = -1
        m_internal_cache_id = -1
        m_speed = 1
        m_syncms = 0
        m_show = Null
        m_color_lookup = Empty
        m_block_queue = False
        m_events_when_completed = Array()
        Set m_tokens = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class


Class GlfShow

    Private m_name
    Private m_steps
    Private m_total_step_time

    Public Property Get Name(): Name = m_name: End Property
    
    Public Property Get Steps() : Set Steps = m_steps : End Property

    Public Function StepAtIndex(index) : Set StepAtIndex = m_steps.Items()(index) : End Function
    
    Public default Function init(name)
        m_name = name
        m_total_step_time = 0
        Set m_steps = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

    Public Function AddStep(absolute_time, relative_time, duration)
        Dim new_step : Set new_step = (new GlfShowStep)()
        new_step.Duration = duration
        new_step.RelativeTime = relative_time
        new_step.AbsoluteTime = absolute_time
        new_step.IsLastStep = True
        
        'Add a empty first step if if show does not start right away
        If UBound(m_steps.Keys) = -1 Then
            If Not IsNull(new_step.Time) And new_step.Time <> 0 Then
                Dim empty_step : Set empty_step = (new GlfShowStep)()
                empty_step.Duration = new_step.Time
                m_total_step_time = new_step.Time
                m_steps.Add CStr(UBound(m_steps.Keys())+1), empty_step        
            End If
        End If
        

        

        If UBound(m_steps.Keys()) > -1 Then
            Dim prevStep : Set prevStep = m_steps.Items()(UBound(m_steps.Keys()))
            prevStep.IsLastStep = False
            'need to work out previous steps duration.
            If IsNull(prevStep.Duration) Then
                'The previous steps duration needs calculating.
                'If this step has a relative time then the last steps duration is that time.
                If Not IsNull(new_step.Time) Then
                    If new_step.IsRelativeTime Then
                        prevStep.Duration = new_step.Time
                    Else
                        prevStep.Duration = new_step.Time - m_total_step_time
                    End If
                Else
                    prevStep.Duration = 1
                End If
            End If
            m_total_step_time = m_total_step_time + prevStep.Duration
        Else
            If IsNull(new_step.Duration) Then
                m_total_step_time = m_total_step_time + 1
            Else
                m_total_step_time = m_total_step_time + new_step.Time
            End If
        End If

        m_steps.Add CStr(UBound(m_steps.Keys())+1), new_step
        Set AddStep = new_step
    End Function

    Public Function ToYaml()
        'Dim yaml
        'yaml = yaml & "  " & Replace(m_name, "shotprofile_", "") & ":" & vbCrLf
        'yaml = yaml & "    states: " & vbCrLf
        'Dim token,evt,x : x = 0
        'For Each evt in m_states.Keys
        '    yaml = yaml & "     - name: " & StateName(x) & vbCrLf
            'yaml = yaml & "       show: " & m_states(evt).Show & vbCrLf
            'yaml = yaml & "       loops: " & m_states(evt).Loops & vbCrLf
            'yaml = yaml & "       speed: " & m_states(evt).Speed & vbCrLf
            'yaml = yaml & "       sync_ms: " & m_states(evt).SyncMs & vbCrLf

            'If Ubound(m_states(evt).Tokens().Keys)>-1 Then
            '    yaml = yaml & "       show_tokens: " & vbCrLf
            '    For Each token in m_states(evt).Tokens().Keys()
            '        yaml = yaml & "         " & token & ": " & m_states(evt).Tokens(token) & vbCrLf
            '    Next
            'End If

            'yaml = yaml & "     block: " & m_block & vbCrLf
            'yaml = yaml & "     advance_on_hit: " & m_advance_on_hit & vbCrLf
            'yaml = yaml & "     loop: " & m_loop & vbCrLf
            'yaml = yaml & "     rotation_pattern: " & m_rotation_pattern & vbCrLf
            'yaml = yaml & "     state_names_to_not_rotate: " & m_states_not_to_rotate & vbCrLf
            'yaml = yaml & "     state_names_to_rotate: " & m_states_to_rotate & vbCrLf
         '   x = x +1
        'Next
        'ToYaml = yaml
    End Function

End Class

Class GlfRunningShow

    Private m_key
    Private m_show_name
    Private m_show_settings
    Private m_current_step
    Private m_priority
    Private m_total_steps
    Private m_tokens
    Private m_internal_cache_id
    Private m_loops
    Private m_shows_added

    Public Property Get CacheName(): CacheName = m_show_name & "_" & m_internal_cache_id & "_" & ShowSettings.InternalCacheId: End Property
    Public Property Get Tokens(): Set Tokens = m_tokens : End Property

    Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Priority(): Priority = m_priority End Property
    Public Property Let Priority(input): m_priority = input End Property        

    Public Property Get CurrentStep(): CurrentStep = m_current_step End Property
    Public Property Let CurrentStep(input): m_current_step = input End Property        

    Public Property Get TotalSteps(): TotalSteps = m_total_steps End Property
    Public Property Let TotalSteps(input): m_total_steps = input End Property        

    Public Property Get ShowName(): ShowName = m_show_name: End Property
    Public Property Let ShowName(input): m_show_name = input: End Property

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property
        
    Public Property Get ShowSettings(): Set ShowSettings = m_show_settings: End Property
    Public Property Let ShowSettings(input)
        Set m_show_settings = input
        m_loops = m_show_settings.Loops
    End Property

    Public Property Get ShowsAdded()
        If IsNull(m_shows_added) Then
            ShowsAdded = Null
        Else
            Set ShowsAdded = m_shows_added
        End If
    End Property
    Public Property Let ShowsAdded(input)
        If IsNull(input) Then
            m_shows_added = Null
        Else
            Set m_shows_added = input
        End If
    End Property
    
    Public default Function init(rname, rkey, show_settings, priority, tokens, cache_id)
        m_show_name = rname
        m_key = rkey
        m_current_step = 0
        m_priority = priority
        m_internal_cache_id = cache_id
        m_loops=show_settings.Loops
        Set m_show_settings = show_settings
        m_shows_added = Null
        Dim key
        Dim mergedTokens : Set mergedTokens = CreateObject("Scripting.Dictionary")
        If Not IsNull(m_show_settings.Tokens) Then
            For Each key In m_show_settings.Tokens.Keys()
                mergedTokens.Add key, m_show_settings.Tokens()(key)
            Next
        End If
        If Not IsNull(tokens) Then
            For Each key In tokens.Keys
                If mergedTokens.Exists(key) Then
                    mergedTokens(key) = tokens(key)
                Else
                    mergedTokens.Add key, tokens(key)
                End If
            Next
        End If
        Set m_tokens = mergedTokens
        
        m_total_steps = UBound(m_show_settings.Show.Steps.Keys())
        If glf_running_shows.Exists(m_show_name) Then
            glf_running_shows(m_show_name).StopRunningShow()
            glf_running_shows.Add m_show_name, Me
        Else
            glf_running_shows.Add m_show_name, Me
        End If 
        Play
        Set Init = Me
	End Function

    Public Sub Play()
        'Play the show.
        Log "Playing show: " & m_show_name & " With Key: " & m_key
        GlfShowStepHandler(Array(Me))
    End Sub

    Public Sub StopRunningShow()
        Log "Removing show: " & m_show_name & " With Key: " & m_key
        Dim cached_show,light, cached_show_lights
        If glf_cached_shows.Exists(CacheName) Then
            cached_show = glf_cached_shows(CacheName)
            Set cached_show_lights = cached_show(1)
        Else
            msgbox "show " & running_show.CacheName & " not cached! Problem with caching"
        End If
        Dim lightStack
        For Each light in cached_show_lights.Keys()

            Set lightStack = glf_lightStacks(light)
            
            If Not lightStack.IsEmpty() Then
                lightStack.PopByKey(m_show_name & "_" & m_key)
            End If

            Dim show_key
            For Each show_key in glf_running_shows.Keys
                If Left(show_key, Len("fade_" & m_show_name & "_" & m_key & "_" & light)) = "fade_" & m_show_name & "_" & m_key & "_" & light Then
                    glf_running_shows(show_key).StopRunningShow()
                End If
            Next
            
            If Not lightStack.IsEmpty() Then
                ' Set the light to the next color on the stack
                Dim nextColor
                Set nextColor = lightStack.Peek()
                Glf_SetLight light, nextColor("Color")
            Else
                ' Turn off the light since there's nothing on the stack
                Glf_SetLight light, "000000"
            End If
        Next

        RemoveDelay Me.ShowName & "_" & Me.Key
        glf_running_shows.Remove m_show_name
    End Sub

    Public Sub Log(message)
        If glf_debug_level = "Debug" Then
            glf_debugLog.WriteToLog "Running Show", message
        End If
    End Sub
End Class

Function GlfShowStepHandler(args)
    Dim running_show : Set running_show = args(0)
    Dim nextStep : Set nextStep = running_show.ShowSettings.Show.StepAtIndex(running_show.CurrentStep)
    If UBound(nextStep.Lights) > -1 Then
        Dim cached_show, cached_show_seq
        If glf_cached_shows.Exists(running_show.CacheName) Then
            cached_show = glf_cached_shows(running_show.CacheName)
            cached_show_seq = cached_show(0)
        Else
            msgbox running_show.CacheName & " show not cached! Problem with caching"
        End If

        If Not IsNull(running_show.ShowsAdded) Then
            Dim show_added
            For Each show_added in running_show.ShowsAdded.Keys()
                If glf_running_shows.Exists(show_added) Then 
                    glf_running_shows(show_added).StopRunningShow()
                End If
            Next
            running_show.ShowsAdded = Null
        End If  

        Dim shows_added, replacement_color
        replacement_color = Empty
        If Not IsEmpty(running_show.ShowSettings.ColorLookup) Then
            'msgbox ubound(running_show.ShowSettings.ColorLookup())
            'MsgBox UBound(cached_show_seq)
            replacement_color = running_show.ShowSettings.ColorLookup()(running_show.CurrentStep)
        End If
        Set shows_added = LightPlayerCallbackHandler(running_show.Key, Array(cached_show_seq(running_show.CurrentStep)), running_show.ShowName, running_show.Priority + running_show.ShowSettings.Priority, True, running_show.ShowSettings.Speed, replacement_color)
        If IsObject(shows_added) Then
            'Fade shows were added, log them agains the current show.
            running_show.ShowsAdded = shows_added
        End If
    End If
    If UBound(nextStep.ShowsInStep().Keys())>-1 Then
        Dim show_item
        Dim show_items : show_items = nextStep.ShowsInStep().Items()
        For Each show_item in show_items
            If show_item.Action = "stop" Then
                If glf_running_shows.Exists(running_show.Key & "_" & show_item.Show & "_" & show_item.Key) Then 
                    glf_running_shows(running_show.Key & "_" & show_item.Show & "_" & show_item.Key).StopRunningShow()
                End If
            Else
                Dim new_running_show
                'MsgBox running_show.Priority + running_show.ShowSettings.Priority
                'msgbox running_show.Key & "_" & show_item.Key
                Set new_running_show = (new GlfRunningShow)(show_item.Key, show_item.Key, show_item, running_show.Priority + running_show.ShowSettings.Priority, Null, Null)
            End If
        Next
    End If
    If UBound(nextStep.DOFEventsInStep().Keys())>-1 Then
        Dim dof_item
        Dim dof_items : dof_items = nextStep.DOFEventsInStep().Items()
        For Each dof_item in dof_items
            DOF dof_item.DOFEvent, dof_item.Action
        Next
    End If
    If UBound(nextStep.SlidesInStep().Keys())>-1 Then
        Dim slide_item
        Dim slide_items : slide_items = nextStep.SlidesInStep().Items()
        For Each slide_item in slide_items
            
        Next
    End If

    If nextStep.Duration = -1 Then
        'glf_debugLog.WriteToLog "Running Show", "HOLD"
        Exit Function
    End If
    running_show.CurrentStep = running_show.CurrentStep + 1
    If nextStep.IsLastStep = True Then
        'msgbox "last step"
        If IsNull(nextStep.Duration) Then
            'msgbox "5!"
            nextStep.Duration = 1
        End If
    End If
    If running_show.CurrentStep > running_show.TotalSteps Then
        'End of Show
        'glf_debugLog.WriteToLog "Running Show", "END OF SHOW"
        If running_show.Loops = -1 Or running_show.Loops > 1 Then
            If running_show.Loops > 1 Then
                running_show.Loops = running_show.Loops - 1
            End If
            running_show.CurrentStep = 0
            SetDelay running_show.ShowName & "_" & running_show.Key, "GlfShowStepHandler", Array(running_show), (nextStep.Duration / running_show.ShowSettings.Speed) * 1000
        Else
'            glf_debugLog.WriteToLog "Running Show", "STOPPING SHOW, NO Loops"
            If UBound(running_show.ShowSettings().EventsWhenCompleted) > -1 Then
                Dim evt_when_completed
                For Each evt_when_completed in running_show.ShowSettings().EventsWhenCompleted
                    DispatchPinEvent evt_when_completed, Null
                Next
            End If
            DispatchPinEvent running_show.ShowName & "_" & running_show.Key & "_unblock_queue", Null
            running_show.StopRunningShow()
        End If
    Else
'        glf_debugLog.WriteToLog "Running Show", "Scheduling Next Step"
        SetDelay running_show.ShowName & "_" & running_show.Key, "GlfShowStepHandler", Array(running_show), (nextStep.Duration / running_show.ShowSettings.Speed) * 1000
    End If
End Function

Class GlfShowStep

    Private m_lights, m_shows, m_dofs, m_slides, m_time, m_duration, m_isLastStep, m_absTime, m_relTime

    Public Property Get Lights(): Lights = m_lights: End Property
    Public Property Let Lights(input) : m_lights = input: End Property

    Public Property Get ShowsInStep(): Set ShowsInStep = m_shows: End Property
    Public Property Get Shows(name)
        Dim new_show : Set new_show = (new GlfShowPlayerItem)()
        new_show.Show = name
        m_shows.Add name & CStr(UBound(m_shows.Keys)), new_show
        Set Shows = new_show
    End Property

    Public Property Get DOFEventsInStep(): Set DOFEventsInStep = m_dofs: End Property
    Public Property Get DOFEvent(dof_event)
        Dim new_dof : Set new_dof = (new GlfDofPlayerItem)()
        new_dof.DOFEvent = dof_event
        m_dofs.Add dof_event & CStr(UBound(m_dofs.Keys)), new_dof
        Set DOFEvent = new_dof
    End Property

    Public Property Get SlidesInStep(): Set SlidesInStep = m_slides: End Property
        Public Property Get Slides(slide)
            Dim new_slide : Set new_slide = (new GlfSlidePlayerItem)()
            new_slide.Slide = slide
            m_slides.Add slide & CStr(UBound(m_slides.Keys)), new_slide
            Set Slides = new_slide
        End Property

    Public Property Get Time()
        If IsNull(m_relTime) Then
            Time = m_absTime
        Else
            Time = m_relTime
        End If
    End Property

    Public Property Get IsRelativeTime()
        If Not IsNull(m_relTime) Then
            IsRelativeTime = True
        Else
            IsRelativeTime = False
        End If
    End Property

    Public Property Let RelativeTime(input) : m_relTime = input: End Property
    Public Property Let AbsoluteTime(input) : m_absTime = input: End Property

    Public Property Get Duration(): Duration = m_duration: End Property
    Public Property Let Duration(input) : m_duration = input: End Property

    Public Property Get IsLastStep(): IsLastStep = m_isLastStep: End Property
    Public Property Let IsLastStep(input) : m_isLastStep = input: End Property
        
    Public default Function init()
        m_lights = Array()
        m_duration = Null
        m_time = Null
        m_absTime = Null
        m_relTime = Null
        m_isLastStep = False
        Set m_shows = CreateObject("Scripting.Dictionary")
        Set m_dofs = CreateObject("Scripting.Dictionary")
        Set m_slides = CreateObject("Scripting.Dictionary")
        Set Init = Me
	End Function

    Public Function ToYaml()
        'Dim yaml
        'yaml = yaml & "  " & Replace(m_name, "shotprofile_", "") & ":" & vbCrLf
        'yaml = yaml & "    states: " & vbCrLf
        'Dim token,evt,x : x = 0
        'For Each evt in m_states.Keys
        '    yaml = yaml & "     - name: " & StateName(x) & vbCrLf
            'yaml = yaml & "       show: " & m_states(evt).Show & vbCrLf
            'yaml = yaml & "       loops: " & m_states(evt).Loops & vbCrLf
            'yaml = yaml & "       speed: " & m_states(evt).Speed & vbCrLf
            'yaml = yaml & "       sync_ms: " & m_states(evt).SyncMs & vbCrLf

            'If Ubound(m_states(evt).Tokens().Keys)>-1 Then
            '    yaml = yaml & "       show_tokens: " & vbCrLf
            '    For Each token in m_states(evt).Tokens().Keys()
            '        yaml = yaml & "         " & token & ": " & m_states(evt).Tokens(token) & vbCrLf
            '    Next
            'End If

            'yaml = yaml & "     block: " & m_block & vbCrLf
            'yaml = yaml & "     advance_on_hit: " & m_advance_on_hit & vbCrLf
            'yaml = yaml & "     loop: " & m_loop & vbCrLf
            'yaml = yaml & "     rotation_pattern: " & m_rotation_pattern & vbCrLf
            'yaml = yaml & "     state_names_to_not_rotate: " & m_states_not_to_rotate & vbCrLf
            'yaml = yaml & "     state_names_to_rotate: " & m_states_to_rotate & vbCrLf
         '   x = x +1
        'Next
        'ToYaml = yaml
    End Function

End Class


Class GlfSlidePlayer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_debug
    private m_base_device
    Private m_events
    Private m_eventValues

    Public Property Get Name() : Name = "slide_player" : End Property

    Public Property Get EventName(name)
        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_slide : Set new_slide = (new GlfSlidePlayerItem)()
        m_eventValues.Add newEvent.Raw, new_slide
        Set EventName = new_slide
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "slide_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "slide_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_slide_player_play", "SlidePlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & evt & "_slide_player_play"
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            'Fire Slide
            bcpController.PlaySlide m_eventValues(evt).Slide, m_mode, m_events(evt).EventName, m_priority
        End If
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_slide_player", message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function SlidePlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim slide_player : Set slide_player = ownProps(1)
    Select Case evt
        Case "activate"
            slide_player.Activate
        Case "deactivate"
            slide_player.Deactivate
        Case "play"
            slide_player.Play(ownProps(2))
    End Select
    If IsObject(args(1)) Then
        Set SlidePlayerEventHandler = kwargs
    Else
        SlidePlayerEventHandler = kwargs
    End If
End Function

Class GlfSlidePlayerItem
	Private m_slide, m_action, m_expire, m_max_queue_time, m_method, m_priority, m_target, m_tokens
    
    Public Property Get Slide(): Slide = m_slide: End Property
    Public Property Let Slide(input)
        m_slide = input
    End Property
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input)
        m_action = input
    End Property

    Public Property Get Expire(): Expire = m_expire: End Property
    Public Property Let Expire(input)
        m_expire = input
    End Property

    Public Property Get MaxQueueTime(): MaxQueueTime = m_max_queue_time: End Property
    Public Property Let MaxQueueTime(input)
        m_max_queue_time = input
    End Property

    Public Property Get Method(): Method = m_method: End Property
    Public Property Let Method(input)
        m_method = input
    End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input)
        m_priority = input
    End Property

    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input)
        m_target = input
    End Property

    Public Property Get Tokens(): Tokens = m_tokens: End Property
    Public Property Let Tokens(input)
        m_tokens = input
    End Property

	Public default Function init()
        m_action = "play"
        m_slide = Empty
        m_expire = Empty
        m_max_queue_time = Empty
        m_method = Empty
        m_priority = Empty
        m_target = Empty
        m_tokens = Empty
        Set Init = Me
	End Function

End Class


Class GlfSoundPlayer

    Private m_priority
    Private m_mode
    Private m_events
    Private m_eventValues
    Private m_debug
    Private m_name
    Private m_value
    private m_base_device

    Public Property Get Name() : Name = "sound_player" : End Property
    Public Property Get EventSounds() : EventSounds = m_eventValues.Items() : End Property
    Public Property Get EventName(name)

        Dim newEvent : Set newEvent = (new GlfEvent)(name)
        m_events.Add newEvent.Raw, newEvent
        Dim new_sound : Set new_sound = (new GlfSoundPlayerItem)()
        m_eventValues.Add newEvent.Raw, new_sound
        Set EventName = new_sound
        
    End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "sound_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_eventValues = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "sound_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_sound_player_play", "SoundPlayerEventHandler", m_priority+m_events(evt).Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).EventName, m_mode & "_" & m_eventValues(evt).Key & "_sound_player_play"
            PlayOff evt
        Next
    End Sub

    Public Function Play(evt)
        Play = Empty
        If m_events(evt).Evaluate() Then
            If m_eventValues(evt).Action = "stop" Then
                PlayOff evt
            Else
                glf_sound_buses(m_eventValues(evt).Sound.Bus).Play m_eventValues(evt)
            End If
        End If
    End Function

    Public Sub PlayOff(evt)
        glf_sound_buses(m_eventValues(evt).Sound.Bus).StopSoundWithKey m_eventValues(evt).Sound.File
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

    Public Function ToYaml()
        Dim yaml
        Dim evt
        If UBound(m_events.Keys) > -1 Then
            For Each key in m_events.keys
                yaml = yaml & "  " & key & ": " & vbCrLf
                yaml = yaml & m_events(key).ToYaml
            Next
            yaml = yaml & vbCrLf
        End If
        ToYaml = yaml
    End Function

End Class

Function SoundPlayerEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim sound_player : Set sound_player = ownProps(1)
    Select Case evt
        Case "activate"
            sound_player.Activate
        Case "deactivate"
            sound_player.Deactivate
        Case "play"
            'Dim block_queue
            sound_player.Play(ownProps(2))
            'If Not IsEmpty(block_queue) Then
            '    kwargs.Add "wait_for", block_queue
            'End If
    End Select
    If IsObject(args(1)) Then
        Set SoundPlayerEventHandler = kwargs
    Else
        SoundPlayerEventHandler = kwargs
    End If
End Function


Class GlfSoundPlayerItem
	Private m_sound, m_action, m_key, m_volume, m_loops
    
    Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property

    Public Property Get Key(): Key = m_key: End Property
    Public Property Let Key(input): m_key = input: End Property

    Public Property Get Sound()
        If IsNull(m_sound) Then
            Sound = Null
        Else
            Set Sound = m_sound
        End If
    End Property
	Public Property Let Sound(input)
        If glf_sounds.Exists(input) Then
            Set m_sound = glf_sounds(input)
        End If
    End Property
  
	Public default Function init()
        m_action = "play"
        m_sound = Null
        m_key = Empty
        m_volume = Empty
        m_loops = Empty
        Set Init = Me
	End Function

End Class

Class GlfStateMachine
    Private m_name
    Private m_player_var_name
    Private m_mode
    Private m_debug
    Private m_priority
    Private m_states
    Private m_transitions
    private m_base_device
 
    Private m_state
    Private m_persist_state
    Private m_starting_state
 
    Public Property Get Name(): Name = m_name: End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    
        
    Public Property Get GetValue(value)
        Select Case value
            Case "state":
                GetValue = State()
        End Select
    End Property

    Public Property Get State()
        If m_persist_state = True Then
            Dim s : s = GetPlayerState(m_player_var_name)
            If s=False Then
                State = Null
            Else
                State = s
            End If
        Else
            State = m_state
        End If
    End Property

    Public Property Let State(value)
        If m_persist_state = True Then
            SetPlayerState m_player_var_name, value
            m_state = value
        Else
            m_state = value
        End If
    End Property
    
    Public Property Get States(name)
        If m_states.Exists(name) Then
            Set States = m_states(name)
        Else
            Dim new_state : Set new_state = (new GlfStateMachineState)(name)
            m_states.Add name, new_state
            Set States = new_state
        End If
    End Property
    Public Property Get StateItems(): StateItems = m_states.Items(): End Property
 
    Public Property Get Transitions()
        Dim count : count = UBound(m_transitions.Keys)
        Dim new_transition : Set new_transition = (new GlfStateMachineTranistion)()
        m_transitions.Add CStr(count), new_transition
        Set Transitions = new_transition
    End Property
 
    Public Property Get PersistState(): PersistState = m_persist_state : End Property
    Public Property Let PersistState(value) : m_persist_state = value : End Property

    Public Property Get StartingState(): StartingState = m_starting_state : End Property
    Public Property Let StartingState(value) : m_starting_state = value : End Property
 
    Public default Function init(name, mode)
        m_name = name
        m_player_var_name = "state_machine_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_debug = False
        m_persist_state = False
        m_starting_state = "start"
        m_state = Null
 
        Set m_states = CreateObject("Scripting.Dictionary")
        Set m_transitions = CreateObject("Scripting.Dictionary")
 
        Set m_base_device = (new GlfBaseModeDevice)(mode, "state_machine", Me)
        glf_state_machines.Add name, Me
        Set Init = Me
    End Function

    Public Sub Activate()
        Enable()
    End Sub

    Public Sub Deactivate()
        Disable()
    End Sub

    Public Sub Enable()
        If IsNull(State()) Then
            StartState m_starting_state
        Else
            AddHandlersForCurrentState()
            RunShowForCurrentState()
        End If

    End Sub

    Public Sub Disable()
        RemoveHandlers()
        StopShowForCurrentState()
        m_state = Null
    End Sub

    Public Sub StartState(start_state)
        Log("Starting state " & start_state)
        If Not m_states.Exists(start_state) Then
            Log("Invalid state " & start_state)
            Exit Sub
        End If

        Dim state_config : Set state_config = m_states(start_state)

        State() = start_state
        If UBound(state_config.EventsWhenStarted().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStarted().Items()
                If evt.Evaluate() = True Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        AddHandlersForCurrentState()
        RunShowForCurrentState()
    End Sub

    Public Sub StopCurrentState()
        Log "Stopping state " & State()
        RemoveHandlers()
        Dim state_config : Set state_config = m_states(state)

        If UBound(state_config.EventsWhenStopped().Keys()) > -1 Then
            Dim evt
            For Each evt in state_config.EventsWhenStopped().Items()
                If evt.Evaluate() = True Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If

        StopShowForCurrentState()

        State() = Null
    End Sub

    Public Sub RunShowForCurrentState()
        Log state
        Dim state_config : Set state_config = m_states(state)
        If Not IsNull(state_config.ShowWhenActive().Show) Then
            Dim show : Set show = state_config.ShowWhenActive
            Log "Starting show %s" & m_name & "_" & show.Key
            Dim new_running_show
            Set new_running_show = (new GlfRunningShow)(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key, show.Key, show, m_priority, Null, state_config.InternalCacheId)
        End If
    End Sub

    Public Sub StopShowForCurrentState()
        Dim state_config : Set state_config = m_states(state)
        If Not IsNull(state_config.ShowWhenActive().Show) Then
            Dim show : Set show = state_config.ShowWhenActive
            Log "Stopping show %s" & m_name & "_" & show.Key
            If glf_running_shows.Exists(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key) Then 
                glf_running_shows(m_mode & "_" & m_name & "_" & state_config.Name & "_" & show.Key).StopRunningShow()
            End If
        End If
    End Sub

    Public Sub AddHandlersForCurrentState()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            If transition.Source.Exists(State()) Then
                For Each evt in transition.Events.Items()
                    AddPinEventListener evt.EventName, m_name & "_" & transition.Target & "_" & evt.EventName & "_transition", "StateMachineTransitionHandler", m_priority+evt.Priority, Array("transition", Me, evt, transition)
                Next
            End If
        Next
    End Sub

    Public Sub RemoveHandlers()
        Dim transition, evt
        For Each transition in m_transitions.Items()
            For Each evt in transition.Events.Items()
                RemovePinEventListener evt.EventName, m_name & "_" & transition.Target & "_" & evt.EventName & "_transition"
            Next
        Next
    End Sub

    Public Sub MakeTransition(transition)

        Log "Transitioning from " & State() & " to " & transition.Target
        StopCurrentState()
        If UBound(transition.EventsWhenTransitioning().Keys()) > -1 Then
            Dim evt
            For Each evt in transition.EventsWhenTransitioning().Items()
                If evt.Evaluate() = True Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
        End If
        
        StartState transition.Target

    End Sub

    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class
 
Class GlfStateMachineState
	Private m_name, m_label, m_show_when_active, m_events_when_started, m_events_when_stopped, m_internal_cache_id
 

    Public Property Get InternalCacheId(): InternalCacheId = m_internal_cache_id: End Property
    Public Property Let InternalCacheId(input): m_internal_cache_id = input: End Property

	Public Property Get Name(): Name = m_name: End Property
    Public Property Let Name(input): m_name = input: End Property
 
    Public Property Get Label(): Label = m_label: End Property
    Public Property Let Label(input): m_label = input: End Property
 
    Public Property Get ShowWhenActive()
        Set ShowWhenActive = m_show_when_active
    End Property

    Public Property Get EventsWhenStarted(): Set EventsWhenStarted = m_events_when_started: End Property
    Public Property Let EventsWhenStarted(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_started.Add newEvent.Name, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenStopped(): Set EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(input)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_stopped.Add newEvent.Name, newEvent
        Next
    End Property
 
	Public default Function init(name)
        m_name = name
        m_label = Empty
        m_internal_cache_id = -1
        Set m_show_when_active = (new GlfShowPlayerItem)()
        Set m_events_when_started = CreateObject("Scripting.Dictionary")
        Set m_events_when_stopped = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class
 
Class GlfStateMachineTranistion
	Private m_name, m_sources, m_target, m_events, m_events_when_transitioning

    Public Property Get Source(): Set Source = m_sources: End Property
    Public Property Let Source(value)
        Dim x
        For x=0 to UBound(value)
            m_sources.Add value(x), True
        Next    
    End Property
 
    Public Property Get Target(): Target = m_target: End Property
    Public Property Let Target(input): m_target = input: End Property  

    Public Property Get Events(): Set Events = m_events: End Property
    Public Property Let Events(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events.Add newEvent.Name, newEvent
        Next    
    End Property
 
    Public Property Get EventsWhenTransitioning(): Set EventsWhenTransitioning = m_events_when_transitioning: End Property
    Public Property Let EventsWhenTransitioning(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_transitioning.Add newEvent.Name, newEvent
        Next    
    End Property
 
	Public default Function init()
        Set m_sources = CreateObject("Scripting.Dictionary")
        m_target = Empty
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_events_when_transitioning = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function
 
End Class


Public Function StateMachineTransitionHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim state_machine : Set state_machine = ownProps(1)
    Select Case evt
        Case "transition"
            Dim glf_event : Set glf_event = ownProps(2)
            If glf_event.Evaluate() = True Then
                state_machine.MakeTransition ownProps(3)
            Else
                If glf_debug_level = "Debug" Then
                    glf_debugLog.WriteToLog "State machine transition",  "failed condition: " & glf_event.Raw
                End If
            End If
    End Select
    If IsObject(args(1)) Then
        Set StateMachineTransitionHandler = kwargs
    Else
        StateMachineTransitionHandler = kwargs
    End If
End Function
Class GlfTilt

    Private m_name
    Private m_priority
    Private m_base_device
    Private m_reset_warnings_events
    Private m_tilt_events
    Private m_tilt_warning_events
    Private m_tilt_slam_tilt_events
    Private m_settle_time
    Private m_warnings_to_tilt
    Private m_multiple_hit_window
    Private m_tilt_warning_switch
    Private m_tilt_switch
    Private m_slam_tilt_switch
    Private m_last_tilt_warning_switch 
    Private m_last_warning
    Private m_balls_to_collect
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "tilt_settle_ms_remaining":
                GetValue = TiltSettleMsRemaining()
            Case "tilt_warnings_remaining":
                GetValue = TiltWarningsRemaining()
        End Select
    End Property

    Public Property Let ResetWarningEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_reset_warnings_events.Add newEvent.Raw, newEvent
        Next
    End Property
    'Public Property Let TiltEvents(value)
    '    Dim x
    '    For x=0 to UBound(value)
    '        Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
    '        m_tilt_events.Add newEvent.Raw, newEvent
    '    Next
    'End Property
    'Public Property Let TiltWarningEvents(value)
    '    Dim x
    '    For x=0 to UBound(value)
    '        Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
    '        m_tilt_warning_events.Add newEvent.Raw, newEvent
    '    Next
    'End Property
    'Public Property Let SlamTiltEvents(value)
    '    Dim x
    '    For x=0 to UBound(value)
    '        Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
    '        m_tilt_slam_tilt_events.Add newEvent.Raw, newEvent
    '    Next
    'End Property
    Public Property Let SettleTime(value): Set m_settle_time = CreateGlfInput(value): End Property
    Public Property Let WarningsToTilt(value): Set m_warnings_to_tilt = CreateGlfInput(value): End Property
    Public Property Let MultipleHitWindow(value): Set m_multiple_hit_window = CreateGlfInput(value): End Property
    'Public Property Let TiltWarningSwitch(value): m_tilt_warning_switch = value: End Property
    'Public Property Let TiltSwitch(value): m_tilt_switch = value: End Property
    'Public Property Let SlamTiltSwitch(value): m_slam_tilt_switch = value: End Property

    Private Property Get TiltSettleMsRemaining()
        TiltSettleMsRemaining = 0
        If m_last_tilt_warning_switch > 0 Then
            Dim delta
            delta = m_settle_time.Value - (gametime - m_last_tilt_warning_switch)
            If delta > 0 Then
                TiltSettleMsRemaining = delta
            End If
        End If
    End Property

    Private Property Get TiltWarningsRemaining() 
        TiltWarningsRemaining = 0

        If glf_gameStarted Then
            TiltWarningsRemaining = m_warnings_to_tilt.Value() - GetPlayerState("tilt_warnings")
        End If   
    End Property
    
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(mode)
        m_name = "tilt_" & mode.name
        m_priority = mode.Priority
        Set m_reset_warnings_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_warning_events = CreateObject("Scripting.Dictionary")
        Set m_tilt_slam_tilt_events = CreateObject("Scripting.Dictionary")
        Set m_settle_time = CreateGlfInput(0)
        Set m_warnings_to_tilt = CreateGlfInput(0)
        Set m_multiple_hit_window = CreateGlfInput(0)
        m_tilt_switch = Empty
        m_tilt_warning_switch = Empty
        m_slam_tilt_switch = Empty
        m_last_tilt_warning_switch = 0
        m_last_warning = 0
        m_balls_to_collect = 0
        Set m_base_device = (new GlfBaseModeDevice)(mode, "tilt", Me)
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_reset_warnings_events.Keys()
            AddPinEventListener m_reset_warnings_events(evt).EventName, m_name & "_" & evt & "_reset_warnings", "TiltHandler", m_priority+m_reset_warnings_events(evt).Priority, Array("reset_warnings", Me, m_reset_warnings_events(evt))
        Next
        For Each evt in m_tilt_events.Keys()
            AddPinEventListener m_tilt_events(evt).EventName, m_name & "_" & evt & "_tilt", "TiltHandler", m_priority+m_tilt_events(evt).Priority, Array("tilt", Me, m_tilt_events(evt))
        Next
        For Each evt in m_tilt_warning_events.Keys()
            AddPinEventListener m_tilt_warning_events(evt).EventName, m_name & "_" & evt & "_tilt_warning", "TiltHandler", m_priority+m_tilt_warning_events(evt).Priority, Array("tilt_warning", Me, m_tilt_warning_events(evt))
        Next
        For Each evt in m_tilt_slam_tilt_events.Keys()
            AddPinEventListener m_tilt_slam_tilt_events(evt).EventName, m_name & "_" & evt & "_slam_tilt", "TiltHandler", m_priority+m_tilt_slam_tilt_events(evt).Priority, Array("slam_tilt", Me, m_tilt_slam_tilt_events(evt))
        Next
        
        AddPinEventListener  "s_tilt_warning_active", m_name & "_tilt_warning_switch_active", "TiltHandler", m_priority, Array("_tilt_warning_switch_active", Me)
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_reset_warnings_events.Keys()
            RemoveEventListener m_reset_warnings_events(evt).EventName, m_name & "_" & evt & "_reset_warnings"
        Next
        For Each evt in m_tilt_events.Keys()
            RemoveEventListener m_tilt_events(evt).EventName, m_name & "_" & evt & "_tilt"
        Next
        For Each evt in m_tilt_warning_events.Keys()
            RemoveEventListener m_tilt_warning_events(evt).EventName, m_name & "_" & evt & "_tilt_warning"
        Next
        For Each evt in m_tilt_slam_tilt_events.Keys()
            RemoveEventListener m_tilt_slam_tilt_events(evt).EventName, m_name & "_" & evt & "_slam__tilt"
        Next

        RemoveEventListener "s_tilt_warning_active", m_name & "_tilt_warning_switch_active"

    End Sub

    Public Sub TiltWarning()
        'Process a tilt warning.
        'If the number of warnings is than the number to cause a tilt, a tilt will be
        'processed.

        m_last_tilt_warning_switch = gametime
        If glf_gameStarted = False Or glf_gameTilted = True Then
            Exit Sub
        End If
        Log "Tilt Warning"
        m_last_warning = gametime
        SetPlayerState "tilt_warnings", GetPlayerState("tilt_warnings")+1
        Dim warnings : warnings = GetPlayerState("tilt_warnings")
        Dim warnings_to_tilt : warnings_to_tilt = m_warnings_to_tilt.Value()
        If warnings>=warnings_to_tilt Then
            Tilt()
        Else
            Dim kwargs
            Set kwargs = GlfKwargs()
            With kwargs
                .Add "warnings", warnings
                .Add "warnings_remaining", warnings_to_tilt - warnings
            End With
            DispatchPinEvent "tilt_warning", kwargs
            DispatchPinEvent "tilt_warning_" & warnings, Null
        End If
    End Sub

    Public Sub ResetWarnings()
        'Reset the tilt warnings for the current player.
        If glf_gamestarted = False or glf_gameEnding = True Then
            Exit Sub
        End If
        SetPlayerState "tilt_warnings", 0
    End Sub

    Public Sub Tilt()
        'Cause the ball to tilt.
        'This will post an event called *tilt*, set the game mode's ``tilted``
        'attribute to *True*, disable the flippers and autofire devices, end the
        'current ball, and wait for all the balls to drain.
        If glf_gameStarted = False or glf_gameTilted=True or glf_gameEnding = True Then
            Exit Sub
        End If
        glf_gametilted = True
        m_balls_to_collect = glf_BIP
        Log "Processing Tilt. Balls to collect: " & m_balls_to_collect
        DispatchPinEvent "tilt", Null
        AddPinEventListener GLF_BALL_ENDING, m_name & "_ball_ending", "TiltHandler", 20, Array("tilt_ball_ending", Me)
        AddPinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain", "TiltHandler", 999999, Array("tilt_ball_drain", Me)
        Glf_EndBall()
    End Sub

    Public Sub TiltedBallDrain(unclaimed_balls)
        Log "Tilted ball drain, unclaimed balls: " & unclaimed_balls
        m_balls_to_collect = m_balls_to_collect - unclaimed_balls
        Log "Tilted ball drain. Balls to collect: " & m_balls_to_collect
        If m_balls_to_collect <= 0 Then
            TiltDone()
        End If
    End Sub

    Public Sub HandleTiltWarningSwitch()
        Log "Handling Tilt Warning Switch"
        If m_last_warning = 0 Or (m_last_warning + m_multiple_hit_window.Value() * 0.001) <= gametime Then
            TiltWarning()
        End If
    End Sub

    Public Sub BallEndingTilted()
        If m_balls_to_collect<=0 Then
            TiltDone()
        End If
    End Sub

    Public Sub TiltDone()
        If TiltSettleMsRemaining() > 0 Then
            SetDelay "delay_tilt_clear", "TiltHandler", Array(Array("tilt_done", Me), Null), TiltSettleMsRemaining()
            Exit Sub
        End If
        Log "Tilt Done"
        RemovePinEventListener GLF_BALL_ENDING, m_name & "_ball_ending"
        RemovePinEventListener GLF_BALL_DRAIN, m_name & "_ball_drain"
        glf_gameTilted = False
        DispatchPinEvent m_name & "_clear", Null
    End Sub
    
    Public Sub SlamTilt()
        'Process a slam tilt.
        'This method posts the *slam_tilt* event and (if a game is active) sets
        'the game mode's ``slam_tilted`` attribute to *True*.
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TiltHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim tilt : Set tilt = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set TiltHandler = kwargs
                Else
                    TiltHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "_tilt_warning_switch_active":
            tilt.HandleTiltWarningSwitch
        Case "tilt_ball_ending"
            kwargs.Add "wait_for", tilt.Name & "_clear"
            tilt.BallEndingTilted
        Case "tilt_ball_drain"
            tilt.TiltedBallDrain kwargs
            kwargs = kwargs -1
        Case "tilt_done"
            tilt.TiltDone
        Case "reset_warnings"
            tilt.ResetWarnings
    End Select

    If IsObject(args(1)) Then
        Set TiltHandler = kwargs
    Else
        TiltHandler = kwargs
    End If
End Function

Class GlfTimedSwitches

    Private m_name
    Private m_priority
    Private m_base_device
    Private m_time
    Private m_switches
    Private m_events_when_active
    Private m_events_when_released
    Private m_active_switches
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        'Select Case value
            'Case   
        'End Select
        GetValue = True
    End Property

    Public Property Let Time(value): Set m_time = CreateGlfInput(value): End Property
    Public Property Get Time(): Time = m_time.Value(): End Property
    
    Public Property Let Switches(value): m_switches = value: End Property
    
    Public Property Let EventsWhenActive(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_active.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let EventsWhenReleased(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_released.Add newEvent.Raw, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

    Public default Function init(name, mode)
        m_name = "timed_switch_" & name
        m_priority = mode.Priority
        Set m_events_when_active = CreateObject("Scripting.Dictionary")
        Set m_events_when_released = CreateObject("Scripting.Dictionary")
        Set m_time = CreateGlfInput(0)
        m_switches = Array()
        Set m_active_switches = CreateObject("Scripting.Dictionary")
        Set m_base_device = (new GlfBaseModeDevice)(mode, "timed_switch", Me)
        Set Init = Me
    End Function

    Public Sub Activate()
        Dim switch
        For Each switch in m_switches
            AddPinEventListener switch & "_active", m_name & "_active", "TimedSwitchHandler", m_priority, Array("active", Me, switch)
            AddPinEventListener switch & "_inactive", m_name & "_inactive", "TimedSwitchHandler", m_priority, Array("inactive", Me, switch)
        Next
    End Sub

    Public Sub Deactivate()
        Dim switch
        For Each switch in m_switches
            RemovePinEventListener switch & "_active", m_name & "_active"
            RemovePinEventListener switch & "_inactive", m_name & "_inactive"
        Next
    End Sub

    Public Sub SwitchActive(switch)
        If UBound(m_active_switches.Keys()) = -1 Then
            Dim evt
            For Each evt in m_events_when_active.Keys()
                Log "Switch Active: " & switch & ". Event: " & m_events_when_active(evt).EventName
                DispatchPinEvent m_events_when_active(evt).EventName, Null
            Next
        End If
        If Not m_active_switches.Exists(switch) Then
            m_active_switches.Add switch, True
        End If
    End Sub

    Public Sub SwitchInactive(switch)
        RemoveDelay m_name & "_" & switch & "_active"
        If m_active_switches.Exists(switch) Then
            m_active_switches.Remove switch
            If UBound(m_active_switches.Keys()) = -1 Then
                Dim evt
                For Each evt in m_events_when_released.Keys()
                    Log "Switch Release: " & switch & ". Event: " & m_events_when_released(evt).EventName
                    DispatchPinEvent m_events_when_released(evt).EventName, Null
                Next
            End If
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TimedSwitchHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim timed_switch : Set timed_switch = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set TimedSwitchHandler = kwargs
                Else
                    TimedSwitchHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "active"
            SetDelay timed_switch.Name & "_" & ownProps(2) & "_active" , "TimedSwitchHandler" , Array(Array("passed_time", timed_switch, ownProps(2)),Null), timed_switch.Time
        Case "passed_time"
            timed_switch.SwitchActive ownProps(2)
        Case "inactive"
            timed_switch.SwitchInactive ownProps(2)
    End Select

    If IsObject(args(1)) Then
        Set TimedSwitchHandler = kwargs
    Else
        TimedSwitchHandler = kwargs
    End If
End Function



Class GlfTimer

    Private m_name
    Private m_priority
    Private m_mode
    Private m_base_device
    Private m_debug

    Private m_control_events
    Private m_running
    Private m_ticks
    Private m_ticks_remaining
    Private m_start_value
    Private m_end_value
    Private m_direction
    Private m_tick_interval
    Private m_starting_tick_interval
    Private m_max_value
    Private m_restart_on_complete
    Private m_start_running

    Public Property Get Name() : Name = m_name : End Property
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property
    

    Public Property Get ControlEvents()
        Dim count : count = UBound(m_control_events.Keys) 
        Dim newEvent : Set newEvent = (new GlfTimerControlEvent)()
        m_control_events.Add CStr(count), newEvent
        Set ControlEvents = newEvent
    End Property
    Public Property Get StartValue() : StartValue = m_start_value : End Property
    Public Property Get EndValue() : EndValue = m_end_value : End Property
    Public Property Get Direction() : Direction = m_direction : End Property
    Public Property Let StartValue(value) : Set m_start_value = CreateGlfInput(value) : End Property
    Public Property Let EndValue(value) : Set m_end_value = CreateGlfInput(value) : End Property
    Public Property Let Direction(value) : m_direction = value : End Property
    Public Property Let MaxValue(value) : m_max_value = value : End Property
    Public Property Let RestartOnComplete(value) : m_restart_on_complete = value : End Property
    Public Property Let StartRunning(value) : m_start_running = value : End Property
    Public Property Let TickInterval(value)
        m_tick_interval = value
        m_starting_tick_interval = value
    End Property

    Public Property Get GetValue(value)
        Select Case value
            Case "ticks"
                GetValue = m_ticks
            Case "ticks_remaining"
              GetValue = m_ticks_remaining
        End Select
    End Property

	Public default Function init(name, mode)
        m_name = "timer_" & name
        m_mode = mode.Name
        m_priority = mode.Priority
        m_direction = "up"
        m_ticks = 0
        m_ticks_remaining = 0
        m_tick_interval = 1000
        m_starting_tick_interval = 1000
        m_restart_on_complete = False
        m_start_running = False
        Set m_start_value = CreateGlfInput(0)
        Set m_end_value = CreateGlfInput(-1)

        Set m_control_events = CreateObject("Scripting.Dictionary")
        m_running = False

        Set m_base_device = (new GlfBaseModeDevice)(mode, "timer", Me)

        glf_timers.Add name, Me

        Set Init = Me
	End Function

    Public Sub Activate()
        Dim evt
        For Each evt in m_control_events.Keys
            AddPinEventListener m_control_events(evt).EventName.EventName, m_name & "_action", "TimerEventHandler", m_priority+m_control_events(evt).EventName.Priority, Array("action", Me, m_control_events(evt))
        Next
        m_ticks = m_start_value.Value
        m_ticks_remaining = m_ticks
        Log "Activating Timer"
        If m_start_running = True Then
            StartTimer()
        End If
    End Sub

    Public Sub Deactivate()
        Dim evt
        For Each evt in m_control_events.Keys
            RemovePinEventListener m_control_events(evt).EventName.EventName, m_name & "_action"
        Next
        RemoveDelay m_name & "_tick"
        m_running = False
    End Sub

    Public Sub Action(controlEvent)

        If Not IsNull(controlEvent.EventName) Then
            If controlEvent.EventName().Evaluate() = False Then
                Exit Sub
            End IF
        End If

        dim value : value = controlEvent.Value
        Select Case controlEvent.Action
            Case "add"
                Add value
            Case "subtract"
                Subtract value
            Case "jump"
                Jump value
            Case "start"
                StartTimer()
            Case "stop"
                StopTimer()
            Case "reset"
                Reset()
            Case "restart"
                Restart()
            Case "pause"
                Pause value
            Case "set_tick_interval"
                SetTickInterval value
            Case "change_tick_interval"
                ChangeTickInterval value
            Case "reset_tick_interval"
                SetTickInterval m_starting_tick_interval
        End Select

    End Sub

    Private Sub StartTimer()
        If m_running Then
            Exit Sub
        End If

        Log "Starting Timer"
        m_running = True
        RemoveDelay m_name & "_unpause"
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_started", kwargs
        PostTickEvents()
        SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval
    End Sub

    Private Sub StopTimer()
        Log "Stopping Timer"
        m_running = False
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_stopped", kwargs
        RemoveDelay m_name & "_tick"
    End Sub

    Public Sub Pause(pause_ms)
        Log "Pausing Timer for "&pause_ms&" ms"
        m_running = False
        RemoveDelay m_name & "_tick"
        
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_paused", kwargs

        If pause_ms > 0 Then
            Dim startControlEvent : Set startControlEvent = (new GlfTimerControlEvent)()
            startControlEvent.Action = "start"
            SetDelay m_name & "_unpause", "TimerEventHandler", Array(Array("action", Me, startControlEvent), Null), pause_ms
        End If
    End Sub 

    Public Sub Tick()
        Log "Timer Tick"
        If Not m_running Then
            Log "Timer is not running. Will remove."
            Exit Sub
        End If

        Dim newValue
        If m_direction = "down" Then
            newValue = m_ticks - 1
        Else
            newValue = m_ticks + 1
        End If
        
        Log "ticking: old value: "& m_ticks & ", new Value: " & newValue & ", target: "& m_end_value.Value
        m_ticks = newValue
        If Not PostTickEvents() Then
            SetDelay m_name & "_tick", "TimerEventHandler", Array(Array("tick", Me), Null), m_tick_interval    
        End If
    End Sub

    Private Function CheckForDone

        ' Checks to see if this timer is done. Automatically called anytime the
        ' timer's value changes.
        Log "Checking to see if timer is done. Ticks: "&m_ticks&", End Value: "&m_end_value.Value&", Direction: "& m_direction

        if m_direction = "up" And m_end_value.Value<>-1 And m_ticks >= m_end_value.Value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If m_direction = "down" And m_ticks <= m_end_value.Value Then
            TimerComplete()
            CheckForDone = True
            Exit Function
        End If

        If m_end_value.Value<>-1 Then 
            m_ticks_remaining = abs(m_end_value.Value - m_ticks)
        End If
        Log "Timer is not done"

        CheckForDone = False

    End Function

    Private Sub TimerComplete

        Log "Timer Complete"

        StopTimer()
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_complete", kwargs
        
        If m_restart_on_complete Then
            Log "Restart on complete: True"
            Restart()
        End If
    End Sub

    Private Sub Restart
        Reset()
        If Not m_running Then
            StartTimer()
        Else
            PostTickEvents()
        End If
    End Sub

    Private Sub Reset
        Log "Resetting timer. New value: "& m_start_value.Value
        Jump m_start_value.Value
    End Sub

    Private Sub Jump(timer_value)
        m_ticks = timer_value

        If m_max_value and m_ticks > m_max_value Then
            m_ticks = m_max_value
        End If

        CheckForDone()
    End Sub

    Public Sub ChangeTickInterval(change)
        m_tick_interval = m_tick_interval * change
    End Sub

    Public Sub SetTickInterval(timer_value)
        m_tick_interval = timer_value
    End Sub
        
    Private Function PostTickEvents()
        PostTickEvents = True
        If Not CheckForDone() Then
            PostTickEvents = False
            Dim kwargs : Set kwargs = GlfKwargs()
            With kwargs
                .Add "ticks", m_ticks
                .Add "ticks_remaining", m_ticks_remaining
            End With
            DispatchPinEvent m_name & "_tick", kwargs
            Log "Ticks: "&m_ticks&", Remaining: " & m_ticks_remaining
        End If
    End Function

    Public Sub Add(timer_value) 
        Dim new_value

        new_value = m_ticks + timer_value

        If Not IsEmpty(m_max_value) And new_value > m_max_value Then
            new_value = m_max_value
        End If
        m_ticks = new_value
        timer_value = new_value - timer_value

        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_added", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_added", kwargs
        CheckForDone()
    End Sub

    Public Sub Subtract(timer_value)
        m_ticks = m_ticks - timer_value
        Dim kwargs : Set kwargs = GlfKwargs()
        With kwargs
            .Add "ticks", m_ticks
            .Add "ticks_subtracted", timer_value
            .Add "ticks_remaining", m_ticks_remaining
        End With
        DispatchPinEvent m_name & "_time_subtracted", kwargs
        
        CheckForDone()
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function TimerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim timer : Set timer = ownProps(1)
    'debug.print "TimerEventHandler: " & timer.Name & ": " & evt
    Select Case evt
        Case "action"
            Dim controlEvent : Set controlEvent = ownProps(2)
            timer.Action controlEvent
        Case "tick"
            timer.Tick 
    End Select
    If IsObject(args(1)) Then
        Set TimerEventHandler = kwargs
    Else
        TimerEventHandler = kwargs
    End If
End Function

Class GlfTimerControlEvent
	Private m_event, m_action, m_value
  
	Public Property Get EventName(): Set EventName = m_event: End Property
    Public Property Let EventName(input)
        Dim newEvent : Set newEvent = (new GlfEvent)(input)
        Set m_event = newEvent
    End Property

    Public Property Get Action(): Action = m_action : End Property
    Public Property Let Action(input): m_action = input : End Property

    Public Property Get Value()
        If Not IsNull(m_value) Then
            Value = m_value.Value
        Else
            Value = 0
        End If
    End Property
    Public Property Let Value(input)
        Set m_value = CreateGlfInput(input)
    End Property

	Public default Function init()
        m_event = Null
        m_action = Empty
        m_value = Null
	    Set Init = Me
	End Function

End Class
Class GlfVariablePlayer

    Private m_priority
    Private m_name
    Private m_mode
    Private m_events
    Private m_debug
    private m_base_device

    Private m_value

    Public Property Get Name() : Name = m_name : End Property

    Public Property Get EventName(name)
        Dim newEvent : Set newEvent = (new GlfVariablePlayerEvent)(name)
        m_events.Add newEvent.BaseEvent.Raw, newEvent
        Set EventName = newEvent
    End Property
   
    Public Property Let Debug(value)
        m_debug = value
        m_base_device.Debug = value
    End Property
    Public Property Get IsDebug()
        If m_debug Then : IsDebug = 1 : Else : IsDebug = 0 : End If
    End Property

	Public default Function init(mode)
        m_name = "variable_player_" & mode.name
        m_mode = mode.Name
        m_priority = mode.Priority

        Set m_events = CreateObject("Scripting.Dictionary")
        m_debug = False
        Set m_base_device = (new GlfBaseModeDevice)(mode, "variable_player", Me)
        Set Init = Me
	End Function

    Public Sub Activate()
        Log "Activating"
        Dim evt
        For Each evt In m_events.Keys()
            AddPinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_variable_player_play", "VariablePlayerEventHandler", m_priority+m_events(evt).BaseEvent.Priority, Array("play", Me, evt)
        Next
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        Dim evt
        For Each evt In m_events.Keys()
            RemovePinEventListener m_events(evt).BaseEvent.EventName, m_mode & "_variable_player_play"
        Next
    End Sub

    Public Sub Play(evt)
        Log "Playing: " & evt
        If m_events(evt).BaseEvent.Evaluate() = False Then
            Exit Sub
        End If
        Dim vKey, v
        For Each vKey in m_events(evt).Variables.Keys
            Set v = m_events(evt).Variable(vKey)
            Dim varValue : varValue = v.VariableValue
            Select Case v.Action
                Case "add"
                    Log "Add Variable " & vKey & ". New Value: " & CStr(GetPlayerState(vKey) + varValue) & " Old Value: " & CStr(GetPlayerState(vKey))
                    SetPlayerState vKey, GetPlayerState(vKey) + varValue
                Case "add_machine"
                    Log "Add Machine Variable " & vKey & ". New Value: " & CStr(GetPlayerState(vKey) + varValue) & " Old Value: " & CStr(GetPlayerState(vKey))
                    'SetPlayerState vKey, GetPlayerState(vKey) + varValue
                    glf_machine_vars(vkey).Value = glf_machine_vars(vkey).Value + varValue
                Case "set"
                    Log "Setting Variable " & vKey & ". New Value: " & CStr(varValue)
                    SetPlayerState vKey, varValue
                Case "set_machine"
                    Log "Setting Machine Variable " & vKey & ". New Value: " & CStr(varValue)
                    glf_machine_vars(vkey).Value = varValue
        End Select
        Next
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_mode & "_variable_player", message
        End If
    End Sub
End Class

Class GlfVariablePlayerEvent

    Private m_event
	Private m_variables

    Public Property Get BaseEvent() : Set BaseEvent = m_event : End Property
  
	Public Property Get Variable(name)
        If m_variables.Exists(name) Then
            Set Variable = m_variables(name)
        Else
            Dim new_variable : Set new_variable = (new GlfVariablePlayerItem)()
            m_variables.Add name, new_variable
            Set Variable = new_variable
        End If
    End Property
    
    Public Property Get Variables(): Set Variables = m_variables End Property

	Public default Function init(evt)
        Set m_event = (new GlfEvent)(evt)
        Set m_variables = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

End Class

Class GlfVariablePlayerItem
	Private m_block, m_show, m_float, m_int, m_string, m_player, m_action, m_type
  
	Public Property Get Action(): Action = m_action: End Property
    Public Property Let Action(input): m_action = input: End Property

    Public Property Get Block(): Block = m_block End Property
    Public Property Let Block(input): m_block = input End Property

	Public Property Let Float(input): m_float = Glf_ParseInput(input): m_type = "float" : End Property
  
	Public Property Let Int(input): m_int = Glf_ParseInput(input): m_type = "int" : End Property
  
	Public Property Let String(input): m_string = input: m_type = "string" : End Property

    Public Property Get VariableType(): VariableType = m_type: End Property
    Public Property Get VariableValue()
        Select Case m_type
            Case "float"
                VariableValue = GetRef(m_float(0))()
            Case "int"
                VariableValue = GetRef(m_int(0))()
            Case "string"
                VariableValue = m_string
            Case Else
                VariableValue = Empty
        End Select
    End Property

    Public Property Get Player(): Player = m_player: End Property
    Public Property Let Player(input): m_player = input: End Property

	Public default Function init()
        m_action = "add"
        m_type = Empty
        m_block = False
        m_float = Empty
        m_int = Empty
        m_string = Empty
        m_player = Empty
	    Set Init = Me
	End Function

End Class


Function CreateMachineVar(name)
	Dim machine_var : Set machine_var = (new GlfMachineVars)(name)
	Set CreateMachineVar = machine_var
End Function

Class GlfMachineVars

    Private m_name
	Private m_persist
    Private m_value_type
    Private m_value

    Public Property Get Name(): Name = m_name : End Property
    Public Property Let Name(input): m_name = input : End Property

    Public Property Get Value(): Value = m_value : End Property
    Public Property Let Value(input): m_value = input : End Property

    Public Property Let InitialValue(input)
        m_value = input
    End Property

    Public Property Get Persist(): Persist = m_persist : End Property
    Public Property Let Persist(input): m_persist = input : End Property

    Public Property Get ValueType(): ValueType = m_value_type : End Property
    Public Property Let ValueType(input): m_value_type = input : End Property

    Public Function GetValue()
        GetValue = m_value
    End Function

	Public default Function init(name)
        m_name = name
        m_persist = True
        m_value_type = "int"
        m_value = 0
        glf_machine_vars.Add name, Me
	    Set Init = Me
	End Function
End Class

Function VariablePlayerEventHandler(args)
    
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1)
    End If
    Dim evt : evt = ownProps(0)
    Dim variablePlayer : Set variablePlayer = ownProps(1)
    Select Case evt
        Case "play"
            variablePlayer.Play ownProps(2)
    End Select
    If IsObject(args(1)) Then
        Set VariablePlayerEventHandler = kwargs
    Else
        VariablePlayerEventHandler = kwargs
    End If
    
End Function


Class DelayObject
	Private m_name, m_callback, m_ttl, m_args
  
	Public Property Get Name(): Name = m_name: End Property
	Public Property Let Name(input): m_name = input: End Property
  
	Public Property Get Callback(): Callback = m_callback: End Property
	Public Property Let Callback(input): m_callback = input: End Property
  
	Public Property Get TTL(): TTL = m_ttl: End Property
	Public Property Let TTL(input): m_ttl = input: End Property
  
	Public Property Get Args(): Args = m_args: End Property
	Public Property Let Args(input): m_args = input: End Property
  
	Public default Function init(name, callback, ttl, args)
	  m_name = name
	  m_callback = callback
	  m_ttl = ttl
	  m_args = args

	  Set Init = Me
	End Function
End Class

Dim delayQueue : Set delayQueue = CreateObject("Scripting.Dictionary")
Dim delayQueueMap : Set delayQueueMap = CreateObject("Scripting.Dictionary")
Dim delayCallbacks : Set delayCallbacks = CreateObject("Scripting.Dictionary")

Sub SetDelay(name, callbackFunc, args, delayInMs)
    Dim executionTime
    executionTime = gametime + delayInMs
    
    RemoveDelay name
    If delayQueueMap.Exists(name) Then
        delayQueueMap.Remove name
    End If
    

    If delayQueue.Exists(executionTime) Then
        If delayQueue(executionTime).Exists(name) Then
            delayQueue(executionTime).Remove name
        End If
    Else
        delayQueue.Add executionTime, CreateObject("Scripting.Dictionary")
    End If
    'Glf_WriteDebugLog "Delay", "Adding delay for " & name & ", callback: " & callbackFunc & ", ExecutionTime: " & executionTime
    delayQueue(executionTime).Add name, (new DelayObject)(name, callbackFunc, executionTime, args)
    delayQueueMap.Add name, executionTime
    
End Sub

Function AlignToNearest10th(timeMs)
    AlignToNearest10th = Int(timeMs / 100) * 100
End Function

Function RemoveDelay(name)
    If delayQueueMap.Exists(name) Then
        If delayQueue.Exists(delayQueueMap(name)) Then
            If delayQueue(delayQueueMap(name)).Exists(name) Then
                'Glf_WriteDebugLog "Delay", "Removing delay for " & name & " and  Execution Time: " & delayQueueMap(name)
                delayQueue(delayQueueMap(name)).Remove name
            End If
            delayQueueMap.Remove name
            RemoveDelay = True
            'Glf_WriteDebugLog "Delay", "Removing delay for " & name
            Exit Function
        End If
    End If
    RemoveDelay = False
End Function

Sub DelayTick()
    Dim queueItem, key, delayObject
    For Each queueItem in delayQueue.Keys()
        If Int(queueItem) < gametime Then
            For Each key In delayQueue(queueItem).Keys()
                If IsObject(delayQueue(queueItem)(key)) Then
                            Set delayObject = delayQueue(queueItem)(key)
                            'Glf_WriteDebugLog "Delay", "Executing delay: " & key & ", callback: " & delayObject.Callback
                            GetRef(delayObject.Callback)(delayObject.Args)    
                End If
            Next
            delayQueue.Remove queueItem
        End If
    Next
End Sub

Function CreateGlfAutoFireDevice(name)
	Dim flipper : Set flipper = (new GlfAutoFireDevice)(name)
	Set CreateGlfAutoFireDevice = flipper
End Function

Class GlfAutoFireDevice

    Private m_name
    Private m_enable_events
    Private m_disable_events
    Private m_enabled
    Private m_switch
    Private m_action_cb
    Private m_disabled_cb
    Private m_enabled_cb
    Private m_debug

    Public Property Let Switch(value)
        m_switch = value
    End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let DisabledCallback(value) : m_disabled_cb = value : End Property
    Public Property Let EnabledCallback(value) : m_enabled_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "AutoFireDeviceEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "AutoFireDeviceEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "auto_fire_coil_" & name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_enabled = False
        m_action_cb = Empty
        m_disabled_cb = Empty
        m_enabled_cb = Empty
        m_switch = Empty
        m_debug = False
        glf_autofiredevices.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        If Not IsEmpty(m_enabled_cb) Then
            GetRef(m_enabled_cb)()
        End If
        If Not IsEmpty(m_switch) Then
            AddPinEventListener m_switch & "_active", m_name & "_active", "AutoFireDeviceEventHandler", 1000, Array("activate", Me)
            AddPinEventListener m_switch & "_inactive", m_name & "_inactive", "AutoFireDeviceEventHandler", 1000, Array("deactivate", Me)
        End If
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        If Not IsEmpty(m_disabled_cb) Then
            GetRef(m_disabled_cb)()
        End If
        Deactivate(Null)
        RemovePinEventListener m_switch & "_active", m_name & "_active"
        RemovePinEventListener m_switch & "_inactive", m_name & "_inactive"
    End Sub

    Public Sub Activate(active_ball)
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(1, active_ball))
        End If
        DispatchPinEvent m_name & "_activate", Null
    End Sub

    Public Sub Deactivate(active_ball)
        Log "Deactivating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(0, active_ball))
        End If
        DispatchPinEvent m_name & "_deactivate", Null
    End Sub

    Public Sub BallSearch(phase)
        Log "Ball Search, phase " & phase
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(Array(1, Null))
        End If
        SetDelay m_name & "ball_search_deactivate", "AutoFireDeviceEventHandler", Array(Array("deactivate", Me), Null), 150
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function AutoFireDeviceEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim flipper : Set flipper = ownProps(1)
    Select Case evt
        Case "enable"
            flipper.Enable
        Case "disable"
            flipper.Disable
        Case "activate"
            flipper.Activate kwargs
        Case "deactivate"
            flipper.Deactivate kwargs
    End Select
    If IsObject(args(1)) Then
        Set AutoFireDeviceEventHandler = kwargs
    Else
        AutoFireDeviceEventHandler = kwargs
    End If
End Function
Function CreateGlfBallDevice(name)
	Dim device : Set device = (new GlfBallDevice)(name)
	Set CreateGlfBallDevice = device
End Function

Class GlfBallDevice

    Private m_name
    Private m_ball_switches
    Private m_player_controlled_eject_event
    Private m_eject_timeout
    Private m_eject_enable_time
    Private m_balls
    Private m_balls_in_device
    Private m_default_device
    Private m_eject_callback
    Private m_eject_all_events
    Private m_balls_to_eject
    Private m_ejecting_all
    Private m_ejecting
    Private m_mechanical_eject
    Private m_eject_targets
    Private m_entrance_count_delay
    Private m_incoming_balls
    Private m_lost_balls
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "balls":
                GetValue = m_balls_in_device
        End Select
    End Property
    Public Property Let DefaultDevice(value)
        m_default_device = value
        If m_default_device = True Then
            Set glf_plunger = Me
        End If
    End Property
	Public Property Get HasBall()
        HasBall = (Not IsNull(m_balls(0)) And m_ejecting = False)
    End Property
    
    Public Property Get Balls(): Balls = m_balls_in_device : End Property

    Public Property Let AddIncomingBalls(value) : m_incoming_balls = m_incoming_balls + value : End Property
    Public Property Get IncomingBalls() : IncomingBalls = m_incoming_balls : End Property

    Public Property Let EjectCallback(value) : m_eject_callback = value : End Property
    Public Property Let EjectEnableTime(value) : m_eject_enable_time = value : End Property
        
    Public Property Let EjectTimeout(value) : m_eject_timeout = value : End Property
    Public Property Let EntranceCountDelay(value) : m_entrance_count_delay = value : End Property
    Public Property Let EjectAllEvents(value)
        m_eject_all_events = value
        Dim evt
        For Each evt in m_eject_all_events
            AddPinEventListener evt, m_name & "_eject_all", "BallDeviceEventHandler", 1000, Array("ball_eject_all", Me)
        Next
    End Property
    Public Property Let EjectTargets(value)
        m_eject_targets = value
        Dim evt
        For Each evt in m_eject_targets
            AddPinEventListener evt & "_active", m_name & "_eject_target", "BallDeviceEventHandler", 1000, Array("eject_timeout", Me)
        Next
    End Property
    Public Property Let PlayerControlledEjectEvent(value)
        m_player_controlled_eject_event = value
        AddPinEventListener m_player_controlled_eject_event, m_name & "_eject_attempt", "BallDeviceEventHandler", 1000, Array("ball_eject", Me)
    End Property
    Public Property Let BallSwitches(value)
        m_ball_switches = value
        ReDim m_balls(Ubound(m_ball_switches))
        Dim x
        For x=0 to UBound(m_ball_switches)
            m_balls(x) = Null
            AddPinEventListener m_ball_switches(x)&"_active", m_name & "_ball_enter", "BallDeviceEventHandler", 1000, Array("ball_entering", Me, x)
            AddPinEventListener m_ball_switches(x)&"_inactive", m_name & "_ball_exiting", "BallDeviceEventHandler", 1000, Array("ball_exiting", Me, x)
        Next
    End Property
    Public Property Let MechanicalEject(value) : m_mechanical_eject = value : End Property


    Public Property Let Debug(value) : m_debug = value : End Property
        
	Public default Function init(name)
        m_name = "balldevice_" & name
        m_ball_switches = Array()
        m_eject_all_events = Array()
        m_eject_targets = Array()
        m_balls = Array()
        m_debug = False
        m_default_device = False
        m_ejecting = False
        m_eject_callback = Null
        m_ejecting_all = False
        m_balls_to_eject = 0
        m_balls_in_device = 0
        m_lost_balls = 0
        m_mechanical_eject = False
        m_eject_timeout = 1000
        m_eject_enable_time = 0
        m_entrance_count_delay = 500
        m_incoming_balls = 0
        glf_ball_devices.Add name, Me
	    Set Init = Me
	End Function

    Public Sub BallEntering(ball, switch)
        Log "Ball Entering" 
        If m_default_device = False Then
            SetDelay m_name & "_" & switch & "_ball_enter", "BallDeviceEventHandler", Array(Array("ball_enter", Me, switch), ball), m_entrance_count_delay
        Else
            BallEnter ball, switch
        End If
    End Sub

    Public Sub BallEnter(ball, switch)
        RemoveDelay m_name & "_switch" & switch & "_eject_timeout"
        Set m_balls(switch) = ball
        m_balls_in_device = m_balls_in_device + 1
        Log "Ball Entered"
        If m_lost_balls > 0 Then
            m_lost_balls = m_lost_balls - 1
            Log "Lost Ball Found"
            If m_lost_balls = 0 Then
                RemoveDelay m_name & "_clear_lost_balls"
            End If
            Exit Sub
        End If

        Dim unclaimed_balls: unclaimed_balls = 1
        If m_incoming_balls > 0 Then
            unclaimed_balls = 0
        End If
        If unclaimed_balls > 0 Then
            unclaimed_balls = DispatchRelayPinEvent(m_name & "_ball_enter", 1)
        End If        
        Log "Unclaimed Balls: " & unclaimed_balls
        DispatchPinEvent m_name & "_ball_entered", Null
        If unclaimed_balls > 0 Then
            SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 500
        End If
    End Sub

    Public Sub BallExiting(ball, switch)
        RemoveDelay m_name & "_" & switch & "_ball_enter"
        If m_ejecting = False And m_mechanical_eject = False Then
            Log "Ball Lost, Wasn't Ejecting"
            m_lost_balls = m_lost_balls + 1
            SetDelay m_name & "_clear_lost_balls", "BallDeviceEventHandler", Array(Array("clear_lost_balls", Me), Null), 3000
        End If
        m_balls(switch) = Null
        m_balls_in_device = m_balls_in_device - 1
        DispatchPinEvent m_name & "_ball_exiting", Null
        If m_mechanical_eject = True And m_eject_timeout > 0 Then
            SetDelay m_name & "_switch" & switch & "_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), ball), m_eject_timeout
        End If
        Log "Ball Exiting"
    End Sub

    Public Sub BallExitSuccess(ball)
        m_ejecting = False

        If m_incoming_balls > 0 Then
            m_incoming_balls = m_incoming_balls - 1
        End If
        DispatchPinEvent m_name & "_ball_eject_success", Null
        Log "Ball successfully exited"
        If m_ejecting_all = True Then
            If m_balls_to_eject = 0 Then
                m_ejecting_all = False
                Exit Sub
            End If
            If Not IsNull(m_balls(0)) Then
                m_balls_to_eject = m_balls_to_eject - 1
                Eject()
            Else
                SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), ball), 600
            End If
        End If
    End Sub

    Public Sub Eject
        
        If Not IsNull(m_eject_callback) Then
            If Not IsNull(m_balls(0)) Then
                Log "Ejecting."
                SetDelay m_name & "_switch0_eject_timeout", "BallDeviceEventHandler", Array(Array("eject_timeout", Me), m_balls(0)), m_eject_timeout
                m_ejecting = True
            
                GetRef(m_eject_callback)(m_balls(0))
                If m_eject_enable_time > 0 Then
                    SetDelay m_name & "_eject_enable_time", "BallDeviceEventHandler", Array(Array("eject_enable_complete", Me), m_balls(0)), m_eject_enable_time
                End If
            Else
                SetDelay m_name & "_eject_attempt", "BallDeviceEventHandler", Array(Array("ball_eject", Me), Null), 600
            End If
        End If
    End Sub

    Public Sub EjectEnableComplete
        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(Null)
        End If
    End Sub

    Public Sub EjectBalls(balls)
        Log "Ejecting "&balls&" Balls."
        m_ejecting_all = True
        m_balls_to_eject = balls - 1
        Eject()
    End Sub

    Public Sub EjectAll
        Log "Ejecting All." 
        m_ejecting_all = True
        m_balls_to_eject = m_balls_in_device - 1 
        Eject()
    End Sub

    Public Sub ClearLostBalls()
        m_lost_balls = 0
    End Sub

    Public Sub BallSearch(phase)
        Log "Ball Search, phase " & phase
        If m_default_device = True Then
            Exit Sub
        End If
        If phase = 1 And HasBall() Then
            Exit Sub
        End If
        If Not IsNull(m_eject_callback) Then
            GetRef(m_eject_callback)(m_balls(0))
        End If
    End Sub

    Public Function ToYaml
        Dim yaml
        yaml = "  " & m_name & ":" & vbCrLf
        yaml = yaml + "    ball_switches: " & Join(m_ball_switches, ",") & vbCrLf
        yaml = yaml + "    mechanical_eject: " & m_mechanical_eject & vbCrLf
        
        ToYaml = yaml
    End Function

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function BallDeviceEventHandler(args)
    Dim ownProps, ball
    ownProps = args(0)
    Dim evt : evt = ownProps(0)
    Dim ballDevice : Set ballDevice = ownProps(1)
    Dim switch
    'debug.print "Ball Device: " & ballDevice.Name & ". Event: " & evt
    Select Case evt
        Case "ball_entering"
            Set ball = args(1)
            switch = ownProps(2)
            ballDevice.BallEntering ball, switch
        Case "ball_enter"
            Set ball = args(1)
            switch = ownProps(2)
            ballDevice.BallEnter ball, switch
        Case "ball_eject"
            ballDevice.Eject
        Case "ball_eject_all"
            ballDevice.EjectAll
        Case "ball_exiting"
            switch = ownProps(2)
            If RemoveDelay(ballDevice.Name & "_" & switch & "_ball_enter") = False Then
                Set ball = args(1)
                ballDevice.BallExiting ball, switch
            End If
        Case "eject_timeout"
            Set ball = args(1)
            ballDevice.BallExitSuccess ball
        Case "eject_enable_complete"
            ballDevice.EjectEnableComplete
        Case "clear_lost_balls"
            ballDevice.ClearLostBalls
    End Select
End Function
Function CreateGlfDiverter(name)
	Dim diverter : Set diverter = (new GlfDiverter)(name)
	Set CreateGlfDiverter = diverter
End Function

Class GlfDiverter

    Private m_name
    Private m_activate_events
    Private m_deactivate_events
    Private m_activation_time
    Private m_enable_events
    Private m_disable_events
    Private m_activation_switches
    Private m_action_cb
    Private m_enabled
    Private m_active
    Private m_ball_search_hold_time
    Private m_debug

    Public Property Get Name(): Name = m_name : End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "enabled":
                GetValue = m_enabled
            Case "active":
                GetValue = m_active
        End Select
    End Property

    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "DiverterEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "DiverterEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let ActivateEvents(value) 
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_activate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let DeactivateEvents(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_deactivate_events.Add newEvent.Raw, newEvent
        Next
    End Property
    Public Property Let ActivationTime(value) : Set m_activation_time = CreateGlfInput(value) : End Property
    Public Property Let ActivationSwitches(value) : m_activation_switches = value : End Property
    Public Property Let BallSearchHoldTime(value) : Set m_ball_search_hold_time = CreateGlfInput(value) : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "diverter_" & name
        m_enable_events = Array()
        m_disable_events = Array()
        Set m_activate_events = CreateObject("Scripting.Dictionary")
        Set m_deactivate_events = CreateObject("Scripting.Dictionary")
        m_activation_switches = Array()
        Set m_activation_time = CreateGlfInput(0)
        Set m_ball_search_hold_time = CreateGlfInput(1000)
        m_debug = False
        m_enabled = False
        m_active = False
        glf_diverters.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        For Each evt in m_activate_events.Keys()
            AddPinEventListener m_activate_events(evt).EventName, m_name & "_" & evt & "_activate", "DiverterEventHandler", 1000, Array("activate", Me, m_activate_events(evt))
        Next
        For Each evt in m_deactivate_events.Keys()
            AddPinEventListener m_deactivate_events(evt).EventName, m_name & "_" & evt & "_deactivate", "DiverterEventHandler", 1000, Array("deactivate", Me, m_deactivate_events(evt))
        Next
        For Each evt in m_activation_switches
            AddPinEventListener evt & "_active", m_name & "_activate", "DiverterEventHandler", 1000, Array("activate", Me)
        Next
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Dim evt
        For Each evt in m_activate_events.Keys()
            RemovePinEventListener m_activate_events(evt).EventName, m_name & "_" & evt & "_activate"
        Next
        For Each evt in m_deactivate_events.Keys()
            RemovePinEventListener m_deactivate_events(evt).EventName, m_name & "_" & evt & "_deactivate"
        Next
        For Each evt in m_activation_switches
            RemovePinEventListener evt & "_active", m_name & "_activate"
        Next
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
    End Sub

    Public Sub Activate()
        Log "Activating"
        m_active = True
        GetRef(m_action_cb)(1)
        If m_activation_time.Value > 0 Then
            SetDelay m_name & "_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_activation_time.Value
        End If
        DispatchPinEvent m_name & "_activating", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        m_active = False
        RemoveDelay m_name & "_deactivate"
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_deactivating", Null
    End Sub

    Public Sub BallSearch(phase)
        Log "Ball Search, phase " & phase
        If m_active = False Then
            If Not IsEmpty(m_action_cb) Then
                m_active = True
                GetRef(m_action_cb)(1)
            End If
            SetDelay m_name & "ball_search_deactivate", "DiverterEventHandler", Array(Array("deactivate", Me), Null), m_ball_search_hold_time.Value
        Else
            If Not IsEmpty(m_action_cb) Then
                m_active = False
                GetRef(m_action_cb)(0)
            End If
            SetDelay m_name & "ball_search_activate", "DiverterEventHandler", Array(Array("activate", Me), Null), m_ball_search_hold_time.Value
        End If
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DiverterEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim diverter : Set diverter = ownProps(1)
    'Check if the evt has a condition to evaluate    
    If UBound(ownProps) = 2 Then
        If IsObject(ownProps(2)) Then
            If ownProps(2).Evaluate() = False Then
                If IsObject(args(1)) Then
                    Set DiverterEventHandler = kwargs
                Else
                    DiverterEventHandler = kwargs
                End If
                Exit Function
            End If
        End If
    End If
    Select Case evt
        Case "enable"
            diverter.Enable
        Case "disable"
            diverter.Disable
        Case "activate"
            diverter.Activate
        Case "deactivate"
            diverter.Deactivate
    End Select
    If IsObject(args(1)) Then
        Set DiverterEventHandler = kwargs
    Else
        DiverterEventHandler = kwargs
    End If
End Function
Function CreateGlfDroptarget(name)
	Dim droptarget : Set droptarget = (new GlfDroptarget)(name)
	Set CreateGlfDroptarget = droptarget
End Function

Class GlfDroptarget

    Private m_name
	Private m_switch
    Private m_enable_keep_up_events
    Private m_disable_keep_up_events
	Private m_action_cb
	Private m_knockdown_events
	Private m_reset_events
    Private m_complete

    
    Private m_debug

	Public Property Let Switch(value)
		m_switch = value
		AddPinEventListener m_switch & "_active", m_name & "_switch_active", "DroptargetEventHandler", 1000, Array("switch_active", Me)
		AddPinEventListener m_switch & "_inactive", m_name & "_switch_inactive", "DroptargetEventHandler", 1000, Array("switch_inactive", Me)
	End Property
    Public Property Let EnableKeepUpEvents(value)
        Dim evt
        If IsArray(m_enable_keep_up_events) Then
            For Each evt in m_enable_keep_up_events
                RemovePinEventListener evt, m_name & "_enable_keepup"
            Next
        End If
        m_enable_keep_up_events = value
        For Each evt in m_enable_keep_up_events
            AddPinEventListener evt, m_name & "_enable_keepup", "DroptargetEventHandler", 1000, Array("enable_keepup", Me)
        Next
    End Property
    Public Property Let DisableKeepUpEvents(value)
        Dim evt
        If IsArray(m_disable_keep_up_events) Then
            For Each evt in m_disable_keep_up_events
                RemovePinEventListener evt, m_name & "_disable_keepup"
            Next
        End If
        m_disable_keep_up_events = value
        For Each evt in m_disable_keep_up_events
            AddPinEventListener evt, m_name & "_disable_keepup", "DroptargetEventHandler", 1000, Array("disable_keepup", Me)
        Next
    End Property

    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
	Public Property Let KnockdownEvents(value)
		Dim evt
		If IsArray(m_knockdown_events) Then
			For Each evt in m_knockdown_events
				RemovePinEventListener evt, m_name & "_knockdown"
			Next
		End If
		m_knockdown_events = value
		For Each evt in m_knockdown_events
			AddPinEventListener evt, m_name & "_knockdown", "DroptargetEventHandler", 1000, Array("knockdown", Me)
		Next
	End Property
	Public Property Let ResetEvents(value)
		Dim evt
		If IsArray(m_reset_events) Then
			For Each evt in m_reset_events
				RemovePinEventListener evt, m_name & "_reset"
			Next
		End If
		m_reset_events = value
		For Each evt in m_reset_events
			AddPinEventListener evt, m_name & "_reset", "DroptargetEventHandler", 1000, Array("reset", Me)
		Next
	End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "drop_target_" & name
		m_switch = Empty
        EnableKeepUpEvents = Array()
        DisableKeepUpEvents = Array()
		m_action_cb = Empty
		KnockdownEvents = Array()
		ResetEvents = Array()
        m_complete = 0
		m_debug = False
        glf_droptargets.Add name, Me
        Set Init = Me
	End Function

    Public Sub UpdateStateFromSwitch(is_complete)

		Log "Drop target " & m_name & " switch " & m_switch & " has active value " & is_complete & " compared to drop complete " & m_complete

		If is_complete <> m_complete Then
			If is_complete = 1 Then
				Down()
			Else
				Up()
			End	If
		End If
		'UpdateBanks()
    End Sub

    Public Sub Up()
        m_complete = 0
        DispatchPinEvent m_name & "_up", Null
    End Sub

	Public Sub Down()
        m_complete = 1
        DispatchPinEvent m_name & "_down", Null
    End Sub

	Public Sub EnableKeepup()
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(3)
		End If
    End Sub

	Public Sub DisableKeepup()
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(4)
		End If
    End Sub

	Public Sub Knockdown()
        If Not IsEmpty(m_action_cb) And m_complete = 0 Then
            GetRef(m_action_cb)(1)
		End If
    End Sub

	Public Sub Reset()
        If Not IsEmpty(m_action_cb) And m_complete = 1 Then
            GetRef(m_action_cb)(0)
		End If
    End Sub

    Public Sub BallSearch(phase)
        Log "Ball Search, phase " & phase
        If Not IsEmpty(m_action_cb) And m_complete = 0 Then
            GetRef(m_action_cb)(1) 'Knockdown
            SetDelay m_name & "ball_search_deactivate", "DroptargetEventHandler", Array(Array("reset", Me), Null), 100
		Else
            If Not IsEmpty(m_action_cb) And m_complete = 1 Then
                GetRef(m_action_cb)(0) 'Reset
                SetDelay m_name & "ball_search_deactivate", "DroptargetEventHandler", Array(Array("knockdown", Me), Null), 100
            End If
        End If
    End Sub
    
    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function DroptargetEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim droptarget : Set droptarget = ownProps(1)
    Select Case evt
        Case "switch_active"
            droptarget.UpdateStateFromSwitch 1
        Case "switch_inactive"
            droptarget.UpdateStateFromSwitch 0
        Case "enable_keepup"
            droptarget.EnableKeepup
        Case "disable_keepup"
            droptarget.DisableKeepup
        Case "knockdown"
            droptarget.Knockdown
        Case "reset"
            droptarget.Reset
    End Select
    If IsObject(args(1)) Then
        Set DroptargetEventHandler = kwargs
    Else
        DroptargetEventHandler = kwargs
    End If
    
End Function

Function CreateGlfFlipper(name)
	Dim flipper : Set flipper = (new GlfFlipper)(name)
	Set CreateGlfFlipper = flipper
End Function

Class GlfFlipper

    Private m_name
    Private m_enable_events
    Private m_disable_events
    Private m_enabled
    Private m_switch
    Private m_action_cb
    Private m_debug

    Public Property Let Switch(value)
        m_switch = value
    End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "FlipperEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "FlipperEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "flipper_" & name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_enabled = False
        m_action_cb = Empty
        m_switch = Empty
        m_debug = False
        glf_flippers.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        m_enabled = True
        Dim evt
        If Not IsEmpty(m_switch) Then
            AddPinEventListener m_switch & "_active", m_name & "_active", "FlipperEventHandler", 1000, Array("activate", Me)
            AddPinEventListener m_switch & "_inactive", m_name & "_inactive", "FlipperEventHandler", 1000, Array("deactivate", Me)
        End If
    End Sub

    Public Sub Disable()
        Log "Disabling"
        m_enabled = False
        Deactivate()
        Dim evt
        RemovePinEventListener m_switch & "_active", m_name & "_active"
        RemovePinEventListener m_switch & "_inactive", m_name & "_inactive"
    End Sub

    Public Sub Activate()
        Log "Activating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(1)
        End If
        DispatchPinEvent m_name & "_activate", Null
    End Sub

    Public Sub Deactivate()
        Log "Deactivating"
        If Not IsEmpty(m_action_cb) Then
            GetRef(m_action_cb)(0)
        End If
        DispatchPinEvent m_name & "_deactivate", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function FlipperEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim flipper : Set flipper = ownProps(1)
    Select Case evt
        Case "enable"
            flipper.Enable
        Case "disable"
            flipper.Disable
        Case "activate"
            flipper.Activate
        Case "deactivate"
            flipper.Deactivate
    End Select
    If IsObject(args(1)) Then
        Set FlipperEventHandler = kwargs
    Else
        FlipperEventHandler = kwargs
    End If
End Function
Function CreateGlfLightSegmentDisplay(name)
	Dim segment_display : Set segment_display = (new GlfLightSegmentDisplay)(name)
	Set CreateGlfLightSegmentDisplay = segment_display
End Function

Class GlfLightSegmentDisplay
    private m_name
    private m_flash_on
    private m_flashing
    private m_flash_mask

    private m_text
    private m_current_text
    private m_display_state
    private m_current_state
    private m_current_flashing
    Private m_current_flash_mask
    private m_lights
    private m_light_group
    private m_light_groups
    private m_segmentmap
    private m_segment_type
    private m_size
    private m_update_method
    private m_text_stack
    private m_current_text_stack_entry
    private m_integrated_commas
    private m_integrated_dots
    private m_use_dots_for_commas
    private m_display_flash_duty
    private m_display_flash_display_flash_frequency
    private m_default_transition_update_hz
    private m_color
    private m_flex_dmd_index

    Public Property Get Name() : Name = m_name : End Property

    Public Property Get SegmentType() : SegmentType = m_segment_type : End Property
    Public Property Let SegmentType(input)
        m_segment_type = input
        If m_segment_type = "14Segment" Then
            Set m_segmentmap = FOURTEEN_SEGMENTS
        ElseIf m_segment_type = "7Segment" Then
            Set m_segmentmap = SEVEN_SEGMENTS
        End If
        CalculateLights()
    End Property

    Public Property Get LightGroup() : LightGroup = m_light_group : End Property
    Public Property Let LightGroup(input)
        m_light_group = input
        CalculateLights()
    End Property

    Public Property Get LightGroups() : LightGroups = m_light_groups : End Property
    Public Property Let LightGroups(input)
        m_light_groups = input
        CalculateLights()
    End Property

    Public Property Get UpdateMethod() : UpdateMethod = m_update_method : End Property
    Public Property Let UpdateMethod(input) : m_update_method = input : End Property

    Public Property Get SegmentSize() : SegmentSize = m_size : End Property
    Public Property Let SegmentSize(input)
        m_size = input
        CalculateLights()
    End Property

    Public Property Get IntegratedCommas() : IntegratedCommas = m_integrated_commas : End Property
    Public Property Let IntegratedCommas(input) : m_integrated_commas = input : End Property

    Public Property Get IntegratedDots() : IntegratedDots = m_integrated_dots : End Property
    Public Property Let IntegratedDots(input) : m_integrated_dots = input : End Property

    Public Property Get UseDotsForCommas() : UseDotsForCommas = m_use_dots_for_commas : End Property
    Public Property Let UseDotsForCommas(input) : m_use_dots_for_commas = input : End Property

    Public Property Get DefaultColor() : DefaultColor = m_color : End Property
    Public Property Let DefaultColor(input) : m_color = input : End Property

    Public Property Let ExternalFlexDmdSegmentIndex(input)
        m_flex_dmd_index = input
    End Property

    Public Property Get DefaultTransitionUpdateHz() : DefaultTransitionUpdateHz = m_default_transition_update_hz : End Property
    Public Property Let DefaultTransitionUpdateHz(input) : m_default_transition_update_hz = input : End Property

    Public default Function init(name)
        m_name = name
        m_flash_on = True
        m_flashing = "no_flash"
        m_flash_mask = Empty
        m_text = Empty
        m_size = 0
        m_segment_type = Empty
        m_segmentmap = Null
        m_light_group = Empty
        m_light_groups = Array()
        m_current_text = Empty
        m_display_state = Empty
        m_current_state = Null
        m_current_flashing = Empty
        m_current_flash_mask = Empty
        m_current_text_stack_entry = Null
        Set m_text_stack = (new GlfTextStack)()
        m_update_method = "replace"
        m_lights = Array()  
        m_integrated_commas = False
        m_integrated_dots = False
        m_use_dots_for_commas = False
        m_flex_dmd_index = -1

        m_display_flash_duty = 30
        m_default_transition_update_hz = 30
        m_display_flash_display_flash_frequency = 60

        m_color = "ffffff"

        SetDelay m_name & "software_flash", "Glf_SegmentDisplaySoftwareFlashEventHandler", Array(True, Me), 100

        glf_segment_displays.Add name, Me
        Set Init = Me
    End Function

    Private Sub CalculateLights()
        If Not IsEmpty(m_segment_type) And m_size > 0 And (Not IsEmpty(m_light_group) Or Ubound(m_light_groups)>-1) Then
            m_lights = Array()
            If m_segment_type = "14Segment" Then
                ReDim m_lights((m_size * 15)-1)
            ElseIf m_segment_type = "7Segment" Then
                ReDim m_lights((m_size * 8)-1)
            End If

            Dim i, group_idx, current_light_group
            If Not IsEmpty(m_light_group) Then
                current_light_group = m_light_group
            ElseIf UBound(m_light_groups)>-1 Then
                current_light_group = m_light_groups(0)
                group_idx = 0
            End If
            Dim k : k = 0
            For i=0 to UBound(m_lights)
                'On Error Resume Next
                If typename(Eval(current_light_group & CStr(k+1))) = "Light" Then
                    m_lights(i) = current_light_group & CStr(k+1)
                    k=k+1
                Else
                    'msgbox typename(Eval(current_light_group & CStr(k+1)))
                    'msgbox current_light_group & CStr(k+1)
                    current_light_group = m_light_groups(group_idx+1)
                    'msgbox current_light_group
                    group_idx = group_idx + 1
                    k = 0
                    m_lights(i) = current_light_group & CStr(k+1)
                    k=k+1
                End If

            Next
        End If
    End Sub

    Public Sub SetVirtualDMDLights(input)
        If m_flex_dmd_index>-1 Then
            Dim x
            For x=0 to UBound(m_lights)
                glf_lightNames(m_lights(x)).Visible = input
            Next
        End If
    End Sub

    Private Sub SetText(text, flashing, flash_mask)
        'Set a text to the display.
        Exit Sub


        'If flashing = "no_flash" Then
        '    m_flash_on = True
        'ElseIf flashing = "flash_mask" Then
            'm_flash_mask = flash_mask.rjust(len(text))
        'End If

        'If flashing = "no_flash" or m_flash_on = True or Not IsNull(text) Then
        '    If text <> m_display_state Then
        '        m_display_state = text
                'Set text to lights.
        '        If text="" Then
        '            text = Glf_FormatValue(text, " >" & CStr(m_size))
        '        Else
        '            text = Right(text, m_size)
        '        End If
        '        If text <> m_current_text Then
        '            m_current_text = text
        '            UpdateText()
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub UpdateDisplay(segment_text, flashing, flash_mask)
        Set m_current_state = segment_text
        m_flashing = flashing
        m_flash_mask = flash_mask
        'SetText m_current_state.ConvertToString(), flashing, flash_mask
        UpdateText()
    End Sub

    Private Sub UpdateText()
        'iterate lights and chars
        Dim mapped_text, segment
        If m_flash_on = True Or m_flashing = "no_flash" Then
            mapped_text = MapSegmentTextToSegments(m_current_state, m_size, m_segmentmap)
        Else
            If m_flashing = "mask" Then
                mapped_text = MapSegmentTextToSegments(m_current_state.BlankSegments(m_flash_mask), m_size, m_segmentmap)
            ElseIf m_flashing = "match" Then
                mapped_text = MapSegmentTextToSegments(m_current_state.BlankSegments(String(m_size, "F")), m_size, m_segmentmap)
            Else
                mapped_text = MapSegmentTextToSegments(m_current_state.BlankSegments(String(m_size, "F")), m_size, m_segmentmap)
            End If
        End If
        Dim segment_idx, i : segment_idx = 0 : i = 0
        For Each segment in mapped_text
            
            If m_segment_type = "14Segment" Then
                Glf_SetLight m_lights(segment_idx), SegmentColor(segment.a)
                Glf_SetLight m_lights(segment_idx + 1), SegmentColor(segment.b)
                Glf_SetLight m_lights(segment_idx + 2), SegmentColor(segment.c)
                Glf_SetLight m_lights(segment_idx + 3), SegmentColor(segment.d)
                Glf_SetLight m_lights(segment_idx + 4), SegmentColor(segment.e)
                Glf_SetLight m_lights(segment_idx + 5), SegmentColor(segment.f)
                Glf_SetLight m_lights(segment_idx + 6), SegmentColor(segment.g1)
                Glf_SetLight m_lights(segment_idx + 7), SegmentColor(segment.g2)
                Glf_SetLight m_lights(segment_idx + 8), SegmentColor(segment.h)
                Glf_SetLight m_lights(segment_idx + 9), SegmentColor(segment.j)
                Glf_SetLight m_lights(segment_idx + 10), SegmentColor(segment.k)
                Glf_SetLight m_lights(segment_idx + 11), SegmentColor(segment.n)
                Glf_SetLight m_lights(segment_idx + 12), SegmentColor(segment.m)
                Glf_SetLight m_lights(segment_idx + 13), SegmentColor(segment.l)
                Glf_SetLight m_lights(segment_idx + 14), SegmentColor(segment.dp)
                If m_flex_dmd_index > -1 Then
                    'debug.print segment.CharMapping
                    dim hex
                    hex = segment.CharMapping
                    'debug.print typename(hex)
                    On Error Resume Next
				    glf_flex_alphadmd_segments(m_flex_dmd_index+i) = hex
				    If Err Then Debug.Print "Error: " & Err
                    'glf_flex_alphadmd_segments(m_flex_dmd_index+i) = segment.CharMapping '&h2A0F '0010101000001111
                    glf_flex_alphadmd.Segments = glf_flex_alphadmd_segments
                End If
                segment_idx = segment_idx + 15
            ElseIf m_segment_type = "7Segment" Then
                Glf_SetLight m_lights(segment_idx), SegmentColor(segment.a)
                Glf_SetLight m_lights(segment_idx + 1), SegmentColor(segment.b)
                Glf_SetLight m_lights(segment_idx + 2), SegmentColor(segment.c)
                Glf_SetLight m_lights(segment_idx + 3), SegmentColor(segment.d)
                Glf_SetLight m_lights(segment_idx + 4), SegmentColor(segment.e)
                Glf_SetLight m_lights(segment_idx + 5), SegmentColor(segment.f)
                Glf_SetLight m_lights(segment_idx + 6), SegmentColor(segment.g)
                Glf_SetLight m_lights(segment_idx + 7), SegmentColor(segment.dp)
                segment_idx = segment_idx + 8
            End If
            i = i + 1
        Next
    End Sub

    Private Function SegmentColor(value)
        If value = 1 Then
            SegmentColor = m_color
        Else
            SegmentColor = "000000"
        End If
    End Function

    Public Sub AddTextEntry(text, color, flashing, flash_mask, transition, transition_out, priority, key)
    
        If m_update_method = "stack" Then
            m_text_stack.Push text,color,flashing,flash_mask,transition,transition_out,priority,key
            UpdateStack()
        Else
            Dim new_text : new_text = Glf_SegmentTextCreateCharacters(text.Value(), m_size, m_integrated_commas, m_integrated_dots, m_use_dots_for_commas, Array())
            Dim display_text : Set display_text = (new GlfSegmentDisplayText)(new_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, flashing, flash_mask
        End If
    End Sub

    Public Sub UpdateTransition(transition_runner)
        Dim display_text
        display_text = transition_runner.NextStep()
        If IsNull(display_text) Then
            UpdateStack() 
        Else
            Set display_text = (new GlfSegmentDisplayText)(display_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, m_flashing, m_flash_mask
            SetDelay m_name & "_update_transition", "Glf_SegmentDisplayUpdateTransition", Array(Me, transition_runner), 1000/m_default_transition_update_hz
        End If
    End Sub

    Public Sub UpdateStack()

        Dim top_text_stack_entry, top_is_current
        top_is_current = False
        If m_text_stack.IsEmpty() Then
            Dim empty_text : Set empty_text = (new GlfInput)("""" & String(m_size, " ") & """")
            Set top_text_stack_entry = (new GlfTextStackEntry)(empty_text,Null,"no_flash","",Null,Null,999999,"")
        Else
            Set top_text_stack_entry = m_text_stack.Peek()
        End If

        Dim previous_text_stack_entry : previous_text_stack_entry = Null
        If Not IsNull(m_current_text_stack_entry) Then
            Set previous_text_stack_entry = m_current_text_stack_entry
            If previous_text_stack_entry.text.IsPlayerState() Then
                RemovePlayerStateEventListener previous_text_stack_entry.text.PlayerStateValue(), m_name
            ElseIf previous_text_stack_entry.text.IsDeviceState() Then
                RemovePinEventListener top_text_stack_entry.text.DeviceStateEvent() , m_name
            End If

            If m_current_text_stack_entry.Key = top_text_stack_entry.Key Then
                top_is_current = True
            End If
        End If
        
        Set m_current_text_stack_entry = top_text_stack_entry

        'determine if the new key is different than the previous key (out transitions are only applied when changing keys)
        Dim transition_config : transition_config = Null
        If Not IsNull(previous_text_stack_entry) Then
            If top_text_stack_entry.key <> previous_text_stack_entry.key And Not IsNull(previous_text_stack_entry.transition_out) Then
                Set transition_config = previous_text_stack_entry.transition_out
            End If
        End If
        'determine if new text entry has a transition, if so, apply it (overrides any outgoing transition)
        If Not IsNull(top_text_stack_entry.transition) Then
            Set transition_config = top_text_stack_entry.transition
        End If
        'start transition (if configured)
        Dim flashing, flash_mask, display_text
        If Not IsNull(transition_config) And Not top_is_current Then
            'msgbox "starting transition"
            Dim transition_runner
            Select Case transition_config.TransitionType()
                case "push":
                    Set transition_runner = (new GlfPushTransition)(m_size, True, True, True)
                    transition_runner.Direction = transition_config.Direction()
                    transition_runner.TransitionText = transition_config.Text()
                case "cover":
                    Set transition_runner = (new GlfCoverTransition)(m_size, True, True, True)
                    transition_runner.Direction = transition_config.Direction()
                    transition_runner.TransitionText = transition_config.Text()
            End Select

            Dim previous_text
            If Not IsNull(previous_text_stack_entry) Then
                previous_text = previous_text_stack_entry.text.Value()
            Else
                previous_text = String(m_size, " ")
            End If

            If Not IsEmpty(top_text_stack_entry.flashing) Then
                flashing = top_text_stack_entry.flashing
                flash_mask = top_text_stack_entry.flash_mask
            Else
                flashing = m_current_state.flashing
                flash_mask = m_current_state.flash_mask
            End If
            display_text = transition_runner.StartTransition(previous_text, top_text_stack_entry.text.Value(), Array(), Array())
            Set display_text = (new GlfSegmentDisplayText)(display_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, flashing, flash_mask
            SetDelay m_name & "_update_transition", "Glf_SegmentDisplayUpdateTransition", Array(Me, transition_runner), 1000/m_default_transition_update_hz
        Else
            'no transition - subscribe to text template changes and update display
            If top_text_stack_entry.text.IsPlayerState() Then
                AddPlayerStateEventListener top_text_stack_entry.text.PlayerStateValue(), m_name, top_text_stack_entry.text.PlayerStatePlayer(), "Glf_SegmentTextStackEventHandler", top_text_stack_entry.priority, Me
            ElseIf top_text_stack_entry.text.IsDeviceState() Then
                AddPinEventListener top_text_stack_entry.text.DeviceStateEvent() , m_name, "Glf_SegmentTextStackEventHandler", top_text_stack_entry.priority, Me
            End If

            'set any flashing state specified in the entry
            If Not IsEmpty(top_text_stack_entry.flashing) Then
                flashing = top_text_stack_entry.flashing
                flash_mask = top_text_stack_entry.flash_mask
            Else
                flashing = m_current_state.flashing
                flash_mask = m_current_state.flash_mask
            End If

            'update the display
            Dim text_value : text_value = top_text_stack_entry.text.Value()

            If text_value = False Then
                text_value = String(m_size, " ")
            End If
            Dim new_text : new_text = Glf_SegmentTextCreateCharacters(text_value, m_size, m_integrated_commas, m_integrated_dots, m_use_dots_for_commas, Array())
            Set display_text = (new GlfSegmentDisplayText)(new_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
            UpdateDisplay display_text, flashing, flash_mask
        End If
    End Sub

    Public Sub CurrentPlaceholderChanged()
        Dim text_value : text_value = m_current_text_stack_entry.text.Value()
        'msgbox text_value
        If text_value = False Then
            text_value = String(m_size, " ")
        End If
        Dim new_text : new_text = Glf_SegmentTextCreateCharacters(text_value, m_size, m_integrated_commas, m_integrated_dots, m_use_dots_for_commas, Array())
        Dim display_text : Set display_text = (new GlfSegmentDisplayText)(new_text,m_integrated_commas, m_integrated_dots, m_use_dots_for_commas) 
        UpdateDisplay display_text, m_current_text_stack_entry.flashing, m_current_text_stack_entry.flash_mask
    End Sub

    Public Sub RemoveTextByKey(key)
        m_text_stack.PopByKey key
        UpdateStack()
    End Sub

    Public Sub RemoveTextByKeyNoUpdate(key)
        m_text_stack.PopByKey key
    End Sub

    Public Sub SetFlashing(flash_type)

    End Sub

    Public Sub SetFlashingMask(mask)

    End Sub

    Public Sub SetColor(color)

    End Sub

    Public Sub SetSoftwareFlash(enabled)
        m_flash_on = enabled

        If m_flashing = "no_flash" Then
            Exit Sub
        End If

        If IsNull(m_current_state) Then
            Exit Sub
        End If
        UpdateText        
    End Sub

End Class

Sub Glf_SegmentDisplaySoftwareFlashEventHandler(args)
    Dim display, enabled
    Set display = args(1)
    enabled = args(0)
    If enabled = True Then
        SetDelay display.Name & "software_flash", "Glf_SegmentDisplaySoftwareFlashEventHandler", Array(False, display), 100
        display.SetSoftwareFlash True
    Else
        SetDelay display.Name & "software_flash", "Glf_SegmentDisplaySoftwareFlashEventHandler", Array(True, display), 100
        display.SetSoftwareFlash False
    End If

    
End Sub

Sub Glf_SegmentTextStackEventHandler(args)
    Dim segment
    Set segment = args(0) 
    'kwargs = args(1) 
    'Dim player_var : player_var = kwargs(0)
    'Dim value : value = kwargs(1)
    'Dim prevValue : prevValue = kwargs(2)
    segment.CurrentPlaceholderChanged()
End Sub

Sub Glf_SegmentDisplayUpdateTransition(args)
    Dim display, runner
    Set display = args(0) 
    Set runner = args(1)
    display.UpdateTransition runner
End Sub

Class GlfTextStackEntry
    Public text, colors, flashing, flash_mask, transition, transition_out, priority, key

    Public default Function init(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
        Set Me.text = text
        Me.colors = colors
        Me.flashing = flashing
        Me.flash_mask = flash_mask
        If Not IsNull(transition) Then
            Set Me.transition = transition
        Else
            Me.transition = Null
        End If
        If Not IsNull(transition_out) Then
            Set Me.transition_out = transition_out
        Else
            Me.transition_out = Null
        End If
        Me.priority = priority
        Me.key = key
        Set Init = Me
    End Function
End Class

Class GlfTextStack
    Private stack

    ' Initialize an empty stack
    Public default Function Init()
        ReDim stack(-1)  ' Initialize an empty array
        Set Init = Me
    End Function

    ' Push a new text entry onto the stack or update an existing one
    Public Sub Push(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
        Dim found : found = False
        Dim i

        ' Check if the key already exists in the stack and update it
        For i = LBound(stack) To UBound(stack)
            If stack(i).key = key Then
                ' Replace the existing item if the key matches
                Set stack(i) = CreateTextStackEntry(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
                found = True
                Exit For
            End If
        Next
        
        If Not found Then
            ' Insert the new item into the array maintaining priority order
            ReDim Preserve stack(UBound(stack) + 1)
            Set stack(UBound(stack)) = CreateTextStackEntry(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
            SortStackByPriority
        End If
    End Sub

    ' Pop a specific entry from the stack by key
    Public Function PopByKey(key)
        Dim i, removedItem, found
        found = False
        Set removedItem = Nothing
    
        ' Loop through the stack to find the item with the matching key
        For i = LBound(stack) To UBound(stack)
            If stack(i).key = key Then
                ' Store the item to be removed
                Set removedItem = stack(i)
                found = True
    
                ' Shift all elements after the removed item to the left
                Dim j
                For j = i To UBound(stack) - 1
                    Set stack(j) = stack(j + 1)
                Next
    
                ' Resize the array to remove the last element
                ReDim Preserve stack(UBound(stack) - 1)
                Exit For
            End If
        Next
    
        ' Return the removed item (or Nothing if not found)
        If found Then
            Set PopByKey = removedItem
        Else
            Set PopByKey = Nothing
        End If
    End Function

    ' Peek at the top entry of the stack without popping it
    Public Function Peek()
        If UBound(stack) >= 0 Then
            Set Peek = stack(LBound(stack))
        Else
            Set Peek = Nothing
        End If
    End Function

    ' Check if the stack is empty
    Public Function IsEmpty()
        IsEmpty = (UBound(stack) < 0)
    End Function

    ' Create a new GlfTextStackEntry object
    Private Function CreateTextStackEntry(text, colors, flashing, flash_mask, transition, transition_out, priority, key)
        Dim entry
        Set entry = New GlfTextStackEntry
        entry.init text, colors, flashing, flash_mask, transition, transition_out, priority, key
        Set CreateTextStackEntry = entry
    End Function

    ' Sort the stack by priority (descending)
    Private Sub SortStackByPriority()
        Dim i, j
        Dim temp
        For i = LBound(stack) To UBound(stack) - 1
            For j = i + 1 To UBound(stack)
                If stack(i).priority < stack(j).priority Then
                    ' Swap the elements
                    Set temp = stack(i)
                    Set stack(i) = stack(j)
                    Set stack(j) = temp
                End If
            Next
        Next
    End Sub
End Class


Class FourteenSegments
    Public dp, l, m, n, k, j, h, g2, g1, f, e, d, c, b, a, char, hexcode, hexcode_dp

    Public Function CloneMapping()
        Set CloneMapping = (new FourteenSegments)(dp,l,m,n,k,j,h,g2,g1,f,e,d,c,b,a,char)
    End Function

    Public Property Get CharMapping()
        If dp = 1 Then
            CharMapping = hexcode_dp
        Else
            CharMapping = hexcode
        End If
    End Property

    Public default Function init(dp, l, m, n, k, j, h, g2, g1, f, e, d, c, b, a, char)
        Me.dp = dp
        Me.a = a
        Me.b = b
        Me.c = c
        Me.d = d
        Me.e = e
        Me.f = f
        Me.g1 = g1
        Me.g2 = g2
        Me.h = h
        Me.j = j
        Me.k = k
        Me.n = n
        Me.m = m
        Me.l = l
        Me.char = char

        Dim binaryString, decimalValue, i
        binaryString = CStr("0" & n & m & l & g2 & k & j & h & dp & g1 & f & e & d & c & b & a)
        If binaryString = "0000000000000000" Then
            hexcode = 0
        Else
            decimalValue = 0
            For i = 1 To Len(binaryString)
                decimalValue = decimalValue * 2 + Mid(binaryString, i, 1)
            Next
            hexcode = CInt("&H" & Right("0000" & UCase(Hex(decimalValue)), 4))
        End If
        binaryString = CStr("0" & n & m & l & g2 & k & j & h & 1 & g1 & f & e & d & c & b & a)        
        decimalValue = 0
        For i = 1 To Len(binaryString)
            decimalValue = decimalValue * 2 + Mid(binaryString, i, 1)
        Next
        hexcode_dp = CInt("&H" & Right("0000" & UCase(Hex(decimalValue)), 4))
        Set Init = Me
    End Function
End Class



Class SevenSegments
    Public dp, g, f, e, d, c, b, a, char

    Public Function CloneMapping()
        Set CloneMapping = (new SevenSegments)(dp,g,f,e,d,c,b,a,char)
    End Function

    Public default Function init(dp, g, f, e, d, c, b, a, char)
        Me.dp = dp
        Me.a = a
        Me.b = b
        Me.c = c
        Me.d = d
        Me.e = e
        Me.f = f
        Me.g = g
        Me.char = char
        Set Init = Me
    End Function
End Class

Dim FOURTEEN_SEGMENTS
Set FOURTEEN_SEGMENTS = CreateObject("Scripting.Dictionary")

FOURTEEN_SEGMENTS.Add Null, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "?")
FOURTEEN_SEGMENTS.Add 32, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, " ")
FOURTEEN_SEGMENTS.Add 33, (New FourteenSegments)(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, "!")
FOURTEEN_SEGMENTS.Add 34, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, Chr(34)) ' Character "
FOURTEEN_SEGMENTS.Add 35, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 0, 1, 1, 1, 0, "#")
FOURTEEN_SEGMENTS.Add 36, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 1, 1, 0, 1, 1, 0, 1, "$")
FOURTEEN_SEGMENTS.Add 37, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 1, 1, 1, 0, 0, 1, 0, 0, "%")
FOURTEEN_SEGMENTS.Add 38, (New FourteenSegments)(0, 1, 0, 0, 0, 1, 1, 0, 1, 0, 1, 1, 0, 0, 1, "&")
FOURTEEN_SEGMENTS.Add 39, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "'")
FOURTEEN_SEGMENTS.Add 40, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "(")
FOURTEEN_SEGMENTS.Add 41, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, ")")
FOURTEEN_SEGMENTS.Add 42, (New FourteenSegments)(0, 1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, "*")
FOURTEEN_SEGMENTS.Add 43, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, "+")
FOURTEEN_SEGMENTS.Add 44, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ",")
FOURTEEN_SEGMENTS.Add 45, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, "-")
FOURTEEN_SEGMENTS.Add 46, (New FourteenSegments)(1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ".")
FOURTEEN_SEGMENTS.Add 47, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "/")
FOURTEEN_SEGMENTS.Add 48, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "0")
FOURTEEN_SEGMENTS.Add 49, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, "1")
FOURTEEN_SEGMENTS.Add 50, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 1, 1, "2")
FOURTEEN_SEGMENTS.Add 51, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 1, "3")
FOURTEEN_SEGMENTS.Add 52, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, 1, 0, "4")
FOURTEEN_SEGMENTS.Add 53, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 0, 1, "5")
FOURTEEN_SEGMENTS.Add 54, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 1, "6")
FOURTEEN_SEGMENTS.Add 55, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, "7")
FOURTEEN_SEGMENTS.Add 56, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 1, 1, "8")
FOURTEEN_SEGMENTS.Add 57, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1, 1, "9")
FOURTEEN_SEGMENTS.Add 58, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, ":")
FOURTEEN_SEGMENTS.Add 59, (New FourteenSegments)(0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, ";")
FOURTEEN_SEGMENTS.Add 60, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, "<")
FOURTEEN_SEGMENTS.Add 61, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 1, 0, 0, 0, "=")
FOURTEEN_SEGMENTS.Add 62, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, ">")
FOURTEEN_SEGMENTS.Add 63, (New FourteenSegments)(1, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, "?")
FOURTEEN_SEGMENTS.Add 64, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 1, "@")
FOURTEEN_SEGMENTS.Add 65, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 1, "A")
FOURTEEN_SEGMENTS.Add 66, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 1, 0, 0, 0, 1, 1, 1, 1, "B")
FOURTEEN_SEGMENTS.Add 67, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, "C")
FOURTEEN_SEGMENTS.Add 68, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, "D")
FOURTEEN_SEGMENTS.Add 69, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, "E")
FOURTEEN_SEGMENTS.Add 70, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 1, "F")
FOURTEEN_SEGMENTS.Add 71, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 1, 1, 1, 0, 1, "G")
FOURTEEN_SEGMENTS.Add 72, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 1, 0, "H")
FOURTEEN_SEGMENTS.Add 73, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 1, "I")
FOURTEEN_SEGMENTS.Add 74, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, "J")
FOURTEEN_SEGMENTS.Add 75, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, "K")
FOURTEEN_SEGMENTS.Add 76, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, "L")
FOURTEEN_SEGMENTS.Add 77, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, "M")
FOURTEEN_SEGMENTS.Add 78, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 1, 0, 1, 1, 0, "N")
FOURTEEN_SEGMENTS.Add 79, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "O")
FOURTEEN_SEGMENTS.Add 80, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, 1, "P")
FOURTEEN_SEGMENTS.Add 81, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, "Q")
FOURTEEN_SEGMENTS.Add 82, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 1, 1, "R")
FOURTEEN_SEGMENTS.Add 83, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 0, 1, "S")
FOURTEEN_SEGMENTS.Add 84, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, "T")
FOURTEEN_SEGMENTS.Add 85, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 0, "U")
FOURTEEN_SEGMENTS.Add 86, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, "V")
FOURTEEN_SEGMENTS.Add 87, (New FourteenSegments)(0, 1, 0, 1, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, "W")
FOURTEEN_SEGMENTS.Add 88, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "X")
FOURTEEN_SEGMENTS.Add 89, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 1, 1, 1, 0, "Y")
FOURTEEN_SEGMENTS.Add 90, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, "Z")
FOURTEEN_SEGMENTS.Add 91, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, "[")
FOURTEEN_SEGMENTS.Add 92, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, Chr(92)) ' Character \
FOURTEEN_SEGMENTS.Add 93, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, "]")
FOURTEEN_SEGMENTS.Add 94, (New FourteenSegments)(0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "^")
FOURTEEN_SEGMENTS.Add 95, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 1, "_")
FOURTEEN_SEGMENTS.Add 96, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "`")
FOURTEEN_SEGMENTS.Add 97, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 1, 0, 0, 0, "a")
FOURTEEN_SEGMENTS.Add 98, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, 0, "b")
FOURTEEN_SEGMENTS.Add 99, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 1, 0, 0, 0, "c")
FOURTEEN_SEGMENTS.Add 100, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, "d")
FOURTEEN_SEGMENTS.Add 101, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 1, 1, 0, 0, 0, "e")
FOURTEEN_SEGMENTS.Add 102, (New FourteenSegments)(0, 0, 1, 0, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, "f")
FOURTEEN_SEGMENTS.Add 103, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 1, 1, 1, 1, "g")
FOURTEEN_SEGMENTS.Add 104, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 1, 0, 0, "h")
FOURTEEN_SEGMENTS.Add 105, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, "i")
FOURTEEN_SEGMENTS.Add 106, (New FourteenSegments)(0, 0, 0, 1, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, "j")
FOURTEEN_SEGMENTS.Add 107, (New FourteenSegments)(0, 1, 1, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "k")
FOURTEEN_SEGMENTS.Add 108, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "l")
FOURTEEN_SEGMENTS.Add 109, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 0, 0, "m")
FOURTEEN_SEGMENTS.Add 110, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 1, 0, 0, "n")
FOURTEEN_SEGMENTS.Add 111, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 1, 1, 0, 0, "o")
FOURTEEN_SEGMENTS.Add 112, (New FourteenSegments)(0, 0, 0, 0, 1, 0, 0, 0, 1, 1, 1, 0, 0, 1, 1, "p")
FOURTEEN_SEGMENTS.Add 113, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0, 1, 1, 1, 1, "q")
FOURTEEN_SEGMENTS.Add 114, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 1, 0, 0, 0, 0, "r")
FOURTEEN_SEGMENTS.Add 115, (New FourteenSegments)(0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, "s")
FOURTEEN_SEGMENTS.Add 116, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 0, 0, "t")
FOURTEEN_SEGMENTS.Add 117, (New FourteenSegments)(0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 0, 0, "u")
FOURTEEN_SEGMENTS.Add 118, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, "v")
FOURTEEN_SEGMENTS.Add 119, (New FourteenSegments)(0, 0, 1, 0, 0, 0, 0, 0, 1, 1, 1, 0, 0, 1, 0, "w")
FOURTEEN_SEGMENTS.Add 120, (New FourteenSegments)(0, 1, 0, 1, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, "x")
FOURTEEN_SEGMENTS.Add 121, (New FourteenSegments)(0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 1, 1, 1, 1, 0, "y")
FOURTEEN_SEGMENTS.Add 122, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 1, 0, 0, 1, "z")
FOURTEEN_SEGMENTS.Add 123, (New FourteenSegments)(0, 0, 0, 1, 0, 0, 1, 0, 1, 0, 0, 1, 0, 0, 1, "{")
FOURTEEN_SEGMENTS.Add 124, (New FourteenSegments)(0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, "|")
FOURTEEN_SEGMENTS.Add 125, (New FourteenSegments)(0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 1, 1, 0, 0, 1, "}")
FOURTEEN_SEGMENTS.Add 126, (New FourteenSegments)(0, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, "~")


Dim SEVEN_SEGMENTS
Set SEVEN_SEGMENTS = CreateObject("Scripting.Dictionary")

SEVEN_SEGMENTS.Add Null, (New SevenSegments)(0, 0, 0, 0, 0, 0, 0, 0, "?")
SEVEN_SEGMENTS.Add 32, (New SevenSegments)(0, 0, 0, 0, 0, 0, 0, 0, " ")
SEVEN_SEGMENTS.Add 33, (New SevenSegments)(1, 0, 0, 0, 0, 1, 1, 0, "!")
SEVEN_SEGMENTS.Add 34, (New SevenSegments)(0, 0, 1, 0, 0, 0, 1, 0, Chr(34)) ' Character "
SEVEN_SEGMENTS.Add 35, (New SevenSegments)(0, 1, 1, 1, 1, 1, 1, 0, "#")
SEVEN_SEGMENTS.Add 36, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "$")
SEVEN_SEGMENTS.Add 37, (New SevenSegments)(1, 1, 0, 1, 0, 0, 1, 0, "%")
SEVEN_SEGMENTS.Add 38, (New SevenSegments)(0, 1, 0, 0, 0, 1, 1, 0, "&")
SEVEN_SEGMENTS.Add 39, (New SevenSegments)(0, 0, 1, 0, 0, 0, 0, 0, "'")
SEVEN_SEGMENTS.Add 40, (New SevenSegments)(0, 0, 1, 0, 1, 0, 0, 1, "(")
SEVEN_SEGMENTS.Add 41, (New SevenSegments)(0, 0, 0, 0, 1, 0, 1, 1, ")")
SEVEN_SEGMENTS.Add 42, (New SevenSegments)(0, 0, 1, 0, 0, 0, 0, 1, "*")
SEVEN_SEGMENTS.Add 43, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 0, "+")
SEVEN_SEGMENTS.Add 44, (New SevenSegments)(0, 0, 0, 1, 0, 0, 0, 0, ",")
SEVEN_SEGMENTS.Add 45, (New SevenSegments)(0, 1, 0, 0, 0, 0, 0, 0, "-")
SEVEN_SEGMENTS.Add 46, (New SevenSegments)(1, 0, 0, 0, 0, 0, 0, 0, ".")
SEVEN_SEGMENTS.Add 47, (New SevenSegments)(0, 1, 0, 1, 0, 0, 1, 0, "/")
SEVEN_SEGMENTS.Add 48, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 1, "0")
SEVEN_SEGMENTS.Add 49, (New SevenSegments)(0, 0, 0, 0, 0, 1, 1, 0, "1")
SEVEN_SEGMENTS.Add 50, (New SevenSegments)(0, 1, 0, 1, 1, 0, 1, 1, "2")
SEVEN_SEGMENTS.Add 51, (New SevenSegments)(0, 1, 0, 0, 1, 1, 1, 1, "3")
SEVEN_SEGMENTS.Add 52, (New SevenSegments)(0, 1, 1, 0, 0, 1, 1, 0, "4")
SEVEN_SEGMENTS.Add 53, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "5")
SEVEN_SEGMENTS.Add 54, (New SevenSegments)(0, 1, 1, 1, 1, 1, 0, 1, "6")
SEVEN_SEGMENTS.Add 55, (New SevenSegments)(0, 0, 0, 0, 0, 1, 1, 1, "7")
SEVEN_SEGMENTS.Add 56, (New SevenSegments)(0, 1, 1, 1, 1, 1, 1, 1, "8")
SEVEN_SEGMENTS.Add 57, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 1, "9")
SEVEN_SEGMENTS.Add 58, (New SevenSegments)(0, 0, 0, 0, 1, 0, 0, 1, ":")
SEVEN_SEGMENTS.Add 59, (New SevenSegments)(0, 0, 0, 0, 1, 1, 0, 1, ";")
SEVEN_SEGMENTS.Add 60, (New SevenSegments)(0, 1, 1, 0, 0, 0, 0, 1, "<")
SEVEN_SEGMENTS.Add 61, (New SevenSegments)(0, 1, 0, 0, 1, 0, 0, 0, "=")
SEVEN_SEGMENTS.Add 62, (New SevenSegments)(0, 1, 0, 0, 0, 0, 1, 1, ">")
SEVEN_SEGMENTS.Add 63, (New SevenSegments)(1, 1, 0, 1, 0, 0, 1, 1, "?")
SEVEN_SEGMENTS.Add 64, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 1, "@")
SEVEN_SEGMENTS.Add 65, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 1, "A")
SEVEN_SEGMENTS.Add 66, (New SevenSegments)(0, 1, 1, 1, 1, 1, 0, 0, "B")
SEVEN_SEGMENTS.Add 67, (New SevenSegments)(0, 0, 1, 1, 1, 0, 0, 1, "C")
SEVEN_SEGMENTS.Add 68, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 0, "D")
SEVEN_SEGMENTS.Add 69, (New SevenSegments)(0, 1, 1, 1, 1, 0, 0, 1, "E")
SEVEN_SEGMENTS.Add 70, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 1, "F")
SEVEN_SEGMENTS.Add 71, (New SevenSegments)(0, 0, 1, 1, 1, 1, 0, 1, "G")
SEVEN_SEGMENTS.Add 72, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 0, "H")
SEVEN_SEGMENTS.Add 73, (New SevenSegments)(0, 0, 1, 1, 0, 0, 0, 0, "I")
SEVEN_SEGMENTS.Add 74, (New SevenSegments)(0, 0, 0, 1, 1, 1, 1, 0, "J")
SEVEN_SEGMENTS.Add 75, (New SevenSegments)(0, 1, 1, 1, 0, 1, 0, 1, "K")
SEVEN_SEGMENTS.Add 76, (New SevenSegments)(0, 0, 1, 1, 1, 0, 0, 0, "L")
SEVEN_SEGMENTS.Add 77, (New SevenSegments)(0, 0, 0, 1, 0, 1, 0, 1, "M")
SEVEN_SEGMENTS.Add 78, (New SevenSegments)(0, 0, 1, 1, 0, 1, 1, 1, "N")
SEVEN_SEGMENTS.Add 79, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 1, "O")
SEVEN_SEGMENTS.Add 80, (New SevenSegments)(0, 1, 1, 1, 0, 0, 1, 1, "P")
SEVEN_SEGMENTS.Add 81, (New SevenSegments)(0, 1, 1, 0, 1, 0, 1, 1, "Q")
SEVEN_SEGMENTS.Add 82, (New SevenSegments)(0, 0, 1, 1, 0, 0, 1, 1, "R")
SEVEN_SEGMENTS.Add 83, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "S")
SEVEN_SEGMENTS.Add 84, (New SevenSegments)(0, 1, 1, 1, 1, 0, 0, 0, "T")
SEVEN_SEGMENTS.Add 85, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 0, "U")
SEVEN_SEGMENTS.Add 86, (New SevenSegments)(0, 0, 1, 1, 1, 1, 1, 0, "V")
SEVEN_SEGMENTS.Add 87, (New SevenSegments)(0, 0, 1, 0, 1, 0, 1, 0, "W")
SEVEN_SEGMENTS.Add 88, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 0, "X")
SEVEN_SEGMENTS.Add 89, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 0, "Y")
SEVEN_SEGMENTS.Add 90, (New SevenSegments)(0, 1, 0, 1, 1, 0, 1, 1, "Z")
SEVEN_SEGMENTS.Add 91, (New SevenSegments)(0, 0, 1, 1, 1, 0, 0, 1, "[")
SEVEN_SEGMENTS.Add 92, (New SevenSegments)(0, 1, 1, 0, 0, 1, 0, 0, Chr(92)) ' Character \
SEVEN_SEGMENTS.Add 93, (New SevenSegments)(0, 0, 0, 0, 1, 1, 1, 1, "]")
SEVEN_SEGMENTS.Add 94, (New SevenSegments)(0, 0, 1, 0, 0, 0, 1, 1, "^")
SEVEN_SEGMENTS.Add 95, (New SevenSegments)(0, 0, 0, 0, 1, 0, 0, 0, "_")
SEVEN_SEGMENTS.Add 96, (New SevenSegments)(0, 0, 0, 0, 0, 0, 1, 0, "`")
SEVEN_SEGMENTS.Add 97, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 1, "a")
SEVEN_SEGMENTS.Add 98, (New SevenSegments)(0, 1, 1, 1, 1, 1, 0, 0, "b")
SEVEN_SEGMENTS.Add 99, (New SevenSegments)(0, 1, 0, 1, 1, 0, 0, 0, "c")
SEVEN_SEGMENTS.Add 100, (New SevenSegments)(0, 1, 0, 1, 1, 1, 1, 0, "d")
SEVEN_SEGMENTS.Add 101, (New SevenSegments)(0, 1, 1, 1, 1, 0, 1, 1, "e")
SEVEN_SEGMENTS.Add 102, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 1, "f")
SEVEN_SEGMENTS.Add 103, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 1, "g")
SEVEN_SEGMENTS.Add 104, (New SevenSegments)(0, 1, 1, 1, 0, 1, 0, 0, "h")
SEVEN_SEGMENTS.Add 105, (New SevenSegments)(0, 0, 0, 1, 0, 0, 0, 0, "i")
SEVEN_SEGMENTS.Add 106, (New SevenSegments)(0, 0, 0, 0, 1, 1, 0, 0, "j")
SEVEN_SEGMENTS.Add 107, (New SevenSegments)(0, 1, 1, 1, 0, 1, 0, 1, "k")
SEVEN_SEGMENTS.Add 108, (New SevenSegments)(0, 0, 1, 1, 0, 0, 0, 0, "l")
SEVEN_SEGMENTS.Add 109, (New SevenSegments)(0, 0, 0, 1, 0, 1, 0, 0, "m")
SEVEN_SEGMENTS.Add 110, (New SevenSegments)(0, 1, 0, 1, 0, 1, 0, 0, "n")
SEVEN_SEGMENTS.Add 111, (New SevenSegments)(0, 1, 0, 1, 1, 1, 0, 0, "o")
SEVEN_SEGMENTS.Add 112, (New SevenSegments)(0, 1, 1, 1, 0, 0, 1, 1, "p")
SEVEN_SEGMENTS.Add 113, (New SevenSegments)(0, 1, 1, 0, 0, 1, 1, 1, "q")
SEVEN_SEGMENTS.Add 114, (New SevenSegments)(0, 1, 0, 1, 0, 0, 0, 0, "r")
SEVEN_SEGMENTS.Add 115, (New SevenSegments)(0, 1, 1, 0, 1, 1, 0, 1, "s")
SEVEN_SEGMENTS.Add 116, (New SevenSegments)(0, 1, 1, 1, 1, 0, 0, 0, "t")
SEVEN_SEGMENTS.Add 117, (New SevenSegments)(0, 0, 0, 1, 1, 1, 0, 0, "u")
SEVEN_SEGMENTS.Add 118, (New SevenSegments)(0, 0, 0, 1, 1, 1, 0, 0, "v")
SEVEN_SEGMENTS.Add 119, (New SevenSegments)(0, 0, 0, 1, 0, 1, 0, 0, "w")
SEVEN_SEGMENTS.Add 120, (New SevenSegments)(0, 1, 1, 1, 0, 1, 1, 0, "x")
SEVEN_SEGMENTS.Add 121, (New SevenSegments)(0, 1, 1, 0, 1, 1, 1, 0, "y")
SEVEN_SEGMENTS.Add 122, (New SevenSegments)(0, 1, 0, 1, 1, 0, 1, 1, "z")
SEVEN_SEGMENTS.Add 123, (New SevenSegments)(0, 1, 0, 0, 0, 1, 1, 0, "{")
SEVEN_SEGMENTS.Add 124, (New SevenSegments)(0, 0, 1, 1, 0, 0, 0, 0, "|")
SEVEN_SEGMENTS.Add 125, (New SevenSegments)(0, 1, 1, 1, 0, 0, 0, 0, "}")
SEVEN_SEGMENTS.Add 126, (New SevenSegments)(0, 0, 0, 0, 0, 0, 0, 1, "~")


Function MapSegmentTextToSegments(text_state, display_width, segment_mapping)
    'Map a segment display text to a certain display mapping.

    Dim text : text = text_state.Text
    Dim segments()
    ReDim segments(UBound(text))

    Dim charCode, char, mapping, i, new_mapping
    For i = 0 To UBound(text)
        Set char = text(i)
        If segment_mapping.Exists(char("char_code")) Then
            Set mapping = segment_mapping(char("char_code"))
            Set new_mapping = mapping.CloneMapping()
            If char("dot") = True Then
                new_mapping.dp = 1
            Else
                new_mapping.dp = 0
            End If
        Else
            Set new_mapping = segment_mapping(Null)
        End If

        Set segments(i) = new_mapping
    Next

    MapSegmentTextToSegments = segments
End Function


Class GlfSegmentDisplayText

    Private m_embed_dots
    Private m_embed_commas
    Private m_use_dots_for_commas
    Private m_text

    Public Property Get Text() : Text = m_text : End Property

    ' Initialize the class
    Public default Function Init(char_list, embed_dots, embed_commas, use_dots_for_commas)
        m_embed_dots = embed_dots
        m_embed_commas = embed_commas
        m_use_dots_for_commas = use_dots_for_commas
        m_text = char_list
        Set Init = Me
    End Function

    ' Get the length of the text
    Public Function Length()
        Length = UBound(m_text) + 1
    End Function

    ' Get a character or a slice of the text
    Public Function GetItem(index)
        If IsArray(index) Then
            Dim slice, i
            slice = Array()
            For i = LBound(index) To UBound(index)
                slice = AppendArray(slice, m_text(index(i)))
            Next
            GetItem = slice
        Else
            GetItem = m_text(index)
        End If
    End Function

    ' Extend the text with another list
    Public Sub Extend(other_text)
        Dim i
        For i = LBound(other_text) To UBound(other_text)
            m_text = AppendArray(m_text, other_text(i))
        Next
    End Sub

    ' Convert the text to a string
    Public Function ConvertToString()
        Dim text, char, i
        text = ""
        For i = LBound(m_text) To UBound(m_text)
            Set char = m_text(i)
            text = text & Chr(char("char_code"))
            If char("dot") Then text = text & "."
            If char("comma") Then text = text & ","
        Next
        ConvertToString = text
    End Function

    ' Get colors (to be implemented in subclasses)
    Public Function GetColors()
        GetColors = Null
    End Function

    Public Function BlankSegments(flash_mask)
        Dim arrFlashMask, i
        ReDim arrFlashMask(Len(flash_mask) - 1)
        For i = 1 To Len(flash_mask)
            arrFlashMask(i - 1) = Mid(flash_mask, i, 1)
        Next


        Dim new_text, char, mask
        new_text = Array()
    
        ' Iterate over characters and the flash mask
        For i = LBound(m_text) To UBound(m_text)
            Set char = m_text(i)
            mask = arrFlashMask(i)
    
            ' If mask is "F", blank the character
            If mask = "F" Then
                new_text = AppendArray(new_text, Glf_SegmentTextCreateDisplayCharacter(32, False, False, char("color")))
            Else
                ' Otherwise, keep the character as is
                new_text = AppendArray(new_text, char)
            End If
        Next
    
        ' Create a new GlfSegmentDisplayText object with the updated text
        Dim blanked_text
        Set blanked_text = (new GlfSegmentDisplayText)(new_text, m_embed_commas, m_embed_dots, m_use_dots_for_commas)
        Set BlankSegments = blanked_text
    End Function
    

End Class

' Helper function to append to an array
Function AppendArray(arr, value)
    Dim newArr, i
    ReDim newArr(UBound(arr) + 1)
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            Set newArr(i) = arr(i)
        Else
            newArr(i) = arr(i)
        End If
    Next
    If IsObject(value) Then
        Set newArr(UBound(newArr)) = value
    Else
        newArr(UBound(newArr)) = value
    End If
    AppendArray = newArr
End Function

' Helper function to slice an array
Function SliceArray(arr, start_idx, end_idx)
    Dim sliced, i, j
    ReDim sliced(end_idx - start_idx)
    j = 0
    For i = start_idx To end_idx
        If IsObject(arr(i)) Then
            Set sliced(j) = arr(i)
        Else
            sliced(j) = arr(i)
        End If
        j = j + 1
    Next
    SliceArray = sliced
End Function

' Helper function to prepend an element to an array
Function PrependArray(arr, value)
    Dim newArr, i
    ReDim newArr(UBound(arr) + 1)
    If IsObject(value) Then
        Set newArr(0) = value
    Else
        newArr(0) = value
    End If
    
    For i = LBound(arr) To UBound(arr)
        If IsObject(arr(i)) Then
            Set newArr(i + 1) = arr(i)
        Else
            newArr(i + 1) = arr(i)
        End If
    Next
    PrependArray = newArr
End Function


Function Glf_SegmentTextCreateCharacters(text, display_size, collapse_dots, collapse_commas, use_dots_for_commas, colors)
            


    Dim char_list, uncolored_chars, left_pad_color, default_right_color, i, char, color, current_length
    char_list = Array()

    ' Determine padding and default colors
    If IsArray(colors) And UBound(colors) >= 0 Then

        left_pad_color = colors(0)
        default_right_color = colors(UBound(colors))

    Else
        left_pad_color = Null
        default_right_color = Null
    End If

    ' Embed dots and commas
    uncolored_chars = Glf_SegmentTextEmbedDotsAndCommas(text, collapse_dots, collapse_commas, use_dots_for_commas)
 
    ' Adjust colors to match the uncolored characters
    If IsArray(colors) And UBound(colors) >= 0 Then
        Dim adjusted_colors
        adjusted_colors = SliceArray(colors, UBound(colors) - UBound(uncolored_chars) + 1, UBound(colors))
    Else
        adjusted_colors = Array()
    End If

    ' Create display characters
    For i = LBound(uncolored_chars) To UBound(uncolored_chars)
        char = uncolored_chars(i)
        
        If IsArray(adjusted_colors) And UBound(adjusted_colors) >= 0 Then
            color = adjusted_colors(0)
            adjusted_colors = SliceArray(adjusted_colors, 1, UBound(adjusted_colors))
        Else
            color = default_right_color
        End If
        char_list = AppendArray(char_list, Glf_SegmentTextCreateDisplayCharacter(char(0), char(1), char(2), color))
    Next

    ' Adjust the list size to match the display size
    current_length = UBound(char_list) + 1
    
    If current_length > display_size Then
        ' Truncate characters from the left
        char_list = SliceArray(char_list, current_length - display_size, UBound(char_list))
    ElseIf current_length < display_size Then
        ' Pad with spaces to the left
        Dim padding
        padding = display_size - current_length
        For i = 1 To padding
            char_list = PrependArray(char_list, Glf_SegmentTextCreateDisplayCharacter(32, False, False, left_pad_color))
        Next
    End If
    'msgbox ">"&text&"<"
    'msgbox UBound(char_list)
    Glf_SegmentTextCreateCharacters = char_list
End Function

Function Glf_SegmentTextEmbedDotsAndCommas(text, collapse_dots, collapse_commas, use_dots_for_commas)
    Dim char_has_dot, char_has_comma, char_list
    Dim i, char_code

    char_has_dot = False
    char_has_comma = False
    char_list = Array()

    ' Iterate through the text in reverse
    For i = Len(text) To 1 Step -1
        char_code = Asc(Mid(text, i, 1))
        
        ' Check for dots and commas and handle collapsing rules
        If (collapse_dots And char_code = Asc(".")) Or (use_dots_for_commas And char_code = Asc(",")) Then
            char_has_dot = True
        ElseIf collapse_commas And char_code = Asc(",") Then
            char_has_comma = True
        Else
            ' Insert the character at the start of the list
            char_list = PrependArray(char_list, Array(char_code, char_has_dot, char_has_comma))
            char_has_dot = False
            char_has_comma = False
        End If
    Next

    Glf_SegmentTextEmbedDotsAndCommas = char_list
End Function

' Helper function to create a display character
Function Glf_SegmentTextCreateDisplayCharacter(char_code, has_dot, has_comma, color)
    Dim display_character
    Set display_character = CreateObject("Scripting.Dictionary")
    display_character.Add "char_code", char_code
    display_character.Add "dot", has_dot
    display_character.Add "comma", has_comma
    display_character.Add "color", color
    Set Glf_SegmentTextCreateDisplayCharacter = display_character
End Function

Function CreateGlfMagnet(name)
	Dim magnet : Set magnet = (new GlfMagnet)(name)
	Set CreateGlfMagnet = magnet
End Function

Class GlfMagnet

    Private m_name
    Private m_enable_events
    Private m_disable_events
    Private m_fling_ball_events
    Private m_fling_drop_time
    Private m_fling_regrab_time
    Private m_grab_ball_events
    Private m_grab_switch
    Private m_grab_time
    Private m_release_ball_events
    Private m_release_time
    Private m_reset_events
    Private m_action_cb

    Private m_active
    Private m_release_in_progress

    Private m_debug

    Public Property Let EnableEvents(value)
        Dim evt
        If IsArray(m_enable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_enable"
            Next
        End If
        m_enable_events = value
        For Each evt in m_enable_events
            AddPinEventListener evt, m_name & "_enable", "MagnetEventHandler", 1000, Array("enable", Me)
        Next
    End Property
    Public Property Let DisableEvents(value)
        Dim evt
        If IsArray(m_disable_events) Then
            For Each evt in m_enable_events
                RemovePinEventListener evt, m_name & "_disable"
            Next
        End If
        m_disable_events = value
        For Each evt in m_disable_events
            AddPinEventListener evt, m_name & "_disable", "MagnetEventHandler", 1000, Array("disable", Me)
        Next
    End Property
    Public Property Let ActionCallback(value) : m_action_cb = value : End Property
    Public Property Let FlingBallEvents(value) : m_fling_ball_events = value : End Property
    Public Property Let FlingDropTime(value) : Set m_fling_drop_time = CreateGlfInput(value) : End Property
    Public Property Let FlingRegrabTime(value) : Set m_fling_regrab_time = CreateGlfInput(value) : End Property
    Public Property Let GrabBallEvents(value) : m_grab_ball_events = value : End Property
    Public Property Let GrabSwitch(value)
        m_grab_switch = value
    End Property
    Public Property Let GrabTime(value) : Set m_grab_time = CreateGlfInput(value) : End Property
    Public Property Let ReleaseBallEvents(value) : m_release_ball_events = value : End Property
    Public Property Let ReleaseTime(value) : Set m_release_time = CreateGlfInput(value) : End Property
    Public Property Let ResetEvents(value) : m_reset_events = value : End Property
    Public Property Let Debug(value) : m_debug = value : End Property

	Public default Function init(name)
        m_name = "magnet_" & name
        EnableEvents = Array("ball_started")
        DisableEvents = Array("ball_will_end", "service_mode_entered")
        m_action_cb = Empty
        m_fling_ball_events = Array()
        Set m_fling_drop_time = CreateGlfInput(250)
        Set m_fling_regrab_time = CreateGlfInput(50)
        m_grab_ball_events = Array()
        m_grab_switch = Empty
        Set m_grab_time = CreateGlfInput(1500)
        m_release_ball_events = Array()
        Set m_release_time = CreateGlfInput(500)
        m_reset_events = Array("machine_reset_phase_3", "ball_starting")
        m_active = False
        m_release_in_progress = False
        m_debug = False
        glf_magnets.Add name, Me
        Set Init = Me
	End Function

    Public Sub Enable()
        Log "Enabling"
        Dim evt
        For Each evt in m_fling_ball_events
            AddPinEventListener evt, m_name & "_fling", "MagnetEventHandler", 1000, Array("fling", Me)
        Next
        For Each evt in m_grab_ball_events
            AddPinEventListener evt, m_name & "_grab", "MagnetEventHandler", 1000, Array("grab", Me)
        Next
        For Each evt in m_release_ball_events
            AddPinEventListener evt, m_name & "_release", "MagnetEventHandler", 1000, Array("release", Me)
        Next
        AddPinEventListener m_grab_switch & "_active", m_name & "_grab_switch", "MagnetEventHandler", 1000, Array("grab", Me)
    End Sub

    Public Sub Disable()
        Log "Disabling"
        Dim evt
        For Each evt in m_fling_ball_events
            RemovePinEventListener evt, m_name & "_fling"
        Next
        For Each evt in m_grab_ball_events
            RemovePinEventListener evt, m_name & "_grab"
        Next
        For Each evt in m_release_ball_events
            RemovePinEventListener evt, m_name & "_release"
        Next
        RemovePinEventListener m_grab_switch & "_active", m_name & "_grab_switch"
    End Sub

    Public Sub AddBall(ball)
        m_magnet.AddBall ball
    End Sub

    Public Sub RemoveBall(ball)
        m_magnet.RemoveBall ball
    End Sub

    Public Sub Fling()
        If m_active = False or m_release_in_progress = True Then
            Exit Sub
        End If
        m_active = False
        m_release_in_progress = True
        Log "Flinging Ball"
        DispatchPinEvent m_name & "flinging_ball", Null
        GetRef(m_action_cb)(0)
        SetDelay m_name & "_fling_reenable", "MagnetEventHandler" , Array(Array("fling_reenable", Me),Null), m_fling_drop_time.Value
    End Sub

    Public Sub FlingReenable()
        GetRef(m_action_cb)(1)
        SetDelay m_name & "_fling_regrab", "MagnetEventHandler" , Array(Array("fling_regrab", Me),Null), m_fling_regrab_time.Value
    End Sub

    Public Sub FlingRegrab()
        m_release_in_progress = False
        GetRef(m_action_cb)(0)
        DispatchPinEvent m_name & "_flinged_ball", Null
    End Sub

    Public Sub Grab()
        If m_active = True Or m_release_in_progress = True Then
            Exit Sub
        End If
        Log "Grabbing Ball"
        m_active = True
        GetRef(m_action_cb)(1)
        DispatchPinEvent m_name & "_grabbing_ball", Null
        SetDelay m_name & "_grabbing_done", "MagnetEventHandler" , Array(Array("grabbing_done", Me),Null), m_grab_time.Value
    End Sub

    Public Sub GrabbingDone()
        DispatchPinEvent m_name & "_grabbed_ball", Null
    End Sub

    Public Sub Release()
        If m_active = False or m_release_in_progress = True Then
            Exit Sub
        End If
        m_active = False
        m_release_in_progress = True
        Log "Releasing Ball"
        DispatchPinEvent m_name & "releasing_ball", Null
        GetRef(m_action_cb)(0)
        SetDelay m_name & "_release_done", "MagnetEventHandler" , Array(Array("release_done", Me),Null), m_release_time.Value
    End Sub

    Public Sub ReleaseDone()
        m_release_in_progress = False
        DispatchPinEvent m_name & "released_ball", Null
    End Sub

    Private Sub Log(message)
        If m_debug = True Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub
End Class

Function MagnetEventHandler(args)
    Dim ownProps, kwargs : ownProps = args(0)
    If IsObject(args(1)) Then
        Set kwargs = args(1) 
    Else
        kwargs = args(1) 
    End If
    Dim evt : evt = ownProps(0)
    Dim magnet : Set magnet = ownProps(1)
    Select Case evt
        Case "enable"
            magnet.Enable
        Case "disable"
            magnet.Disable
        Case "fling"
            magnet.Fling
        Case "grab"
            magnet.Grab
        Case "release"
            magnet.Release
        Case "grabbing_done"
            magnet.GrabbingDone
        Case "release_done"
            magnet.ReleaseDone
        Case "fling_reenable"
            magnet.FlingReenable
        Case "fling_regrab"
            magnet.FlingRegrab
    End Select
    If IsObject(args(1)) Then
        Set MagnetEventHandler = kwargs
    Else
        MagnetEventHandler = kwargs
    End If
    
End Function
Class GlfCoverTransition

    Private m_output_length
    Private m_collapse_dots
    Private m_collapse_commas
    Private m_use_dots_for_commas

    Private m_current_step
    Private m_total_steps
    Private m_direction
    Private m_transition_text
    Private m_transition_color

    Private m_current_text
    Private m_new_text
    Private m_current_colors
    Private m_new_colors

    ' Initialize the transition class
    Public Default Function Init(output_length, collapse_dots, collapse_commas, use_dots_for_commas)
        m_output_length = output_length
        m_collapse_dots = collapse_dots
        m_collapse_commas = collapse_commas
        m_use_dots_for_commas = use_dots_for_commas

        m_current_step = 0
        m_total_steps = 0
        m_direction = "right"
        m_transition_text = "" ' Default empty transition text
        m_transition_color = Empty

        Set Init = Me
    End Function

    ' Set transition direction
    Public Property Let Direction(value)
        If value = "left" Or value = "right" Then
            m_direction = value
        Else
            m_direction = "right" ' Default to right
        End If
    End Property

    ' Set transition text
    Public Property Let TransitionText(value)
        m_transition_text = value
    End Property

    ' Start transition
    Public Function StartTransition(current_text, new_text, current_colors, new_colors)
        ' Store text and colors
        m_current_text = current_text
        m_new_text = Space(m_output_length - Len(new_text)) & new_text
        m_current_colors = current_colors
        m_new_colors = new_colors
        ' Calculate total steps for transition
        m_total_steps = m_output_length + Len(m_transition_text)
        'If m_total_steps > 0 Then m_total_steps = m_total_steps
        ' Reset step counter
        m_current_step = 0
        StartTransition = NextStep()
    End Function

    ' Manually call this to progress the transition
    Public Function NextStep()
        If m_current_step >= m_total_steps Then
            NextStep = Null
            Exit Function
        End If
        ' Get the correct transition text for this step
        Dim result
        result = GetTransitionStep(m_current_step)
        ' Increment step counter
        m_current_step = m_current_step + 1

        ' Return the current frame's text
        NextStep = result
    End Function

    ' Get the text for the current step
    Private Function GetTransitionStep(current_step)
        Dim transition_sequence, start_idx, end_idx

        If m_direction = "right" Then
            ' Right cover transition: new_text + transition_text moves in
            transition_sequence = m_new_text & m_transition_text & m_current_text
            start_idx = Len(transition_sequence) - (current_step + m_output_length)
            end_idx = start_idx + m_output_length - 1
        ElseIf m_direction = "left" Then
            ' Left cover transition: transition_text + new_text moves in
            transition_sequence = m_current_text & m_transition_text & m_new_text
            start_idx = current_step
            end_idx = start_idx + m_output_length - 1
        End If

        ' Ensure valid slice indices
        If start_idx < 0 Then start_idx = 0
        If end_idx > Len(transition_sequence) - 1 Then end_idx = Len(transition_sequence) - 1

        ' Extract the correct frame of text
        Dim sliced_text
        sliced_text = Mid(transition_sequence, start_idx + 1, end_idx - start_idx + 1)
        If m_output_length>m_current_step Then
            If m_direction = "right" Then
                sliced_text = m_new_text & m_transition_text & Right(m_current_text, m_output_length-current_step)
            ElseIf m_direction = "left" Then
                sliced_text = Left(m_current_text, m_output_length - m_current_step) & Left(m_transition_text & m_new_text, current_step)
            End If
        End If
        'MsgBox "transition_text-"&transition_sequence&", current_step=" & current_step & ", start_idx=" & start_idx & ", end_idx=" & end_idx & ", text=>" & sliced_text &"<, text_len=>" & Len(sliced_text) &"<, Total Steps: "&m_total_steps
        
        ' Convert only the final sliced text to segment display characters
        GetTransitionStep = Glf_SegmentTextCreateCharacters(sliced_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, m_new_colors)
    End Function

End Class
Class GlfPushTransition

    Private m_output_length
    Private m_collapse_dots
    Private m_collapse_commas
    Private m_use_dots_for_commas

    Private m_current_step
    Private m_total_steps
    Private m_direction
    Private m_transition_text
    Private m_transition_color

    Private m_current_text
    Private m_new_text
    Private m_current_colors
    Private m_new_colors

    ' Initialize the transition class
    Public Default Function Init(output_length, collapse_dots, collapse_commas, use_dots_for_commas)
        m_output_length = output_length
        m_collapse_dots = collapse_dots
        m_collapse_commas = collapse_commas
        m_use_dots_for_commas = use_dots_for_commas

        m_current_step = 0
        m_total_steps = 0
        m_direction = "right"
        m_transition_text = "" ' Default empty transition text
        m_transition_color = Empty

        Set Init = Me
    End Function

    ' Set transition direction
    Public Property Let Direction(value)
        If value = "left" Or value = "right" Then
            m_direction = value
        Else
            m_direction = "right" ' Default to right
        End If
    End Property

    ' Set transition text
    Public Property Let TransitionText(value)
        m_transition_text = value
    End Property

    ' Start transition
    Public Function StartTransition(current_text, new_text, current_colors, new_colors)
        ' Store text and colors
        m_current_text = current_text
        m_new_text = Space(m_output_length - Len(new_text)) & new_text
        m_current_colors = current_colors
        m_new_colors = new_colors
        ' Calculate total steps for transition
        m_total_steps = m_output_length + Len(m_transition_text)
        If m_total_steps > 0 Then m_total_steps = m_total_steps + 1
        'm_total_steps=(m_output_length*2)+1
        ' Reset step counter
        m_current_step = 0
        StartTransition = NextStep()
    End Function

    ' Manually call this to progress the transition
    Public Function NextStep()
        If m_current_step >= m_total_steps Then
            NextStep = Null ' Transition complete
            Exit Function
        End If
        ' Get the correct transition text for this step
        Dim result
        result = GetTransitionStep(m_current_step)
        ' Increment step counter
        m_current_step = m_current_step + 1

        ' Return the current frame's text
        NextStep = result
    End Function

    ' Get the text for the current step
    Private Function GetTransitionStep(current_step)
        Dim transition_sequence, start_idx, end_idx
        'msgbox "Step"
        ' Construct the full transition sequence as plain text
        If m_direction = "right" Then
            ' Right push: [NEW_TEXT + TRANSITION_TEXT + OLD_TEXT] moves LEFT
            transition_sequence = m_new_text & m_transition_text & m_current_text
    
            ' Calculate slice indices
            start_idx = Len(transition_sequence) - (current_step + m_output_length)
            end_idx = start_idx + m_output_length - 1
    
        ElseIf m_direction = "left" Then
            ' Left push: [OLD_TEXT + TRANSITION_TEXT + NEW_TEXT] moves RIGHT
            transition_sequence = m_current_text & m_transition_text & m_new_text
    
            ' Calculate slice indices
            start_idx = current_step
            end_idx = start_idx + m_output_length - 1
        End If
    
        ' Ensure valid slice indices
        If start_idx < 0 Then start_idx = 0
        If end_idx > Len(transition_sequence) - 1 Then end_idx = Len(transition_sequence) - 1
    
        ' Extract the correct frame of text
        Dim sliced_text
        sliced_text = Mid(transition_sequence, start_idx + 1, end_idx - start_idx + 1)
    
        ' Debugging output
        'MsgBox "transition_text-"&transition_sequence&", current_step=" & current_step & ", start_idx=" & start_idx & ", end_idx=" & end_idx & ", text=>" & sliced_text &"<"
        ' Convert only the final sliced text to segment display characters
        GetTransitionStep = Glf_SegmentTextCreateCharacters(sliced_text, m_output_length, m_collapse_dots, m_collapse_commas, m_use_dots_for_commas, m_new_colors)
    End Function

End Class


Function CreateGlfSoundBus(name)
	Dim bus : Set bus = (new GlfSoundBus)(name)
	Set CreateGlfSoundBus = bus
End Function

Class GlfSoundBus

    Private m_name
    Private m_simultaneous_sounds
    Private m_current_sounds
    Private m_volume
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "simultaneous_sounds":
                GetValue = m_simultaneous_sounds
            Case "volume":
                GetValue = m_volume
        End Select
    End Property

    Public Property Get SimultaneousSounds(): SimultaneousSounds = m_simultaneous_sounds: End Property
    Public Property Let SimultaneousSounds(input): m_simultaneous_sounds = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Let Debug(value)
        m_debug = value
    End Property

    Public default Function Init(name)
        m_name = "sound_bus_" & name
        m_simultaneous_sounds = 8
        m_volume = 0.5
        Set m_current_sounds = CreateObject("Scripting.Dictionary")
        glf_sound_buses.Add name, Me
        Set Init = Me
    End Function

    Public Sub Play(sound_settings)
        If (UBound(m_current_sounds.Keys)-1) > m_simultaneous_sounds Then
            'TODO: Queue Sound
        Else
            If m_current_sounds.Exists(sound_settings.Sound.File) Then
                m_current_sounds.Remove sound_settings.Sound.File
                RemoveDelay m_name & "_stop_sound_" & sound_settings.Sound.File       
            End If
            m_current_sounds.Add sound_settings.Sound.File, sound_settings
            Dim volume : volume = m_volume
            If Not IsEmpty(sound_settings.Sound.Volume) Then
                volume = sound_settings.Sound.Volume
            End If
            If Not IsEmpty(sound_settings.Volume) Then
                volume = sound_settings.Volume
            End If
            Dim loops : loops = 0
            If Not IsEmpty(sound_settings.Sound.Loops) Then
                loops = sound_settings.Sound.Loops
            End If
            If Not IsEmpty(sound_settings.Loops) Then
                loops = sound_settings.Loops
            End If

            PlaySound sound_settings.Sound.File, loops, volume, 0,0,0,0,0,0
            If loops = 0 Then
                SetDelay m_name & "_stop_sound_" & sound_settings.Sound.File, "Glf_SoundBusStopSoundHandler", Array(sound_settings.Sound.File, Me), sound_settings.Sound.Duration
            ElseIf loops>0 Then
                SetDelay m_name & "_stop_sound_" & sound_settings.Sound.File, "Glf_SoundBusStopSoundHandler", Array(sound_settings.Sound.File, Me), sound_settings.Sound.Duration*loops
            End If
        End If
    End Sub

    Public Sub StopSoundWithKey(sound_key)
        If m_current_sounds.Exists(sound_key) Then
            Dim sound_settings : Set sound_settings = m_current_sounds(sound_key)
            StopSound(sound_key)
            Dim evt
            For Each evt in sound_settings.Sound.EventsWhenStopped.Items()
                If evt.Evaluate() Then
                    DispatchPinEvent evt.EventName, Null
                End If
            Next
            m_current_sounds.Remove sound_key
        End If
    End Sub

    Private Sub Log(message)
        If m_debug Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Function Glf_SoundBusStopSoundHandler(args)
    Dim sound_key : sound_key = args(0)
    Dim sound_bus : Set sound_bus = args(1)
    sound_bus.StopSoundWithKey sound_key
End Function

Function CreateGlfSound(name)
	Dim sound : Set sound = (new GlfSound)(name)
	Set CreateGlfSound = sound
End Function

Class GlfSound

    Private m_name
    Private m_file
    Private m_events_when_stopped
    Private m_bus
    Private m_volume
    Private m_priority
    Private m_max_queue_time
    Private m_duration
    Private m_loops
    Private m_debug

    Public Property Get Name(): Name = m_name: End Property
    Public Property Get GetValue(value)
        Select Case value
            Case "file":
                GetValue = m_file
            Case "volume":
                GetValue = m_volume
            Case "events_when_stopped":
                Set GetValue = m_events_when_stopped
            Case "bus":
                GetValue = m_bus
            Case "priority":
                GetValue = m_priority
            Case "max_queue_time":
                GetValue = m_max_queue_time
            Case "duration":
                GetValue = m_duration
        End Select
    End Property

    Public Property Get File(): File = m_file: End Property
    Public Property Let File(input): m_file = input: End Property

    Public Property Get Bus(): Bus = m_bus: End Property
    Public Property Let Bus(input): m_bus = input: End Property

    Public Property Get Volume(): Volume = m_volume: End Property
    Public Property Let Volume(input): m_volume = input: End Property

    Public Property Get Loops(): Loops = m_loops: End Property
    Public Property Let Loops(input): m_loops = input: End Property

    Public Property Get Priority(): Priority = m_priority: End Property
    Public Property Let Priority(input): m_priority = input: End Property

    Public Property Get MaxQueueTime(): MaxQueueTime = m_max_queue_time: End Property
    Public Property Let MaxQueueTime(input): m_max_queue_time = input: End Property

    Public Property Get Duration(): Duration = m_duration: End Property
    Public Property Let Duration(input): m_duration = input: End Property

    Public Property Get EventsWhenStopped(): Set EventsWhenStopped = m_events_when_stopped: End Property
    Public Property Let EventsWhenStopped(value)
        Dim x
        For x=0 to UBound(value)
            Dim newEvent : Set newEvent = (new GlfEvent)(value(x))
            m_events_when_stopped.Add x, newEvent
        Next
    End Property

    Public Property Let Debug(value)
        m_debug = value
    End Property

    Public default Function Init(name)
        m_name = "sound_bus_" & name
        m_file = Empty
        m_bus = Empty
        m_volume = Empty
        m_priority = 0
        m_duration = 0
        m_max_queue_time = -1 
        m_loops = 0
        Set m_events_when_stopped = CreateObject("Scripting.Dictionary")
        glf_sounds.Add name, Me
        Set Init = Me
    End Function

    Private Sub Log(message)
        If m_debug Then
            glf_debugLog.WriteToLog m_name, message
        End If
    End Sub

End Class

Class GlfEvent
	Private m_raw, m_name, m_event, m_condition, m_delay, m_priority, has_condition
  
    Public Property Get Name() : Name = m_name : End Property
    Public Property Get EventName() : EventName = m_event : End Property
    Public Property Get Condition() : Condition = m_condition : End Property
    Public Property Get Delay() : Delay = m_delay : End Property
    Public Property Get Raw() : Raw = m_raw : End Property
    Public Property Get Priority() : Priority = m_priority : End Property

    Public Function Evaluate()
        If has_condition = True Then
            Evaluate = GetRef(m_condition)()
        Else
            Evaluate = True
        End If
    End Function

	Public default Function init(evt)
        m_raw = evt
        Dim parsedEvent : parsedEvent = Glf_ParseEventInput(evt)
        m_name = parsedEvent(0)
        m_event = parsedEvent(1)
        m_condition = parsedEvent(2)
        If Not IsNull(m_condition) Then
            has_condition = True
        Else
            has_condition = False
        End If
        m_delay = parsedEvent(3)
        m_priority = parsedEvent(4)
	    Set Init = Me
	End Function

End Class

Class GlfRandomEvent
	
    Private m_parent_key
    Private m_key
    Private m_mode
    Private m_events
    Private m_weights
    Private m_eventIndexMap
    Private m_fallback_event
    Private m_force_all
    Private m_force_different
    Private m_disable_random
    Private m_total_weights

    Public Property Let FallbackEvent(value) : m_fallback_event = value : End Property
    Public Property Let ForceAll(value) : m_force_all = value : End Property
    Public Property Let ForceDifferent(value) : m_force_different = value : End Property
    Public Property Let DisableRandom(value) : m_disable_random = value : End Property

	Public default Function init(evt, mode, key)
        m_parent_key = evt
        m_key = key
        m_mode = mode
        m_fallback_event = Empty
        m_force_all = True
        m_force_different = True
        m_disable_random = False
        m_total_weights = 0
        Set m_events = CreateObject("Scripting.Dictionary")
        Set m_weights = CreateObject("Scripting.Dictionary")
        Set m_eventIndexMap = CreateObject("Scripting.Dictionary")
	    Set Init = Me
	End Function

    Public Sub Add(evt, weight)
        Dim newEvent : Set newEvent = (new GlfEvent)(evt)
        m_events.Add newEvent.Raw, newEvent
        m_weights.Add newEvent.Raw, weight
        m_total_weights = m_total_weights + weight
        m_eventIndexMap.Add newEvent.Raw, UBound(m_events.Keys)
    End Sub

    Public Function GetNextRandomEvent()

        Dim valid_events, event_to_fire
        Dim event_keys, event_items
        Dim i, count, key
        
        Set valid_events = CreateObject("Scripting.Dictionary")
        event_keys = m_events.Keys
        For i = 0 To UBound(event_keys)
            If m_events(event_keys(i)).Evaluate Then
                valid_events.Add event_keys(i), m_events(event_keys(i))
            End If
        Next
        event_to_fire = CheckFallback(valid_events)
        If Not IsEmpty(event_to_fire) Then
            GetNextRandomEvent = event_to_fire
            Exit Function
        End If

        If m_force_all = True Then
            event_keys = valid_events.Keys
            event_items = valid_events.Items
            valid_events.RemoveAll
            For i=0 to UBound(event_keys)
                If GetPlayerState("random_" & m_mode & "_" & m_key & "_" & m_eventIndexMap(event_keys(i))) = False Then
                    valid_events.Add event_keys(i), event_items(i)
                End If
            Next
        End If

        event_to_fire = CheckFallback(valid_events)
        If Not IsEmpty(event_to_fire) Then
            GetNextRandomEvent = event_to_fire
            Exit Function
        End If

        If m_force_different = True Then
            If valid_events.Exists(GetPlayerState("random_" & m_mode & "_" & m_key & "_last")) Then
                valid_events.Remove GetPlayerState("random_" & m_mode & "_" & m_key & "_last")
            End If
        End If

        event_to_fire = CheckFallback(valid_events)
        If Not IsEmpty(event_to_fire) Then
            GetNextRandomEvent = event_to_fire
            Exit Function
        End If

        If UBound(valid_events.Keys) = -1 Then
            GetNextRandomEvent = Empty
            Exit Function
        End If

        'Random Selection From remaining valid events
        Dim chosenKey
        If m_disable_random = False Then
            Dim total_weight
            For Each key In valid_events.Keys
                total_weight = total_weight + m_weights(key)
            Next

            Randomize
            'randomIdx = Int(Rnd() * (UBound(valid_events.Keys)-LBound(valid_events.Keys) + 1) + LBound(valid_events.Keys))
            Dim randVal
            randVal = Rnd() * total_weight
            Dim cumulativeWeight
            cumulativeWeight = 0
            
            For Each key In valid_events.Keys
                cumulativeWeight = cumulativeWeight + m_weights(key)
                If randVal <= cumulativeWeight Then
                    chosenKey = key
                    Exit For
                End If
            Next


        Else
            event_keys = m_events.Keys
            count = 0
            For i = 0 To UBound(event_keys)
                If GetPlayerState("random_" & m_mode & "_" & m_key & "_" & m_events(event_keys(i)).Raw) = True Then
                    If valid_events.Exists(m_events(event_keys(i)).Raw) Then
                        valid_events.Remove m_events(event_keys(i)).Raw
                    End If
                End If
            Next
            chosenKey = valid_events.keys()(0)
        End If
        
        SetPlayerState "random_" & m_mode & "_" & m_key & "_last", valid_events(chosenKey).Raw
        SetPlayerState "random_" & m_mode & "_" & m_key & "_" & valid_events(chosenKey).Raw, True

        event_keys = m_events.Keys
        count = 0
        For i = 0 To UBound(event_keys)
            If GetPlayerState("random_" & m_mode & "_" & m_key & "_" & m_events(event_keys(i)).Raw) = True Then
                count = count + 1
            End If
        Next
        If count = (UBound(event_keys) + 1) Then
            For i = 0 To UBound(event_keys)
                SetPlayerState "random_" & m_mode & "_" & m_key & "_" & m_events(event_keys(i)).Raw, False
            Next
        End If

        GetNextRandomEvent = valid_events(chosenKey).EventName

    End Function

    Public Function CheckFallback(valid_events)
        If UBound(valid_events.Keys()) = -1 Then
            If Not IsEmpty(m_fallback_event) Then
                CheckFallback = m_fallback_event
            Else
                CheckFallback = Empty
            End If
        Else
            CheckFallback = Empty
        End If
    End Function

End Class


'******************************************************
'*****  Player Setup                               ****
'******************************************************

Sub Glf_AddPlayer()
    Dim kwargs : Set kwargs = GlfKwargs()
    With kwargs
        .Add "num", -1
    End With
    Select Case UBound(glf_playerState.Keys())
        Case -1:
            kwargs("num") = 1
            DispatchPinEvent "player_added", kwargs
            glf_playerState.Add "PLAYER 1", Glf_InitNewPlayer()
            SetPlayerStateByPlayer GLF_SCORE, 0, 0
            SetPlayerStateByPlayer "number", 1, 0
            Glf_BcpAddPlayer 1
            glf_currentPlayer = "PLAYER 1"
        Case 0:     
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 2
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 2", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 1
                SetPlayerStateByPlayer "number", 2, 1
                Glf_BcpAddPlayer 2
            End If
        Case 1:
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 3
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 3", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 2
                SetPlayerStateByPlayer "number", 3, 2
                Glf_BcpAddPlayer 3
            End If     
        Case 2:   
            If GetPlayerState(GLF_CURRENT_BALL) = 1 Then
                kwargs("num") = 4
                DispatchPinEvent "player_added", kwargs
                glf_playerState.Add "PLAYER 4", Glf_InitNewPlayer()
                SetPlayerStateByPlayer GLF_SCORE, 0, 3
                SetPlayerStateByPlayer "number", 4, 3
                Glf_BcpAddPlayer 4
            End If  
            glf_canAddPlayers = False
    End Select
End Sub

Function Glf_InitNewPlayer()
    Dim state : Set state = CreateObject("Scripting.Dictionary")
    state.Add GLF_SCORE, 0
    Glf_MonitorPlayerStateUpdate GLF_SCORE, 0
    state.Add GLF_INITIALS, ""
    Glf_MonitorPlayerStateUpdate GLF_INITIALS, ""
    state.Add GLF_CURRENT_BALL, 1
    Glf_MonitorPlayerStateUpdate GLF_CURRENT_BALL, 1
    state.Add "extra_balls", 0
    Glf_MonitorPlayerStateUpdate "extra_balls", 0
    Dim i
    For i=0 To UBound(glf_initialVars.Keys())
        state.Add glf_initialVars.Keys()(i), glf_initialVars.Items()(i)
        Glf_MonitorPlayerStateUpdate glf_initialVars.Keys()(i), glf_initialVars.Items()(i)
    Next
    Set Glf_InitNewPlayer = state
End Function


'****************************
' Setup Player
' Event Listeners:  
    AddPinEventListener GLF_GAME_STARTED,   "start_game_setup",   "Glf_SetupPlayer", 1000, Null
    AddPinEventListener GLF_NEXT_PLAYER,    "next_player_setup",  "Glf_SetupPlayer", 1000, Null
'
'*****************************
Function Glf_SetupPlayer(args)
    EmitAllglf_playerEvents()
End Function

'****************************
' StartGame
'
'*****************************

AddPinEventListener "request_to_start_game", "request_to_start_game_ball_controller", "Glf_BallController", 30, Null
Function Glf_BallController(args)
    Dim balls_in_trough : balls_in_trough = 0
    If glf_troughSize = 1 Then
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 2 Then
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 3 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 4 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 5 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 6 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 7 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough7.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
    If glf_troughSize = 8 Then 
        If swTrough1.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough2.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough3.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough4.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough5.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough6.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If swTrough7.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
        If Drain.BallCntOver = 1 Then balls_in_trough = balls_in_trough + 1
    End If
  
    If glf_troughSize <> balls_in_trough Then
        Glf_BallController = False
        Exit Function
    End If

    Glf_BallController = True
End Function

AddPinEventListener "request_to_start_game", "request_to_start_game_result", "Glf_StartGame", 20, Null

Function Glf_StartGame(args)
    If args(1) = True And glf_gameStarted = False Then
        Glf_AddPlayer()
        glf_gameStarted = True
        DispatchPinEvent GLF_GAME_START, Null
        If useBcp Then
            bcpController.Send "player_turn_start?player_num=int:1"
            bcpController.Send "ball_start?player_num=int:1&ball=int:1"
            bcpController.SendPlayerVariable "number", 1, 0
        End If
        SetDelay GLF_GAME_STARTED, "Glf_DispatchGameStarted", Null, 50
    End If
End Function

Sub Glf_EndBall()

    glf_BIP = 0
    DispatchPinEvent GLF_BALL_WILL_END, Null
    DispatchQueuePinEvent GLF_BALL_ENDING, Null
    Dim device
    For Each device in glf_flippers.Items()
        device.Disable()
    Next
    For Each device in glf_autofiredevices.Items()
        device.Disable()
    Next

End Sub

Public Function Glf_DispatchGameStarted(args)
    DispatchPinEvent GLF_GAME_STARTED, Null
End Function


'******************************************************
'*****   Ball Release                              ****
'******************************************************

'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener GLF_GAME_STARTED, "start_game_release_ball",   "Glf_ReleaseBall", 20, True
AddPinEventListener GLF_NEXT_PLAYER, "next_player_release_ball",   "Glf_ReleaseBall", 20, True
'
'*****************************
Function Glf_ReleaseBall(args)
    Dim kwargs
    Set kwargs = GlfKwargs()
    If Not IsNull(args) Then
        If args(0) = True Then
            kwargs.Add "new_ball", True
        End If
    End If
    DispatchQueuePinEvent "balldevice_trough_ball_eject_attempt", kwargs
End Function


'****************************
' Release Ball
' Event Listeners:  
AddPinEventListener "balldevice_trough_ball_eject_attempt", "trough_eject",  "Glf_TroughReleaseBall", 20, Null
'
'*****************************
Function Glf_TroughReleaseBall(args)

    If Not IsNull(args) Then
        If IsObject(args(1)) Then
            If args(1)("new_ball") = True Then
                glf_BIP = glf_BIP + 1
                DispatchPinEvent GLF_BALL_STARTED, Null
                If useBcp Then
                    bcpController.SendPlayerVariable GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL), GetPlayerState(GLF_CURRENT_BALL)-1
                    bcpController.SendPlayerVariable GLF_SCORE, GetPlayerState(GLF_SCORE), GetPlayerState(GLF_SCORE)
                End If
            End If
        End If
    End If
    Glf_WriteDebugLog "Release Ball", "swTrough1: " & swTrough1.BallCntOver
    glf_plunger.AddIncomingBalls = 1
    swTrough1.kick 90, 10
    DispatchPinEvent "trough_eject", Null
    Glf_WriteDebugLog "Release Ball", "Just Kicked"
End Function

'****************************
' Ball Drain
' Event Listeners:      
    AddPinEventListener GLF_BALL_DRAIN, "ball_drain", "Glf_Drain", 20, Null
'
'*****************************
Function Glf_Drain(args)
    
    If Not glf_gameTilted Then
        Dim ballsToSave : ballsToSave = args(1) 
        Glf_WriteDebugLog "end_of_ball, unclaimed balls", CStr(ballsToSave)
        Glf_WriteDebugLog "end_of_ball, balls in play", CStr(glf_BIP)
        If ballsToSave <= 0 Then
            Exit Function
        End If

        glf_BIP = glf_BIP - 1
        glf_debugLog.WriteToLog "Trough", "Ball Drained: BIP: " & glf_BIP

        If glf_BIP > 0 Then
            Exit Function
        End If
        
        DispatchPinEvent GLF_BALL_WILL_END, Null
        DispatchQueuePinEvent GLF_BALL_ENDING, Null
    End If
    
End Function

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener GLF_BALL_ENDING, "ball_will_end", "Glf_BallWillEnd", 10, Null
'
'*****************************
Function Glf_BallWillEnd(args)
    DispatchPinEvent GLF_BALL_ENDED, Null
End Function

'****************************
' End Of Ball
' Event Listeners:      
AddPinEventListener GLF_BALL_ENDED, "end_of_ball", "Glf_EndOfBall", 20, Null
'
'*****************************
Function Glf_EndOfBall(args)

    If GetPlayerState("extra_balls") > 0 Then
        'self.debug_log("Awarded extra ball to Player %s. Shoot Again", self.player.index + 1)
        'self.player.extra_balls -= 1
        SetPlayerState "extra_balls", GetPlayerState("extra_balls") - 1
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000
        Exit Function
    End If
    
    SetPlayerState GLF_CURRENT_BALL, GetPlayerState(GLF_CURRENT_BALL) + 1

    Dim previousPlayerNumber : previousPlayerNumber = Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            If UBound(glf_playerState.Keys()) > 0 Then
                glf_currentPlayer = "PLAYER 2"
            End If
        Case "PLAYER 2":
            If UBound(glf_playerState.Keys()) > 1 Then
                glf_currentPlayer = "PLAYER 3"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 3":
            If UBound(glf_playerState.Keys()) > 2 Then
                glf_currentPlayer = "PLAYER 4"
            Else
                glf_currentPlayer = "PLAYER 1"
            End If
        Case "PLAYER 4":
            glf_currentPlayer = "PLAYER 1"
    End Select
    
    If useBcp Then
        bcpController.SendPlayerVariable "number", Getglf_currentPlayerNumber(), previousPlayerNumber
    End If
    If GetPlayerState(GLF_CURRENT_BALL) > glf_ballsPerGame Then
        DispatchPinEvent GLF_GAME_OVER, Null
    Else
        SetDelay "end_of_ball_delay", "EndOfBallNextPlayer", Null, 1000 
    End If
    
End Function

'****************************
' Game Over
' Event Listeners:      
AddPinEventListener GLF_GAME_OVER, "glf_game_over", "Glf_GameOver", 20, Null
'
'*****************************
Function Glf_GameOver(args)
    If GetPlayerStateForPlayer("0", "score") = False Then
        glf_machine_vars("player1_score").Value = 0
    Else
        glf_machine_vars("player1_score").Value = GetPlayerStateForPlayer("0", "score")
    End If
    If GetPlayerStateForPlayer("1", "score") = False Then
        glf_machine_vars("player2_score").Value = 0
    Else
        glf_machine_vars("player2_score").Value = GetPlayerStateForPlayer("1", "score")
    End If
    If GetPlayerStateForPlayer("2", "score") = False Then
        glf_machine_vars("player3_score").Value = 0
    Else
        glf_machine_vars("player3_score").Value = GetPlayerStateForPlayer("2", "score")
    End If
    If GetPlayerStateForPlayer("3", "score") = False Then
        glf_machine_vars("player4_score").Value = 0
    Else
        glf_machine_vars("player4_score").Value = GetPlayerStateForPlayer("3", "score")
    End If
    glf_gameStarted = False
    glf_currentPlayer = Null
    glf_playerState.RemoveAll()

    Dim device
    For Each device in glf_ball_devices
        If device.HasBall() Then
            device.EjectAll()
        End If
    Next
End Function

Public Function EndOfBallNextPlayer(args)
    DispatchPinEvent GLF_NEXT_PLAYER, Null
End Function

'*****************************************************************************************************************************************
'  ERROR LOGS by baldgeek
'*****************************************************************************************************************************************
Class GlfDebugLogFile
	Private Filename
	Private TxtFileStream

	Public default Function init()
        Dim timestamp : timestamp = GetTimeStamp(True)
        Filename = cGameName + "_" & timestamp & "_debug_log.txt"
		TxtFileStream = Null
		Set Init = Me
	End Function

	Public Sub EnableLogs()
		If glf_debugEnabled = True Then
			DisableLogs()
			Dim FormattedMsg, Timestamp, fso, logFolder
			Set fso = CreateObject("Scripting.FileSystemObject")
			logFolder = "glf_logs"
			If Not fso.FolderExists(logFolder) Then
				fso.CreateFolder logFolder
			End If
			Set TxtFileStream = fso.OpenTextFile(logFolder & "\" & Filename, 8, True)
		End If
	End Sub

	Public Sub DisableLogs()
		If Not IsNull(TxtFileStream) Then
			TxtFileStream.Close
		End If
	End Sub
	
	Private Function LZ(ByVal Number, ByVal Places)
		Dim Zeros
		Zeros = String(CInt(Places), "0")
		LZ = Right(Zeros & CStr(Number), Places)
	End Function
	
	Private Function GetTimeStamp(full)
		Dim CurrTime, Elapsed, MilliSecs
		CurrTime = Now()
		Elapsed = Timer()
		MilliSecs = Int((Elapsed - Int(Elapsed)) * 1000)
        If full = True Then
            GetTimeStamp = _
            LZ(Year(CurrTime),   4) & "-" _
            & LZ(Month(CurrTime),  2) & "-" _
            & LZ(Day(CurrTime),	2) & "_" _
            & LZ(Hour(CurrTime),   2) & "_" _
            & LZ(Minute(CurrTime), 2) & "_" _
            & LZ(Second(CurrTime), 2) & "_" _
            & LZ(MilliSecs, 4)
        Else
            GetTimeStamp = _
            LZ(Hour(CurrTime),   2) & "_" _
            & LZ(Minute(CurrTime), 2) & "_" _
            & LZ(Second(CurrTime), 2)
        End If
	End Function
	
	' *** Debug.Print the time with milliseconds, and a message of your choice
	Public Sub WriteToLog(label, message)
		If glf_debugEnabled = True Then
			Dim FormattedMsg, Timestamp
			Timestamp = GetTimeStamp(False)
			FormattedMsg = Timestamp & ": " & label & ": " & message
			TxtFileStream.WriteLine FormattedMsg
			Debug.Print label & ": " & message
		End If
		Glf_MonitorEventStream label, message
	End Sub
End Class

'*****************************************************************************************************************************************
'  END ERROR LOGS by baldgeek
'*****************************************************************************************************************************************


Dim glf_lastPinEvent : glf_lastPinEvent = Null

Dim glf_dispatch_parent : glf_dispatch_parent = 0
Dim glf_dispatch_q : Set glf_dispatch_q = CreateObject("Scripting.Dictionary")

Dim glf_frame_dispatch_count : glf_frame_dispatch_count = 0
Dim glf_frame_handler_count : glf_frame_handler_count = 0

Dim glf_dispatch_queue_int : glf_dispatch_queue_int = 0

Sub DispatchPinEvent(e, kwargs)
    AddToDispatchEvents e, kwargs, 1
End Sub

Sub AddToDispatchEvents(e, kwargs, event_type)
    If glf_dispatch_await.Exists(e) Then
        glf_dispatch_await.Remove e
    End If
    glf_dispatch_await.Add e & ";" & glf_dispatch_queue_int, Array(kwargs, event_type)
    glf_dispatch_queue_int = glf_dispatch_queue_int + 1
End Sub

Function DispatchPinHandlers(e, args)
    DispatchPinHandlers = Empty
    Dim handler : handler = args(0)
    Dim event_args, retArgs
    event_args = args(1)
    glf_frame_handler_count = glf_frame_handler_count + 1
    If IsNull(event_args(0)) Then
        Set retArgs = GlfKwargs()
    Else
        On Error Resume Next
        retArgs = event_args(0)
        If Err Then 
        Set retArgs = event_args(0)
        End If
    End If
    On Error Resume Next
        glf_dispatch_current_kwargs = retArgs	
    If Err Then 
        Set glf_dispatch_current_kwargs = retArgs
    End If
    If event_args(1) = 2 Then 'Queue Event
        Set retArgs = GetRef(handler(0))(Array(handler(2), retArgs, args(2)))
        If IsObject(retArgs) Then
            If retArgs.Exists("wait_for") Then
                DispatchPinHandlers = retArgs("wait_for")
            End If
        End If
    Else
        GetRef(handler(0))(Array(handler(2), event_args(0), args(2)))
    End If
End Function

Sub RunDispatchPinEvent(eKey, kwargs)
    Dim e
    e=Split(eKey,";")(0)
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If

    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If

    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            glf_frame_dispatch_count = glf_frame_dispatch_count + 1
            glf_dispatch_handlers_await.Add e&"_"&k(1), Array(handler, kwargs, e)
        Else
            Glf_WriteDebugLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
End Sub

Sub RunAutoFireDispatchPinEvent(e, kwargs)

    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchPinEvent", e & " has no listeners"
        Exit Sub
    End If

    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchPinEvent", e
    Dim handler
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        If handlers.Exists(k(1)) Then
            handler = handlers(k(1))
            glf_frame_dispatch_count = glf_frame_dispatch_count + 1
            'debug.print "Adding Handler for: " & e&"_"&k(1)
            'glf_dispatch_handlers_await.Add e&"_"&k(1), Array(handler, kwargs, e)
            'If SwitchHandler(handler(0), Array(handler(2), kwargs, e)) = False Then
                'debug.print e&"_"&k(1)
                GetRef(handler(0))(Array(handler(2), kwargs, e))
            'End If
        Else
            Glf_WriteDebugLog "DispatchPinEvent_"&e, "Handler does not exist: " & k(1)
        End If
    Next
    Glf_EventBlocks(e).RemoveAll

End Sub

Function DispatchRelayPinEvent(e, kwargs)
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchRelayPinEvent", e & " has no listeners"
        DispatchRelayPinEvent = kwargs
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k
    Dim handlers : Set handlers = glf_pinEvents(e)
    Glf_WriteDebugLog "DispatchReplayPinEvent", e
    For Each k In glf_pinEventsOrder(e)
        Glf_WriteDebugLog "DispatchReplayPinEvent_"&e, "key: " & k(1) & ", priority: " & k(0)
        kwargs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
    Next
    DispatchRelayPinEvent = kwargs
    Glf_EventBlocks(e).RemoveAll
End Function

Function DispatchQueuePinEvent(e, kwargs)
    
    AddToDispatchEvents e, kwargs, 2
    Exit Function
    
    If Not glf_pinEvents.Exists(e) Then
        Glf_WriteDebugLog "DispatchQueuePinEvent", e & " has no listeners"
        Exit Function
    End If
    If Not Glf_EventBlocks.Exists(e) Then
        Glf_EventBlocks.Add e, CreateObject("Scripting.Dictionary")
    End If
    glf_lastPinEvent = e
    Dim k,i,retArgs
    Dim handlers : Set handlers = glf_pinEvents(e)
    If IsNull(kwargs) Then
        Set kwargs = GlfKwargs()
    End If
    Glf_WriteDebugLog "DispatchQueuePinEvent", e
    Dim glf_dis_events : glf_dis_events = glf_pinEventsOrder(e)
    For i=0 to UBound(glf_dis_events)
        k = glf_dis_events(i)
        Glf_WriteDebugLog "DispatchQueuePinEvent"&e, "key: " & k(1) & ", priority: " & k(0)
        'msgbox "DispatchQueuePinEvent: " & e & " , key: " & k(1) & ", priority: " & k(0)
        'msgbox handlers(k(1))(0)
        'Call the handlers.
        'The handlers might return a waitfor command.
        'If NO wait for command, continue calling handlers.
        'IF wait for command, then AddPinEventListener for the waitfor event. The callback handler needs to be ContinueDispatchQueuePinEvent.
        Set retArgs = GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), kwargs, e))
        If retArgs.Exists("wait_for") And i<Ubound(glf_dis_events) Then
            'pause execution of handlers at index I. 
            Glf_WriteDebugLog "DispatchQueuePinEvent"&e, k(1) & "_wait_for"
            Dim wait_for : wait_for = retArgs("wait_for")
            kwargs.Remove "wait_for" 
            AddPinEventListener wait_for, k(1) & "_wait_for", "ContinueDispatchQueuePinEvent", k(0), Array(e, kwargs, i+1)
            Exit For
            'add event listener for the wait_for event.
            'pass in the index and handlers from this.
            'in the handler for resume queue event, process from the index the remaining handlers.
        End If
    Next
    Glf_EventBlocks(e).RemoveAll
End Function


'args Array(3)
' Array(original_event, orignal_kwargs, index)
' wait_for kwargs
' event
Function ContinueDispatchQueuePinEvent(args)
    
    Dim ownProps : ownProps = args(0)
    Dim kwargs
    If IsObject(args(1)) Then
        Set kwargs = args(1)
    Else
        kwargs = args(1)
    End If


    Dim i,key,keys,items
    keys=ownProps(0)
    items=ownProps(1)

    'Inject handlers back into dispatch
    For i=0 to UBound(keys)
        glf_dispatch_handlers_await.Add keys(i), items(i)
    Next
    Exit Function


End Function

Sub AddPinEventListener(evt, key, callbackName, priority, args)
    Dim i, inserted, tempArray
    If Not glf_pinEvents.Exists(evt) Then
        glf_pinEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_pinEvents(evt).Exists(key) Then
        glf_pinEvents(evt).Add key, Array(callbackName, priority, args)
        Sortglf_pinEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePinEventListener(evt, key)
    If glf_pinEvents.Exists(evt) Then
        If glf_pinEvents(evt).Exists(key) Then
            glf_pinEvents(evt).Remove key
            Sortglf_pinEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_pinEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_pinEventsOrder to maintain order based on priority
    If Not glf_pinEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_pinEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_pinEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_pinEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the glf_pinEventsOrder with the newly ordered or modified list
        glf_pinEventsOrder(evt) = tempArray
    End If
End Sub
Function GetPlayerState(key)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If glf_playerState(glf_currentPlayer).Exists(key)  Then
        GetPlayerState = glf_playerState(glf_currentPlayer)(key)
    Else
        GetPlayerState = False
    End If
End Function

Function GetPlayerStateForPlayer(player, key)
    dim p
    Select Case player
        Case 0:
            p = "PLAYER 1"
        Case 1:
            p = "PLAYER 2"
        Case 2:
            p = "PLAYER 3"
        Case 3:
            p = "PLAYER 4"
    End Select

    If glf_playerState.Exists(p) Then
        GetPlayerStateForPlayer = glf_playerState(p)(key)
    Else
        GetPlayerStateForPlayer = False
    End If
End Function

Function Getglf_currentPlayerNumber()
    Select Case glf_currentPlayer
        Case "PLAYER 1":
            Getglf_currentPlayerNumber = 0
        Case "PLAYER 2":
            Getglf_currentPlayerNumber = 1
        Case "PLAYER 3":
            Getglf_currentPlayerNumber = 2
        Case "PLAYER 4":
            Getglf_currentPlayerNumber = 3
    End Select
End Function

Function Getglf_PlayerName(player)
    Select Case player
        Case 0:
            Getglf_PlayerName = "PLAYER 1"
        Case 1:
            Getglf_PlayerName = "PLAYER 2"
        Case 2:
            Getglf_PlayerName = "PLAYER 3"
        Case 3:
            Getglf_PlayerName = "PLAYER 4"
    End Select
End Function

Function SetPlayerState(key, value)
    If IsNull(glf_currentPlayer) Then
        Exit Function
    End If

    If IsArray(value) Then
        If Join(GetPlayerState(key)) = Join(value) Then
            Exit Function
        End If
    Else
        If VarType(GetPlayerState(key)) <> vbBoolean Then
            If GetPlayerState(key) = value Then
                Exit Function
            End If
        End If
    End If   
    Dim prevValue
    If glf_playerState(glf_currentPlayer).Exists(key) Then
        prevValue = glf_playerState(glf_currentPlayer)(key)
        glf_playerState(glf_currentPlayer).Remove key
    End If
    glf_playerState(glf_currentPlayer).Add key, value
    If glf_debug_level = "Debug" Then
        Dim p,v : p = prevValue : v = value
        If IsNull(prevValue) Then
            p=""
        End If
        If IsNull(value) Then
            v=""
        End If
        Glf_WriteDebugLog "Player State", "Variable "& key &" changed from " & CStr(p) & " to " & CStr(v)
    End If
    Glf_MonitorPlayerStateUpdate key, value
    If glf_playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue, -1
    End If
    
    SetPlayerState = Null
End Function

Function SetPlayerStateByPlayer(key, value, player)

    If IsArray(value) Then
        If Join(GetPlayerStateForPlayer(player, key)) = Join(value) Then
            Exit Function
        End If
    Else
        If GetPlayerStateForPlayer(player, key) = value Then
            Exit Function
        End If
    End If   
    Dim prevValue, player_name
    player_name = Getglf_PlayerName(player)
    If glf_playerState(player_name).Exists(key) Then
        prevValue = glf_playerState(player_name)(key)
        glf_playerState(player_name).Remove key
    End If
    glf_playerState(player_name).Add key, value
    
    If glf_playerEvents.Exists(key) Then
        FirePlayerEventHandlers key, value, prevValue, player
    End If
    
    SetPlayerStateByPlayer = Null
End Function

Sub FirePlayerEventHandlers(evt, value, prevValue, player)
    If Not glf_playerEvents.Exists(evt) Then
        Exit Sub
    End If    
    Dim k
    Dim handlers : Set handlers = glf_playerEvents(evt)
    For Each k In glf_playerEventsOrder(evt)
        If handlers(k(1))(3) = player or handlers(k(1))(3) = Getglf_currentPlayerNumber() Then
            GetRef(handlers(k(1))(0))(Array(handlers(k(1))(2), Array(evt,value,prevValue)))
        End If
    Next
End Sub

Sub AddPlayerStateEventListener(evt, key, player, callbackName, priority, args)
    If Not glf_playerEvents.Exists(evt) Then
        glf_playerEvents.Add evt, CreateObject("Scripting.Dictionary")
    End If
    If Not glf_playerEvents(evt).Exists(key) Then
        glf_playerEvents(evt).Add key, Array(callbackName, priority, args, player)
        Sortglf_playerEventsByPriority evt, priority, key, True
    End If
End Sub

Sub RemovePlayerStateEventListener(evt, key)
    If glf_playerEvents.Exists(evt) Then
        If glf_playerEvents(evt).Exists(key) Then
            glf_playerEvents(evt).Remove key
            Sortglf_playerEventsByPriority evt, Null, key, False
        End If
    End If
End Sub

Sub Sortglf_playerEventsByPriority(evt, priority, key, isAdding)
    Dim tempArray, i, inserted, foundIndex
    
    ' Initialize or update the glf_playerEventsOrder to maintain order based on priority
    If Not glf_playerEventsOrder.Exists(evt) Then
        ' If the event does not exist in glf_playerEventsOrder, just add it directly if we're adding
        If isAdding Then
            glf_playerEventsOrder.Add evt, Array(Array(priority, key))
        End If
    Else
        tempArray = glf_playerEventsOrder(evt)
        If isAdding Then
            ' Prepare to add one more element if adding
            ReDim Preserve tempArray(UBound(tempArray) + 1)
            inserted = False
            
            For i = 0 To UBound(tempArray) - 1
                If priority > tempArray(i)(0) Then ' Compare priorities
                    ' Move existing elements to insert the new callback at the correct position
                    Dim j
                    For j = UBound(tempArray) To i + 1 Step -1
                        tempArray(j) = tempArray(j - 1)
                    Next
                    ' Insert the new callback
                    tempArray(i) = Array(priority, key)
                    inserted = True
                    Exit For
                End If
            Next
            
            ' If the new callback has the lowest priority, add it at the end
            If Not inserted Then
                tempArray(UBound(tempArray)) = Array(priority, key)
            End If
        Else
            ' Code to remove an element by key
            foundIndex = -1 ' Initialize to an invalid index
            
            ' First, find the element's index
            For i = 0 To UBound(tempArray)
                If tempArray(i)(1) = key Then
                    foundIndex = i
                    Exit For
                End If
            Next
            
            ' If found, remove the element by shifting others
            If foundIndex <> -1 Then
                For i = foundIndex To UBound(tempArray) - 1
                    tempArray(i) = tempArray(i + 1)
                Next
                
                ' Resize the array to reflect the removal
                ReDim Preserve tempArray(UBound(tempArray) - 1)
            End If
        End If
        
        ' Update the glf_playerEventsOrder with the newly ordered or modified list
        glf_playerEventsOrder(evt) = tempArray
    End If
End Sub

Sub EmitAllglf_playerEvents()
    Dim key
    For Each key in glf_playerState(glf_currentPlayer).Keys()
        FirePlayerEventHandlers key, GetPlayerState(key), GetPlayerState(key), -1
    Next
End Sub

'*******************************************
'  Drain, Trough, and Ball Release
'*******************************************
' It is best practice to never destroy balls. This leads to more stable and accurate pinball game simulations.
' The following code supports a "physical trough" where balls are not destroyed.
' To use this, 
'   - The trough geometry needs to be modeled with walls, and a set of kickers needs to be added to 
'	 the trough. The number of kickers depends on the number of physical balls on the table.
'   - A timer called "UpdateTroughTimer" needs to be added to the table. It should have an interval of 100 and be initially disabled.
'   - The balls need to be created within the Table1_Init sub. A global ball array (gBOT) can be created and used throughout the script


Dim ballInReleasePostion : ballInReleasePostion = False
'TROUGH 
Sub swTrough1_Hit
	ballInReleasePostion = True
	UpdateTrough
End Sub
Sub swTrough1_UnHit
	ballInReleasePostion = False
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
Sub swTrough6_Hit
	UpdateTrough
End Sub
Sub swTrough6_UnHit
	UpdateTrough
End Sub
Sub swTrough7_Hit
	UpdateTrough
End Sub
Sub swTrough7_UnHit
	UpdateTrough
End Sub
Sub Drain_Hit
	UpdateTrough
    If glf_gameStarted = True Then
        DispatchRelayPinEvent GLF_BALL_DRAIN, 1
    End If
End Sub
Sub Drain_UnHit
	UpdateTrough
End Sub

Sub UpdateTrough
	SetDelay "update_trough", "UpdateTroughDebounced", Null, 100
End Sub

Sub UpdateTroughDebounced(args)
	If glf_troughSize > 1 Then
		If swTrough1.BallCntOver = 0 Then swTrough2.kick 57, 10
	End If
	If glf_troughSize > 2 Then
		If swTrough2.BallCntOver = 0 Then swTrough3.kick 57, 10
	End If
	If glf_troughSize > 3 Then 
		If swTrough3.BallCntOver = 0 Then swTrough4.kick 57, 10
	End If
	If glf_troughSize > 4 Then 
		If swTrough4.BallCntOver = 0 Then swTrough5.kick 57, 10
	End If
	If glf_troughSize > 5 Then
		If swTrough5.BallCntOver = 0 Then swTrough6.kick 57, 10
	End If
	If glf_troughSize > 6 Then
		If swTrough6.BallCntOver = 0 Then swTrough7.kick 57, 10
	End If

	If glf_lastTroughSw.BallCntOver = 0 Then Drain.kick 57, 10
End Sub
