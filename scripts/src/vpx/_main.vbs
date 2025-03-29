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