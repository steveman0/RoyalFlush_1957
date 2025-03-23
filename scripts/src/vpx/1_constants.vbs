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
