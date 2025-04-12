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
