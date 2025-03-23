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
