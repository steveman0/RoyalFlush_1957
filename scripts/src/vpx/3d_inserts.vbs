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
