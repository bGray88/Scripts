
'******************************************************************************
' Purpose:	Organize the contents of an array list and return it
'
' Inputs:	An array
'
' Returns:	An alphabetically sorted array
'******************************************************************************
Function BubbleSortList(arrList)
	If IsArray(arrList) And UBound(arrList) > 0 Then
		Dim intCounter : intCounter = LBound(arrList)
		Dim intBubbleCounter : intBubbleCounter = LBound(arrList)
		Do While intCounter <> UBound(arrList)
			intBubbleCounter = intCounter
			Do While arrList(intBubbleCounter) > arrList(intBubbleCounter + 1)
				Call SwapPlaces(arrList(intBubbleCounter), arrList(intBubbleCounter + 1))
				If intBubbleCounter > LBound(arrList) Then
					intBubbleCounter = intBubbleCounter - 1
				End If
			Loop
			intCounter = intCounter + 1
		Loop
		BubbleSortList = arrList
	Else
		BubbleSortList = arrList
	End If
End Function
