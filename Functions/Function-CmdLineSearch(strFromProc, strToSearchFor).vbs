
'******************************************************************************
' Purpose:	
'
' Inputs:	
'
' Returns:	
'******************************************************************************
Function CmdLineSearch(strFromProc, strToSearchFor)
	Dim arrLines : arrLines = Split(strFromProc, vbCrLf)
	Dim intInString : intInString = 0
	Dim i : i = 0
	For i To UBound(arrLines)
		intInString = InStr(arrLines(i), strToSearchFor)
		If intInString = 0 Or intInString = Null Then
			CmdLineSearch = False
		Else
			CmdLineSearch = True
		End If
	Next
End Function
