
'******************************************************************************
' Purpose:	
'
' Inputs:	
'
' Returns:	
'******************************************************************************
Function CmdLineExtract(strFromProc, strToSearchFor)
	Dim arrLines : arrLines = Split(strFromProc, vbCrLf)
	Dim intInString = intInString = 0
	Dim i : i = 0
	For i To UBound(arrLines)
		intInString = InStr(arrLines(i), strToSearchFor)
		If Not intInString = 0 Or intInString = Null Then
			CmdLineExtract = arrLines(i)
		End If
	Next
End Function
