
'******************************************************************************
' Purpose:	Finding the multi character extension value of a file
'
' Inputs:	The full path to a file
'
' Returns:	The multi character extension value of a file
'******************************************************************************
Function GetExtension(strFileName)	
	Dim intNextDot : intNextDot = 0
	Dim intLastDot : intLastDot = 0
	Dim intCounter : intCounter = 1
	Do While intCounter <> Len(strFileName)
		intNextDot = InStr(intCounter, strFileName, ".")
		If intNextDot = 0 Then
			Exit Do
		Else
			intCounter = intNextDot + 1
		End If
		intLastDot = intNextDot
	Loop
	Dim strFileBase : strFileBase = Left(strFileName, intLastDot)
	GetExtension = Replace(strFileName, strFileBase, "")
End Function
