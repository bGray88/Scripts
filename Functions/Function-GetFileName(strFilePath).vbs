
'******************************************************************************
' Purpose:	Determine the name of the file name in path
'
' Inputs:	String format file ex: calc.exe"
'
' Returns:	String of the name of the file name in path
'******************************************************************************
Function GetFileName(strFilePath)
	Dim intFirstSlash : intFirstSlash = 0
	Dim intSecondSlash : intSecondSlash = 0
	Dim intCounter : intCounter = 1
	Do While intCounter <> Len(strFilePath)
		intFirstSlash = InStr(intCounter, strFilePath, "\")
		If intFirstSlash = 0 Then
			Exit Do
		End If
		intSecondSlash = intFirstSlash
		intCounter = intCounter + 1
	Loop
	Dim strFolderPath : strFolderPath = Left(strFilePath, intSecondSlash)
	GetFileName = Replace(strFilePath, strFolderPath, "")
End Function
