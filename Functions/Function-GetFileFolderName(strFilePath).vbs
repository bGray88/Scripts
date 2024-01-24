
'******************************************************************************
' Purpose:	Determine the name of the folder name in path
'
' Inputs:	String format tree path ex: "C:\WINDOWS\System32"
'
' Returns:	String of the name of the folder name in path
'******************************************************************************
Function GetFileFolderName(strFilePath)
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
	GetFileFolderName = Left(strFilePath, intSecondSlash)
End Function
