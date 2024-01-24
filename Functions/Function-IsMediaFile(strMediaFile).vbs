
'******************************************************************************
' Purpose:	Finds whether the extension of a file is media associated
'
' Inputs:	The file name including file type extension to check for
'
' Returns:	Boolean value of determined file existence
'******************************************************************************
Function IsMediaFile(strMediaFile)
	Dim intNextDot : intNextDot = 0
	Dim intLastDot : intLastDot = 0
	Dim intCounter : intCounter = 1
	Do While intCounter <> Len(strMediaFile)
		intNextDot = InStr(intCounter, strMediaFile, ".")
		If intNextDot = 0 Then
			Exit Do
		Else
			intCounter = intNextDot + 1
		End If
		intLastDot = intNextDot
	Loop
	Dim strFileBase : strFileBase = Left(strMediaFile, intLastDot)
	Dim strFileExt : strFileExt = Replace(strMediaFile, strFileBase, "")
	
	Select Case strFileExt
		Case "mp4"
			IsMediaFile = True
		Case "mp3"
			IsMediaFile = True
		Case "flac"
			IsMediaFile = True
		Case "wav"
			IsMediaFile = True
		Case Else
			IsMediaFile = False
	End Select
End Function
