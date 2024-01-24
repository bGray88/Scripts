
'******************************************************************************
' Purpose:	Removing all blank lines from a text file
'
' Inputs:	The full path to a file
'
' Returns:	New string value devoid of blank lines, and File line count
'******************************************************************************
Function RemoveBlankLines(strFileName)
	Const ForReading = 1
	intLineCount = 0
	Dim strLine
	Dim strNoBlankLines
	Dim objTxtFile : Set objTxtFile = objFSO.OpenTextFile(strFileName, _
	ForReading)
	Do Until objTxtFile.AtEndOfStream
		strLine = objTxtFile.ReadLine
		strLine = Trim(strLine)
		If Len(strLine) > 0 Then
			strNoBlankLines = strNoBlankLines & strLine & vbCrLf
		End If
		intLineCount = intLineCount + 1
	Loop
	objTxtFile.Close
	RemoveBlankLines = array(strNoBlankLines, intLineCount)
End Function
