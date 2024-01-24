
'******************************************************************************
' Purpose:	Counts the lines before EOF in text file
'
' Inputs:	The full path including file name of text file
'
' Returns:	Integer value of lines counted
'******************************************************************************
	Function CountLinesText(strTextPath)
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Const blCreateFalse = False, blCreateTrue = True
		Dim intLineCount : intLineCount = 0
		Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(strTextPath) Then
			Dim ObjTextFile : Set objTextFile = objFSO.OpenTextFile(strTextPath, ForReading, blCreateFalse)
			
			Do While Not objTextFile.AtEndOfStream
				objTextFile.ReadLine
				intLineCount = intLineCount + 1
			Loop
			objTextFile.Close()
			Set objTextFile = Nothing
		End If
		CountLinesText = intLineCount
		Set objFSO = Nothing
	End Function
