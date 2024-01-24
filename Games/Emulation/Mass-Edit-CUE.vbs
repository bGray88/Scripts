Option Explicit
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const blnCreateFileFalse = False, blnCreateFileTrue = True

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objShell : Set objShell = CreateObject("WScript.Shell")

Dim	CurrentDirectory : CurrentDirectory = objShell.ExpandEnvironmentStrings("%__CD__%")

Dim objFolder : Set objFolder = objFSO.GetFolder(CurrentDirectory)
Dim colFolders : Set colFolders = objFolder.SubFolders
Dim objRegEx : Set objRegEx = CreateObject("VBScript.RegExp")
Dim strCue
Dim strCueNew
Dim folderItem
Dim objCueFile
Dim strNoBlankLines
Dim intStrExists
Dim arrCue

For Each folderItem in colFolders
	If objFSO.FileExists(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".cue") Then
		arrCue = RemoveBlankLines(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".cue")
		strNoBlankLines = arrCue(0)
		Set objCueFile = objFSO.OpenTextFile(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".cue", _
		ForWriting)
		objCueFile.Write strNoBlankLines
		objCueFile.Close
		objRegEx.IgnoreCase = True
		objRegEx.Pattern = "CDI/2352"
		Set objCueFile = objFSO.OpenTextFile(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".cue", _
		ForReading)
		strCue = objCueFile.ReadAll
		intStrExists = objRegEx.Test(strCue)
		objCueFile.Close
		If intStrExists Then
			Set objCueFile = objFSO.OpenTextFile(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".cue", _
			ForReading)
			strCue = objCueFile.ReadAll
			objCueFile.Close
			strCueNew = Replace(strCue, "CDI/2352", _
			"MODE2/2352")
			Set objCueFile = objFSO.OpenTextFile(CurrentDirectory & "\" & folderItem.Name & "\" & folderItem.Name & ".cue", ForWriting)
			objCueFile.Write strCueNew
			objCueFile.Close
		End If
	End If
Next

'******************************************************************************
' Purpose:	Removing all blank lines from a text file
'
' Inputs:	The full path to a file
'
' Returns:	New string value devoid of blank lines, and File line count
'******************************************************************************
Function RemoveBlankLines(strFileName)
	Const ForReading = 1
	Dim intLineCount : intLineCount = 0
	Dim strLine
	Dim strNoBlankLines
	Dim objTxtFile : Set objTxtFile = objFSO.OpenTextFile(strFileName, _
	ForReading)
	Do Until objTxtFile.AtEndOfStream
		strLine = objTxtFile.ReadLine
		If Len(strLine) > 0 Then
			strNoBlankLines = strNoBlankLines & strLine & vbCrLf
		End If
		intLineCount = intLineCount + 1
	Loop
	objTxtFile.Close
	RemoveBlankLines = array(strNoBlankLines, intLineCount)
End Function
