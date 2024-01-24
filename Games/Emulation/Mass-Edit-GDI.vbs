Option Explicit
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const blnCreateFileFalse = False, blnCreateFileTrue = True

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim objShell : Set objShell = CreateObject("WScript.Shell")

Dim	CurrentDirectory : CurrentDirectory = objShell.ExpandEnvironmentStrings("%__CD__%")

Dim objFolder : Set objFolder = objFSO.GetFolder(CurrentDirectory)
Dim colFolders : Set colFolders = objFolder.SubFolders
Dim objSubFolder
Dim colFiles
Dim objFile
Dim strGdi
Dim strGdiNew
Dim intCurrCount
Dim intLineCount
Dim folderItem
Dim objGdiFile
Dim intGdiUpdate
Dim extType
Dim strNoBlankLines
Dim arrGdi

For Each folderItem in colFolders
	Set objSubFolder = objFSO.GetFolder(folderItem.Name)
	If objFSO.FileExists(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".gdi") Then
		Set colFiles = objSubFolder.Files
		intCurrCount = 1
		arrGdi = RemoveBlankLines(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".gdi")
		strNoBlankLines = arrGdi(0)
		intLineCount = arrGdi(1)
		Set objGdiFile = objFSO.OpenTextFile(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".gdi", _
		ForWriting)
		objGdiFile.Write strNoBlankLines
		objGdiFile.Close
		For Each objFile in colFiles
			intGdiUpdate = 0
			Set objGdiFile = objFSO.OpenTextFile(CurrentDirectory & folderItem.Name & "\" & folderItem.Name & ".gdi", _
			ForReading)
			strGdi = objGdiFile.ReadAll
			objGdiFile.Close
			extType = GetExtension(objFile.Name)
			If Lcase(extType) = "bin" Or Lcase(extType) = "raw" Then
				If intLineCount > 9 And intCurrCount < 10 Then
					strGdiNew = Replace(strGdi, folderItem.Name & " (Track 0" & intCurrCount & ").bin", _
					"track0" & intCurrCount & "." & extType)
					intGdiUpdate = 1
				ElseIf intCurrCount > 9 Then
					strGdiNew = Replace(strGdi, folderItem.Name & " (Track " & intCurrCount & ").bin", _
					"track" & intCurrCount & "." & extType)
					intGdiUpdate = 1
				Else
					strGdiNew = Replace(strGdi, folderItem.Name & " (Track " & intCurrCount & ").bin", _
					"track0" & intCurrCount & "." & extType)
					intGdiUpdate = 1
				End If
			End If
			If intGdiUpdate = 1 Then
				Set objGdiFile = objFSO.OpenTextFile(CurrentDirectory & "\" & folderItem.Name & "\" & folderItem.Name & ".gdi", ForWriting)
				objGdiFile.Write strGdiNew
				objGdiFile.Close
				intCurrCount = intCurrCount + 1
			End If
		Next
	End If
Next

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

'******************************************************************************
' Purpose:	Removing all blank lines from a text file
'
' Inputs:	The full path to a file
'
' Returns:	New string value devoid of blank lines, and File line count
'******************************************************************************
Function RemoveBlankLines(strFileName)
	Const ForReading = 1
	Dim intLineCount = 0
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
