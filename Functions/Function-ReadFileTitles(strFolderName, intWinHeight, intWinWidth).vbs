
'******************************************************************************
' Purpose:	Reads titles of files from a folder
'
' Inputs:	A full path string to a folder
'
' Returns:	An array of strings
'
' Optional:	Function HtmlProgress can be added for progress bar display
'******************************************************************************
Function ReadFileTitles(strFolderName, intWinHeight, intWinWidth)
	Dim arrList : arrList = Array()
	Dim strAppTitle : strAppTitle = "Reading..."
	Dim strAppTask : strAppTask = "Reading Files"
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderName)
	Dim colFiles : Set colFiles = objFolder.Files
	Dim intColSize : intColSize = colFiles.Count
	if intColSize > 0 Then
		Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
		Dim objFile
		Dim intSize : intSize = 0
		For Each objFile in colFiles
			ReDim Preserve arrList(intSize)
			arrList(intSize) = objFile.Name
			intSize = intSize + 1
			HtmlProgress strAppTask, objExplorer, intSize, intColSize
		Next
		WindowClose(objExplorer)
	End If
	ReadFileTitles = arrList
	Set objFolder = Nothing
	Set colFiles = Nothing
	Set objFSO = Nothing
End Function
