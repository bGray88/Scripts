
'******************************************************************************
' Purpose:	Reads titles of folders from a folder
'
' Inputs:	A full path string to a folder
'
' Returns:	An array of strings
'
' Optional:	Function HtmlProgress can be added for progress bar display
'******************************************************************************
Function ReadFolderTitles(strFolderName, intWinHeight, intWinWidth)
	Dim arrList : arrList = Array()
	Dim strAppTitle : strAppTitle = "Reading..."
	Dim strAppTask : strAppTask = "Reading Files"
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim intSize : intSize = 0
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderName)
	Dim colFolders : Set colFolders = objFolder.Subfolders
	Dim intColSize : intColSize = colFolders.Count
	If intColSize > 0 Then
		Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
		Dim folderItem
		For Each folderItem in colFolders
			ReDim Preserve arrList(intSize)
			arrList(intSize) = folderItem.Name
			intSize = intSize + 1
			HtmlProgress strAppTask, objExplorer, intSize, intColSize
		Next
		WindowClose(objExplorer)
	End If
	ReadFolderTitles = arrList
	Set objFolder = Nothing
	Set colFolders = Nothing
	Set objFSO = Nothing
End Function
