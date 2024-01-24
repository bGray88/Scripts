
'******************************************************************************
' Purpose:	To recursively traverse all available folder levels and write
'			names of said folders to file on stream
'
' Inputs:	File to write to, Path to base folder, Initial tab count
'******************************************************************************
Sub WriteRecurseFileName(objWriteStream, strFolderPath, intTabInit)
	Dim intTabCount : intTabCount = intTabInit
	Dim arrFileList : arrFileList = Array()
	Dim arrSubList : arrSubList = Array()
	Dim strFolderName : strFolderName = GetFileFolderName(strFolderPath)
	Dim objFile, objFileName, objSubFolder
	Dim i : i = 0
	Do While i < intTabCount
		objWriteStream.Write(vbTab)
		i = i + 1
	Loop
	objWriteStream.WriteLine strFolderName
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
	Dim colFolders : Set colFolders = objFolder.SubFolders
	Dim colFiles : Set colFiles = objFolder.Files
	Dim intSize : intSize = 0
	For Each objFile in colFiles
		ReDim Preserve arrFileList(intSize)
		arrFileList(intSize) = objFile.Name
		intSize = intSize + 1
	Next
	If colFiles.Count > 0 Then
		arrFileList = BubbleSortList(arrFileList)
		For Each objFileName in arrFileList
			i = 0
			Do While i < intTabCount
				objWriteStream.Write(vbTab)
				i = i + 1
			Loop
			objWriteStream.Write(vbTab)
			objWriteStream.WriteLine objFileName
		Next
	End If
	If TypeName(colFolders) <> "Nothing" Then
		If colFolders.Count > 0 Then
			Dim j : j = 0
			For Each objSubFolder in colFolders
				ReDim Preserve arrSubList(j)
				arrSubList(j) = objSubFolder.Name
				j = j + 1
			Next
			arrSubList = BubbleSortList(arrSubList)
			For Each objSubFolder in arrSubList
				WriteRecurseFileName objWriteStream, strFolderPath & "\" & _
					objSubFolder, intTabInit + 1
			Next
		End If
	End If
	
	Set objFolder = Nothing
	Set colFolders = Nothing
End Sub
