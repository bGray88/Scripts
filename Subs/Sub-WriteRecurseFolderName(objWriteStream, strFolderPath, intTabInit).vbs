
'******************************************************************************
' Purpose:	To recursively traverse all available folder levels and write
'			names of said folders to file on stream
'
' Inputs:	File to write to, Path to base folder, Initial tab count
'******************************************************************************
Sub WriteRecurseFolderName(objWriteStream, strFolderPath, intTabInit)
	Dim intTabCount : intTabCount = intTabInit
	Dim arrSubList : arrSubList = Array()
	Dim strFolderName : strFolderName = GetFileFolderName(strFolderPath)	
	Dim i : i = 0
	Do While i < intTabCount
		objWriteStream.Write(vbTab)
		i = i + 1
	Loop
	objWriteStream.WriteLine strFolderName
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
	Dim colFolders : Set colFolders = objFolder.SubFolders
	If TypeName(colFolders) <> "Nothing" Then
		If colFolders.Count > 0 Then
			Dim j : j = 0
			Dim objSubFolder
			For Each objSubFolder in colFolders
				ReDim Preserve arrSubList(j)
				arrSubList(j) = objSubFolder.Name
				j = j + 1
			Next
			arrSubList = BubbleSortList(arrSubList)
			For Each objSubFolder in arrSubList
				WriteRecurseFolderName objWriteStream, strFolderPath & "\" & _
					objSubFolder, intTabInit + 1
			Next
		End If
	End If
	
	Set objFolder = Nothing
	Set colFolders = Nothing
End Sub
