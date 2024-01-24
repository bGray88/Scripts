
'******************************************************************************
' Purpose:	To recursively traverse all available folder levels and write
'			names of all folders passed to file on stream
'
' Inputs:	File to write to, Path to base folder, Initial tab count
'
' Required:	Sub tab WriteMultTab
'******************************************************************************
Sub WriteRecurseFolderFileName(objWriteStream, strFolderPath, intTabInit)
	Dim intTabCount : intTabCount = intTabInit
	Dim arrSubList : arrSubList = Array()
	Dim arrListTitles : arrListTitles = Array()
	Dim strFolderName : strFolderName = GetFileFolderName(strFolderPath)
	Dim objSubFolder, ItemTitle, SubFolder
	Dim blnFileFound : blnFileFound = False
	
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
	Dim colFolders : Set colFolders = objFolder.SubFolders
	Dim intColSize : intColSize = colFolders.Count
	
	arrListTitles = ReadFileTitles(strFolderPath, intWinHeight, intWinWidth)
	
	If UBound(arrListTitles) >= 0 Then
		arrListTitles = BubbleSortList(arrListTitles)
		For Each ItemTitle in arrListTitles
			If (Right(ItemTitle, 3)) = "mp4" And blnFileFound = False Then
				blnFileFound = True
				objWriteStream.Write(VbCrLf)
				WriteMultTab objWriteStream, intTabCount
				objWriteStream.WriteLine strFolderName
				WriteMultTab objWriteStream, intTabCount
				objWriteStream.WriteLine "-------------------------"
				WriteMultTab objWriteStream, intTabCount
				objWriteStream.WriteLine ItemTitle
			ElseIf (Right(ItemTitle, 3)) = "mp4" Then
				WriteMultTab objWriteStream, intTabCount
				objWriteStream.WriteLine ItemTitle
			End If
		Next
	End If
	If TypeName(colFolders) <> "Nothing" Then
		If intColSize > 0 Then
			If blnFileFound = False Then
				objWriteStream.Write(VbCrLf)
				WriteMultTab objWriteStream, intTabCount
				objWriteStream.WriteLine strFolderName
				WriteMultTab objWriteStream, intTabCount
				objWriteStream.WriteLine "-------------------------"
			End If
			Dim j : j = 0
			For Each objSubFolder in colFolders
				ReDim Preserve arrSubList(j)
				arrSubList(j) = objSubFolder.Name
				j = j + 1
			Next
			arrSubList = BubbleSortList(arrSubList)
			For Each SubFolder in arrSubList
				WriteRecurseFolderFileName objWriteStream, strFolderPath & "\" & _
					SubFolder, intTabCount + 1
			Next
		End If
	End If
	
	Set objFolder = Nothing
	Set colFolders = Nothing
End Sub
