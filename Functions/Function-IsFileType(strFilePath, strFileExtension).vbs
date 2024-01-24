
'******************************************************************************
' Purpose:	Finds whether the existence of a file plus extension is absolute
'
' Inputs:	The full path to possible file including the file name
' 			and the file type extension to check for
'
' Returns:	Boolean value of determined file existence
'******************************************************************************
Function IsFileType(strFilePath, strFileExtension)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim strFileName : strFileName = objFSO.GetBaseName(strFilePath)
	Dim strPathName : strPathName = objFSO.GetParentFolderName(strFilePath)
	Dim strFullPathName : strFullPathName = strPathName & "\" & strFileName & _
		strFileExtension
	Select Case True
		Case objFSO.FileExists(strFullPathName)
			IsFileType = True
		Case Else
			IsFileType = False
	End Select
End Function
