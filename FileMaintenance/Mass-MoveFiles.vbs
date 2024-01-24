Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim fso : Set fso=CreateObject("Scripting.FileSystemObject")

Dim procFolder
Dim subfolderName
Dim newFolderName : newFolderName = "Wallpapers"
Dim currFolder : currFolder = CurrentDirectory()
Dim baseFolder : Set baseFolder = fso.GetFolder(currFolder)

If baseFolder.Files.Count > 0 Then
	For Each folder In baseFolder.Subfolders
		Set procFolder = fso.GetFolder(folder)
		If Not fso.FolderExists(folder & "\" & newFolderName) Then
			fso.CreateFolder folder & "\" & newFolderName
		End If
		If procFolder.Files.Count > 0 Then
			fso.MoveFile folder & "\*", folder & "\" & newFolderName & "\"
		End If
		If procFolder.Subfolders.Count > 1 Then
			For Each subfolder In procFolder.Subfolders
				If fso.FolderExists(folder & "\" & newFolderName) Then
					If Not GetFileFolderName(subfolder) = newFolderName Then
						subfolderName = GetFileFolderName(subfolder)
						fso.MoveFolder subfolder, folder & "\" & newFolderName & "\"
					End If
				End If
			Next
		End If
	Next
End If

WScript.Echo "Task Completed"
WScript.Quit

'******************************************************************************
' Purpose:	Gathers the current working directory
'
' Inputs:	None
'
' Returns:	A string value of the full path to working directory
'******************************************************************************
Function CurrentDirectory()
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	CurrentDirectory = objShell.ExpandEnvironmentStrings("%__CD__%")
	Set objShell = Nothing
End Function

'******************************************************************************
' Purpose:	Determines whether strFileName is an empty folder
'
' Inputs:	The full path to a folder
'
' Returns:	Boolean value of determined folder contents
'******************************************************************************
Function FolderEmpty(strFolderPath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FolderExists(strFolderPath) Then
		Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
		If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then
			FolderEmpty = True
		Else
			FolderEmpty = False
		End If
	End If
	Set objFSO = Nothing
	Set objFolder = Nothing
End Function

'******************************************************************************
' Purpose:	Determine the name of the lowest level folder name in path
'
' Inputs:	String format tree path ex: "C:\WINDOWS\System32"
'
' Returns:	String of the name of the lowest level folder name in path
'******************************************************************************
Function GetFileFolderName(strFolderPath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
	GetFileFolderName = objFolder.Name
	Set objFSO = Nothing
	Set objFolder = Nothing
End Function
