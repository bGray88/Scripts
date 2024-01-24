Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim fso : Set fso=CreateObject("Scripting.FileSystemObject")

Dim currFileName
Dim currFileExt
Dim newFolderName
Dim currFolder : currFolder = CurrentDirectory()
Dim baseFolder : Set baseFolder = fso.GetFolder(currFolder)

If baseFolder.Files.Count > 0 Then
	For Each file In baseFolder.Files
		currFileName = fso.GetFileName(file)
		currFileExt = LCase(fso.GetExtensionName(file))
		newFolderName = fso.GetBaseName(file)
		If Not currFileExt = "vbs" Then
			If Not fso.FolderExists(currFolder & newFolderName) Then
				fso.CreateFolder currFolder & newFolderName
			End If
			If newFolderName = fso.GetBaseName(currFileName) Then
				If fso.FolderExists(currFolder & newFolderName) Then
					fso.MoveFile currFolder & currFileName, currFolder & newFolderName & "\"
				End If
			End If
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
