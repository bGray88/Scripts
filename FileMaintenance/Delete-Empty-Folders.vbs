Option Explicit
Dim strWindowTitle : strWindowTitle = "Select Folder to Scan"
Dim strFolderStart : strFolderStart = ""

DeleteFolders BrowseForFolder(strWindowTitle, strFolderStart)
MsgBox("The Delete Process Has Completed")
	
Sub DeleteFolders(strFolderSelected)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objSelectedFolder : Set objSelectedFolder = objFSO.GetFolder(strFolderSelected)
	Dim colFolders : Set colFolders = objSelectedFolder.SubFolders
	Dim Folder
	For Each Folder In colFolders
		If FolderEmpty(Folder) Then
			DeleteFolder Folder
		End If
	Next
	Set objFSO = Nothing
	Set objSelectedFolder = Nothing
	Set colFolders = Nothing
End Sub

'******************************************************************************
' Purpose:	Displays a dialogue box prompting for folder selection
'
' Inputs:	Title bar header text, starting directory for dialogue box
'
' Returns:	A string object of the full path to selected folder
'******************************************************************************
Function BrowseForFolder(strTitleBar, strStartDir)
	Const WinHandle = 0, Options = 0
	Dim objShell : Set objShell = WScript.CreateObject("Shell.Application")
    Dim objItem : Set objItem = objShell.BrowseForFolder(WinHandle, strTitleBar, _
		Options, strStartDir)
    Set objShell = Nothing
	
	If objItem Is Nothing Then
		WScript.Quit(1)
	End If

    BrowseForFolder = objItem.Self.path 
    Set objItem = Nothing
End Function

'******************************************************************************
' Purpose:	Delete specified folder
'
' Inputs:	The full path to select folder
'******************************************************************************
Sub DeleteFolder(strFolderPath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFolder strFolderPath, True
	Set objFSO = Nothing
End Sub

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
