
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
