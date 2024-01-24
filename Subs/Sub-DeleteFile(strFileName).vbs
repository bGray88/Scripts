
'******************************************************************************
' Purpose:	Delete specified file
'
' Inputs:	The full path to select file
'******************************************************************************
Sub DeleteFile(strFileName)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile(strFileName)
	Set objFSO = Nothing
End Sub
