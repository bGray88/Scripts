
'******************************************************************************
' Purpose:	Finds whether the existence of a file is absolute
'
' Inputs:	The full path to possible file including the file name
'
' Returns:	Boolean value of determined file existence
'******************************************************************************
Function FileExists(strFilePath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath) Then
		FileExists = True
	Else
		FileExists = False
	End If
	Set objFSO = Nothing
End Function
