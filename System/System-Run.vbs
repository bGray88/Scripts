Dim objShell
Dim strCurrentDirectory, strProgramName, strProgramExt
Dim strConfigExt

Set objShell = CreateObject("WScript.Shell")

strCurrentDirectory = objShell.ExpandEnvironmentStrings("%__CD__%")
strProgramName = "*PROGRAMNAME*"
strProgramExt = ".exe"
'strConfigExt = ".ini"

'If FileExists(strCurrentDirectory & strProgramName & strConfigExt) Then
'	DeleteFile(strCurrentDirectory & strProgramName & strConfigExt)
'End If

objShell.Run DoubleQuotes(strCurrentDirectory & strProgramName & strProgramExt)

Set objShell = Nothing

'******************************************************************************
' Purpose:	Delete specified file
'
' Inputs:	The full path to select file
'******************************************************************************
Sub DeleteFile(strFileName)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.DeleteFile(strFileName)
	Set objFSO = Nothing
End Sub

'******************************************************************************
' Purpose:	Finds whether the existence of a file is absolute
'
' Inputs:	The full path to possible file including the file name
'
' Returns:	Boolean value of determined file existence
'******************************************************************************
Function FileExists(strFilePath)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath) Then
		FileExists = True
	Else
		FileExists = False
	End If
	Set objFSO = Nothing
End Function

'******************************************************************************
' Purpose:	Surrounds string in escaped quote for literal string value
'
' Inputs:	The string to be encapsulated
'
' Returns:	The encapsulated string value
'******************************************************************************
Function DoubleQuotes(strIn)
    DoubleQuotes = """" & strIn & """"
End Function
