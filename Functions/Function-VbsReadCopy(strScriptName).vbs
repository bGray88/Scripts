
'******************************************************************************
' Purpose:	Reads another VBScript
'
' Inputs:	A string value
'
' Returns:	String statement to be executed
'******************************************************************************
Function VbsReadCopy(strScriptName)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objTextFile : Set objTextFile = objFSO.OpenTextFile(strScriptName & ".vbs")
	Dim strBodyText : strBodyText = objTextFile.ReadAll
	objTextFile.Close
	Set objTextFile = Nothing
	VbsReadCopy = strBodyText
	Set objFSO = Nothing
End Function
