
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
