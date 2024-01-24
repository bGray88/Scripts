
'******************************************************************************
' Purpose:	
'
' Inputs:	
'******************************************************************************
Sub CallVbs(strScriptName, strArgs)
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim strCurrentDirectory : strCurrentDirectory = _
		objShell.ExpandEnvironmentStrings("%__CD__%")
	objShell.Run("cscript " & strScriptName ".vbs" & strArgs)
	Set objShell = Nothing
End Sub
