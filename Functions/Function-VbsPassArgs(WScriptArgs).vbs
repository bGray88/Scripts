
'******************************************************************************
' Purpose:	Receives arguments sent to the VBScript
'
' Inputs:	WScriptArgs should be any set of arguments sent to the script
'
' Returns:	An array of items processed from the arguments
'******************************************************************************
Function VbsPassArgs(WScriptArgs)
	Dim argsVarArr : argsVarArr = Array()
	Dim intArgsCount : intArgsCount = 0
	Dim objArgs : Set objArgs = WScriptArgs
	If WScriptArgs.Count <> 0 Then
		intArgsCount = WScriptArgs.Count
		For i = 0 To intArgsCount
			ReDim Preserve argsVarArr(i)
			argsVarArr(i) = objArgs(i)
		Next
		VbsPassArgs = argsVarArr
	Else
		VbsPassArgs = argsVarArr
	End If
	Set objArgs = Nothing
End Function
