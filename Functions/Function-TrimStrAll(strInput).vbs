
'******************************************************************************
' Purpose:	Removes all white space (including Tabs) from strInput
'
' Inputs:	Title bar header text, starting directory for dialogue box
'
' Returns:	A string object of the full path to selected folder
'******************************************************************************
Function TrimAll(strInput)
	Dim strRegExp : Set strRegExp = New RegExp
	strRegExp.Pattern = "^\s*"
	strRegExp.Multiline = False
	TrimAll = strRegExp.Replace(strInput, "")
End Function
