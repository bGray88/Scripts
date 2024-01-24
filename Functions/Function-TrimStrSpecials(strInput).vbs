
'******************************************************************************
' Purpose:	Removes all special character from strInput
'
' Inputs:	Title bar header text, starting directory for dialogue box
'
' Returns:	A string object of the full path to selected folder
'******************************************************************************
Function TrimSpecials(strInput)
	Dim strRegExp : Set strRegExp = New RegExp
	strRegExp.Global = True
	strRegExp.Pattern = "[^0-9a-zA-Z\s]"
	strRegExp.Multiline = False
	TrimSpecials = strRegExp.Replace(strInput, "")
End Function
