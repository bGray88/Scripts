
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
