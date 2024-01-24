
'******************************************************************************
' Purpose:	To add a recurring number of tab characters to any writestream
'
' Inputs:	File to write to, Initial tab count
'******************************************************************************
Sub WriteMultTab(objTabStream, intInit)
	Dim intCount : intCount = intInit
	Dim i : i = 0
	
	Do While i < intCount
		objTabStream.Write(vbTab)
		i = i + 1
	Loop
End Sub
