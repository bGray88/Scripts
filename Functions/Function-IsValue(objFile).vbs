
'******************************************************************************
' Purpose:	Check whether a value has been returned
'
' Inputs:	An object
'
' Returns:	Boolean value of determined object status
'******************************************************************************
Function IsValue(objFile)
	On Error Resume Next
	Dim objTemp : objTemp = " " & objFile
	If Err <> 0 Then
		IsValue = False
	Else
		IsValue = True
	End If
	On Error GoTo 0
End Function
