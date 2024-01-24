
'******************************************************************************
' Purpose:	HTA sleep sub requires next function to be successful
'
' Inputs:	The time to wait in seconds, the next function to be called
'******************************************************************************
Sub WaitSeconds(intSeconds, strNextFunction)
	Dim intTimeWait : intTimeWait = intSeconds * 1000
	Dim idTimer : idTimer = Window.SetTimeOut(strNextFunction, _
		intTimeWait, "VBScript")
End Sub
