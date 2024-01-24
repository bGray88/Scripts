
'******************************************************************************
' Purpose:	To manually manage errors
'
' Inputs:	An integer value to indicate error
'******************************************************************************
Sub ErrHandling(intErrNumber)
	Select Case intErrNumber
		Case 1
			MsgBox("The file selected is not of the required type, quitting")
			Window.Close
			WaitSeconds 5
		Case Else
			Window.Close
			WaitSeconds 5
	End Select
End Sub
