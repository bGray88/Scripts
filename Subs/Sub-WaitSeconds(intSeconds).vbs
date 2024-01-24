
'******************************************************************************
' Purpose:	HTA or WScript sleep function
'
' Inputs:	The time to sleep in seconds
'******************************************************************************
Sub WaitSeconds(intSeconds)
	Const HideWindow = 0, ShowWindow = 1
	Const WaitForCompletion = True, SkipPastCompletion = False
	Dim intTimeWaitSeconds : intTimeWaitSeconds = intSeconds
	Dim objWShell : Set objWShell = CreateObject("WScript.Shell")
	Dim strPingCommand : strPingCommand = "cmd /c ping 1.1.1.1 -n 1 -w " & intTimeWaitSeconds
	objWShell.Run strPingCommand, HideWindow, WaitForCompletion
	Set objWShell = Nothing
End Sub
