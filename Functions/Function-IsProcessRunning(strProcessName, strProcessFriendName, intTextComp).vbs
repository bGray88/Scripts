
'******************************************************************************
' Purpose:	
'
' Inputs:	
'
' Returns:	
'******************************************************************************
Function IsProcessRunning(strProcessName, strProcessFriendName, intTextComp)
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:")
	Dim blFoundProc : blFoundProc = False
	Dim strProcName : strProcName = strProcessName & ".exe"
	Dim strProcFriendName : strProcFriendName = strProcessFriendName
	Dim vbTextCompare : vbTextCompare = intTextComp
	Dim Process
	For Each Process In objWMIService.InstancesOf("Win32_Process")
		If StrComp(Process.Name, strProcName, vbTextCompare) = 0 Then
			blFoundProc = True
		End If
	Next
	IsProcessRunning = blFoundProc
End Function