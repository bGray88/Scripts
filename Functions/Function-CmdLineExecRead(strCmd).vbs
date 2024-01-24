
'******************************************************************************
' Purpose:	
'
' Inputs:	
'
' Returns:	
'******************************************************************************
Function CmdLineExecRead(strCmd)
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim objExec : Set objExec = objShell.Exec(strCmd)
	Dim strCmdOutput 
	strCmdOutput = objExec.StdOut.ReadAll()
	CmdLineExecRead = strCmdOutput
	Set objShell = Nothing
	Set objExec = Nothing
End Function
