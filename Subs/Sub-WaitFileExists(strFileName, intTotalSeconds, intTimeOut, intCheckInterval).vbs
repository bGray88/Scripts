
'******************************************************************************
' Purpose:	Sleep function that halts extra processes until file exists
'
' Inputs:	Check interval the duration slept until existence check
'******************************************************************************
' 
Sub WaitFileExists(strFileName, intTotalSeconds, intTimeOut, intCheckInterval)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim intInterval : intInterval = intCheckInterval
	If intInterval < 1 Then
		intInterval = 1
	End If
	Do While Not objFSO.FileExists(strFileName)
		If intTimeOut = 0 Then
			Exit Do
		End If
		objShell.Run "cmd /c ping localhost -n " & intInterval, 0, True
		If intTimeOut > 0 Then
			intTimeOut = intTimeOut - 1
		End If
	Loop
	Set objFSO = Nothing
	Set objShell = Nothing
End Sub
