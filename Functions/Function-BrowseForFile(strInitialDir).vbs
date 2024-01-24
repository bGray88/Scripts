
'******************************************************************************
' Purpose:	Browse for file dialogue box
'
' Inputs:	The starting directory for the dialogue box
'
' Returns:	A string that consists of the full path to file selected
'******************************************************************************
Function BrowseForFile(strInitialDir)
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Dim colItems : Set colItems = _
		objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	Dim strMshta : strMshta = "mshta.exe ""about:" & _
		"<" & "input type=file id=FILE>" & "<" & "script>FILE.click();" & _
		"new ActiveXObject('Scripting.FileSystemObject')" & _
		 ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & _
		 "<" & "/script>"""
    Dim intWinVersion
	Dim objItem
    For Each objItem in colItems
        intWinVersion = CInt(Left(objItem.version, 1))
    Next
    Set objWMIService = Nothing
    Set colItems = Nothing

    If (intWinVersion <= 5) Then
        ' Running XP and can use the original mechanism
		Dim objDialog : Set objDialog = CreateObject("UserAccounts.CommonDialog")
        objDialog.InitialDir = strInitialDir
        objDialog.Filter = "All Files|*.*"
		Do While True
			If objDialog.ShowOpen = True Then
				BrowseForFile = objDialog.FileName
				Exit Do
			Else
				BrowseForFile = ""
			End If
		Loop
        Set objDialog = Nothing
    Else
        ' Running Windows 7 or later
        Dim objShell : Set objShell = CreateObject("WScript.Shell")
        Dim objFilePath : Set objFilePath = objShell.Exec(strMshta)
		
		Dim strFilePath : strFilePath = _
			Replace(objFilePath.StdOut.ReadAll, VbCrLf, "")
		' Cancel was pressed
		If strFilePath = "" Then
			BrowseForFile = ""
		Else
			BrowseForFile = strFilePath
		End If
        Set objShell = Nothing
        Set objFilePath = Nothing
    End If
End Function
