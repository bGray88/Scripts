Option Explicit
Const ShowWindow = 1, DontShowWindow = 0
Const WaitUntilFinished = True, DontWaitUntilFinished = False

Dim strUserFolder : strUserFolder = ExpandEnvironmentStrings("%USERPROFILE%")

Dim strSelectedFolder : strSelectedFolder = ""
Do While strSelectedFolder = ""
	strSelectedFolder = BrowseForFolder( "Find the file directory", "" )
	If strSelectedFolder = "" Then
		WScript.Quit()
	ElseIf Not FolderExists(strSelectedFolder) Then
		MsgBox("The Folder Does Not Exist")
		strSelectedFolder = ""
	End If
Loop

If Not FolderEmpty(strSelectedFolder) Then
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim intStartSeconds, intEndSeconds
	Do While True
		intStartSeconds = AnswerBox("Enter start time in seconds")
		If IsNumeric(intStartSeconds) Then
			intEndSeconds = AnswerBox("Enter end time in seconds")
			If IsNumeric(intEndSeconds) Then
				If intEndSeconds <= intStartSeconds Then
					MsgBox "End Time cannot be less than or equal to Start Time"
				Else
					Exit Do
				End If
			ElseIf intEndSeconds = "" Then
				MsgBox "Quitting"
				Set objFSO = Nothing
				Set objShell = Nothing
				WScript.Quit(0)
			End If
		ElseIf intStartSeconds = "" Then
			MsgBox "Quitting"
			Set objFSO = Nothing
			Set objShell = Nothing
			WScript.Quit(0)
		End If
	Loop

	Dim strNewFolderName : strNewFolderName = strUserFolder & "\Videos\" & _
		GetFileFolderName(strSelectedFolder) & "\"
	If FolderExists(strNewFolderName) Then
		Dim strRunDelete : strRunDelete = "cmd /c rd /s /q """ & strNewFolderName & """"  
		objShell.run strRunDelete, ShowWindow, WaitUntilFinished
	Else
		objFSO.CreateFolder(strNewFolderName)
	End If

	Dim objSelectedFolder : Set objSelectedFolder = objFSO.GetFolder(strSelectedFolder)
	Dim colMovieFiles : Set colMovieFiles = objSelectedFolder.Files
	Dim MovieTitle
	Dim intReturnCode
	Dim ConverterProgramPath : ConverterProgramPath = _
		"C:\Program Files (x86)\FreeTime\FormatFactory\FormatFactory.exe"
	For Each MovieTitle In colMovieFiles
		If objFSO.GetExtensionName(MovieTitle) = "mp4" Then
			intReturnCode = objShell.Run _
				(DoubleQuotes(ConverterProgramPath) & " " & DoubleQuotes("Custom") & _
				" " & DoubleQuotes("Custom AVC MP4") & _
				" " & DoubleQuotes(MovieTitle) & _
				" " & DoubleQuotes(strNewFolderName & objFSO.GetBaseName(MovieTitle) & ".mp4") & _
				" " & "/st=" & intStartSeconds & " /et=" & intEndSeconds, ShowWindow, WaitUntilFinished)
		End If
	Next
	Set objFSO = Nothing
	Set objShell = Nothing
	Set objSelectedFolder = Nothing
	Set colMovieFiles = Nothing
Else
	MsgBox "The folder selected is empty, Quitting"
End If

MsgBox "Conversion Complete!" & vbCrLf & _
		"Clips Can Be Found In User Videos Directory"

'******************************************************************************
' Purpose:	Accepts text input in text box
'
' Inputs:	Prompt for specific input, Boolean value indicating
'			whether empty values should be allowed
'
' Returns:	Returns path or empty value as string
'******************************************************************************
Function AnswerBox(strPrompt)
	Dim strAnswer
	Do While True
		strAnswer = InputBox(strPrompt)
		Select Case True
			Case IsEmpty(strAnswer) 'Cancel has been pressed
				AnswerBox = ""
				Exit Do
			Case "" = Trim(strAnswer) 'Empty answer has been entered
				AnswerBox = ""
				Exit Do
			Case Else 'Valid answer has been entered
				AnswerBox = strAnswer
				Exit Do
			End Select
	Loop
End Function

'******************************************************************************
' Purpose:	Displays a dialogue box prompting for folder selection
'
' Inputs:	Title bar header text, starting directory for dialogue box
'
' Returns:	A string object of the full path to selected folder
'******************************************************************************
Function BrowseForFolder(strTitleBar, strStartDir)
	Const WinHandle = 0, Options = 0
	Dim objShell : Set objShell = WScript.CreateObject("Shell.Application")
    Dim objItem : Set objItem = objShell.BrowseForFolder(WinHandle, strTitleBar, _
		Options, strStartDir)
    Set objShell = Nothing
	
	If objItem Is Nothing Then
		WScript.Quit(1)
	End If

    BrowseForFolder = objItem.Self.path 
    Set objItem = Nothing
End Function

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

'******************************************************************************
' Purpose:	Allows for the use of environment variables
'
' Inputs:	The string variable value to be processed
'
' Returns:	Returns expanded path in the form of a string
'******************************************************************************
Function ExpandEnvironmentStrings(strVariable)
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	ExpandEnvironmentStrings = objShell.ExpandEnvironmentStrings(strVariable)
	Set objShell = Nothing
End Function

'******************************************************************************
' Purpose:	Determines whether strFileName is an empty folder
'
' Inputs:	The full path to a folder
'
' Returns:	Boolean value of determined folder contents
'******************************************************************************
Function FolderEmpty(strFolderPath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

	If objFSO.FolderExists(strFolderPath) Then
		Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
		If objFolder.Files.Count = 0 And objFolder.SubFolders.Count = 0 Then
			FolderEmpty = True
		Else
			FolderEmpty = False
		End If
	End If
	Set objFSO = Nothing
	Set objFolder = Nothing
End Function

'******************************************************************************
' Purpose:	Determines whether strFileName is an existing folder
'
' Inputs:	The full path to a possible folder
'
' Returns:	Boolean value of determined folder existence
'******************************************************************************
Function FolderExists(strFolderName)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	FolderExists = objFSO.FolderExists(strFolderName)
	Set objFSO = Nothing
End Function

'******************************************************************************
' Purpose:	Determine the name of the lowest level folder name in path
'
' Inputs:	String format tree path ex: "C:\WINDOWS\System32"
'
' Returns:	String of the name of the lowest level folder name in path
'******************************************************************************
Function GetFileFolderName(strFolderPath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderPath)
	GetFileFolderName = objFolder.Name
	Set objFSO = Nothing
	Set objFolder = Nothing
End Function
