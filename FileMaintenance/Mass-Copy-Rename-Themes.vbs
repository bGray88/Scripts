Option Explicit
Const OverwriteFile = False

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strInitFolder : strInitFolder = ""
Dim strBrowseTitle : strBrowseTitle = ""

Do
	MsgBox "Please select the folder containing the default.zip"
	Dim strSelectedZipFolder : strSelectedZipFolder = _
		BrowseForFolder(strBrowseTitle, strInitFolder)
	If strSelectedZipFolder = "" Then
		WScript.Quit(0)
	ElseIf Not objFSO.FileExists(strSelectedZipFolder & "\Default.zip") Then
		MsgBox "Default.zip Not Found"
		strSelectedZipFolder = ""
	End If
Loop While strSelectedZipFolder = ""

Do
	MsgBox "Please select the folder containing the media files"
	Dim strSelectedMediaFolder : strSelectedMediaFolder = _
		BrowseForFolder(strBrowseTitle, strInitFolder)
	If strSelectedMediaFolder = "" Then
		WScript.Quit(0)
	Else
		Dim objSelectedFolder : Set objSelectedFolder = _
			objFSO.GetFolder(strSelectedMediaFolder)
		Dim colMediaFiles : Set colMediaFiles = objSelectedFolder.Files
		Dim intColSize : intColSize = colMediaFiles.Count
		If intColSize < 1 Then
			MsgBox "Media Files Not Found"
			strSelectedMediaFolder = ""
		End If
		Set objSelectedFolder = Nothing
	End If
Loop While strSelectedMediaFolder = ""

If Not intColSize = 0 Then
	Dim strUserFolder : strUserFolder = ExpandEnvironmentStrings("%USERPROFILE%")
	Dim strZipFile : strZipFile = strSelectedZipFolder & "\Default.zip"
	Dim strOutputFolder : strOutputFolder = strUserFolder & _
		"\Desktop\" & GetFileFolderName(strSelectedZipFolder) & "\"

	If Not objFSO.FolderExists(strOutputFolder) Then
		objFSO.CreateFolder(strOutputFolder)
	End If

	Dim intWinHeight : intWinHeight = 200
	Dim intWinWidth : intWinWidth = 400
	Dim strAppTitle : strAppTitle = "Writing Files"
	Dim strApptask : strAppTask = "Writing..."
	intColSize = colMediaFiles.Count
	Dim intWriteCounter : intWriteCounter = 0
	Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
		intWinHeight, intWinWidth)
	Dim ItemTitle
	For Each ItemTitle In colMediaFiles
		intWriteCounter = intWriteCounter + 1
		HtmlProgress strAppTask, objExplorer, intWriteCounter, intColSize
		Dim strItemName : strItemName = objFSO.GetBaseName(ItemTitle)
		objFSO.CopyFile strZipFile, strOutputFolder, OverwriteFile
		objFSO.MoveFile strOutputFolder & "Default.zip", _
			strOutputFolder & strItemName & ".zip"
	Next

	WindowClose(objExplorer)
	Set objExplorer = Nothing
Else
	MsgBox "File does not contain any files to be copied"
End If

Set objFSO =  Nothing
Set colMediaFiles = Nothing

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
' Purpose:	Checks rounding eligibility of number to the next whole number
'
' Inputs:	A number value
'
' Returns:	An integer value
'******************************************************************************
Function DblRoundUp(intDouble)
	If intDouble = Round(intDouble) Then
		DblRoundUp = intDouble
	Else
		DblRoundUp = Round(intDouble + 0.5)
	End If
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

'******************************************************************************
' Purpose:	Adds a progress bar to any task that can be measured by 
' 			integer variables
'
' Inputs:	Current task name, IE Application object
'
' Optional:	If Round is used instead of custom function DblRoundUp 100
'			percent may be unattainable in some circumstances
'******************************************************************************
Sub HtmlProgress(strCurrentTask, objHtmlExplorer, intCurrentCount, intTotalCount)
	Dim dblBarSize : dblBarSize = objHtmlExplorer.Width * .10
	Dim intCompleteProgress : intCompleteProgress = 100
	Dim dblProgressMod : dblProgressMod = .78
	If intTotalCount > 0 Then
		Dim intCurrentProgress : intCurrentProgress = DblRoundUp(intCurrentCount / intTotalCount * intCompleteProgress)
		'intCurrentProgress = Round(intCurrentCount / intTotalCount * intCompleteProgress)
		Dim intBarProgress : intBarProgress = DblRoundUp(intCurrentCount / intTotalCount * dblBarSize) * dblProgressMod
		'intBarProgress = Round(intCurrentCount / intTotalCount * dblBarSize) * dblProgressMod
		Dim strTaskMessage : strTaskMessage = "Completing requested task - " & strCurrentTask & "...<BR>" & _
			"This might take several minutes to complete.<BR>" & _
			"<BR>"
		Dim strProgressNum : strProgressNum = "Current Status " & intCurrentProgress & "% Complete<BR>"
		Dim strProgressBar : strProgressBar = "<" & "input align='left' " & _
			"size=" & dblBarSize & " type='text' " & _
			"value=" & ChrW(9632) & String(intBarProgress, ChrW(9632)) & " disabled='disabled'>"
		objHtmlExplorer.Document.Body.InnerHTML = strTaskMessage & strProgressNum & strProgressBar
		If intCurrentCount = intTotalCount Then
			objHtmlExplorer.Document.Body.InnerHTML = strTaskMessage & _
				"Current Status " & intCompleteProgress & "% Complete<BR>" & _
				strProgressBar & "<BR>" & _
				"Task Complete"
		End If
	End If
End Sub

'******************************************************************************
' Purpose:	Closes the currently opened IE window
'
' Inputs:	The IE object
'******************************************************************************
Sub WindowClose(objDoneExplorer)
	If Not objDoneExplorer Is Nothing then
		If IsObject(objDoneExplorer) Then
			On Error Resume Next
			WScript.Sleep 2000
			objDoneExplorer.Quit
			On Error GoTo 0
		End If
	End If
End Sub

'******************************************************************************
' Purpose:	Creates a non-resizeable window via IE 
'
' Inputs:	strTitle is the header title bar string, intWindowHeight is the
' 			overall height of the window, intWindowWidth does the same for the
'			the width of the window
'
' Returns:	An IE Application object
'******************************************************************************
Function WindowCreate(strTitle, intWindowHeight, intWindowWidth)
	Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
	Dim objCreateExplorer : Set objCreateExplorer = WScript.CreateObject("InternetExplorer.Application", "IE_")
	objCreateExplorer.Navigate "about:blank"
	objCreateExplorer.Document.Title = strTitle & String(80, ".")
	objCreateExplorer.ToolBar = 0
	objCreateExplorer.MenuBar = 0
	objCreateExplorer.StatusBar = 0
	objCreateExplorer.Resizable = False
	objCreateExplorer.Width = intWindowWidth
	objCreateExplorer.Height = intWindowHeight 
	objCreateExplorer.Left = 0
	objCreateExplorer.Top = 0
	
	Do While (objCreateExplorer.Busy)
		WScript.Sleep 200
	Loop
	
	objCreateExplorer.Visible = True
	objShell.AppActivate objCreateExplorer
	Dim objDocument : Set objDocument = objCreateExplorer.Document
	objDocument.Body.Scroll = "No"
	
	Set WindowCreate = objCreateExplorer
	Set objShell = Nothing
End Function
