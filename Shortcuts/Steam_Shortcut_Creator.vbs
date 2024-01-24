
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
''''''''''''''''''''''''   SECTION: SETUP VARIABLES   '''''''''''''''''''''''''
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim strDesktopPath, strUserProfile, strSteamMediaPath, strSteamID
Dim strCurrentLine, strSteamGameTitle, strGameIcon, strIconCode, strGameTitle
Dim strLogFile, strSaveShortcuts
Dim objLink, ObjTextFile, objExplorer 
Dim arrFileNames, arrIconNames
Dim intIconArrSize, intTitleArrSize, intCurrentIndex
Dim listSteamIds

Dim arrNoIconFile : arrNoIconFile = Array()
Dim arrSteamTitles : arrSteamTitles = Array()
Dim dictSteamIds : Set dictSteamIds = CreateObject("Scripting.Dictionary")
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strIDFileExt : strIDFileExt = "acf"
Dim strIconFileExt : strIconFileExt = "ico"
Dim strIconUrlPath : strIconUrlPath = "https://steamdb.info/app"
Dim strAppTitle : strAppTitle = "Processing..."
Dim strAppTask : strAppTask = "Processing Files"
Dim intWinHeight : intWinHeight = 300
Dim intWinWidth : intWinWidth = 500

Dim strErrorPathNotFound : strErrorPathNotFound = "Steam may not be installed " & _
	"on this machine. " & vbCrLf & _
	"Please check the default location in your Program Files directory." & vbCrLf & _
	"Quitting."
Dim objXmlHttp : Set objXmlHttp = CreateObject("MSXML2.XMLHTTP.3.0")
strDesktopPath = ExpandEnvironmentStrings("%USERPROFILE%")
strDesktopPath = strDesktopPath & "\Desktop"
strSteamMediaPath = BrowseForFolder("Choose Steam Base Directory", "")
strSaveShortcuts = strDesktopPath & "\Steam Shortcuts"
strLogFile = "\log.txt"

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
''''''''''''''''''''''''         SECTION: MAIN        '''''''''''''''''''''''''
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'

If Not FolderExists(strSteamMediaPath) Then
	MsgBox strErrorPathNotFound
	WScript.Quit(0)
End If
If FolderEmpty(strSteamMediaPath) Then
	MsgBox strErrorPathNotFound
	WScript.Quit(0)
End If

arrFileNames = ReadFileTitles(strSteamMediaPath & "\steamapps\", intWinHeight, intWinWidth)
arrIconNames = ReadFileTitles(strSteamMediaPath & "\steam\games\", intWinHeight, intWinWidth)

Set objExplorer = WindowCreate(strAppTitle, _
	intWinHeight, intWinWidth)
		
intIconArrSize = 0
intCurrentIndex = 0
intTitleArrSize = UBound(arrFileNames)
For Each strFileName in arrFileNames
	' Get the Steam ID from the file title
	' ------------------------------------
	Redim arrSteamTitles(1)
	
	If strIDFileExt = GetExtension(strFileName) Then
		strSteamID = Left(strFileName, InStr(strFileName, "." & strIDFileExt)-1)
		strSteamID = Right(strSteamID, Len(strSteamID) - InStr(strSteamID, "_"))
		strSteamID = Replace(strSteamID, "_", "")
				
		If Not IsNumeric(strSteamID) Then
			MsgBox strSteamID & " Is Not A Valid Steam ID"
			WScript.Quit(0)
		End If
		
		' Get the Steam game's title
		' --------------------------
		If objFSO.FileExists(strSteamMediaPath & "\steamapps\" & strFileName) Then
			Set objTextFile = objFSO.OpenTextFile(strSteamMediaPath & "\steamapps\" & strFileName,_
				ForReading)
			
			Do While Not objTextFile.AtEndOfStream
				strCurrentLine = objTextFile.ReadLine
				If InStr(strCurrentLine, "name") Then
					strSteamGameTitle = Replace(strCurrentLine, "name", "")
					strSteamGameTitle = Replace(strSteamGameTitle, """", "")
					strSteamGameTitle = TrimSpecials(strSteamGameTitle)
					strSteamGameTitle = TrimAll(strSteamGameTitle)
					arrSteamTitles(0) = strSteamGameTitle
				End If
			Loop
			
			objTextFile.Close()
		Else
			MsgBox strSteamID & " Does Not Have A Corresponding ACF File"
			WScript.Quit(0)
		End If
	End If
	
	' Get the Steam Icon code from SteamDB
	' ------------------------------------
	objXmlHttp.Open "GET", strIconUrlPath & "/" & strSteamID, False
	objXmlHttp.Send
	If UBound(split(objXmlHttp.ResponseText, "clienticon")) > 0 Then
		strIconCode = split(objXmlHttp.ResponseText, "clienticon")(1)
		strIconCode = split(strIconCode, "</a>")(0)
		strIconCode = split(strIconCode, """nofollow"">")(1)
	End If
	
	For Each strIconFile in arrIconNames
		If strIconCode = Replace(strIconFile, "." & strIconFileExt, "") Then
			arrSteamTitles(1) = strIconFile
		Else
			ReDim Preserve arrNoIconFile(intIconArrSize)
			arrNoIconFile(intIconArrSize) = strIconCode
			intIconArrSize = intIconArrSize + 1
		End If
	Next
	If UBound(arrIconNames) < 0 Then
		ReDim Preserve arrNoIconFile(intIconArrSize)
		arrNoIconFile(intIconArrSize) = strSteamGameTitle
		intIconArrSize = intIconArrSize + 1
	End If
	
	' Dictionary to contain apps in format Key:AppID; Item:GameTitle; Item:GameIconCode
	' -------------------------------------------------------------------------------
	If strIDFileExt = GetExtension(strFileName) Then
		dictSteamIds.add strSteamID, arrSteamTitles
	End If
	
	HtmlProgress strAppTask, objExplorer, intCurrentIndex, intTitleArrSize
	intCurrentIndex = intCurrentIndex + 1
Next

WindowClose(objExplorer)

If dictSteamIds.Count = 0 Then
	MsgBox "There are no applications in your selected Steam directory" & vbCrLf & _
		"Quitting"
	WScript.Quit(0)
Else
	If Not FolderExists(strSaveShortcuts) Then
		objFSO.CreateFolder(strSaveShortcuts)
	End If

	listSteamIds = dictSteamIds.Keys
	
	Set objTextFile = objFSO.OpenTextFile(strSaveShortcuts & strLogFile,_
											ForWriting, True)
	objTextFile.WriteLine("Steam Shortcuts Created" & vbCrLf & _
							"----------------------" & vbCrLf)
							
	For Each objGameID in listSteamIds
		strGameTitle = dictSteamIds.Item(objGameID)(0)
		strGameIcon = dictSteamIds.Item(objGameID)(1)
		Set objLink = objShell.CreateShortcut(strSaveShortcuts & "\" & strGameTitle & ".lnk")
		objLink.TargetPath = "steam://run/" & objGameID
		objLink.Description = strGameTitle
		objLink.IconLocation = strSteamMediaPath & "\steam\games\" & strGameIcon & ", 0"
		objLink.WorkingDirectory = ""
		objLink.Save
		objTextFile.WriteLine(strGameTitle & " - Created")
	Next
		
	objTextFile.WriteLine(vbCrLf & "No Icons" & vbCrLf & _
									"--------" & vbCrLf)
	For Each strIcon in arrNoIconFile
		objTextFile.WriteLine(strIcon & " - No Icon")
	Next

	objTextFile.Close()

	MsgBox "Task Completed"
End If

' Clean up 
Set objLink = Nothing
Set dictSteamIds = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objTextFile = Nothing
Set objXmlHttp = Nothing

':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'
''''''''''''''''''''''''      SECTION: FUNCTIONS      '''''''''''''''''''''''''
':::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::'

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
' Purpose:	Finding the multi character extension value of a file
'
' Inputs:	The full path to a file
'
' Returns:	The multi character extension value of a file
'******************************************************************************
Function GetExtension(strFileName)	
	Dim intNextDot : intNextDot = 0
	Dim intLastDot : intLastDot = 0
	Dim intCounter : intCounter = 1
	Do While intCounter <> Len(strFileName)
		intNextDot = InStr(intCounter, strFileName, ".")
		If intNextDot = 0 Then
			Exit Do
		Else
			intCounter = intNextDot + 1
		End If
		intLastDot = intNextDot
	Loop
	Dim strFileBase : strFileBase = Left(strFileName, intLastDot)
	GetExtension = Replace(strFileName, strFileBase, "")
End Function

'******************************************************************************
' Purpose:	Determine the name of the file name in path
'
' Inputs:	String format file ex: calc.exe"
'
' Returns:	String of the name of the file name in path
'******************************************************************************
Function GetFileName(strFilePath)
	Dim intFirstSlash : intFirstSlash = 0
	Dim intSecondSlash : intSecondSlash = 0
	Dim intCounter : intCounter = 1
	Do While intCounter <> Len(strFilePath)
		intFirstSlash = InStr(intCounter, strFilePath, "\")
		If intFirstSlash = 0 Then
			Exit Do
		End If
		intSecondSlash = intFirstSlash
		intCounter = intCounter + 1
	Loop
	Dim strFolderPath : strFolderPath = Left(strFilePath, intSecondSlash)
	GetFileName = Replace(strFilePath, strFolderPath, "")
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
' Purpose:	Reads titles of files from a folder
'
' Inputs:	A full path string to a folder
'
' Returns:	An array of strings
'
' Optional:	Function HtmlProgress can be added for progress bar display
'******************************************************************************
Function ReadFileTitles(strFolderName, intWinHeight, intWinWidth)
	Dim arrList : arrList = Array()
	Dim strAppTitle : strAppTitle = "Reading..."
	Dim strAppTask : strAppTask = "Reading Files"
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderName)
	Dim colFiles : Set colFiles = objFolder.Files
	Dim intColSize : intColSize = colFiles.Count
	if intColSize > 0 Then
		Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
		Dim objFile
		Dim intSize : intSize = 0
		For Each objFile in colFiles
			ReDim Preserve arrList(intSize)
			arrList(intSize) = objFile.Name
			intSize = intSize + 1
			HtmlProgress strAppTask, objExplorer, intSize, intColSize
		Next
		WindowClose(objExplorer)
	End If
	ReadFileTitles = arrList
	Set objFolder = Nothing
	Set colFiles = Nothing
	Set objFSO = Nothing
End Function

'******************************************************************************
' Purpose:	Removes all white space (including Tabs) from strInput
'
' Inputs:	Title bar header text, starting directory for dialogue box
'
' Returns:	A string object of the full path to selected folder
'******************************************************************************
Function TrimAll(strInput)
	Dim strRegExp : Set strRegExp = New RegExp
	strRegExp.Pattern = "^\s*"
	strRegExp.Multiline = False
	TrimAll = strRegExp.Replace(strInput, "")
End Function

'******************************************************************************
' Purpose:	Removes all special character from strInput
'
' Inputs:	Title bar header text, starting directory for dialogue box
'
' Returns:	A string object of the full path to selected folder
'******************************************************************************
Function TrimSpecials(strInput)
	Dim strRegExp : Set strRegExp = New RegExp
	strRegExp.Global = True
	strRegExp.Pattern = "[^0-9a-zA-Z\s]"
	strRegExp.Multiline = False
	TrimSpecials = strRegExp.Replace(strInput, "")
End Function

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
