Option Explicit
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const blnCreateFileFalse = False, blnCreateFileTrue = True

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strUserFolder : strUserFolder = ExpandEnvironmentStrings("%USERPROFILE%")

Dim strFileFolder : strFileFolder = ""
Do While strFileFolder = ""
	strFileFolder = BrowseForFolder( "Find the file directory", "" )
	If Not FolderExists(strFileFolder) Then
		MsgBox("The Folder Does Not Exist")
		strFileFolder = ""
	ElseIf FolderEmpty(strFileFolder) Then
		MsgBox("The Folder Is Empty")
		strFileFolder = ""
	End If
Loop

Dim strPathMediaFolderName : strPathMediaFolderName = GetFileFolderName(strFileFolder)
Dim strListTitle : strListTitle = "Missing ROM's (" & strPathMediaFolderName & ")"
Dim strListTextFile : strListTextFile = strUserFolder & "\Desktop\" & strListTitle & ".txt"

Dim arrScreenSize : arrScreenSize = GetScreenDimensions()
Dim intWinHeight : intWinHeight = (arrScreenSize(0) / 3)
Dim intWinWidth : intWinWidth = (arrScreenSize(1) / 3)

' Read titles from the folder
Dim arrListFiles : arrListFiles = ReadFileTitles(strFileFolder, intWinHeight, intWinWidth)
Dim arrListFolders : arrListFolders = ReadFolderTitles(strFileFolder, intWinHeight, intWinWidth)

If UBound(arrListFiles) >= 0 And UBound(arrListFolders) >= 0 Then

	' Organize the List
	arrListFiles = BubbleSortListProgress(arrListFiles, intWinHeight, intWinWidth)
	arrListFolders = BubbleSortListProgress(arrListFolders, intWinHeight, intWinWidth)

	Dim objTextStream
	' Check for currently existing lists and overwrite if necessary
	If objFSO.FileExists(strListTextFile) Then
		' Open the text file
		Set objTextStream = objFSO.OpenTextFile(strListTextFile, ForWriting, _
			blnCreateFileFalse)
	Else
		' Create the text file
		Set objTextStream = objFSO.CreateTextFile(strListTextFile, ForWriting, _
			blnCreateFileTrue)
	End If
	'Write Title to the contents of the text file
	objTextStream.WriteLine strListTitle
	objTextStream.WriteLine "-------------------------" & VbCrLf
	objTextStream.WriteLine "File Listing"
	objTextStream.WriteLine "-------------------------" & VbCrLf
	
	Dim intWriteCounter : intWriteCounter = 0
	Dim strAppTitle : strAppTitle = "Writing Files"
	Dim strAppTask : strAppTask = "Writing..."
	Dim intArrSize : intArrSize = UBound(arrListFolders) + 1
	Dim objFile, objFolder
	Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, intWinHeight, intWinWidth)
	HtmlProgress strAppTask, objExplorer, intWriteCounter, intArrSize
	For Each objFile in arrListFolders
		Dim blnFoundMatch : blnFoundMatch = false
		Dim strFileExt : strFileExt = GetExtension(objFile)
		Dim strTitleName : strTitleName = Replace(objFile, "." & strFileExt, "")
		For Each objFolder in arrListFiles
			strFileExt = GetExtension(objFolder)
			Dim strFileName : strFileName = Replace(objFolder, "." & strFileExt, "")
			If StrComp(strTitleName, strFileName, 1) = 0 Then
				blnFoundMatch = true
				Exit For
			End If
		Next
		If blnFoundMatch = false Then
			objTextStream.WriteLine(strTitleName)
		End If
		intWriteCounter = intWriteCounter + 1
		HtmlProgress strAppTask, objExplorer, intWriteCounter, intArrSize
	Next
	
	WindowClose(objExplorer)
	objTextStream.Close
	Set objTextStream = Nothing
	Set objExplorer = Nothing
Else
	MsgBox("The folder selected is clean, Quitting")
End If

Set objFSO = Nothing

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
' Purpose:	Organize the contents of an array list and return it
'
' Inputs:	An array
'
' Optional:	Function HtmlProgress can be added for progress bar display
'
' Returns:	An alphabetically sorted array
'******************************************************************************
Function BubbleSortListProgress(arrList, intWinHeight, intWinWidth)
	Dim strAppTitle : strAppTitle = "Sorting..."
	Dim strAppTask : strAppTask = "Sorting Files"
	Dim objExplorer
	If IsArray(arrList) And UBound(arrList) > 0 Then
		Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
		Dim intCounter : intCounter = LBound(arrList)
		HtmlProgress strAppTask, objExplorer, intCounter, UBound(arrList)
		Dim intBubbleCounter : intBubbleCounter = LBound(arrList)
		Do While intCounter <> UBound(arrList)
			intBubbleCounter = intCounter
			Do While arrList(intBubbleCounter) > arrList(intBubbleCounter + 1)
				Call SwapPlaces(arrList(intBubbleCounter), arrList(intBubbleCounter + 1))
				If intBubbleCounter > LBound(arrList) Then
					intBubbleCounter = intBubbleCounter - 1
				End If
			Loop
			intCounter = intCounter + 1
			HtmlProgress strAppTask, objExplorer, intCounter, UBound(arrList)
		Loop
		WindowClose(objExplorer)
		BubbleSortListProgress = arrList
	ElseIf UBound(arrList) = 0 Then
		Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
		intCounter = LBound(arrList)
		HtmlProgress strAppTask, objExplorer, intCounter, UBound(arrList)
		intCounter = intCounter + 1
		HtmlProgress strAppTask, objExplorer, intCounter, UBound(arrList) + 1
		WindowClose(objExplorer)
		BubbleSortListProgress = arrList
	Else
		BubbleSortListProgress = arrList
	End If
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
' Purpose:	Collects the screen display dimensions height and width
'
' Inputs:	N/A
'
' Returns:	Array containing int values in height(0) and width(1)
'******************************************************************************
Function GetScreenDimensions()
	Dim strComputer : strComputer = "."
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Dim colItems : Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor",,48)
	Dim arrScreenSize(1)
	Dim objItem
	For Each objItem in colItems
		arrScreenSize(0) = objItem.ScreenHeight
		arrScreenSize(1) = objItem.ScreenWidth
	Next
	
	GetScreenDimensions = arrScreenSize
	Set objWMIService = Nothing
	Set colItems = Nothing
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
' Purpose:	Reads titles of folders from a folder
'
' Inputs:	A full path string to a folder
'
' Returns:	An array of strings
'
' Optional:	Function HtmlProgress can be added for progress bar display
'******************************************************************************
Function ReadFolderTitles(strFolderName, intWinHeight, intWinWidth)
	Dim arrList : arrList = Array()
	Dim strAppTitle : strAppTitle = "Reading..."
	Dim strAppTask : strAppTask = "Reading Files"
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim intSize : intSize = 0
	Dim objFolder : Set objFolder = objFSO.GetFolder(strFolderName)
	Dim colFolders : Set colFolders = objFolder.Subfolders
	Dim intColSize : intColSize = colFolders.Count
	If intColSize > 0 Then
		Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
		Dim folderItem
		For Each folderItem in colFolders
			ReDim Preserve arrList(intSize)
			arrList(intSize) = folderItem.Name
			intSize = intSize + 1
			HtmlProgress strAppTask, objExplorer, intSize, intColSize
		Next
		WindowClose(objExplorer)
	End If
	ReadFolderTitles = arrList
	Set objFolder = Nothing
	Set colFolders = Nothing
	Set objFSO = Nothing
End Function

'******************************************************************************
' Purpose:	Swaps the positions of a and b
'
' Inputs:	Value a and value b
'******************************************************************************
' Exchanges the data in a and b
Sub SwapPlaces(a, b)
	Dim temp : temp = a
	a = b
	b = temp
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
