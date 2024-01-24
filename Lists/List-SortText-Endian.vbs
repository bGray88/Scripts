Option Explicit
Const ForReading = 1, ForWriting = 2, ForAppending = 8, TristateUseDefault = -2
Const blnCreateFileFalse = False, blnCreateFileTrue = True

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strUserFolder : strUserFolder = ExpandEnvironmentStrings("%USERPROFILE%")

Dim objListFile : objListFile = "" 
Do While objListFile = ""
	objListFile = BrowseForFile("")
	If objListFile = "" Then
		WScript.Quit(0)
	ElseIf Not FileExists(objListFile) Then
		MsgBox("The File Does Not Exist")
		objListFile = ""
	ElseIf Not IsFileType(objListFile, ".txt") Then
		MsgBox("The File Is Not A Text File")
		objListFile = ""
	End If
Loop

Dim strFullListPath : strFullListPath = objListFile
Dim strListTitle : strListTitle = "List Files " & objFSO.GetBaseName(objFSO.GetFile(objListFile))

Dim arrScreenSize : arrScreenSize = GetScreenDimensions()
Dim intWinHeight : intWinHeight = (arrScreenSize(0) / 3)
Dim intWinWidth : intWinWidth = (arrScreenSize(1) / 3)

Dim objTextStream
Set objTextStream = objFSO.OpenTextFile(strFullListPath, ForReading, _
		blnCreateFileFalse, TristateUseDefault)

Dim arrListTitles : arrListTitles = Array()
Dim i : i = 0
Dim j : j = 0
Dim strServerFilePath : strServerFilePath = ""
Do Until objTextStream.AtEndOfStream
	ReDim Preserve arrListTitles(i)
	arrListTitles(i) = objTextStream.ReadLine
	If strServerFilePath <> "W:\MP3 Music CD'S\" Then
		If InStr(arrListTitles(i), "MP3") <> 0 Then
			strServerFilePath = "W:\MP3 Music CD'S\"
		Else
			strServerFilePath = "W:\FLAC Music CD'S\"
		End If
	End If
	i = i + 1
Loop

objTextStream.Close()

If UBound(arrListTitles) >= 0 Then

	Dim strListTextFile : strListTextFile = strUserFolder & "\Desktop\" & strListTitle & ".txt"

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
	
	' Reading text file and trimming the artist names
	Dim strCurrentTitle, strArtist
	Dim intIndexFilePath, intFolderSlash
	i = LBound(arrListTitles)
	j = LBound(arrListTitles)
	Dim arrListArtists : arrListArtists = Array()
	Dim objExplorer : Set objExplorer = WindowCreate("Sorting List", intWinHeight, intWinWidth)
	HtmlProgress "Reading...", objExplorer, i, UBound(arrListTitles) + 1
	Do While i <> UBound(arrListTitles)
		strCurrentTitle = arrListTitles(i)
		
		' Find the NAS files
		intIndexFilePath = InStr(strCurrentTitle, strServerFilePath)
			
		If intIndexFilePath <> 0 Then
			HtmlProgress "Reading...", objExplorer, i, UBound(arrListTitles) + 1
			strArtist = Right(strCurrentTitle, intIndexFilePath)
			strArtist = Replace(strArtist, strServerFilePath, "")
			strArtist = Trim(strArtist)
			intFolderSlash = InStr(strArtist, "\")
			If intFolderSlash <> 0 Then
				strArtist = Left(strArtist, intFolderSlash - 1)
			End If
			ReDim Preserve arrListArtists(j)
			arrListArtists(j) = strArtist
			j = j + 1
		End If
		i = i + 1
		HtmlProgress "Reading...", objExplorer, i, UBound(arrListTitles) + 1
	Loop
	WindowClose(objExplorer)
	
	If UBound(arrListArtists) >= 0 Then
		' Alphabetically sorting the artist names
		arrListArtists = BubbleSortListProgress(arrListArtists, intWinHeight, intWinWidth)

		' Creating a new text file with the organized format style:
		'	-ARTIST NAME-
		'		-ALBUM NAME-
		'		-ALBUM NAME-
		i = LBound(arrListArtists)
		Set objExplorer = WindowCreate("Sorting List", intWinHeight, intWinWidth)
		HtmlProgress "Writing...", objExplorer, i, UBound(arrListArtists) + 1
		'Write Title to the contents of the text file
		objTextStream.WriteLine strListTitle
		objTextStream.WriteLine "-------------------------" & VbCrLf
		objTextStream.WriteLine "File Listing"
		objTextStream.WriteLine "-------------------------" & VbCrLf
		Do While i <> UBound(arrListArtists)
			HtmlProgress "Writing...", objExplorer, i, UBound(arrListArtists) + 1
			j = 0
			If i = 0 Then
				objTextStream.WriteLine arrListArtists(i)
				Do While j <> UBound(arrListTitles)
					If InStr(arrListTitles(j), arrListArtists(i)) <> 0 Then
						objTextStream.WriteLine vbTab & arrListTitles(j)
					End If
					j = j + 1
				Loop
			ElseIf i <> 0 And arrListArtists(i) <> arrListArtists(i - 1) Then
				objTextStream.WriteLine arrListArtists(i)
				Do While j <> UBound(arrListTitles)
					If InStr(arrListTitles(j), arrListArtists(i)) <> 0 Then
						objTextStream.WriteLine vbTab & arrListTitles(j)
					End If
					j = j + 1
				Loop
			End If
			i = i + 1
			HtmlProgress "Writing...", objExplorer, i, UBound(arrListArtists) + 1
		Loop
	Else
		MsgBox("The selected file contains no music files, Quitting")
	End If
	
	' Clean-up
	WindowClose(objExplorer)
	Set objExplorer = Nothing
Else
	MsgBox("The selected file contains no music files, Quitting")
End If

objTextStream.Close
Set objTextStream = Nothing

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
' Purpose:	Finds whether the existence of a file is absolute
'
' Inputs:	The full path to possible file including the file name
'
' Returns:	Boolean value of determined file existence
'******************************************************************************
Function FileExists(strFilePath)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath) Then
		FileExists = True
	Else
		FileExists = False
	End If
	Set objFSO = Nothing
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
' Purpose:	Finds whether the existence of a file plus extension is absolute
'
' Inputs:	The full path to possible file including the file name
' 			and the file type extension to check for
'
' Returns:	Boolean value of determined file existence
'******************************************************************************
Function IsFileType(strFilePath, strFileExtension)
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim strFileName : strFileName = objFSO.GetBaseName(strFilePath)
	Dim strPathName : strPathName = objFSO.GetParentFolderName(strFilePath)
	Dim strFullPathName : strFullPathName = strPathName & "\" & strFileName & strFileExtension
	Select Case True
		Case objFSO.FileExists(strFullPathName)
			IsFileType = True
		Case Else
			IsFileType = False
	End Select
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
