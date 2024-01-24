Option Explicit
Const OverwriteFile = False

Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")

Dim strInitFolder : strInitFolder = ""
Dim strBrowseTitle : strBrowseTitle = ""

' Browse for the picture files
' -----------------------------------------------------------------------------

Do
	MsgBox "Please select the folder containing the Title Pictures"
	Dim strSelectedPicFolder : strSelectedPicFolder = _
		BrowseForFolder(strBrowseTitle, strInitFolder)
	If strSelectedPicFolder = "" Then
		WScript.Quit(0)
	Else
		Dim objSelectedPicFolder : Set objSelectedPicFolder = _
			objFSO.GetFolder(strSelectedPicFolder)
		Dim colPicFiles : Set colPicFiles = objSelectedPicFolder.Files
		Dim intPicColSize : intPicColSize = colPicFiles.Count
		If intPicColSize < 1 Then
			MsgBox "The Folder Selected Is Empty"
			strSelectedPicFolder = ""
		Else
			Dim blPNGExist : blPNGExist = False
			Dim Item
			For Each Item In colPicFiles
				If GetExtension(objFSO.GetFileName(Item)) = "png" Then
					blPNGExist = True
					Exit For
				End If
			Next
			If Not blPNGExist Then
				MsgBox "PNG Files Not Found"
				strSelectedPicFolder = ""
			End If
		End If
		
		Set objSelectedPicFolder = Nothing
	End If
Loop While strSelectedPicFolder = ""

Do
	' Browse for the media files
	' -----------------------------------------------------------------------------

	Do
		MsgBox "Please select the folder containing the Media Files"
		Dim strSelectedMediaFolder : strSelectedMediaFolder = _
			BrowseForFolder(strBrowseTitle, strInitFolder)
		If strSelectedMediaFolder = "" Then
			WScript.Quit(0)
		Else
			Dim objSelectedMediaFolder : Set objSelectedMediaFolder = _
				objFSO.GetFolder(strSelectedMediaFolder)
			Dim colMediaFiles : Set colMediaFiles = objSelectedMediaFolder.Files
			Dim intMedColSize : intMedColSize = colMediaFiles.Count
			If intMedColSize < 1 Then
				MsgBox "The Folder Selected Is Empty"
				strSelectedMediaFolder = ""
			End If
			Set objSelectedMediaFolder = Nothing
		End If
	Loop While strSelectedMediaFolder = ""

	' Process files as necessary
	' -----------------------------------------------------------------------------

	If (intPicColSize > 0) And (intMedColSize > 0) Then
		Dim strUserFolder : strUserFolder = ExpandEnvironmentStrings("%USERPROFILE%")
		Dim strOutputFolder : strOutputFolder = strUserFolder & _
			"\Desktop\" & GetFileFolderName(strSelectedMediaFolder) & "\"

		If Not objFSO.FolderExists(strOutputFolder) Then
			objFSO.CreateFolder(strOutputFolder)
		End If
		
		' Process status window
		' -----------------------------------------------------------------------------
		
		Dim intWinHeight : intWinHeight = 200
		Dim intWinWidth : intWinWidth = 400
		Dim strAppTitle : strAppTitle = "Writing Files"
		Dim strApptask : strAppTask = "Writing..."
		Dim intWriteColSize : intWriteColSize = colMediaFiles.Count
		Dim intWriteCounter : intWriteCounter = 0
		Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
			intWinHeight, intWinWidth)
			
		' Parse filenames as necessary to fit naming convention (## - ##)
		' -----------------------------------------------------------------------------
		
		Dim objMedia
		Dim strMediaFileName
		Dim intNextDash
		Dim intCounter
		Dim intIdxCounter
		Dim strMedFinal
		Dim strPicFileName
		For Each objMedia In colMediaFiles
			strMediaFileName = objFSO.GetFileName(objMedia)
			intNextDash = 0
			intCounter = 0
			intIdxCounter = 1
			Do While intCounter <> 2
				intNextDash = InStr(intIdxCounter + 1, strMediaFileName, "-")
				intCounter = intCounter + 1
				If intNextDash = 0 Then
					strMedFinal = strMediaFileName
					Exit Do
				ElseIf intNextDash > 2 Then
					strMedFinal = Left(strMediaFileName, intNextDash - 2)
				Else
					MsgBox "Invalid Filenames"
					WScript.Quit(0)
				End If
				intIdxCounter = intNextDash
			Loop
			
			' Copy picture files that match the media files in selected folder
			' -----------------------------------------------------------------------------
			
			intWriteCounter = intWriteCounter + 1
			HtmlProgress strAppTask, objExplorer, intWriteCounter, intWriteColSize
			Dim objPicture
			For Each objPicture In colPicFiles
				strPicFileName = objFSO.GetFileName(objPicture)
				If GetExtension(strPicFileName) = "png" Then
					If StrComp(objFSO.GetBaseName(strPicFileName), objFSO.GetBaseName(strMedFinal), 1) = 0 Then
						objFSO.CopyFile objFSO.GetAbsolutePathName(objPicture), strOutputFolder, OverwriteFile
						objFSO.MoveFile strOutputFolder & strPicFileName, _
							strOutputFolder & objFSO.GetBaseName(strMediaFileName) & ".png"
						Exit For
					End If
				End If
			Next
		Next

		WindowClose(objExplorer)
		Set objExplorer = Nothing
	Else
		MsgBox "File does not contain any files to be copied"
	End If
Loop While True

Set objFSO =  Nothing
Set colMediaFiles = Nothing
Set colPicFiles = Nothing

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
