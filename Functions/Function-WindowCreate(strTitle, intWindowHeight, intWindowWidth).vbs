
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
