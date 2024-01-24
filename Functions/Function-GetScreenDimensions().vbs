
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
