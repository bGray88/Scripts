
'**********************************************************************************************
' Purpose:	To read data from an XML and store the data in an ArrayList
'
' Inputs:	The path including filename and extension, The XML tag to search for
'
' Returns:	An ArrayList Including any found data
'**********************************************************************************************
Function GetXmlNameList(strXmlPath, strTagName, intWinHeight, intWinWidth)
	Dim objXMLDoc : Set objXMLDoc = CreateObject("Microsoft.XMLDOM") 
	objXMLDoc.async = False
	objXMLDoc.load(strXMLPath)
	
	Dim objRoot : Set objRoot = objXMLDoc.DocumentElement
	Dim objParseErr : Set objParseErr = objXMLDoc.ParseError
	If objParseErr.ErrorCode <> 0 Then
		MsgBox "Error: Line " & objParseErr.Line & VbCrLf & _
			objParseErr.Reason, 16, "Warning"
		WScript.Quit(1)
	End If
	Dim objNodeList : Set objNodeList = objRoot.getElementsByTagName(strTagName)
	
	Dim Elem
	Dim i : i = 0
	Dim arrListNames : arrListNames = Array()
	Dim strAppTitle : strAppTitle = "Reading..."
	Dim strAppTask : strAppTask = "Reading from XML File"
	'Dim objExplorer : Set objExplorer = WindowCreate(strAppTitle, _
	'	intWinHeight, intWinWidth)
	'HtmlProgress strAppTask, objExplorer, i, objNodeList.Length
	For Each Elem In objNodeList
		ReDim Preserve arrListNames(i)
		arrListNames(i) = objNodeList.item(i).Text
		i = i + 1
	'	HtmlProgress strAppTask, objExplorer, i, objNodeList.Length
	Next	
	
	'WindowClose(objExplorer)
	GetXmlNameList = arrListNames
End Function
