Const ForReading = 1, ForWriting = 2, ForAppending = 8, blCreateTextFile = False
Dim objOFS, objTextFile
Dim strSerialTextFile, strComputer
Dim objWMIService, objNetwork
Dim colItems
Dim strDelimiters

Set objOFS = CreateObject("Scripting.FileSystemObject")

strComputer = "."
strSerialTextFile = "\\SCHOOL-alt\eXpress\SCHOOL\Utilities\Excel School Devices\SerialNumber.txt"

Set objTextFile = objOFS.OpenTextFile(strSerialTextFile, ForAppending, blCreateTextFile)
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct")
Set objNetwork = CreateObject("WScript.Network")
Set strDelimiters = New RegExp
strDelimiters.Global = True
strDelimiters.IgnoreCase = True
strDelimiters.Pattern = "[\.\\\/]+"

For Each objItem in colItems
	strSerialNumber = (strDelimiters.Replace(objItem.IdentifyingNumber, ""))
	strSerialNumber = (Left(strSerialNumber, 7))
	strComputerName = (Replace(objNetwork.ComputerName, strSerialNumber, ""))
	
	objTextFile.Write(strSerialNumber & vbTab & strComputerName & vbCrLf)
Next

objTextFile.Close

Set objOFS = Nothing
Set objTextFile = Nothing
Set objWMIService = Nothing
Set colItems = Nothing
Set objNetwork = Nothing
Set strDelimiters = Nothing
