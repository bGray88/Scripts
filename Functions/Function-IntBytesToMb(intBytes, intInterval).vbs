
'******************************************************************************
' Purpose:	Converts byte value to MegaByte value
'
' Inputs:	An integer value
'
' Returns:	An integer value
'******************************************************************************
Function IntBytesToMb(intBytes)
	Dim intMultiByte : intMultiByte = 1000000
	Dim intMegaBytes : intMegaBytes = CInt(FormatNumber(intBytes) / intMultiByte)
	IntBytesToMb = intMegaBytes
End Function
