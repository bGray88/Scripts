
'******************************************************************************
' Purpose:	Counts the lines before EOF in text file
'
' Inputs:	The full path including file name of text file
'
' Returns:	Integer value of lines counted
'******************************************************************************
Function CountLinesExcel(objExcelWorksheet, intStartRow, intActiveCellColumn)
	Dim intCurrentRow : intCurrentRow = intStartRow
	Dim intRowCount : intRowCount = 0
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Do Until objExcelWorksheet.Cells(intCurrentRow, intActiveCellColumn).Value = ""
		intRowCount = intRowCount + 1
		intCurrentRow = intCurrentRow + 1
	Loop
	CountLinesExcel = intRowCount
	Set objFSO = Nothing
End Function
