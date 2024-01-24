
'******************************************************************************
' Purpose:	Finds the cell with the matching string value
'
' Inputs:	Range of cells to be searched
'
' Returns:	The column number of the cell found to contain match
'******************************************************************************
Function MatchColumnValue(colRange, strMatch)
	If TypeName(colRange) <> "Nothing" Then
		If colRange.Count > 0 Then
			For Each Cell In colRange.Cells
				If StrComp(Cell.Value, strMatch, 1) = 0 Then
					MatchColumnValue = Cell.Column
				End If
			Next
		End If
	End If
End Function
