
'******************************************************************************
' Purpose:	Automatically gather error info
'
' Inputs:	None
'******************************************************************************
Sub DisplayErrorInfo()
    WScript.Echo "Error:      : " & Err
    WScript.Echo "Error (hex) : &H" & Hex(Err)
    WScript.Echo "Source      : " & Err.Source
    WScript.Echo "Description : " & Err.Description
    Err.Clear
End Sub
