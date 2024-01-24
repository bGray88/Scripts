
'******************************************************************************
' Purpose:	Accepts text input in text box
'
' Inputs:	Prompt for specific input, Boolean value indicating
'			whether empty values should be allowed
'
' Returns:	Returns path or empty value as string
'******************************************************************************
Function AnswerBox(strPrompt)
	Dim strAnswer
	Do While True
		strAnswer = InputBox(strPrompt)
		Select Case True
			Case IsEmpty(strAnswer) 'Cancel has been pressed
				AnswerBox = ""
				Exit Do
			Case "" = Trim(strAnswer) 'Empty answer has been entered
				AnswerBox = ""
				Exit Do
			Case Else 'Valid answer has been entered
				AnswerBox = strAnswer
				Exit Do
			End Select
	Loop
End Function
