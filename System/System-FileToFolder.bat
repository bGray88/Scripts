for %%A in (*) do (
	if not "%%~xA"==".bat" (
		mkdir "%%~nA"
		move "%%A" "%CD%\%%~nA\"
	)
)
