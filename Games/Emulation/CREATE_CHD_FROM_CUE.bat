for /r %%A in (*.cue) do (
	if exist "%~dp0%%~nA.cue" (
		move /y "%~dp0%%~nA.cue" "%~dp0%%~nA\%%~nA.cue"
	)
	if exist "%~dp0%%~nA\%%~nA.cue" (
		"%~dp0CHDMAN.EXE" createcd -i "%~dp0%%~nA\%%~nA.cue" -o "%~dp0%%~nA\%%~nA.chd"
	)
	if exist "%~dp0%%~nA\%%~nA.chd" (
		del "%~dp0%%~nA\%%~nA*.bin"
		del "%~dp0%%~nA\%%~nA*.cue"
	)
)