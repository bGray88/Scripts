for /r %%A in (*.gdi) do (
	if exist "%~dp0%%~nA.gdi" (
		move /y "%~dp0%%~nA.gdi" "%~dp0%%~nA\%%~nA.gdi"
	)
	if exist "%~dp0%%~nA\%%~nA.gdi" (
		"%~dp0CHDMAN.EXE" createcd -i "%~dp0%%~nA\%%~nA.gdi" -o "%~dp0%%~nA\%%~nA.chd"
	)
	if exist "%~dp0%%~nA\%%~nA.chd" (
		del "%~dp0%%~nA\%%~nA*.bin"
		del "%~dp0%%~nA\%%~nA*.gdi"
	)
)