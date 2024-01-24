for /r %%A in (*.chd) do (
	if exist "%%A" (
		"%~dp0CHDMAN.EXE" extractcd -i "%%A" -o "%~dp0%%~nA\%%~nA.cue" -ob "%~dp0%%~nA\%%~nA.bin"
	)
	if exist "%~dp0%%~nA\%%~nA.cue" (
		del "%%A"
	)
)