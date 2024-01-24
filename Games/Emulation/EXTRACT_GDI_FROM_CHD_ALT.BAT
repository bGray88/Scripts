for /r %%A in (*.chd) do (
	if exist "%%A" (
		"%~dp0CHDMAN.EXE" extractcd -i "%%A" -o "%~dp0%%~nA\track.gdi" -ob "%~dp0%%~nA\track.bin"
	)
	if exist "%~dp0%%~nA\track.gdi" (
		rename "%~dp0%%~nA\track.gdi" "%%~nA.gdi"
		del "%%A"
	)
)