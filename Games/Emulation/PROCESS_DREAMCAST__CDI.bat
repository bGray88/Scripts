for /r %%A in (*.cue) do (
	if exist "%~dp0%%~nA\%%~nA.cue" (
		"%~dp0CHDMAN.EXE" createcd -i "%~dp0%%~nA\%%~nA.cue" -o "%~dp0%%~nA\%%~nA.chd"
	)
	if exist "%~dp0%%~nA\%%~nA.chd" (
		del "%~dp0%%~nA\%%~nA*.bin"
		del "%~dp0%%~nA\%%~nA*.cue"
	)
)
for /r %%A in (*.chd) do (
	if exist "%%A" (
		"%~dp0CHDMAN.EXE" extractcd -i "%%A" -o "%~dp0%%~nA\track.gdi" -ob "%~dp0%%~nA\track.bin"
	)
	if exist "%~dp0%%~nA\track.gdi" (
		rename "%~dp0%%~nA\track.gdi" "%%~nA.gdi"
		del "%%A"
		move /y "%~dp0%%~nA.gdi" "%~dp0%%~nA\%%~nA.gdi"
		pushd %~dp0
		start /wait "" cmd /c cscript Mass-Edit-GDI.vbs
	)
)
for /r %%A in (*.gdi) do (
	if exist "%%~pA*.raw" (
		"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nA.7z" "%%~pA*.*"
	)
	if exist "%~dp0%%~nA.7z" (
		del "%%A"
		del "%%~pA*.bin"
		del "%%~pA*.raw"
		rmdir "%%~pA"
	)
)