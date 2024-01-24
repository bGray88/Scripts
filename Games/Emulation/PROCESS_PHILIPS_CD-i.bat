for /r %%A in (*.7z, *.zip) do (
	if not exist "%~dp0%%~nA" (
		mkdir "%~dp0%%~nA"
	)
	if exist %%A (
		"C:\Program Files\7-Zip\7zG.exe" e "%%A" -o"%~dp0%%~nA"
		del "%%A"
	)
)
for /r %%A in (*.cue) do (
	if exist "%~dp0%%~nA\%%~nA.cue" (
		pushd %~dp0
		start /wait "" cmd /c cscript Mass-Edit-CUE.vbs
		"%~dp0CHDMAN.EXE" createcd -i "%~dp0%%~nA\%%~nA.cue" -o "%~dp0%%~nA\%%~nA.chd"
	)
	if exist "%~dp0%%~nA\%%~nA.chd" (
		del "%~dp0%%~nA\%%~nA*.bin"
		del "%~dp0%%~nA\%%~nA*.cue"
	)
)
for /r %%A in (*.chd) do (
	if exist "%%A" (
		"%~dp0CHDMAN.EXE" extractcd -i "%%A" -o "%~dp0%%~nA\%%~nA.cue" -ob "%~dp0%%~nA\%%~nA.bin"
	)
	if exist "%~dp0%%~nA\%%~nA.cue" (
		del "%%A"
	)
)
for /r %%A in (*.cue) do (
	if exist "%%~pA*.bin" (
		"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nA.7z" "%%~pA*.*"
	)
	if exist "%~dp0%%~nA.7z" (
		del "%%A"
		del "%%~pA*.bin"
		rmdir "%%~pA"
	)
)