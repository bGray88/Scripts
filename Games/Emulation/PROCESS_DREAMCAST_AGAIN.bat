for /r %%A in (*.7z, *.zip) do (
	if not exist "%~dp0%%~nA" (
		mkdir "%~dp0%%~nA"
	)
	if exist %%A (
		"C:\Program Files\7-Zip\7zG.exe" e "%%A" -o"%~dp0%%~nA"
		del "%%A"
	)
)
for %%A in (*.gdi) do (
	move /y "%~dp0%%~nA.gdi" "%~dp0%%~nA\%%~nA.gdi"
	pushd %~dp0
	start /wait "" cmd /c cscript Mass-Edit-GDI.vbs
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