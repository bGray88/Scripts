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