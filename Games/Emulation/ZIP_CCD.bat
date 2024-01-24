for /r %%A in (*.ccd) do (
	if exist "%%~pA*.img" (
		"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nA.7z" "%%~pA*.*"
	)
	if exist "%~dp0%%~nA.7z" (
		del "%%A"
		del "%%~pA*.img"
		del "%%~pA*.sub"
		rmdir "%%~pA"
	)
)