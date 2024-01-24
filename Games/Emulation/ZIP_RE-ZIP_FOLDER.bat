for %%A in (*.7z, *.zip) do (
	if exist %%A (
		"C:\Program Files\7-Zip\7zG.exe" e "%%A" -o"%~dp0%%~nA"
		del "%%A"
	)
)
for /d %%A in (*) do (
	"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nA.7z" "%~dp0%%~nA\*.*"
	if exist "%~dp0%%~nA.7z" (
		del /f /q "%~dp0%%~nA\*.*"
		rmdir "%~dp0%%~nA"
	)
)