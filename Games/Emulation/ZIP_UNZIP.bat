for /r %%A in (*.7z, *.zip) do (
	if not exist "%~dp0%%~nA" (
		mkdir "%~dp0%%~nA"
	)
	if exist %%A (
		"C:\Program Files\7-Zip\7zG.exe" e "%%A" -o"%~dp0%%~nA"
		del "%%A"
	)
)
