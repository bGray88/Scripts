set "fileExt=.nds"
for %%A in (*.7z, *.zip) do (
	set "true="
	if exist %%A (
		if "%%~xA" == ".7z" set "true=1"
		if "%%~xA" == ".zip" set "true=1"
		if defined true (
			"C:\Program Files\7-Zip\7zG.exe" e "%%A" -o"%~dp0"
			if exist "%~dp0%%~nA%fileExt%" (
				del "%%A"
			)
		)
	)
)
