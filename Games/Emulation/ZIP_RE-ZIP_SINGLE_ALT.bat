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
				"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nA.7z" "%~dp0%%~nA%fileExt%"
				if exist "%~dp0%%~nA.7z" (
					del "%~dp0%%~nA%fileExt%"
				)
			)
		)
	)
)
