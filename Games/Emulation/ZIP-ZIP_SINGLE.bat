set "fileExt=.gba"
for %%A in (*) do (
	if exist "%~dp0%%~nA%fileExt%" (
		"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nA.7z" "%~dp0%%~nA%fileExt%"
		if exist "%~dp0%%~nA.7z" (
			del "%~dp0%%~nA%fileExt%"
		)
	)
)
