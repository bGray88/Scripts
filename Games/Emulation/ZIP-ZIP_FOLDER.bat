for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
	if not "%%A" == "" (
		set "true="
		if exist "%~dp0%%~nxA\" (
			set "true=1"
		)
		if defined true (
			for /f "tokens=*" %%G in ('dir /b ^"%~dp0%%~nxA\%%~nxA.ccd^"') do (
				if exist "%~dp0%%~nG\%%~nG.ccd" (
					del /s /q /f "%~dp0%%~nG\*.bin"
					del /s /q /f "%~dp0%%~nG\*.cue"
				)
			)
			REM Check for non-empty folder
			(dir /b "%~dp0%%~nxA\" | findstr .) > nul && (
				REM 7z folder contents
				"C:\Program Files\7-Zip\7zG.exe" a -t7z -mx9 -m0=lzma2 -md=128m -mfb=128 -ms=on -mmt2 "%~dp0%%~nxA.7z" "%~dp0%%~nxA\*"
				if exist "%~dp0%%~nxA.7z" (
					del /s /q /f "%~dp0%%~nxA\*"
					rd "%~dp0%%~nxA\"
				)
			)
		)
	)
)