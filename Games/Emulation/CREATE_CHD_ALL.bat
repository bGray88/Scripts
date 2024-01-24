for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
	if not "%%A" == "" (
		if exist "%%A" (
			set "true="
			if exist "%~dp0%%~nxA\%%~nxA.cue" set "true=1"
			if exist "%~dp0%%~nxA\%%~nxA.gdi" set "true=1"
			if defined true (
				REM SATURN, SEGA CD, PHILIPS rom processing
				if exist "%~dp0%%~nxA.cue" (
					REM Process bad CUE
					move /y "%~dp0%%~nxA.cue" "%~dp0%%~nxA\%%~nxA.cue"
					pushd %~dp0
					start /wait "" cmd /c cscript _Mass-Edit-CUE.vbs
				)
				if exist "%~dp0%%~nxA\%%~nxA.cue" (
					"%~dp0_CHDMAN.exe" createcd -i "%~dp0%%~nxA\%%~nxA.cue" -o "%~dp0%%~nxA\%%~nxA.chd"
				)
				REM Dreamcast rom processing
				if exist "%~dp0%%~nxA.gdi" (
					REM Process bad GDI
					move /y "%~dp0%%~nxA.gdi" "%~dp0%%~nxA\%%~nxA.gdi"
					pushd %~dp0
					start /wait "" cmd /c cscript _Mass-Edit-GDI.vbs
				)
				if exist "%~dp0%%~nxA\%%~nxA.gdi" (
					"%~dp0_CHDMAN.exe" createcd -i "%~dp0%%~nxA\%%~nxA.gdi" -o "%~dp0%%~nxA\%%~nxA.chd"
				)
				if exist "%~dp0%%~nxA\%%~nxA.chd" (
					del /s /q /f "%~dp0%%~nxA\*.gdi"
					del /s /q /f "%~dp0%%~nxA\*.raw"
					del /s /q /f "%~dp0%%~nxA\*.bin"
					del /s /q /f "%~dp0%%~nxA\*.cue"
				)
			)
		)
	)
)