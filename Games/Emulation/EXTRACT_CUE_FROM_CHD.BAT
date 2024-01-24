for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
	if not "%%A" == "" (
		if exist "%%A" (
			set "true="
			if exist "%~dp0%%~nxA\%%~nxA.chd" set "true=1"
			if defined true (
				if exist "%~dp0%%~nxA\%%~nxA.chd" (
					"%~dp0_CHDMAN.exe" extractcd -i "%~dp0%%~nxA\%%~nxA.chd" -o "%~dp0%%~nxA\%%~nxA.cue" -ob "%~dp0%%~nxA\%%~nxA.bin"
					if exist "%~dp0%%~nxA\%%~nxA.cue" (
						del /s /q /f "%~dp0%%~nxA\*.chd"
					)
				)
			)
		)
	)
)