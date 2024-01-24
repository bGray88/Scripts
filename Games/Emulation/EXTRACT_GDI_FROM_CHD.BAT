for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
	if not "%%A" == "" (
		if exist "%%A" (
			set "true="
			if exist "%~dp0%%~nxA\%%~nxA.chd" set "true=1"
			if defined true (
				if exist "%~dp0%%~nxA\%%~nxA.chd" (
					"%~dp0_CHDMAN.exe" extractcd -i "%~dp0%%~nxA\%%~nxA.chd" -o "%~dp0%%~nxA\track.gdi" -ob "%~dp0%%~nxA\track.bin"
					if exist "%~dp0%%~nxA\track.gdi" (
						rename "%~dp0%%~nxA\track.gdi" "%%~nxA.gdi"
						del /s /q /f "%~dp0%%~nxA\*.chd"
					)
				)
			)
		)
	)
)