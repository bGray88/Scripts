@echo off
for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
	if not "%%A" == "" (
		set true=
		if exist "%~dp0%%~nxA\" (
			set true=1
		)
		if defined true (
			for /f "tokens=*" %%G in ('dir /b ^"%~dp0%%~nxA\%%~nxA.ccd^"') do (
				echo "%~dp0%%~nG\%%~nG"
			)
		)
	)
)