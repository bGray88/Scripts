for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
	if not "%%A" == "" (
		if exist "%%A" (	
			set "true="
			if "%%~xA" == ".7z" set "true=1"
			if "%%~xA" == ".zip" set "true=1"
			if defined true (
				if not exist "%~dp0%%~nA\" (
					mkdir "%~dp0%%~nA\"
				)
				if exist "%~dp0%%~nA\" (
					REM Unzip compressed file
					"C:\Program Files\7-Zip\7zG.exe" e "%%A" -o"%~dp0%%~nA\"
					REM Check for non-empty folder
					(dir /b "%~dp0%%~nA\" | findstr .) > nul && (
						del /s /q /f "%%A"
					)
				)
			)
		)
	)
)