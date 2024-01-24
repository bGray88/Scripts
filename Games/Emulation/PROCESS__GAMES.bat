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
			for /f "tokens=*" %%A in ('dir /b ^"%~dp0^"') do (
				REM Dreamcast rom processing
				if exist "%~dp0%%~nxA.gdi" (
					REM Process Redump ROM
					if exist "%~dp0%%~nxA\%%~nxA.cue" (
						"%~dp0_CHDMAN.exe" createcd -i "%~dp0%%~nxA\%%~nxA.cue" -o "%~dp0%%~nxA\%%~nxA.chd"
						if exist "%~dp0%%~nxA\%%~nxA.chd" (
							del /s /q /f "%~dp0%%~nxA\*.bin"
							del /s /q /f "%~dp0%%~nxA\*.cue"
							"%~dp0_CHDMAN.exe" extractcd -i "%~dp0%%~nxA\%%~nxA.chd" -o "%~dp0%%~nxA\track.gdi" -ob "%~dp0%%~nxA\track.bin"
							if exist "%~dp0%%~nxA\track.gdi" (
								del /s /q /f "%~dp0%%~nxA\*.chd"
								rename "%~dp0%%~nxA\track.gdi" "%%~nxA.gdi"
								REM Process bad GDI
								move /y "%~dp0%%~nxA.gdi" "%~dp0%%~nxA\%%~nxA.gdi"
								pushd %~dp0
								start /wait "" cmd /c cscript Mass-Edit-GDI.vbs
							)
						)
					) else (
						REM Process bad GDI
						move /y "%~dp0%%~nxA.gdi" "%~dp0%%~nxA\%%~nxA.gdi"
						pushd %~dp0
						start /wait "" cmd /c cscript Mass-Edit-GDI.vbs
					)
				) else (
					REM SATURN, SEGA CD, PHILIPS rom processing
					if exist "%~dp0%%~nxA.cue" (
						REM Process bad CUE
						move /y "%~dp0%%~nxA.cue" "%~dp0%%~nxA\%%~nxA.cue"
						pushd %~dp0
						start /wait "" cmd /c cscript Mass-Edit-CUE.vbs
					)
					if exist "%~dp0%%~nxA\%%~nxA.cue" (
						"%~dp0_CHDMAN.exe" createcd -i "%~dp0%%~nxA\%%~nxA.cue" -o "%~dp0%%~nxA\%%~nxA.chd"
						if exist "%~dp0%%~nxA\%%~nxA.chd" (
							del /s /q /f "%~dp0%%~nxA\*.bin"
							del /s /q /f "%~dp0%%~nxA\*.cue"
							"%~dp0_CHDMAN.exe" extractcd -i "%~dp0%%~nxA\%%~nxA.chd" -o "%~dp0%%~nxA\%%~nxA.cue" -ob "%~dp0%%~nxA\%%~nxA.bin"
							if exist "%~dp0%%~nxA\%%~nxA.cue" (
								del /s /q /f "%~dp0%%~nxA\*.chd"
							)
						)
					)
				)
				if exist "%~dp0%%~nxA\" (
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
	)
)