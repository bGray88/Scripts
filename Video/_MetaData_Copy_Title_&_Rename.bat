for /D %%F in (*) do (
	for /F "tokens=*" %%A in ('Dir /b "%%~fF\*.mp4"') do (
		ffmpeg -i "%%~fF\%%~nA.mp4" ^
		-map 0 ^
		-c copy ^
		-metadata title="%%~F - %%~nA" ^
		"%%~fF\%%~F - %%~nA.mp4"
		rename "%%~fF\%%~nA.mp4" "_OLD-%%~nA.mp4"
	)
)
