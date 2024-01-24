for /D %%F in (*) do (
	for /F "tokens=*" %%A in ('Dir /b "%%~fF\*.mp4"') do (
		ffmpeg -i "%%~fF\%%~nA.mp4" ^
		-map 0:v:0 ^
		-map 0:a:0? ^
		-map 0:a:1? ^
		-map 0:a:2? ^
		-map 0:a:3? ^
		-map 0:a:4? ^
		-map 0:a:5? ^
		-map 0:s:0? ^
		-map 0:s:1? ^
		-map_metadata 0 ^
		-c:v:0 copy ^
		-c:a:0 copy ^
		-c:a:1 copy ^
		-c:a:2 copy ^
		-c:a:3 copy ^
		-c:a:4 copy ^
		-c:a:5 copy ^
		-disposition:a:0 default ^
		-disposition:a:1 none ^
		-disposition:a:2 none ^
		-disposition:a:3 none ^
		-disposition:a:4 none ^
		-disposition:a:5 none ^
		-metadata title="%%~F - %%~nA" ^
		-c:s:0 copy ^
		-c:s:1 copy ^
		"%%~fF\%%~F - %%~nA.mp4"
		rename "%%~fF\%%~nA.mp4" "_OLD-%%~nA.mp4"
	)
)
