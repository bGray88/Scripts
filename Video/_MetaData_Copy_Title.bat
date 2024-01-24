for /f "delims=" %%A in ('dir /b *.mp4') do ( 
	ffmpeg -i "%%~nA.mp4" ^
	-map 0 ^
	-c copy ^
	-metadata title="%%~nA" ^
	"%%~nA - outfile.mp4"
	rename "%%~fA" "_OLD-%%~nA.mp4"
)