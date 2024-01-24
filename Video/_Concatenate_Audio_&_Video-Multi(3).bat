for %%A in (*.mp4) do (
	if exist "%~dp01.mp4" (
		if not "%%A"=="1.mp4" (
			if not "%%A"=="3.mp4" (
				rename "%%~A" "2.mp4"
				ffmpeg -f concat -safe 0 -i "list(3).txt" ^
				-map 0:v:0 ^
				-map 0:a:0? ^
				-map 0:a:1? ^
				-map 0:a:2? ^
				-map 0:a:3? ^
				-map 0:a:4? ^
				-map 0:a:5? ^
				-map 0:s:0? ^
				-map 0:s:1? ^
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
				-metadata:s:a:0 language=eng -metadata:s:a:0 title="Stereo (AC3)" ^
				-metadata:s:a:1 language=eng -metadata:s:a:1 title="Stereo (AAC)" ^
				-metadata:s:a:2 language=eng -metadata:s:a:2 title="Director Commentary" ^
				-metadata:s:a:3 language=eng -metadata:s:a:3 title="Director Commentary" ^
				-metadata:s:a:4 language=eng -metadata:s:a:4 title="Director Commentary" ^
				-metadata:s:a:5 language=eng -metadata:s:a:5 title="Director Commentary" ^
				-c:s:0 copy ^
				-c:s:1 copy ^
				-metadata:s:s:0 language=eng -metadata:s:s:0 title="English" ^
				-metadata:s:s:1 language=eng -metadata:s:s:1 title="English" ^
				"%%~nA - outfile.mp4"
				rename "2.mp4" "%%~nA - OLD.mp4"
			)
		)
	)
)