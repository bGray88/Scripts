for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0 ^
-filter_complex "yadif, scale=1440x1080" ^
-c:v libx264 -crf 18 -preset slow ^
-c:a copy ^
-c:s copy ^
"%%~nA - outfile.mp4"