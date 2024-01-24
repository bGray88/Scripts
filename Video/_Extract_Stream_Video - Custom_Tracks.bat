for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:v:0 ^
-c:v:0 copy ^
"%%~nA - outfile.mp4"
