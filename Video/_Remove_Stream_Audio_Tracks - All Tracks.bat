for %%A in (*.mp4) do ffmpeg -i "%%A" ^
-map 0 ^
-c:v:0 copy ^
"%%~nA - outfile.mp4"