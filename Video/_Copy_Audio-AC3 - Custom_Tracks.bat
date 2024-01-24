for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:a:0 ^
-c:a:0 copy ^
"%%~nA - outfile.ac3"
