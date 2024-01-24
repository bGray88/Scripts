for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:a:1 ^
-c:a:0 copy ^
"%%~nA - outfile.m4a"