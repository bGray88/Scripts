for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map_metadata -1 ^
-c copy ^
-map 0 ^
"%%~nA - outfile.mp4"