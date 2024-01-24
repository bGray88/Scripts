for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:a:0 ^
-c:a:0 flac ^
"%%~nA - outfile.flac"
