for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:s:0 ^
-c:s:0 copy ^
"%%~nA.ass"
