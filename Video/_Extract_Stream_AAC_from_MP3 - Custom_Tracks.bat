for %%A in (*.mp3) do ffmpeg -i "%%~nA.mp3" ^
-map 0:a:0 ^
-ac:a:0 2 -c:a:0 aac -ar:a:0 48000 -b:a 160k ^
"%%~nA - outfile.m4a"
