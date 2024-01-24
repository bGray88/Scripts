for %%A in (*.ac3) do ffmpeg -i "%%~nA.ac3" ^
-map 0:a:0 ^
-ac:a:0 6 -c:a:0 ac3 -ar:a:0 48000 -b:a:0 640k ^
"%%~nA - outfile.ac3"
