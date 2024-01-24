for %%A in (*.mka) do ffmpeg -i "%%~nA.mka" ^
-map 0:a:0 ^
-c:a:0 copy ^
"%%~nA - outfile.pcm"
