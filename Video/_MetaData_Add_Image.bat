for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" -i "%%~nA.jpg" ^
-map 1 -map 0 ^
-c copy ^
-disposition:0 attached_pic ^
"%%~nA - outfile.mp4"