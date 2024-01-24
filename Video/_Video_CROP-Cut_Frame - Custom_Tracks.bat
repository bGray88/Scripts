for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0 ^
-vf crop=716:476:0:4,scale=720:480 ^
-c:v:0 libx264 -crf 18 -x264-params level=4.0:mixed-refs=0:vbv-bufsize=25000:vbv-maxrate=20000:rc-lookahead=10:ref=6:bframes=5:b-adapt=2:me=umh:merange=32 ^
-c:a:0 copy ^
-c:s:0 copy ^
"%%~nA - outfile.mp4"