for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:v:0 ^
-map 0:a:0 ^
-map 0:a:1? ^
-map 0:s:0? ^
-map 0:s:1? ^
-c:v:0 libx264 ^
-vf crop=952:712:4:4, scale=960:720 ^
-crf:v:0 23 -x264-params level=4.0:mixed-refs=0:vbv-bufsize=25000:vbv-maxrate=20000:rc-lookahead=10:ref=6:bframes=5:b-adapt=2:me=umh:merange=32 ^
-ac:a:0 2 -c:a:0 ac3 -ar:a:0 48000 -b:a:0 640k ^
-ac:a:1 2 -c:a:1 aac -ar:a:1 48000 -b:a:1 160k ^
-disposition:a:0 default ^
-disposition:a:1 none ^
-metadata:s:a:0 language=eng -metadata:s:a:0 title="Stereo (AC3)" ^
-metadata:s:a:1 language=eng -metadata:s:a:1 title="Stereo (AAC)" ^
-c:s:0 copy ^
-c:s:1 copy ^
-metadata:s:s:0 language=eng -metadata:s:s:0 title="English" ^
-metadata:s:s:1 language=eng -metadata:s:s:1 title="English" ^
"%%~nA - outfile.mp4"