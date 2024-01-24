for %%A in (*.mkv) do ffmpeg -i "%%~nA.mkv" ^
-map 0:v:0 ^
-map 0:a:0 ^
-map 0:a:1? ^
-map 0:a:2? ^
-map 0:a:3? ^
-map 0:a:4? ^
-map 0:a:5? ^
-map 0:s:0? ^
-map 0:s:1? ^
-video_track_timescale 90000 -ss 00:01:05.50 -to 00:20:43.50 ^
-c:v:0 libx264 ^
-vf fade=t=in:st=66.0:d=0.5,fade=t=out:st=1242.5:d=0.5 ^
-crf:v:0 18 -x264-params level=4.0:mixed-refs=0:vbv-bufsize=25000:vbv-maxrate=20000:rc-lookahead=10:ref=6:bframes=5:b-adapt=2:me=umh:merange=32 ^
-filter:a "afade=t=in:st=66.0:d=0.5, afade=t=out:st=1242.5:d=0.5" ^
-ac:a:0 2 -c:a:0 ac3 -ar:a:0 48000 -b:a:0 640k ^
-ac:a:1 2 -c:a:1 aac -ar:a:1 48000 -b:a:1 160k ^
-ac:a:2 6 -c:a:2 ac3 -ar:a:2 48000 -b:a:2 640k ^
-ac:a:3 2 -c:a:3 aac -ar:a:3 48000 -b:a:3 160k ^
-ac:a:4 6 -c:a:4 ac3 -ar:a:4 48000 -b:a:4 640k ^
-ac:a:5 2 -c:a:5 aac -ar:a:5 48000 -b:a:5 160k ^
-disposition:a:0 default ^
-disposition:a:1 none ^
-disposition:a:2 none ^
-disposition:a:3 none ^
-disposition:a:4 none ^
-disposition:a:5 none ^
-metadata:s:a:0 language=eng -metadata:s:a:0 title="Stereo (HD)" ^
-metadata:s:a:1 language=eng -metadata:s:a:1 title="Stereo" ^
-metadata:s:a:2 language=eng -metadata:s:a:2 title="Stereo (HD)" ^
-metadata:s:a:3 language=eng -metadata:s:a:3 title="Stereo" ^
-metadata:s:a:4 language=eng -metadata:s:a:4 title="Stereo (HD)" ^
-metadata:s:a:5 language=eng -metadata:s:a:5 title="Stereo" ^
-c:s:0 copy ^
-c:s:1 copy ^
-metadata:s:s:0 language=eng -metadata:s:s:0 title="English" ^
-metadata:s:s:1 language=eng -metadata:s:s:1 title="English" ^
"%%~nA - outfile.mkv"

REM -video_track_timescale 90000 -ss 00:01:05.50 -to 00:20:43.50 ^