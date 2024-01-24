for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-map 0:v:0 ^
-map 0:a:0? ^
-map 0:a:1? ^
-map 0:a:2? ^
-map 0:a:3? ^
-map 0:a:4? ^
-map 0:s:0? ^
-c:v:0 copy ^
-c:a:0 copy ^
-c:a:1 copy ^
-c:a:2 copy ^
-c:a:3 copy ^
-c:a:4 copy ^
-disposition:a:0 default ^
-disposition:a:1 none ^
-disposition:a:2 none ^
-disposition:a:3 none ^
-disposition:a:4 none ^
-metadata:s:a:0 language=eng -metadata:s:a:0 title="Stereo" ^
-metadata:s:a:1 language=eng -metadata:s:a:1 title="Stereo" ^
-metadata:s:a:2 language=eng -metadata:s:a:2 title="Director Commentary" ^
-metadata:s:a:3 language=eng -metadata:s:a:3 title="Director Commentary" ^
-metadata:s:a:4 language=eng -metadata:s:a:4 title="Director Commentary" ^
-c:s:0 copy ^
-metadata:s:s:0 language=eng -metadata:s:s:0 title="English" ^
"%%~nA - outfile.mp4"
