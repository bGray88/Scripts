for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" -sub_charenc ISO-8859-1 -f srt -i "%%~nA.srt" ^
-map 0:v:0 ^
-map 0:a:0? ^
-map 0:a:1? ^
-map 0:a:2? ^
-map 0:a:3? ^
-map 0:a:4? ^
-map 0:a:5? ^
-map 0:a:6? ^
-map 0:a:7? ^
-map 1:s:0? ^
-map 0:s:0? ^
-map 0:s:1? ^
-map 0:s:2? ^
-c:v:0 copy ^
-c:a:0 copy ^
-c:a:1 copy ^
-c:a:2 copy ^
-c:a:3 copy ^
-c:a:4 copy ^
-c:a:5 copy ^
-c:a:6 copy ^
-c:a:7 copy ^
-disposition:a:0 default ^
-disposition:a:1 none ^
-disposition:a:2 none ^
-disposition:a:3 none ^
-disposition:a:4 none ^
-disposition:a:5 none ^
-disposition:a:6 none ^
-disposition:a:7 none ^
-metadata:s:a:0 language=eng -metadata:s:a:0 title="Surround (AC3)" ^
-metadata:s:a:1 language=eng -metadata:s:a:1 title="Stereo (AAC)" ^
-metadata:s:a:2 language=jpn -metadata:s:a:2 title="Surround (AC3)" ^
-metadata:s:a:3 language=jpn -metadata:s:a:3 title="Stereo (AAC)" ^
-metadata:s:a:4 language=eng -metadata:s:a:4 title="Director Commentary" ^
-metadata:s:a:5 language=eng -metadata:s:a:5 title="Director Commentary" ^
-metadata:s:a:6 language=eng -metadata:s:a:6 title="Director Commentary" ^
-metadata:s:a:7 language=eng -metadata:s:a:7 title="Director Commentary" ^
-c:s:0 mov_text ^
-c:s:1 copy ^
-c:s:2 copy ^
-c:s:3 copy ^
-metadata:s:s:0 language=eng -metadata:s:s:0 title="English (SRT)" ^
-metadata:s:s:1 language=eng -metadata:s:s:1 title="English (SUB)" ^
-metadata:s:s:2 language=eng -metadata:s:s:2 title="Japanese (SRT)" ^
-metadata:s:s:3 language=eng -metadata:s:s:3 title="Japanese (SUB)" ^
"%%~nA - outfile.mp4"
