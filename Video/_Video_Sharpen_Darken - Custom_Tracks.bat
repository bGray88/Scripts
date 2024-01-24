for %%A in (*.mkv) do ffmpeg -i "%%~nA.mkv" ^
-map 0 ^
-filter_complex "eq=contrast=1.05:saturation=0.9, unsharp=7:7:1.5:7:7:1.5" ^
-c:v libx264 -crf 18 -preset slow ^
-c:a copy ^
-c:s copy ^
"%%~nA - outfile.mkv"

REM -filter_complex "eq=saturation=0.9, unsharp=7:7:1.5:5:5:0" ^
REM ===Audio-Re-encode===
REM -ac:a:0 2 -c:a:0 ac3 -ar:a:0 48000 -b:a:0 640k ^
REM -ac:a:1 2 -c:a:1 aac -ar:a:1 48000 -b:a:1 160k ^
REM ===Fade-Audio===
REM -filter:a "afade=t=in:st=0.25:d=0.25, afade=t=out:st=61.75:d=0.25" ^

REM ===Unsharp===
REM ===No Filter===
REM ===Light===
REM -filter_complex "unsharp=3:3:0.5:3:3:0.5"
REM ===Medium===
REM -filter_complex "unsharp=5:5:1:5:5:1"
REM ===Hard===
REM -filter_complex "unsharp=7:7:1.5:7:7:1.5"
REM ===With Filter===
REM ===Light===
REM -filter_complex "eq=contrast=1.03:saturation=0.95, "unsharp=3:3:0.5:3:3:0.5"
REM ===Medium===
REM -filter_complex "eq=contrast=1.05:saturation=0.90, "unsharp=5:5:1:5:5:1"
REM ===Hard===
REM -filter_complex "eq=contrast=1.07:saturation=0.85, unsharp=7:7:1.5:7:7:1.5"

REM ===Smart-Blur===
REM ===Sharpen-Light===
REM smartblur=lr=1.5:ls=-0.25:lt=-3.5:cr=0.75:cs=0.25:ct=0.5
REM ===Sharpen-Medium===
REM smartblur=lr=2.0:ls=-0.35:lt=-4.0:cr=0.85:cs=0.35:ct=1.5
REM ===Sharpen-Hard===
REM smartblur=lr=2.5:ls=-0.45:lt=-4.5:cr=0.95:cs=0.45:ct=2.0
REM ===Soften-Light===
REM smartblur=lr=1.5:ls=0.25:lt=3.5:cr=0.75:cs=0.25:ct=0.5
REM ===Soften-Medium===
REM smartblur=lr=2.0:ls=0.35:lt=4.0:cr=0.85:cs=0.35:ct=1.5
REM ===Soften-Hard===
REM smartblur=lr=2.5:ls=0.45:lt=4.5:cr=0.95:cs=0.45:ct=2.0
