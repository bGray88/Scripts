for %%A in (*.mp4) do ffmpeg -i "%%~nA.mp4" ^
-filter_complex "[0:v]setpts=0.5*PTS[v];[0:a]atempo=2.0[a]" -map "[v]" -map "[a]" ^
"%%~nA - outfile.mp4"