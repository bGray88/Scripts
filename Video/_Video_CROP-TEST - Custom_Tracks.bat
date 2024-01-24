for %%A in (*.mp4) do "C:\ffmpeg\bin\ffplay" -i "%%~nA.mp4" ^
-vf crop=1916:1068:2:6,scale=1920:1080