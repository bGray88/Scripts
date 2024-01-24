set arg1=%1
set arg2=%2
set arg3=%3

"C:\Program Files\Java\jdk1.8.0_121\bin\javapackager.exe" -deploy ^
-BsystemWide=true ^
-native exe ^
-srcfiles %arg1% ^
-outdir "%UserProfile%\Desktop\%arg3%" ^
-outfile %arg3% ^
-appclass %arg2% ^
-name %arg3% ^
