taskkill /f /im explorer.exe
cd "C:\Users\%username%\AppData\Local\Microsoft\Windows\Explorer"
del /f /q *.db
start "C:\WINDOWS" explorer.exe