@echo off
start /b "" %~dp0\edit_online_OFF.vbs
start "" "%~dp0..\steam.exe" -silent -nochatui -nofriendsui