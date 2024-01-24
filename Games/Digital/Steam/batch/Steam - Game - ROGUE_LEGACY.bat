@echo off

start /b "" %~dp0\edit_online_OFF.vbs
start "" "%~dp0..\steam.exe" -silent -nochatui -nofriendsui -applaunch 241600 -fullscreen
@timeout /t 10 /nobreak

:SteamStartCheck
tasklist /FI "IMAGENAME eq steam.exe" /NH | find /I /N "steam.exe" >NUL
if not "%ERRORLEVEL%"=="0" goto SteamCheckFail

:UpdateLoopStart
tasklist /FI "IMAGENAME eq RogueLegacy.exe" /NH | find /I /N "RogueLegacy.exe" >NUL
if "%ERRORLEVEL%"=="0" goto UpdateLoopEnd
goto UpdateLoopStart
:UpdateLoopEnd
goto GameLoopStart

:GameLoopStart
tasklist /FI "IMAGENAME eq RogueLegacy.exe" /NH | find /I /N "RogueLegacy.exe" >NUL
if not "%ERRORLEVEL%"=="0" goto GameLoopEnd
goto GameLoopStart
:GameLoopEnd
goto EndSteamProcess

:SteamCheckFail
call :SetError 1
echo ERROR %ERRORLEVEL% - Steam isn't running
@timeout /t 10 /nobreak
goto :eof

:SetError
exit /B %~1

:EndSteamProcess
"%~dp0..\steam.exe" -shutdown
goto GameEndSuccess

:GameEndSuccess
exit /B 0
