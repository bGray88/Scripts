@echo off

set VBoxManage="C:\Program Files\oracle\VirtualBox\VBoxManage.exe"

::%VBoxManage% setextradata "WIN7_32BIT" GUI/Seamless on
%VBoxManage% startvm "WIN7_32BIT" -type GUI
timeout /t 30 /nobreak
%VBoxManage% guestcontrol "WIN7_32BIT" execute --image "C:\GAMES\TRIADA\Freedom Fighters\Freedom.exe" --username Kepler444 --wait-exit
::%VBoxManage% controlvm "WIN7_32BIT" acpipowerbutton