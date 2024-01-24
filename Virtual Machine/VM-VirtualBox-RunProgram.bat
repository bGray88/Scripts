@echo off

set MultiRes="C:\Resources\Display\MultiRes\MultiRes.exe"
set VBoxManage="C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"

::%VBoxManage% setextradata "WINXP_32BIT" GUI/Seamless on
%VBoxManage% setextradata "WINXP_32BIT" "CustomVideoMode1" 800x600x32"
%VBoxManage% startvm "WINXP_32BIT"
timeout /t 20 /nobreak
%VBoxManage% guestcontrol "WINXP_32BIT" start %MultiRes% --username Otto -- /800,600,32
timeout /t 2 /nobreak
%VBoxManage% guestcontrol "WINXP_32BIT" run "C:\WINDOWS\System32\cmd.exe" --username Otto -- "cmd.exe /c cd FOLDER_PATH & EXE_PATH"
%VBoxManage% guestcontrol "WINXP_32BIT" start %MultiRes% --username Otto -- /restore
timeout /t 2 /nobreak
%VBoxManage% controlvm "WINXP_32BIT" acpipowerbutton
