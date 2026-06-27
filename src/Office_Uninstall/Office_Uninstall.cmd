reg add "HKLM\SOFTWARE\Microsoft\Windows Script Host\Settings" /v Enabled /t REG_DWORD /d 1 /f
cscript.exe "%~dp0Office_Uninstall.vbs" /Force

pause
