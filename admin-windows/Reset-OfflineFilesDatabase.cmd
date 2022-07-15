@echo off
cls
REM --- Get admin rights
fsutil dirty query %SystemDrive% >NUL && set admin=true
if NOT "%admin%"=="true" (
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\RunAsAdmin.vbs"
	echo UAC.ShellExecute "%~0", "%~1 %~2 %~3 %~4 %~5 %~6 %~7 %~8 %~9", "", "runas", 1 >> "%temp%\RunAsAdmin.vbs"
	"%temp%\RunAsAdmin.vbs"
	del /Q /S "%temp%\RunAsAdmin.vbs"
	goto :EOF
)


ECHO Danger! All files where not synched to server will be lost!
ping 172.0.0.1 -n 10 >nul 2>&1


REG ADD "HKLM\System\CurrentControlSet\Services\CSC\Parameters" /v FormatDatabase /t REG_DWORD /d 1 /f

ECHO Rebooting in 60s
ping 172.0.0.1 -n 3 >nul 2>&1

shutdown -r -t 60
