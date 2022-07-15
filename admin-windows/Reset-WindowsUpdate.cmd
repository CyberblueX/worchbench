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


net stop wuauserv
ping 127.0.0.1 -n 4 >nul

net stop wuauserv
ping 127.0.0.1 -n 4 >nul

rd /s /q %windir%\SoftwareDistribution

ping 127.0.0.1 -n 4 >nul
net start wuauserv

ping 127.0.0.1 -n 4 >nul

wuauclt /detectnow
wuauclt /reportnow

ping 127.0.0.1 -n 4 >nul
exit
