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


net stop bits
net stop wuauserv
net stop appidsvc
net stop cryptsvc

ping 127.0.0.1 -n 4 >nul

net stop bits
net stop wuauserv
net stop appidsvc
net stop cryptsvc

pause

Del "%ALLUSERSPROFILE%\ApplicationData\Microsoft\Network\Downloader\qmgr*.dat"

DEL /Q /F "%systemroot%\SoftwareDistribution"
DEL /Q /F "%systemroot%\system32\catroot2"

sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)

sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)

pause

cd /d %windir%\system32

regsvr32.exe atl.dll
regsvr32.exe urlmon.dll
regsvr32.exe mshtml.dll
regsvr32.exe shdocvw.dll
regsvr32.exe browseui.dll
regsvr32.exe jscript.dll
regsvr32.exe vbscript.dll
regsvr32.exe scrrun.dll
regsvr32.exe msxml.dll
regsvr32.exe msxml3.dll
regsvr32.exe msxml6.dll
regsvr32.exe actxprxy.dll
regsvr32.exe softpub.dll
regsvr32.exe wintrust.dll
regsvr32.exe dssenh.dll
regsvr32.exe rsaenh.dll
regsvr32.exe gpkcsp.dll
regsvr32.exe sccbase.dll
regsvr32.exe slbcsp.dll
regsvr32.exe cryptdlg.dll
regsvr32.exe oleaut32.dll
regsvr32.exe ole32.dll
regsvr32.exe shell32.dll
regsvr32.exe initpki.dll
regsvr32.exe wuapi.dll
regsvr32.exe wuaueng.dll
regsvr32.exe wuaueng1.dll
regsvr32.exe wucltui.dll
regsvr32.exe wups.dll
regsvr32.exe wups2.dll
regsvr32.exe wuweb.dll
regsvr32.exe qmgr.dll
regsvr32.exe qmgrprxy.dll
regsvr32.exe wucltux.dll
regsvr32.exe muweb.dll
regsvr32.exe wuwebv.dll

netsh winsock reset

netsh winhttp reset proxy

pause

net start bits
net start wuauserv
net start appidsvc
net start cryptsvc

ping 127.0.0.1 -n 4 >nul

wuauclt /detectnow
wuauclt /reportnow
