@echo off
rem  ©CyberblueX cyberbluex@outlook.com  
cls
rem GOTO :Skip_Check
REM --- Get admin rights
fsutil dirty query %SystemDrive% >NUL && set admin=true
if NOT "%admin%"=="true" (
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\RunAsAdmin.vbs"
	echo UAC.ShellExecute "%~0", "%~1 %~2 %~3 %~4 %~5 %~6 %~7 %~8 %~9", "", "runas", 1 >> "%temp%\RunAsAdmin.vbs"
	"%temp%\RunAsAdmin.vbs"
	del /Q /S "%temp%\RunAsAdmin.vbs"
	goto :EOF
)

IF NOT EXIST "C:\ProgramData\chocolatey\choco.exe" powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((new-object net.webclient).DownloadString('https://chocolatey.org/install.ps1'))" && SET PATH=%PATH%;%ALLUSERSPROFILE%\chocolatey\bin

:Skip_Check

SET "installgroup=Business"
pushd %~dp0
SET "choco_list=%CD%\choco_list.csv"

:Start
cls

SET /p header= <%choco_list%

ECHO Choose installation group:
echo.
ECHO    Name:
setlocal EnableDelayedExpansion
SET /A i=0
FOR %%a in (%header%) Do (

	IF !i! GEQ 1 IF !i! LEQ 10 echo !i!: %%~a
	SET /A i=!i!+1	
)
endlocal
ECHO.
ECHO    upgrade (only upgrades allready installed packages)
echo.

SET /P "installgroup= Type in full Name of Group: (%installgroup%) >>"

IF /I "%installgroup%" == "exit" GOTO :Exit_okay
IF /I "%installgroup%" == "upgrade" GOTO :Upgrade_only

setlocal EnableDelayedExpansion
For /F "usebackq tokens=* delims=" %%a in (%choco_list%) Do (
    set "line=%%a"
    rem setlocal EnableDelayedExpansion
    set "line="!line:;=";"!""
    For /F "tokens=1-10  delims=;" %%a in ("!line!") Do (
        rem echo %%~a - %%~b - %%~c - %%~d
		
		IF /I "%%~b" == "%installgroup%" SET "colums=1,2" & goto :Skip
		IF /I "%%~c" == "%installgroup%" SET "colums=1,3" & goto :Skip
		IF /I "%%~d" == "%installgroup%" SET "colums=1,4" & goto :Skip
		IF /I "%%~e" == "%installgroup%" SET "colums=1,5" & goto :Skip
		IF /I "%%~f" == "%installgroup%" SET "colums=1,6" & goto :Skip
		IF /I "%%~g" == "%installgroup%" SET "colums=1,7" & goto :Skip
		IF /I "%%~h" == "%installgroup%" SET "colums=1,8" & goto :Skip
		IF /I "%%~i" == "%installgroup%" SET "colums=1,9" & goto :Skip
		IF /I "%%~j" == "%installgroup%" SET "colums=1,10"	& goto :Skip	
		goto :Start
    )

rem	For /F "tokens=%blubb% delims=;" %%a in ("!line!") Do (
rem        echo %%~a - %%~b
rem    )
	rem endlocal
)
:Skip
endlocal & SET "colums=%colums%"

setlocal EnableDelayedExpansion
SET /A i=1
For /F "usebackq skip=1 tokens=* delims=" %%a in (%choco_list%) Do (
    set "line=%%a"
    rem setlocal EnableDelayedExpansion
    set "line="!line:;=";"!""

	For /F "tokens=%colums% delims=;" %%a in ("!line!") Do (
    rem    IF /I "%%~b" == "WAHR" echo Adding [!i!]: %%~a 
		rem IF /I "%%~b" == "WAHR" choco install %%~a -y
		IF /I "%%~b" == "WAHR" SET "tmp_packages= !tmp_packages! %%~a"
		IF /I "%%~b" == "WAHR" SET /A i=!i!+1
    )
	
)
endlocal & SET "packages=%tmp_packages%"

IF NOT "%packages%" == "" choco install %packages% --limitoutput --yes


:finishing_steps
DEL /F /S /Q "C:\Users\Public\Desktop\*.lnk"

rem choco list -l

ECHO.
ECHO FINISH :^)
ECHO.

popd

pause
:EXIT_okay
popd
EXIT /B 0

:Upgrade_only
choco upgrade all --limitoutput --yes
goto :finishing_steps