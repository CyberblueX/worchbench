rem ########################################################
rem © 2016 - 2019   cyberbluex@outlook.com
rem ########################################################
rem Notes:
rem %errorlevel% after BCU seems to be always 0 ? But IF ERRORLEVEL works!
rem Sometimes IF ERRORLEVEL is negativ -1 ....
rem 

@echo off
SETLOCAL EnableDelayedExpansion
cls
REM GOTO :Skip_ADM_Check
REM --- Get admin rights
fsutil dirty query %SystemDrive% >NUL && set admin=true
if NOT "%admin%"=="true" (
	echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\RunAsAdmin.vbs"
	echo UAC.ShellExecute "%~0", "%~1 %~2 %~3 %~4 %~5 %~6 %~7 %~8 %~9", "", "runas", 1 >> "%temp%\RunAsAdmin.vbs"
	"%temp%\RunAsAdmin.vbs"
	del /Q /S "%temp%\RunAsAdmin.vbs"
	goto :EOF
)
:Skip_ADM_Check
pushd "%~dp0"


rem config cx_log
FOR /F "skip=2 tokens=2 delims=," %%A in ('wmic csproduct get uuid /FORMAT:csv') DO (SET "cx_log_UUID=%%A")
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Starting Script %~n0"
CALL cx_log /UUID:%cx_log_UUID% /msg:"Full Script Path %~0"

rem Settings

rem default exit code
SET EXITCODE=8000

rem Pathes & Files
rem using %cd% because pushd ...
SET "BcuRoot=%CD%\BCU-4.0.26.1"
::SET "BcuOldRoot=%CD%\BCU-2.60.13"
SET "ConfigRoot=%CD%\Configs"
SET "PwdFilesRoot=%CD%\PwdFiles"
SET "bios_actual_file=%CD%\PwdFiles\01_actual.bin"
SET "bios_actual_tmp_file=%temp%\temp.bin"
SET "bios_tmp_file=%temp%\temp2.bin"
SET "bios_test_unicode=%ConfigRoot%\Test.txt"
SET "bios_defaults_unicode=%ConfigRoot%\Defaults.txt"
SET "bios_output=%temp%\BIOS_%computername%.txt"

rem detect custom config file
IF /I NOT '%~1' == '' IF EXIST "%ConfigRoot%\%~1.txt" SET "bios_config_custom=%ConfigRoot%\%~1.txt"
IF NOT "%bios_config_custom%" == "" CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Info: Custom File %bios_config_custom%"

IF /I "%PROCESSOR_ARCHITECTURE%" == "amd64" SET "HPBCU=%BcuRoot%\BiosConfigUtility64.exe"
IF /I "%PROCESSOR_ARCHITECTURE%" == "amd64" SET "HPBCUPSWD=%BcuRoot%\HPQPswd64.exe"
IF /I "%PROCESSOR_ARCHITECTURE%" == "x86" SET "HPBCU=%BcuRoot%\BiosConfigUtility.exe"
IF /I "%PROCESSOR_ARCHITECTURE%" == "x86" SET "HPBCUPSWD=%BcuRoot%\HPQPswd.exe"


rem check pc manufacturer
FOR /F "skip=2 tokens=2 delims=," %%A in ('wmic csproduct get vendor /FORMAT:csv') DO (SET "PC_Manufacturer=%%A")
IF "%PC_Manufacturer%" == "HP" GOTO Initialisation
IF "%PC_Manufacturer%" == "Hewlett-Packard" GOTO Initialisation
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Manufacturer %PC_Manufacturer% not supported. Exiting..."
SET EXITCODE=8001
GOTO EXIT 

:Initialisation
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Manufacturer %PC_Manufacturer% supported."
FOR /F "skip=2 tokens=2 delims=," %%A in ('wmic csproduct get name /FORMAT:csv') DO (SET "PC_Product_Name=%%A")
SET "bios_config=%ConfigRoot%\%PC_Product_Name%.txt"
SET "bios_config_uefi=%ConfigRoot%\%PC_Product_Name%-UEFI.txt"

rem detect UEFI systems
wmic partition where "Type = 'GPT: System'" get BootPartition | findstr TRUE >nul 2>nul
IF NOT ERRORLEVEL 1 SET "bios_config=%bios_config_uefi%"

rem remove older bin files
DEL /Q /F "%bios_actual_tmp_file%"  >>%cx_log_file% 2>&1
DEL /Q /F "%bios_tmp_file%"  >>%cx_log_file% 2>&1

rem Script in GetMode?
IF /I '%~1' == '/get' GOTO GETBIOS 

rem Unicode detection
%HPBCU% /unicode >nul 2>nul
IF ERRORLEVEL 1 GOTO NoUnicodeSystem
GOTO UnicodeSystem


:NoUnicodeSystem
rem No Unicode Support (Old Device)
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"System UNICODE Password Support: No"
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"This Script can't set Passwords to older Systems...exiting.."
SET EXITCODE=8002
GOTO EXIT

:UnicodeSystem
rem Unicode Support, newer device...
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"System UNICODE Password Support: Yes"


CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Step: Setting BIOS Password"

rem Copy actual BIOS Config to TEMP
COPY "%bios_actual_file%" "%bios_actual_tmp_file%" >>%cx_log_file% 2>&1

rem Write Test without password....
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Trying set testvalue without Password ..."
%HPBCU% /SET:"%bios_test_unicode%" >>%cx_log_file% 2>&1
ECHO RETURNCODE: [!errorlevel!]
IF ERRORLEVEL 1 (
	GOTO try_actual_uc
) ELSE (
	GOTO try_set_new_uc
)

:try_actual_uc
rem Write Test with actual password....
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Trying set testvalue actual Password ..."
%HPBCU% /SET:"%bios_test_unicode%" /cpwdfile:"%bios_actual_tmp_file%" >>%cx_log_file% 2>&1
ECHO RETURNCODE: [!errorlevel!]
IF ERRORLEVEL 1 (
	GOTO try_older_uc
) ELSE (
	GOTO Skip_Set_Password_Unicode
)

:try_set_new_uc
rem ERRORCODE 10 wenn schon eins da ist...
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Trying set new Password"
%HPBCU% /npwdfile:"%bios_actual_tmp_file%" >>%cx_log_file% 2>&1
ECHO RETURNCODE: [!errorlevel!]
IF ERRORLEVEL 1 (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Failed]"
	SET EXITCODE=4001
	GOTO EXIT
) ELSE (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"New Password set. Apply defaults ..."
	%HPBCU% /SET:"%bios_test_unicode%" /cpwdfile:"%bios_actual_tmp_file%" >>%cx_log_file% 2>&1
	GOTO Skip_Set_Password_Unicode
)

:try_older_uc
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Failed ... Trying older Passwords"

FOR /R %PwdFilesRoot% %%G IN (*.bin) DO (
	IF NOT "%bios_actual_file%" == "%%~G" (
		COPY "%%~G" "%bios_tmp_file%" >>%cx_log_file% 2>&1
		%HPBCU% /npwdfile:"%bios_actual_tmp_file%" /cpwdfile:"%bios_tmp_file%" >>%cx_log_file% 2>&1
		IF ERRORLEVEL 1 (
			CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Failed] %%~nG"
		) ELSE (
			rem Success, goto Skip
			CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Success] %%~nG"
			DEL /Q /F "%bios_tmp_file%"  >>%cx_log_file% 2>&1
			GOTO Skip_Set_Password_Unicode
		)		
		DEL /Q /F "%bios_tmp_file%"  >>%cx_log_file% 2>&1
	)
)
rem All known older PW failed...exit
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"No old PW matches ... exiting"
SET EXITCODE=4002
GOTO EXIT
	
:Skip_Set_Password_Unicode
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Step: Apply BIOS Config File"

IF NOT "%bios_config_custom%" == "" (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Using custom config %bios_config_custom%"
	%HPBCU% /SET:"%bios_config_custom%" /cpwdfile:"%bios_actual_tmp_file%" >>%cx_log_file% 2>&1
	ECHO RETURNCODE: [!errorlevel!]
	IF NOT ERRORLEVEL 1 (
		CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Sucessfull] set Custom Bios Config file"
		SET EXITCODE=0
		GOTO EXIT
	) ELSE (
		CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Failed] set Custom Bios Config file"
		ping 172.0.0.1 -n 2 >nul 2>nul
		SET EXITCODE=4005
		GOTO EXIT
	)		
)

:try_set_defaults
rem Set Defaults
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Trying set defaults ..."
%HPBCU% /SET:"%bios_defaults_unicode%" /cpwdfile:"%bios_actual_tmp_file%" >>%cx_log_file% 2>&1
ECHO RETURNCODE: [!errorlevel!]
IF NOT ERRORLEVEL 1 (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Sucessfull] set all defaults"
) ELSE (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Failed] set all defaults"
	ping 172.0.0.1 -n 2 >nul 2>nul
)

IF EXIST "%bios_config%" (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Using %bios_config%"
	%HPBCU% /SET:"%bios_config%" /cpwdfile:"%bios_actual_tmp_file%" >>%cx_log_file% 2>&1
	ECHO RETURNCODE: [!errorlevel!]
	IF NOT ERRORLEVEL 1 (
		CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Sucessfull] set Bios Config file"
		SET EXITCODE=0
		GOTO EXIT
	) ELSE (
		CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"[Failed] set Bios Config file"
		ping 172.0.0.1 -n 2 >nul 2>nul
		SET EXITCODE=4003
		GOTO EXIT
	)		
) ELSE (
	CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"No BIOS Config File for %PC_Product_Name% found...."
	%HPBCU% /get:"%bios_output%" >>%cx_log_file% 2>&1
	SET EXITCODE=4004
	GOTO EXIT
)
SET EXITCODE=4100
GOTO EXIT

:GETBIOS

%HPBCU% /get:"%bios_output%" >>%cx_log_file% 2>&1
notepad "%bios_output%"

SET EXITCODE=0
GOTO EXIT


:EXIT
rem Step Exit and Cleaning Tasks
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Finished Script %~n0 - Errorlevel: %EXITCODE%"
DEL /Q /F "%bios_actual_tmp_file%"  >>%cx_log_file% 2>&1
DEL /Q /F "%bios_tmp_file%"  >>%cx_log_file% 2>&1
popd
EXIT /B %EXITCODE%

rem Old Code Snippets
EXIT
EXIT
EXIT

rem Create BIOS File with ERRORLEVEL
rem ERRORLEVEL can be less 0 and greater than 0
CALL cx_log /UUID:%cx_log_UUID% /ToScreen /msg:"Creating Temp-PWD file..."
%HPBCUPSWD% /s /f"%bios_actual_tmp_file%" /p"%bios_pwd_new_uc%"
ECHO RETURNCODE: [!errorlevel!]
IF ERRORLEVEL 1 ( 
	CALL cx_log /UUID:%cx_log_UUID% /ERROR /ToScreen /msg:"[Failed]"
	SET EXITCODE=5000
	GOTO EXIT
) 
IF !ERRORLEVEL! LSS 0 (	
	CALL cx_log /UUID:%cx_log_UUID% /ERROR /ToScreen /msg:"[Failed]"
	SET EXITCODE=5000
	GOTO EXIT
) 