@echo off
SETLOCAL EnableDelayedExpansion
:: =====================================================
:: © 2017 - 2019   cyberbluex@outlook.com
:: Version 1.1 - 1.2.2017
:: Changelog
:: 21.9.2019	- Changed default path from %~dp0 to %temp%
::				- Now the var %cx_log_file% are avaible in calling script
::				- Added Switch /GlobalLog , to enable Logging in Shared Files.
::				  Usefull for Admin Network Shares with Scripts....
::				  ! Need write access to Script Path !

:: 1.2.2017		Initial Version
:: =====================================================


rem CALL :Start_Time
rem Clearing old vars to zero
rem FOR /F "tokens=1 delims==" %%A in ('SET cx_log ^>nul 2^>nul ') DO SET "%%~A="

rem switch for Debugmode
SET "cx_debugmode=0"
rem Quick prefilter
FOR %%z in (%*) do (
	FOR /F "usebackq tokens=* delims=-/" %%A in ('%%~z') do (
rem	 	ECHO %%~A
		IF /I '%%~A' == 'debugmode' SET "cx_debugmode=1"
		IF /I '%%~A' == 'debug' SET "cx_debugmode=1"
	)
)

IF "%cx_debugmode%" == "1" cls

IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% START with Args:
IF "%cx_debugmode%" == "1" ECHO %~n0: [%*]

rem Fast ways
IF /I '%~1' == ''				GOTO arg_not_set
IF /I '%~1' == '/?'				GOTO arg_help
IF /I '%~1' == '-?'				GOTO arg_help
IF /I '%~1' == '?'				GOTO arg_help

rem Global Options

rem Default Logpath
SET "cx_log_path=%TEMP%"
SET "cx_globallog_path=%~dp0"

rem Default Delimiter
SET "cx_log_Delimiter=;"

rem Default Fileextension
SET "cx_log_fileextension=.log"

rem Default Filenames
SET "cx_log_filename_globaluuid=_PP-Log_UUID"
SET "cx_log_filename_globalcomputername=_PP-Log_Computername"

rem Width for Logtyp
SET "cx_log_setting_widthtyp=11"
SET "cx_log_setting_widthcomputername=17"
SET "cx_log_setting_widthuuid=36"
SET "cx_log_setting_widthuser=15"

rem Switches set to Default
SET "cx_log_sw_globallog=0"
SET "cx_log_sw_toscreen=0"
SET "cx_log_sw_error=0"
SET "cx_log_sw_warning=0"
SET "cx_log_sw_softout=0"

:cx_log
:: CALL cx_log "%~nx0" "Info" "Text"
:: CALL cx_log "Programm" "Typ" "Text"

SET /a i=0

IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% Detect Args

:Shift_Loop
SET /a i+=1
IF "%cx_debugmode%" == "1" ECHO.
IF "%cx_debugmode%" == "1" ECHO %~n0: Shiftloop Number: %i%
IF "%cx_debugmode%" == "1" ECHO %~n0: Arg:              %~1

IF '%~1' == '' GOTO Start_Log

FOR /F "usebackq tokens=1,* delims=:" %%A IN ('%~1') DO (

	IF "%cx_debugmode%" == "1" ECHO A: [%%A] B: [%%B]
	

	SET "tmp_arg=%%~A"
	SET "tmp_arg_command=%%~A"
	SET "tmp_arg_data=%%~B"
	
	IF "!tmp_arg:~0,1!" == "-" SET "tmp_arg_command=!tmp_arg:~1,999!"
	IF "!tmp_arg:~0,1!" == "/" SET "tmp_arg_command=!tmp_arg:~1,999!"

	IF "%cx_debugmode%" == "1" ECHO Ziffer1: [!tmp_arg:~0,1!]
	IF "%cx_debugmode%" == "1" ECHO Command: [!tmp_arg_command!]
	IF "%cx_debugmode%" == "1" ECHO DATA:    [!tmp_arg_data!]
	
	IF /I '!tmp_arg_command!' == 'help' 		GOTO arg_help
	IF /I '!tmp_arg_command!' == 'hilfe' 		GOTO arg_help
	IF /I '!tmp_arg_command!' == 'path' 		SET "cx_log_path=!tmp_arg_data!"           & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'logpath' 		SET "cx_log_path=!tmp_arg_data!"	       & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'file' 		SET "cx_log_file=!tmp_arg_data!"           & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'logfile' 		SET "cx_log_file=!tmp_arg_data!"           & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'uuid' 		SET "cx_log_uuid=!tmp_arg_data!"           & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'computername' SET "cx_log_computername=!tmp_arg_data!"   & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'device' 		SET "cx_log_computername=!tmp_arg_data!"   & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'pc' 			SET "cx_log_computername=!tmp_arg_data!"   & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'delimiter' 	SET "cx_log_Delimiter=!tmp_arg_data!"      & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'extension' 	SET "cx_log_fileextension=!tmp_arg_data!"  & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'msg'  		SET "cx_log_msg=!tmp_arg_data!"            & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'message'  	SET "cx_log_msg=!tmp_arg_data!"            & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'text' 		SET "cx_log_msg=!tmp_arg_data!"            & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'ToScreen' 	CALL :Toggle_Switch "cx_log_sw_toscreen"   & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'Error' 		CALL :Toggle_Switch "cx_log_sw_error"      & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'Warning' 		CALL :Toggle_Switch "cx_log_sw_warning"    & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'Softout' 		CALL :Toggle_Switch "cx_log_sw_softout"    & SHIFT /1 & GOTO Shift_Loop
	IF /I '!tmp_arg_command!' == 'GlobalLog'	CALL :Toggle_Switch "cx_log_sw_globallog"  & SHIFT /1 & GOTO Shift_Loop
	
	CALL :arg_unkown "%%~A" & SHIFT /1 & GOTO Shift_Loop
	
)


:Start_Log

rem Check If needed Args avaible
IF "%cx_log_msg%" == "" GOTO arg_missing

IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_toscreen   = %cx_log_sw_toscreen%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_error      = %cx_log_sw_error%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_warning    = %cx_log_sw_warning%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_softout    = %cx_log_sw_softout%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_globallog  = %cx_log_sw_globallog%

IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_path          = %cx_log_path%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_file          = %cx_log_file%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_uuid          = %cx_log_uuid%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_computername  = %cx_log_computername%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_fileextension = %cx_log_fileextension%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_msg           = %cx_log_msg%

IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% Detect Args finish


IF "%cx_debugmode%" == "1" ECHO.
IF "%cx_debugmode%" == "1" ECHO %~n0: Starting Logging

IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% Logging Start

:Start_Log_Preconfig
rem Vorverarbeitung der Daten

rem Logging Switches
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_toscreen = %cx_log_sw_toscreen%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_error    = %cx_log_sw_error%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_warning  = %cx_log_sw_warning%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_softout  = %cx_log_sw_softout%
IF "%cx_debugmode%" == "1" ECHO %~n0: cx_log_sw_globallog  = %cx_log_sw_globallog%

rem Setting Logtyp
SET "cx_log_typ="
IF "%cx_log_sw_warning%" 	== "1" SET "cx_log_typ=WARNING: "
IF "%cx_log_sw_error%" 		== "1" SET "cx_log_typ=ERROR: "

rem Setting Computername
IF "%cx_log_computername%" == "" SET "cx_log_computername=%Computername%"
IF "%cx_log_computername%" == "" SET "cx_log_computername=Unkown"

rem Setting UUID
IF "%cx_log_UUID%" == "" FOR /F "skip=2 tokens=2 delims=," %%A in ('wmic csproduct get uuid /FORMAT:csv') DO (SET "cx_log_UUID=%%A")
IF "%cx_log_UUID%" == "" SET "cx_log_UUID=Unkown"

rem Setting Logfilenames
SET "cx_log_file_globalcomputername=%cx_globallog_path%\%cx_log_filename_globalcomputername%%cx_log_fileextension%"
SET "cx_log_file_globaluuid=%cx_globallog_path%\%cx_log_filename_globaluuid%%cx_log_fileextension%"
SET "cx_log_file=%cx_log_path%\%cx_log_UUID%%cx_log_fileextension%"



IF "%cx_debugmode%" == "1" ECHO %~n0: Selected File: "%~2"


IF EXIST "%cx_log_file%" (
	IF "%cx_debugmode%" == "1" ECHO %~n0: Using existing File: %cx_log_file%
) ELSE (
	IF "%cx_debugmode%" == "1" ECHO %~n0: Creating new File: %cx_log_file%
)

IF "%cx_debugmode%" == "1" ECHO %~n0: Selected Path: "%~2"

IF EXIST "%cx_log_path%" (
	IF "%cx_debugmode%" == "1" ECHO %~n0: Using existing Path: %cx_log_path%
) ELSE (
	IF "%cx_debugmode%" == "1" ECHO %~n0: Creating new Folder: %cx_log_path%
	MD "%cx_log_path%" >nul 2>nul
	IF ERRORLEVEL 1 ECHO %~n0: Failed to create new Folder: %cx_log_path%
)

rem Get Date
IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% Before Date

IF "%systemdrive%" == "X:" SET "tmp_logline_datetime_utc=%DATE%T%time:~0,8% PE"
IF "%systemdrive%" == "X:" CALL :f_expand_Var "%tmp_logline_datetime_utc%" "23" "tmp_logline_datetime_utc"
IF "%systemdrive%" == "X:" GOTO Skip_WMIC


rem Query Date from WMI
rem IF "%cx_getdate_sw_utc%" == "1" ( SET "tmp_query=win32_utctime" ) ELSE ( SET "tmp_query=win32_localtime" )
for /f %%x in ('wmic path win32_utctime get /format:list ^| find "="') do set cx_dt_%%x
rem for /f %%x in ('wmic path %tmp_query% get /format:list ^| findstr "="') do set cx_dt_%%x

rem Add leading zeros then Pad digits
Set cx_dt_Month=00%cx_dt_Month%
Set cx_dt_Day=00%cx_dt_Day%
Set cx_dt_Hour=00%cx_dt_Hour%
Set cx_dt_Minute=00%cx_dt_Minute%
Set cx_dt_Second=00%cx_dt_Second%

Set cx_dt_Month=%cx_dt_Month:~-2%
Set cx_dt_Day=%cx_dt_Day:~-2%
Set cx_dt_Hour=%cx_dt_Hour:~-2%
Set cx_dt_Minute=%cx_dt_Minute:~-2%
Set cx_dt_Second=%cx_dt_Second:~-2%

set "cx_getdate_sw_utc=1"
set "cx_getdate_sw_seconds=1"

rem date/time in ISO 8601 format:
IF "%cx_getdate_sw_utc%" == "1" ( SET "tmp_suffix=Z" ) ELSE ( SET "tmp_suffix=" )
IF "%cx_getdate_sw_utc%" == "1" ( SET "tmp_suffix2= UTC" ) ELSE ( SET "tmp_suffix2=    " )
IF "%cx_getdate_sw_seconds%" == "1" ( SET "tmp_seconds=:%cx_dt_Second%" ) ELSE ( SET "tmp_seconds=" )
SET "cx_dt_format_iso=%cx_dt_Year%-%cx_dt_Month%-%cx_dt_Day%T%cx_dt_Hour%:%cx_dt_Minute%%tmp_seconds%%tmp_suffix%"
SET  "cx_t_format_iso=%cx_dt_Hour%:%cx_dt_Minute%%tmp_seconds%%tmp_suffix%"
SET  "cx_d_format_iso=%cx_dt_Year%-%cx_dt_Month%-%cx_dt_Day%%tmp_suffix%"
SET "tmp_logline_datetime_utc=%cx_dt_Year%-%cx_dt_Month%-%cx_dt_Day%T%cx_dt_Hour%:%cx_dt_Minute%:%cx_dt_Second%%tmp_suffix2%"

:Skip_WMIC

IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% After Date
IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% Expand Vars

CALL :f_expand_Var "%cx_log_typ%" "%cx_log_setting_widthtyp%" "tmp_cx_log_typ"
CALL :f_expand_Var "Warning: " "%cx_log_setting_widthtyp%" "tmp_cx_log_txt_warning"
CALL :f_expand_Var "%cx_log_computername%" "%cx_log_setting_widthcomputername%" "tmp_cx_log_computername"
CALL :f_expand_Var "%cx_log_uuid%" "%cx_log_setting_widthuuid%" "tmp_cx_log_uuid"
CALL :f_expand_Var "%username%" "%cx_log_setting_widthuser%" "tmp_cx_log_username"

IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% Before Logging
IF "%cx_log_sw_toscreen%" == "1" CALL :Write_Console

CALL :Write_PCUUID_File 2>&1 >>%cx_log_file%
IF "%cx_debugmode%" == "1" CALL :Write_PCUUID_File

IF "%cx_log_sw_globallog%" == "1" (
	CALL :Write_GlobalUUID_File 2>&1 >>%cx_log_file_globaluuid%
	IF "%cx_debugmode%" == "1" CALL :Write_GlobalUUID_File

	Call :Write_GlobalComputername_File 2>&1 >>%cx_log_file_globalcomputername%
	IF "%cx_debugmode%" == "1" Call :Write_GlobalComputername_File 
)

:EXIT_okay
SET "f_exitcode=0"
:: ======================================================================================
:: =========================== End Individual Script here ===============================
:EXIT
IF "%cx_debugmode%" == "1" ECHO %~n0: %TIME% ENDE
ENDLOCAL & SET "cx_log_file=%cx_log_file%"
rem Use goto EXIT for quitting the Script to disconnect temp Shares
EXIT /B %f_exitcode%
EXIT


:arg_help
ECHO.
ECHO %~n0: Sry. No Help TXT definded yet. Open Script
ECHO %~n0: %~f0
ECHO %~n0: for more details.
ECHO.
SET "f_exitcode=9000"
GOTO EXIT

:arg_missing
ECHO %~n0: Arguments missing (%*)
EXIT /B 0

:arg_unkown
ECHO %~n0: Unkown Argument (%1)
SET "cx_log_unkownargs=%~1 %cx_log_unkownargs%"
EXIT /B 0

:arg_not_set
ECHO.
ECHO %~n0: No Arguments set 1(%1) 2(%2)
GOTO arg_help

:Toggle_Switch
rem CALL :Toggle_Switch "Variable"
FOR /F "tokens=2 delims==" %%A in ('SET %~1') DO SET "tmp_sw=%%~A"
IF "%tmp_sw%" == "1" SET "%~1=0"
IF "%tmp_sw%" == "0" SET "%~1=1"
SET "tmp_sw="
EXIT /B 0

rem Output to Console
:Write_Console
ECHO [%time:~0,8%] %cx_log_typ%%cx_log_msg%
EXIT /B 0

rem Log to PC UUID File
:Write_PCUUID_File
echo %tmp_logline_datetime_utc%%cx_log_Delimiter%%tmp_cx_log_username%%cx_log_Delimiter%%tmp_cx_log_typ%%cx_log_Delimiter%%cx_log_msg%
IF NOT "%cx_log_unkownargs%" == "" echo %tmp_logline_datetime_utc%%cx_log_Delimiter%%tmp_cx_log_txt_warning%%cx_log_Delimiter%cx_Log Unkown Args: %cx_log_unkownargs%
EXIT /B 0

rem Log to Global Logfile (UUID)
:Write_GlobalUUID_File
echo %tmp_logline_datetime_utc%%cx_log_Delimiter%%tmp_cx_log_uuid%%cx_log_Delimiter%%tmp_cx_log_username%%cx_log_Delimiter%%tmp_cx_log_typ%%cx_log_Delimiter%%cx_log_msg%
IF NOT "%cx_log_unkownargs%" == "" echo %tmp_logline_datetime_utc%%cx_log_Delimiter%%tmp_cx_log_uuid%%cx_log_Delimiter%%tmp_cx_log_txt_warning%%cx_log_Delimiter%cx_Log Unkown Args: %cx_log_unkownargs%
EXIT /B 0

rem Log to Global Logfile (Computername)
:Write_GlobalComputername_File
echo %tmp_logline_datetime_utc%%cx_log_Delimiter%%tmp_cx_log_computername%%cx_log_Delimiter%%tmp_cx_log_username%%cx_log_Delimiter%%tmp_cx_log_typ%%cx_log_Delimiter%%cx_log_msg%
IF NOT "%cx_log_unkownargs%" == "" echo %tmp_logline_datetime_utc%%cx_log_Delimiter%%tmp_cx_log_computername%%cx_log_Delimiter%%tmp_cx_log_txt_warning%%cx_log_Delimiter%cx_Log Unkown Args: %cx_log_unkownargs%
EXIT /B 0

:f_expand_Var
SETLOCAL EnableDelayedExpansion
:: CALL f_expand_Var     %~1        %~2        %~3
:: CALL f_expand_Var "%text_in%" "%max_chars%" "var_out"

IF NOT "%~4" == "" ECHO %~n0: Zu viele Paramter angegeben 1[%~1] 2[%~2] 3[%~3] 4[%~4] 5[%~5] 6[%~6] 7[%~7] 8[%~8] 9[%~9]"

SET "tmp_var=%~1"
IF "%~2" == "" ( SET "max_chars=10" ) ELSE ( SET "max_chars=%~2" )

SET "tmp_var=%tmp_var%                                                                                                                                                                                                                                                  "
SET "tmp_var=!tmp_var:~0,%max_chars%!"

ENDLOCAL & SET "%~3=%tmp_var%"	
EXIT /B 0

:f_expand_Var_Num
SETLOCAL EnableDelayedExpansion
:: CALL :f_expand_Var_Num     %~1        %~2        %~3
:: CALL :f_expand_Var_Num "%text_in%" "max_chars" "var_out"

IF NOT "%~4" == "" ECHO %~n0: Zu viele Paramter angegeben 1[%~1] 2[%~2] 3[%~3] 4[%~4] 5[%~5] 6[%~6] 7[%~7] 8[%~8] 9[%~9]"

SET "tmp_var=%~1"
IF "%~2" == "" ( SET "max_chars=10" ) ELSE ( SET "max_chars=%~2" )

SET "tmp_var=                                                                                                                                                                                                                                                  %tmp_var%"
SET "tmp_var=!tmp_var:~-%max_chars%!"

ENDLOCAL & SET "%~3=%tmp_var%"		
EXIT /B 0

:Start_Time
set /a timerstart= ((1%time:~0,2%-100)*60*60) + ((1%time:~3,2%-100)*60) + (1%time:~6,2%-100)
set /a timerstart2= ((1%time:~0,2%-100)*60*60*100) + ((1%time:~3,2%-100)*60*100) + ((1%time:~6,2%-100)*100) + (1%time:~9,2%-100)
GOTO :EOF

:End_Time
set /a timerstop=((1%time:~0,2%-100)*60*60)+((1%time:~3,2%-100)*60)+(1%time:~6,2%-100)
set /a timerstop2=((1%time:~0,2%-100)*60*60*100) + ((1%time:~3,2%-100)*60*100) + ((1%time:~6,2%-100)*100) + (1%time:~9,2%-100)
set /a timeseks=(%timerstop%-%timerstart%)
set /a timemins=(%timerstop%-%timerstart%)/60
set /a timemseks2=(%timerstop2%-%timerstart2%)
set /a timeseks2=(%timerstop2%-%timerstart2%)/100
set /a timemins2=(%timerstop2%-%timerstart2%)/100/60
echo Laufzeit: %timemins2%:%timeseks2%:%timemseks2% bei Punkt %~1
GOTO :EOF