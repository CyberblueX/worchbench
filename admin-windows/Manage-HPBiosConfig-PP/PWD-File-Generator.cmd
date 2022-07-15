@echo off
SETLOCAL ENABLEDELAYEDEXPANSION
pushd %~dp0
SET "bios_pwd_old_uc=Passwort1;Passwort2;Passwort3"
SET "BcuRoot=%CD%\BCU-4.0.26.1"
SET "HPBCUPSWD=%BcuRoot%\HPQPswd64.exe"

ECHO %BcuRoot%
ECHO %HPBCUPSWD%

	FOR %%A IN (%bios_pwd_old_uc%) do (
		%HPBCUPSWD% /s /p"%%~A" /f"%CD%\PwdFiles\%%~A-!RANDOM!.bin"
		IF ERRORLEVEL 1 (
			ECHO "[Failed] %%~A"
		) ELSE (
			rem Success, goto Skip
			ECHO "[Success] %%~A"
		)
)
popd
pause