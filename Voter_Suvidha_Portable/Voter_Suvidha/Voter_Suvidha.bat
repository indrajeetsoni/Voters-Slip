@echo off
setlocal EnableDelayedExpansion
title Voter Suvidha - Local Server
cls

echo =====================================================================
echo                     VOTER SUVIDHA (Voter Suvidha)
echo             Voter Slip and Electoral Roll Management
echo =====================================================================
echo.
echo Starting Voter Suvidha local server...
echo The application will open in your web browser automatically.
echo (NOTE: Keep this window open while using Voter Suvidha)
echo.

set "SCRIPT_DIR=%~dp0"
set "SERVER_PS1=%SCRIPT_DIR%voter_suvidha\server.ps1"
set "SERVER_PY=%SCRIPT_DIR%voter_suvidha\server.py"

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    py "%SERVER_PY%"
    goto :SERVER_STOPPED
)

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    python "%SERVER_PY%"
    goto :SERVER_STOPPED
)

set "PS_EXE=powershell.exe"
if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
    set "PS_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
) else if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
    set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
)

"%PS_EXE%" -ExecutionPolicy Bypass -NoProfile -File "%SERVER_PS1%"

:SERVER_STOPPED
echo.
echo =====================================================================
echo Voter Suvidha server has stopped.
echo =====================================================================
pause
