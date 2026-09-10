@echo off
setlocal EnableDelayedExpansion
title Voter Suvidha - Local Server
cls

echo =====================================================================
echo                     VOTER SUVIDHA (Voter Suvidha)
echo             Voter Slip and Electoral Roll Management
echo =====================================================================
echo.

set "SCRIPT_DIR=%~dp0"
set "SERVER_PS1=%SCRIPT_DIR%voter_suvidha\server.ps1"
set "SERVER_PY=%SCRIPT_DIR%voter_suvidha\server.py"

if not exist "%SERVER_PY%" (
    echo ERROR: Cannot find "%SERVER_PY%"
    echo Make sure Voter_Suvidha.bat is in the SAME folder as the
    echo "voter_suvidha" subfolder.
    echo.
    pause
    exit /b 1
)

:: ---------------------------------------------------------------------
:: Step 1: find Python
:: ---------------------------------------------------------------------
call :FIND_PYTHON

if "%PY_EXE%"=="" (
    echo ---------------------------------------------------------------------
    echo  PYTHON NOT FOUND
    echo ---------------------------------------------------------------------
    echo.
    echo  The PDF extractor needs Python 3. Without it the app cannot read
    echo  the electoral roll PDF.
    echo.
    echo  Fix: install Python 3, TICKING "Add python.exe to PATH",
    echo       then run SETUP_FIRST_TIME.bat once.
    echo.
    echo       https://www.python.org/downloads/windows/
    echo.
    echo ---------------------------------------------------------------------
    echo.
    pause
    exit /b 1
)

echo Python found: %PY_EXE%

:: ---------------------------------------------------------------------
:: Step 2: verify required packages
:: ---------------------------------------------------------------------
"%PY_EXE%" -c "import pymupdf, openpyxl, fontTools" >nul 2>&1
if errorlevel 1 (
    echo.
    echo Required packages are missing. Installing them now...
    echo.
    "%PY_EXE%" -m pip install pymupdf openpyxl fonttools
    echo.
    "%PY_EXE%" -c "import pymupdf, openpyxl, fontTools" >nul 2>&1
    if errorlevel 1 (
        echo ---------------------------------------------------------------------
        echo  Could not install pymupdf / openpyxl automatically.
        echo  Run SETUP_FIRST_TIME.bat and read the error it prints.
        echo ---------------------------------------------------------------------
        echo.
        pause
        exit /b 1
    )
)

echo Dependencies OK.
echo.
echo Starting server... the app will open in your browser automatically.
echo (Keep THIS window open while using Voter Suvidha.)
echo.

"%PY_EXE%" "%SERVER_PY%"

echo.
echo =====================================================================
echo Voter Suvidha server has stopped.
echo =====================================================================
pause
exit /b 0


:: =====================================================================
:: Locate a usable Python: PATH first, then common install locations
:: =====================================================================
:FIND_PYTHON
set "PY_EXE="

py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%i in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY_EXE=%%i"
    if not "!PY_EXE!"=="" exit /b 0
)

python -c "import sys" >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PY_EXE=%%i"
    if not "!PY_EXE!"=="" exit /b 0
)

python3 -c "import sys" >nul 2>&1
if not errorlevel 1 (
    for /f "delims=" %%i in ('python3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY_EXE=%%i"
    if not "!PY_EXE!"=="" exit /b 0
)

for %%d in (
    "%LOCALAPPDATA%\Programs\Python"
    "C:\Program Files"
    "C:\Program Files (x86)"
    "C:\"
) do (
    if exist "%%~d" (
        for /f "delims=" %%p in ('dir /b /ad "%%~d\Python3*" 2^>nul') do (
            if exist "%%~d\%%p\python.exe" (
                set "PY_EXE=%%~d\%%p\python.exe"
                exit /b 0
            )
        )
    )
)

exit /b 0
