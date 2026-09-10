@echo off
setlocal EnableDelayedExpansion
title Voter Suvidha - One Time Setup
cls

echo =====================================================================
echo            VOTER SUVIDHA - ONE TIME SETUP (Dependencies)
echo =====================================================================
echo.
echo This installs the two Python packages the PDF extractor needs:
echo    pymupdf   (reads the electoral roll PDF)
echo    openpyxl  (writes the 12-column Excel)
echo.
echo Internet is needed ONLY for this one-time step.
echo After this, the app runs fully offline.
echo.

call :FIND_PYTHON
if "%PY_EXE%"=="" goto :NO_PYTHON

echo Using Python: %PY_EXE%
"%PY_EXE%" --version
echo.

echo Upgrading pip...
"%PY_EXE%" -m pip install --upgrade pip

echo.
echo Installing pymupdf and openpyxl...
"%PY_EXE%" -m pip install pymupdf openpyxl fonttools

echo.
echo Verifying installation...
"%PY_EXE%" -c "import pymupdf, openpyxl, fontTools; print('OK  pymupdf', pymupdf.__doc__ or '', '| openpyxl', openpyxl.__version__)"
if errorlevel 1 goto :VERIFY_FAILED

echo.
echo =====================================================================
echo  SETUP COMPLETE. You can now run Voter_Suvidha.bat
echo =====================================================================
echo.
pause
exit /b 0


:VERIFY_FAILED
echo.
echo ---------------------------------------------------------------------
echo  Packages did not import correctly.
echo.
echo  Try running this manually and read the error:
echo      "%PY_EXE%" -m pip install pymupdf openpyxl fonttools
echo.
echo  If you are on a network with a proxy, pip may be blocked.
echo ---------------------------------------------------------------------
echo.
pause
exit /b 1


:NO_PYTHON
echo ---------------------------------------------------------------------
echo  PYTHON NOT FOUND ON THIS COMPUTER
echo ---------------------------------------------------------------------
echo.
echo  Install Python 3 first:
echo    1. Go to  https://www.python.org/downloads/windows/
echo    2. Download "Windows installer (64-bit)" for Python 3.12 or newer
echo    3. IMPORTANT: on the first installer screen, TICK the box
echo         [x] Add python.exe to PATH
echo    4. Click "Install Now"
echo    5. Close this window, then run SETUP_FIRST_TIME.bat again
echo.
echo ---------------------------------------------------------------------
echo.
pause
exit /b 1


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
