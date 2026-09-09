@echo off
chcp 65001 >nul
title वोटर सुविधा (Voter Suvidha) - Server
cls
echo =====================================================================
echo               वोटर सुविधा (Voter Suvidha)
echo     निर्वाचन नामावली एवं मतदाता पर्ची प्रणाली
echo =====================================================================
echo.
echo [1/2] सर्वर प्रारंभ किया जा रहा है (Starting Local Server on port 5000)...

set SCRIPT_DIR=%~dp0
set SERVER_PS1=%SCRIPT_DIR%voter_suvidha\server.ps1
set SERVER_PY=%SCRIPT_DIR%voter_suvidha\server.py

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    start "" py "%SERVER_PY%"
) else (
    where python >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        start "" python "%SERVER_PY%"
    ) else (
        set "PS_EXE=powershell.exe"
        if exist "%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe" (
            set "PS_EXE=%SystemRoot%\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
        ) else if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
            set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
        )
        start "" "%PS_EXE%" -ExecutionPolicy Bypass -NoExit -File "%SERVER_PS1%"
    )
)

echo [2/2] ब्राउज़र में एप्लिकेशन खोला जा रहा है (Opening Web App)...
timeout /t 2 /nobreak >nul

if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" "http://127.0.0.1:5000"
) else (
    start "" "http://127.0.0.1:5000"
)

echo.
echo =====================================================================
echo वोटर सुविधा वेब एप्लिकेशन आपके ब्राउज़र में खुल चुका है!
echo URL: http://127.0.0.1:5000
echo.
echo इस विंडो को बंद न करें। जब काम समाप्त हो जाए, तो इसे बंद कर सकते हैं।
echo =====================================================================
pause
