@echo off
setlocal EnableDelayedExpansion
title Voter Suvidha - Push to GitHub
cls

set "REPO=https://github.com/indrajeetsoni/Voters-Slip.git"
set "BRANCH=main"
set "SCRIPT_DIR=%~dp0"

echo =====================================================================
echo            PUSH THIS FOLDER TO GITHUB
echo =====================================================================
echo.
echo  Folder : %SCRIPT_DIR%
echo  Repo   : %REPO%
echo  Branch : %BRANCH%
echo.
echo  Your existing GitHub history is kept. This adds one new commit
echo  containing whatever is in this folder right now.
echo.
echo  Press Ctrl+C to cancel, or
pause

cd /d "%SCRIPT_DIR%"

:: ------------------------------------------------------------------ git?
git --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ---------------------------------------------------------------------
    echo  GIT IS NOT INSTALLED
    echo ---------------------------------------------------------------------
    echo.
    echo  Install Git for Windows first:  https://git-scm.com/download/win
    echo  Accept all the default options, then run this file again.
    echo.
    pause
    exit /b 1
)

:: -------------------------------------------------- who is committing?
git config user.name  >nul 2>&1 || git config --global user.name  "indrajeetsoni"
git config user.email >nul 2>&1 || git config --global user.email "indrajeetsoni@users.noreply.github.com"

:: ------------------------------------------------- keep junk out of git
if not exist ".gitignore" (
    echo Creating .gitignore ...
    >  ".gitignore" echo # working folders, not source
    >> ".gitignore" echo voter_suvidha/uploads/
    >> ".gitignore" echo voter_suvidha/downloads/
    >> ".gitignore" echo voter_suvidha/web/downloads/
    >> ".gitignore" echo __pycache__/
    >> ".gitignore" echo *.pyc
    >> ".gitignore" echo # backups made by the patcher
    >> ".gitignore" echo *.bak
    >> ".gitignore" echo # generated output
    >> ".gitignore" echo unmapped_glyphs.png
    >> ".gitignore" echo preview_*.html
)

:: ------------------------------------------- first run: attach to repo
if not exist ".git" (
    echo.
    echo This folder came from a ZIP download, so it has no git history yet.
    echo Attaching it to your repository ...
    echo.
    git init -b %BRANCH%                       || goto :FAIL
    git remote add origin "%REPO%"             || goto :FAIL
    echo Fetching your current GitHub history ...
    git fetch origin %BRANCH%                  || goto :FAIL
    :: make the remote history our parent, keeping every local file as-is
    git reset --soft origin/%BRANCH%           || goto :FAIL
) else (
    git remote get-url origin >nul 2>&1 || git remote add origin "%REPO%"
    git fetch origin %BRANCH%
)

:: -------------------------------------------------------------- commit
echo.
echo Changes to be pushed:
echo ---------------------------------------------------------------------
git add -A
git status --short
echo ---------------------------------------------------------------------
echo.

git diff --cached --quiet
if not errorlevel 1 (
    echo Nothing has changed since the last push. Done.
    echo.
    pause
    exit /b 0
)

set "MSG="
set /p MSG=Commit message (press Enter for the default): 
if "%MSG%"=="" set "MSG=Fix Devanagari extraction and new voter slip design"

git commit -m "%MSG%"                          || goto :FAIL

:: ---------------------------------------------------------------- push
echo.
echo Pushing to GitHub ...
echo A browser or sign-in box may open the first time. Approve it.
echo.
git push -u origin %BRANCH%                    || goto :PUSHFAIL

echo.
echo =====================================================================
echo  DONE. Open your repo to confirm:
echo  https://github.com/indrajeetsoni/Voters-Slip/tree/%BRANCH%
echo =====================================================================
echo.
pause
exit /b 0


:PUSHFAIL
echo.
echo ---------------------------------------------------------------------
echo  PUSH WAS REJECTED
echo ---------------------------------------------------------------------
echo.
echo  Usually this means GitHub has commits your folder does not.
echo  Bring them in, then run this file again:
echo.
echo      git pull --rebase origin %BRANCH%
echo.
echo  If it complains about sign-in, install Git Credential Manager
echo  (included with Git for Windows) or use a Personal Access Token
echo  as the password.
echo.
pause
exit /b 1

:FAIL
echo.
echo ---------------------------------------------------------------------
echo  A git command failed - the message above says why.
echo  Send me that message and I will tell you the next step.
echo ---------------------------------------------------------------------
echo.
pause
exit /b 1
