@echo off
title Push to GitHub: Reflecteddark/revops-enterprise-os
cd /d "C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"

echo ======================================================================
echo       REVOPS ENTERPRISE OS V17.6: PUSH TO GITHUB
echo       Repository: https://github.com/Reflecteddark/revops-enterprise-os
echo ======================================================================
echo.

echo [1/3] Syncing latest files from scratch workspace...
if exist sync_to_repo.py python sync_to_repo.py

echo.
echo [2/3] Checking Git status...
git status -s

echo.
set "COMMIT_MSG="
set /p COMMIT_MSG="[Optional] Enter commit message (Press Enter for default): "
if not defined COMMIT_MSG set "COMMIT_MSG=chore: update RevOps Enterprise OS suite"

echo.
echo Adding files to Git index...
git add .

git diff-index --quiet HEAD --
if errorlevel 1 (
    echo Creating commit: "%COMMIT_MSG%"...
    git commit -m "%COMMIT_MSG%"
) else (
    echo No new changes to commit.
)

echo.
echo [3/3] Pushing to GitHub (origin main)...
git push origin main

if errorlevel 1 (
    echo.
    echo ======================================================================
    echo   [ERROR] Push failed. Check your network connection.
    echo ======================================================================
) else (
    echo.
    echo ======================================================================
    echo   SUCCESS! Project successfully updated on GitHub:
    echo   https://github.com/Reflecteddark/revops-enterprise-os
    echo ======================================================================
)

echo.
pause
