@echo off
setlocal
title Install Dreamkas Dependencies
cd /d "%~dp0"

echo ==========================================
echo      Install Dreamkas Dependencies
echo ==========================================
echo.

set "PYTHON_CMD="
if exist "C:\Python314\python.exe" set "PYTHON_CMD=C:\Python314\python.exe"
if "%PYTHON_CMD%"=="" (where py >nul 2>nul && set "PYTHON_CMD=py")
if "%PYTHON_CMD%"=="" (where python >nul 2>nul && set "PYTHON_CMD=python")

if "%PYTHON_CMD%"=="" (
    echo ERROR: Python not found.
    pause
    exit /b 1
)

echo Python: %PYTHON_CMD%
echo.

%PYTHON_CMD% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo pip not found. Trying ensurepip...
    %PYTHON_CMD% -m ensurepip --upgrade
    if errorlevel 1 (
        echo ERROR: Could not install pip.
        pause
        exit /b 1
    )
)

echo Upgrading pip tools...
%PYTHON_CMD% -m pip install --upgrade pip setuptools wheel

echo Installing dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt

echo.
echo Done.
pause
endlocal
