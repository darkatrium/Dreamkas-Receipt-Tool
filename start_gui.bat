@echo off
setlocal
title Dreamkas GUI
cd /d "%~dp0"
set "PYTHON_CMD="
if exist "C:\Python314\python.exe" set "PYTHON_CMD=C:\Python314\python.exe"
if "%PYTHON_CMD%"=="" (where py >nul 2>nul && set "PYTHON_CMD=py")
if "%PYTHON_CMD%"=="" (where python >nul 2>nul && set "PYTHON_CMD=python")
if "%PYTHON_CMD%"=="" (echo ERROR: Python not found.& pause& exit /b 1)
%PYTHON_CMD% -c "import requests, openpyxl, qrcode, PIL, reportlab" >nul 2>nul
if errorlevel 1 (
    echo Installing missing dependencies...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (echo ERROR: Dependency installation failed.& pause& exit /b 1)
)
%PYTHON_CMD% "dreamkas_gui.py"
endlocal
