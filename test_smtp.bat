@echo off
setlocal
title Dreamkas SMTP Test
cd /d "%~dp0"
set "PYTHON_CMD="
if exist "C:\Python314\python.exe" set "PYTHON_CMD=C:\Python314\python.exe"
if "%PYTHON_CMD%"=="" (where py >nul 2>nul && set "PYTHON_CMD=py")
if "%PYTHON_CMD%"=="" (where python >nul 2>nul && set "PYTHON_CMD=python")
if "%PYTHON_CMD%"=="" (echo ERROR: Python not found.& pause& exit /b 1)
if not exist "test_smtp.py" (echo ERROR: test_smtp.py not found.& pause& exit /b 1)
set /p TEST_TO=Recipient email for SMTP test: 
if "%TEST_TO%"=="" (echo Cancelled.& pause& exit /b 1)
%PYTHON_CMD% "test_smtp.py" --to "%TEST_TO%"
endlocal
