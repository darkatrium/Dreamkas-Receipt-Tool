@echo off
setlocal
title Build Dreamkas EXE
cd /d "%~dp0"
set "PYTHON_CMD="
if exist "C:\Python314\python.exe" set "PYTHON_CMD=C:\Python314\python.exe"
if "%PYTHON_CMD%"=="" (where py >nul 2>nul && set "PYTHON_CMD=py")
if "%PYTHON_CMD%"=="" (where python >nul 2>nul && set "PYTHON_CMD=python")
if "%PYTHON_CMD%"=="" (echo ERROR: Python not found.& pause& exit /b 1)
%PYTHON_CMD% -m pip --version >nul 2>nul
if errorlevel 1 %PYTHON_CMD% -m ensurepip --upgrade
echo Installing build dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 (echo ERROR: pip install failed.& pause& exit /b 1)
echo Building console EXE...
%PYTHON_CMD% -m PyInstaller --onefile --name DreamkasReceipt dreamkas_receipt.py
echo Building GUI EXE...
%PYTHON_CMD% -m PyInstaller --onefile --windowed --name DreamkasReceiptGUI dreamkas_gui.py
echo.
echo Done. EXE files are in the dist folder.
pause
endlocal
