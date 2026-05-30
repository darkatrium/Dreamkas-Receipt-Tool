@echo off
setlocal
title Dreamkas Receipt Tool
cd /d "%~dp0"

echo ==========================================
echo        Dreamkas Receipt Tool v6.5
echo ==========================================
echo.

set "PYTHON_CMD="

if exist "C:\Python314\python.exe" set "PYTHON_CMD=C:\Python314\python.exe"

if "%PYTHON_CMD%"=="" (
    where py >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=py"
)

if "%PYTHON_CMD%"=="" (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if "%PYTHON_CMD%"=="" (
    echo ERROR: Python not found.
    echo Install Python and try again.
    pause
    exit /b 1
)

echo Python: %PYTHON_CMD%
echo.

if not exist "dreamkas_receipt.py" (
    echo ERROR: dreamkas_receipt.py not found.
    pause
    exit /b 1
)

if not exist "requirements.txt" (
    echo ERROR: requirements.txt not found.
    pause
    exit /b 1
)

echo Checking pip...
%PYTHON_CMD% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo pip not found. Trying to install pip...
    %PYTHON_CMD% -m ensurepip --upgrade
    if errorlevel 1 (
        echo ERROR: Could not install pip.
        pause
        exit /b 1
    )
)

echo Checking Python packages...
%PYTHON_CMD% -c "import requests, openpyxl, qrcode, PIL, reportlab" >nul 2>nul
if errorlevel 1 (
    echo Missing packages detected.
    echo Installing all dependencies from requirements.txt...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Could not install dependencies.
        echo Try manually:
        echo %PYTHON_CMD% -m pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
)

echo Dependencies OK.
echo.

set "EXCEL_FILE="

if not "%~1"=="" set "EXCEL_FILE=%~1"

if "%EXCEL_FILE%"=="" (
    if exist "dreamkas_receipt_template.xlsx" set "EXCEL_FILE=dreamkas_receipt_template.xlsx"
)

if "%EXCEL_FILE%"=="" (
    echo Excel file not found.
    echo Type Excel file path or drag XLSX file onto this BAT next time.
    set /p EXCEL_FILE=Excel file path: 
)

if "%EXCEL_FILE%"=="" (
    echo Cancelled.
    pause
    exit /b 1
)

echo Excel: %EXCEL_FILE%
echo.
echo Starting program...
echo.

%PYTHON_CMD% "dreamkas_receipt.py" --excel "%EXCEL_FILE%"

echo.
echo Program finished.
pause
endlocal
