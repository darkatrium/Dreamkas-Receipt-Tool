@echo off
setlocal EnableExtensions
title Diagnose Python
cd /d "%~dp0"

echo ==========================================
echo          Diagnose Python
echo ==========================================
echo.

echo Main path:
if exist "C:\Python314\python.exe" (
    echo FOUND: C:\Python314\python.exe
    "C:\Python314\python.exe" -c "import sys, encodings; print(sys.version); print(sys.executable)" 2>&1
) else (
    echo not found
)
echo.

echo Directory listing C:\Python314:
if exist "C:\Python314" (
    dir "C:\Python314"
) else (
    echo not found
)
echo.

echo Saved python_path.txt:
if exist python_path.txt (type python_path.txt) else (echo not found)
echo.

echo Installer log:
if exist "%TEMP%\dreamkas-python-install.log" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Content -Path $env:TEMP\dreamkas-python-install.log -Tail 120"
) else (
    echo not found
)
echo.

echo Pip:
if exist "C:\Python314\python.exe" (
    "C:\Python314\python.exe" -m pip --version 2>&1
)
echo.

echo Import test:
if exist "C:\Python314\python.exe" (
    "C:\Python314\python.exe" -c "import requests, openpyxl, qrcode, PIL, reportlab; print('all imports OK')" 2>&1
)
echo.

echo Done.
pause
endlocal

