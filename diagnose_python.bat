@echo off
setlocal EnableExtensions
title Diagnose Python
cd /d "%~dp0"

echo ==========================================
echo          Diagnose Python
echo ==========================================
echo.

echo where python:
where python
echo.

echo where py:
where py
echo.

echo where winget:
where winget
echo.

echo py -0p:
py -0p
echo.

echo Testing commands:
echo.

for %%C in ("C:\Python314\python.exe" "C:\Python313\python.exe" "C:\Python312\python.exe" "py -3.14" "py -3.13" "py -3.12" "py -3.11" "py -3" "py" "python") do (
    echo --- %%~C
    %%~C -c "import sys; print(sys.version); print(sys.executable)" 2>&1
    echo.
)

echo ensurepip check with python:
python -m ensurepip --version
echo.

echo pip check with python:
python -m pip --version
echo.

echo import test with python:
python -c "import requests, openpyxl, qrcode, PIL, reportlab; print('all imports OK')"
echo.

pause
endlocal

