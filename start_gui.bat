@echo off
setlocal EnableExtensions
title Dreamkas GUI
cd /d "%~dp0"

call :ensure_python
if errorlevel 1 goto fail

call :ensure_pip
if errorlevel 1 goto fail

%PYTHON_CMD% -c "import requests, openpyxl, qrcode, PIL, reportlab" >nul 2>nul
if errorlevel 1 (
    echo Installing missing dependencies...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 goto fail
)

%PYTHON_CMD% "dreamkas_gui.py"
exit /b 0

:fail
echo.
echo ERROR: GUI startup failed.
pause
exit /b 1


:try_python
set "TEST_CMD=%~1"
if "%TEST_CMD%"=="" exit /b 1
set "TEST_VERSION="
set "TEST_MAJOR="
set "TEST_MINOR="

%TEST_CMD% -c "import sys; print(str(sys.version_info.major)+'.'+str(sys.version_info.minor)+'.'+str(sys.version_info.micro)); print(sys.executable)" > "%TEMP%\dreamkas_pycheck.txt" 2>nul
if errorlevel 1 exit /b 1

for /f "usebackq tokens=1 delims=" %%A in ("%TEMP%\dreamkas_pycheck.txt") do (
    if not defined TEST_VERSION set "TEST_VERSION=%%A"
)

if "%TEST_VERSION%"=="" exit /b 1

echo %TEST_VERSION% | findstr /r "^[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*$" >nul
if errorlevel 1 (
    set "TEST_VERSION="
    exit /b 1
)

for /f "tokens=1,2 delims=." %%A in ("%TEST_VERSION%") do (
    set "TEST_MAJOR=%%A"
    set "TEST_MINOR=%%B"
)

if "%TEST_MAJOR%"=="" exit /b 1
if %TEST_MAJOR% LSS 3 exit /b 1
if %TEST_MAJOR% EQU 3 if %TEST_MINOR% LSS 10 exit /b 1

set "PYTHON_CMD=%TEST_CMD%"
set "PYTHON_VERSION=%TEST_VERSION%"
set "TEST_VERSION="
exit /b 0


:find_python
set "PYTHON_CMD="
set "PYTHON_VERSION="

if exist "C:\Python314\python.exe" (
    call :try_python ""C:\Python314\python.exe""
    if not errorlevel 1 goto python_found
)

if exist "C:\Python313\python.exe" (
    call :try_python ""C:\Python313\python.exe""
    if not errorlevel 1 goto python_found
)

if exist "C:\Python312\python.exe" (
    call :try_python ""C:\Python312\python.exe""
    if not errorlevel 1 goto python_found
)

where py >nul 2>nul
if not errorlevel 1 (
    call :try_python "py -3.14"
    if not errorlevel 1 goto python_found

    call :try_python "py -3.13"
    if not errorlevel 1 goto python_found

    call :try_python "py -3.12"
    if not errorlevel 1 goto python_found

    call :try_python "py -3.11"
    if not errorlevel 1 goto python_found

    call :try_python "py -3"
    if not errorlevel 1 goto python_found

    call :try_python "py"
    if not errorlevel 1 goto python_found
)

where python >nul 2>nul
if not errorlevel 1 (
    call :try_python "python"
    if not errorlevel 1 goto python_found
)

exit /b 1

:python_found
echo Python command: %PYTHON_CMD%
echo Python version: %PYTHON_VERSION%
%PYTHON_CMD% -c "import sys; print('Python exe: ' + sys.executable)"
echo.
exit /b 0


:install_python
echo Valid Python 3.10+ was not found.
echo.
echo Trying automatic Python installation with winget...
echo.

where winget >nul 2>nul
if errorlevel 1 (
    echo ERROR: winget was not found.
    echo.
    echo Automatic Python installation is not available on this system.
    echo Install Python manually from python.org.
    echo During installation enable:
    echo - pip
    echo - Add python.exe to PATH
    echo - Python Launcher
    echo.
    echo Download page:
    echo https://www.python.org/downloads/windows/
    exit /b 1
)

echo Installing Python 3.13 with winget...
winget install --id Python.Python.3.13 -e --source winget --accept-package-agreements --accept-source-agreements
if errorlevel 1 (
    echo.
    echo Python 3.13 installation failed. Trying Python 3.12...
    winget install --id Python.Python.3.12 -e --source winget --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo ERROR: Python installation with winget failed.
        echo Install Python manually from:
        echo https://www.python.org/downloads/windows/
        exit /b 1
    )
)

echo.
echo Python installation finished.
echo Trying to detect newly installed Python...
echo.

call :find_python
if errorlevel 1 (
    echo.
    echo Python was installed, but this terminal session cannot detect it yet.
    echo Close this window and run the BAT again.
    exit /b 1
)

exit /b 0


:ensure_python
call :find_python
if not errorlevel 1 exit /b 0

call :install_python
if errorlevel 1 exit /b 1

exit /b 0


:ensure_pip
echo Checking pip...
%PYTHON_CMD% -m pip --version >nul 2>nul
if not errorlevel 1 (
    echo pip OK.
    echo.
    exit /b 0
)

echo pip not found. Trying ensurepip...
%PYTHON_CMD% -m ensurepip --upgrade
%PYTHON_CMD% -m pip --version >nul 2>nul
if not errorlevel 1 (
    echo pip installed with ensurepip.
    echo.
    exit /b 0
)

echo.
echo ensurepip failed or is unavailable.
echo Trying get-pip.py bootstrap...
echo.

set "GETPIP_FILE=%TEMP%\get-pip.py"
if exist "%GETPIP_FILE%" del /q "%GETPIP_FILE%" >nul 2>nul

powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $env:TEMP\get-pip.py -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo ERROR: Could not download get-pip.py.
    echo Check internet connection or download manually:
    echo https://bootstrap.pypa.io/get-pip.py
    exit /b 1
)

%PYTHON_CMD% "%GETPIP_FILE%"
%PYTHON_CMD% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: pip installation failed even with get-pip.py.
    echo.
    echo Most likely this Python installation is incomplete or broken.
    echo Recommended fix:
    echo 1. Uninstall broken Python installations if needed.
    echo 2. Install Python 3.12 or 3.13 from python.org.
    echo 3. Enable "pip" and "Add python.exe to PATH".
    exit /b 1
)

echo pip installed with get-pip.py.
echo.
exit /b 0

