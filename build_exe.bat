@echo off
setlocal EnableExtensions
title Build Dreamkas EXE
cd /d "%~dp0"

call :ensure_python
if errorlevel 1 goto fail

call :ensure_pip
if errorlevel 1 goto fail

echo Installing build dependencies...
%PYTHON_CMD% -m pip install -r requirements.txt
if errorlevel 1 goto fail

echo.
echo Building console EXE...
%PYTHON_CMD% -m PyInstaller --onefile --name DreamkasReceipt --add-data "dreamkas_receipt_template.xlsx;." dreamkas_receipt.py
if errorlevel 1 goto fail

echo.
echo Building GUI EXE...
%PYTHON_CMD% -m PyInstaller --onefile --windowed --name DreamkasReceiptGUI --add-data "dreamkas_receipt_template.xlsx;." dreamkas_gui.py
if errorlevel 1 goto fail

echo.
echo Done. EXE files are in the dist folder.
pause
exit /b 0

:fail
echo.
echo ERROR: Build failed.
pause
exit /b 1


:is_admin
net session >nul 2>nul
if errorlevel 1 exit /b 1
exit /b 0


:try_python_file
set "TEST_FILE=%~1"
if "%TEST_FILE%"=="" exit /b 1
if not exist "%TEST_FILE%" exit /b 1

"%TEST_FILE%" -c "import sys, encodings; print(str(sys.version_info.major)+'.'+str(sys.version_info.minor)+'.'+str(sys.version_info.micro))" > "%TEMP%\dreamkas_pycheck.txt" 2> "%TEMP%\dreamkas_pycheck_error.txt"
if errorlevel 1 (
    echo Python runtime check failed:
    echo %TEST_FILE%
    if exist "%TEMP%\dreamkas_pycheck_error.txt" type "%TEMP%\dreamkas_pycheck_error.txt"
    exit /b 1
)

set "TEST_VERSION="
for /f "usebackq tokens=1 delims=" %%A in ("%TEMP%\dreamkas_pycheck.txt") do (
    if not defined TEST_VERSION set "TEST_VERSION=%%A"
)

if "%TEST_VERSION%"=="" exit /b 1

for /f "tokens=1,2 delims=." %%A in ("%TEST_VERSION%") do (
    set "TEST_MAJOR=%%A"
    set "TEST_MINOR=%%B"
)

if "%TEST_MAJOR%"=="" exit /b 1
if %TEST_MAJOR% LSS 3 exit /b 1
if %TEST_MAJOR% EQU 3 if %TEST_MINOR% LSS 10 exit /b 1

set "PYTHON_CMD=%TEST_FILE%"
set "PYTHON_VERSION=%TEST_VERSION%"
exit /b 0


:find_python
set "PYTHON_CMD="
set "PYTHON_VERSION="

REM 1. Main hardcoded Python path.
if exist "C:\Python314\python.exe" (
    call :try_python_file "C:\Python314\python.exe"
    if not errorlevel 1 goto python_found
)

REM 2. Fallback saved path, if user created it earlier.
if exist "python_path.txt" (
    for /f "usebackq tokens=* delims=" %%P in ("python_path.txt") do (
        if not "%%P"=="" (
            call :try_python_file "%%P"
            if not errorlevel 1 goto python_found
        )
    )
)

REM 3. Fallback common system paths.
for %%P in (
"C:\Python313\python.exe"
"C:\Python312\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
"%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
) do (
    call :try_python_file "%%~P"
    if not errorlevel 1 goto python_found
)

exit /b 1

:python_found
echo Python command: %PYTHON_CMD%
echo Python version: %PYTHON_VERSION%
%PYTHON_CMD% -c "import sys; print('Python exe: ' + sys.executable)"
echo.
echo %PYTHON_CMD%>python_path.txt
exit /b 0


:download_python_installer
set "PY_INSTALLER=%TEMP%\dreamkas-python-3.13.13-amd64.exe"

echo Removing old cached Python installer if it exists...
if exist "%PY_INSTALLER%" del /f /q "%PY_INSTALLER%" >nul 2>nul

echo Downloading Python installer fresh...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.13.13/python-3.13.13-amd64.exe' -OutFile $env:TEMP\dreamkas-python-3.13.13-amd64.exe -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo ERROR: Could not download Python installer.
    exit /b 1
)

if not exist "%PY_INSTALLER%" (
    echo ERROR: Download finished but installer file was not created:
    echo %PY_INSTALLER%
    exit /b 1
)

for %%F in ("%PY_INSTALLER%") do echo Downloaded installer size: %%~zF bytes
exit /b 0


:download_embedded_zip
set "EMBED_ZIP=%TEMP%\dreamkas-python-3.13.13-embed-amd64.zip"

echo Removing old cached embedded Python ZIP if it exists...
if exist "%EMBED_ZIP%" del /f /q "%EMBED_ZIP%" >nul 2>nul

echo Downloading embedded Python ZIP fresh...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.13.13/python-3.13.13-embed-amd64.zip' -OutFile $env:TEMP\dreamkas-python-3.13.13-embed-amd64.zip -UseBasicParsing; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo ERROR: Could not download embedded Python ZIP.
    exit /b 1
)

if not exist "%EMBED_ZIP%" (
    echo ERROR: Download finished but embedded ZIP was not created:
    echo %EMBED_ZIP%
    exit /b 1
)

for %%F in ("%EMBED_ZIP%") do echo Downloaded embedded ZIP size: %%~zF bytes
exit /b 0


:show_installer_log
set "PY_INSTALL_LOG=%TEMP%\dreamkas-python-install.log"
if exist "%PY_INSTALL_LOG%" (
    echo.
    echo Last lines of installer log:
    echo ------------------------------------------------------------
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-Content -Path $env:TEMP\dreamkas-python-install.log -Tail 100"
    echo ------------------------------------------------------------
) else (
    echo Installer log not found:
    echo %PY_INSTALL_LOG%
)
exit /b 0


:clean_c_python
set "PY_TARGET=C:\Python314"
if exist "%PY_TARGET%\python.exe" (
    call :try_python_file "%PY_TARGET%\python.exe"
    if not errorlevel 1 exit /b 0
)
if exist "%PY_TARGET%" (
    echo Removing broken/incomplete folder:
    echo %PY_TARGET%
    rmdir /s /q "%PY_TARGET%" >nul 2>nul
    if exist "%PY_TARGET%" (
        echo ERROR: Could not remove %PY_TARGET%.
        echo Run this BAT as Administrator or delete C:\Python314 manually.
        exit /b 1
    )
)
exit /b 0


:install_python_with_exe
set "PY_TARGET=C:\Python314"
set "PY_INSTALLER=%TEMP%\dreamkas-python-3.13.13-amd64.exe"
set "PY_INSTALL_LOG=%TEMP%\dreamkas-python-install.log"

call :download_python_installer
if errorlevel 1 exit /b 1

if exist "%PY_INSTALL_LOG%" del /f /q "%PY_INSTALL_LOG%" >nul 2>nul
if not exist "%PY_TARGET%" mkdir "%PY_TARGET%" >nul 2>nul

echo Running Python installer with explicit full feature set...
echo TargetDir=%PY_TARGET%
echo Log=%PY_INSTALL_LOG%
echo.

REM Important: Include_exe=1 and Include_lib=1 are explicit.
REM Without Include_exe, the installer can leave python.exe / Lib incomplete in maintenance mode.
"%PY_INSTALLER%" /quiet InstallAllUsers=1 TargetDir="%PY_TARGET%" Include_core=1 Include_exe=1 Include_lib=1 Include_pip=1 Include_dev=0 Include_doc=0 Include_tcltk=0 Include_tools=0 Include_test=0 Include_launcher=0 InstallLauncherAllUsers=0 PrependPath=0 Shortcuts=0 /log "%PY_INSTALL_LOG%"

set "INSTALL_EXIT=%ERRORLEVEL%"
echo Python installer exit code: %INSTALL_EXIT%

timeout /t 5 /nobreak >nul

if not exist "%PY_TARGET%\python.exe" (
    echo.
    echo ERROR: Python installer did not create:
    echo %PY_TARGET%\python.exe
    call :show_installer_log
    exit /b 1
)

call :try_python_file "%PY_TARGET%\python.exe"
if errorlevel 1 (
    echo ERROR: Python installer created invalid Python.
    echo Trying embedded ZIP fallback next.
    call :show_installer_log
    exit /b 1
)

echo C:\Python314\python.exe>python_path.txt
echo Python installed successfully into C:\Python314
echo.
exit /b 0


:install_python_embedded_to_c
set "PY_TARGET=C:\Python314"
set "EMBED_ZIP=%TEMP%\dreamkas-python-3.13.13-embed-amd64.zip"

echo.
echo Installing embedded Python into C:\Python314...
echo.

if exist "%PY_TARGET%" (
    echo Removing current C:\Python314 before embedded install...
    rmdir /s /q "%PY_TARGET%" >nul 2>nul
)

mkdir "%PY_TARGET%" >nul 2>nul
if not exist "%PY_TARGET%" (
    echo ERROR: Could not create C:\Python314.
    echo Run this BAT as Administrator.
    exit /b 1
)

call :download_embedded_zip
if errorlevel 1 exit /b 1

echo Extracting embedded Python to C:\Python314...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Expand-Archive -Path $env:TEMP\dreamkas-python-3.13.13-embed-amd64.zip -DestinationPath 'C:\Python314' -Force; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"
if errorlevel 1 (
    echo ERROR: Could not extract embedded Python ZIP.
    exit /b 1
)

REM Enable site module in python313._pth.
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { $pth = Get-ChildItem -Path 'C:\Python314' -Filter 'python*._pth' | Select-Object -First 1; if($pth){ (Get-Content $pth.FullName) -replace '#import site','import site' | Set-Content $pth.FullName -Encoding ASCII }; exit 0 } catch { Write-Host $_.Exception.Message; exit 1 }"

if not exist "%PY_TARGET%\python.exe" (
    echo ERROR: Embedded Python did not create C:\Python314\python.exe.
    exit /b 1
)

call :try_python_file "%PY_TARGET%\python.exe"
if errorlevel 1 (
    echo ERROR: Embedded Python is invalid.
    exit /b 1
)

echo C:\Python314\python.exe>python_path.txt
echo Embedded Python installed successfully into C:\Python314
echo.
exit /b 0


:install_python_to_c
echo Valid Python 3.10+ was not found.
echo.
echo Installing Python into:
echo C:\Python314
echo.
echo This is a fixed path requested for this tool.
echo For C:\Python314 installation, Administrator rights are recommended.
echo.

call :is_admin
if errorlevel 1 (
    echo WARNING: This BAT is not running as Administrator.
    echo Installation into C:\Python314 may fail.
    echo.
)

call :clean_c_python
if errorlevel 1 exit /b 1

call :install_python_with_exe
if not errorlevel 1 exit /b 0

echo.
echo Standard Python installer did not produce a valid Python.
echo Switching to embedded ZIP fallback.
echo.

call :install_python_embedded_to_c
if not errorlevel 1 exit /b 0

echo.
echo ERROR: Both installer and embedded ZIP fallback failed.
echo.
exit /b 1


:ensure_python
call :find_python
if not errorlevel 1 exit /b 0

call :install_python_to_c
if errorlevel 1 exit /b 1

call :find_python
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
    exit /b 1
)

%PYTHON_CMD% "%GETPIP_FILE%"
%PYTHON_CMD% -m pip --version >nul 2>nul
if errorlevel 1 (
    echo.
    echo ERROR: pip installation failed even with get-pip.py.
    exit /b 1
)

echo pip installed with get-pip.py.
echo.
exit /b 0

