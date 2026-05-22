@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo   Metadata Cleaner - One-Click Build
echo ============================================
echo.

REM ---- Step 1: Try python command ----
echo [1] Checking Python...
python --version >nul 2>&1
if !errorlevel! equ 0 (
    echo [OK] Python found ^(system PATH^)
    python build_windows.py
    pause
    exit /b
)

REM ---- Step 2: Try py launcher ----
py --version >nul 2>&1
if !errorlevel! equ 0 (
    echo [OK] Python found ^(py launcher^)
    py build_windows.py
    pause
    exit /b
)

REM ---- Step 3: Scan registry for Python installs ----
for /f "tokens=2*" %%a in ('reg query "HKCU\SOFTWARE\Python\PythonCore\3.13\InstallPath" /ve 2^>nul ^| find "REG_SZ"') do set "PYEXE=%%bpython.exe"
if defined PYEXE if exist "!PYEXE!" (
    echo [OK] Python found ^(registry^)
    "!PYEXE!" build_windows.py
    pause
    exit /b
)

for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Python\PythonCore\3.13\InstallPath" /ve 2^>nul ^| find "REG_SZ"') do set "PYEXE=%%bpython.exe"
if defined PYEXE if exist "!PYEXE!" (
    echo [OK] Python found ^(registry^)
    "!PYEXE!" build_windows.py
    pause
    exit /b
)

REM ---- Step 4: Scan common install folders ----
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%PROGRAMFILES%\Python313"
    "C:\Python313"
) do (
    if exist "%%~d\python.exe" (
        set "PYEXE=%%~d\python.exe"
    )
)
if defined PYEXE (
    echo [OK] Python found ^(file scan^)
    "!PYEXE!" build_windows.py
    pause
    exit /b
)

REM ---- Python NOT found - download and install ----
echo.
echo Python 3.13 is required but was not found.
echo.
echo The script will now download Python from a China mirror.
echo Make sure to check [v] Add Python to PATH during install!
echo.
choice /c yn /n /m "Download and install Python 3.13? [Y]es [N]o: "
if !errorlevel! equ 2 exit /b 1

echo.
echo Downloading Python 3.13 from Tsinghua mirror...
echo This may take a few minutes...
echo.

set "INSTALLER=%TEMP%\python-3.13.5-amd64.exe"
set "URL=https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/python-3.13.5-amd64.exe"

powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '!URL!' -OutFile '!INSTALLER!'"

if not exist "!INSTALLER!" (
    echo Tsinghua mirror failed, trying backup URL...
    set "URL=https://repo.huaweicloud.com/python/3.13.5/python-3.13.5-amd64.exe"
    powershell -Command "Invoke-WebRequest -Uri '!URL!' -OutFile '!INSTALLER!'"
)

if not exist "!INSTALLER!" (
    echo Download failed. Please install Python manually:
    echo   https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/
    pause
    exit /b 1
)

echo Download complete. Launching installer...
echo [!!] CRITICAL: Check [v] Add Python to PATH
echo.

"!INSTALLER!" /passive PrependPath=1 InstallAllUsers=0 Include_test=0

echo.
echo Installation finished. Press any key to continue...
pause >nul

REM Re-check for Python after install
python --version >nul 2>&1
if !errorlevel! equ 0 (
    python build_windows.py
    pause
    exit /b
)

REM Scan again after fresh install
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%PROGRAMFILES%\Python313"
    "C:\Python313"
) do (
    if exist "%%~d\python.exe" set "PYEXE=%%~d\python.exe"
)
if defined PYEXE (
    "!PYEXE!" build_windows.py
    pause
    exit /b
)

echo Python still not detected after install.
echo Try restarting your computer and run this script again.
echo.
echo Or install manually from:
echo   https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/
pause
exit /b 1
