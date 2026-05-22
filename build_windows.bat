@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo   Metadata Cleaner - One-Click Build
echo ============================================
echo.

REM ---- Step 1: Try direct commands ----
python --version >nul 2>&1
if !errorlevel! equ 0 set "PYTHON=python" & goto :run

py --version >nul 2>&1
if !errorlevel! equ 0 set "PYTHON=py" & goto :run

REM ---- Step 2: where command ----
for /f "delims=" %%i in ('where python 2^>nul') do (
    set "PYTHON=%%i"
    goto :run
)

REM ---- Step 3: Scan registry for any Python 3.7-3.13 ----
for /l %%v in (13,-1,7) do (
    for %%r in (
        "HKCU\SOFTWARE\Python\PythonCore\3.%%v\InstallPath"
        "HKLM\SOFTWARE\Python\PythonCore\3.%%v\InstallPath"
    ) do (
        for /f "tokens=2*" %%a in ('reg query %%~r /ve 2^>nul ^| find "REG_SZ"') do (
            if exist "%%bpython.exe" (
                set "PYTHON=%%bpython.exe"
                goto :run
            )
        )
    )
)

REM ---- Step 4: Scan file system for any Python3* folder ----
for %%d in (
    "%LOCALAPPDATA%\Programs\Python"
    "%PROGRAMFILES%"
    "C:\"
) do (
    for /d %%p in ("%%~d\Python3*") do (
        if exist "%%p\python.exe" (
            set "PYTHON=%%p\python.exe"
            goto :run
        )
    )
)

REM ---- Python NOT found - download and install ----
echo.
echo Python 3 was not found on this computer.
echo.
echo The script will download Python 3.13 from Tsinghua mirror.
echo IMPORTANT: Check [v] Add Python to PATH during install!
echo.
choice /c yn /n /m "Download and install Python 3.13? [Y]es [N]o: "
if !errorlevel! equ 2 exit /b 1

set "URL=https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/python-3.13.5-amd64.exe"
set "INSTALLER=%TEMP%\python-3.13.5-amd64.exe"

echo.
echo Downloading Python 3.13 from Tsinghua mirror...
powershell -Command "Invoke-WebRequest -Uri '!URL!' -OutFile '!INSTALLER!'"

if not exist "!INSTALLER!" (
    echo Tsinghua failed, trying Huawei mirror...
    set "URL=https://repo.huaweicloud.com/python/3.13.5/python-3.13.5-amd64.exe"
    powershell -Command "Invoke-WebRequest -Uri '!URL!' -OutFile '!INSTALLER!'"
)

if not exist "!INSTALLER!" (
    echo.
    echo Download failed. Install Python manually:
    echo   https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/
    pause
    exit /b 1
)

echo.
echo Installing Python 3.13...
echo [!!] Make sure to CHECK: [v] Add Python to PATH
"!INSTALLER!" /passive PrependPath=1 InstallAllUsers=0 Include_test=0

echo.
echo Wait for installation to finish, then press any key...
pause >nul

REM After install - try python directly
python --version >nul 2>&1
if !errorlevel! equ 0 set "PYTHON=python" & goto :run

REM File scan after fresh install
for %%d in (
    "%LOCALAPPDATA%\Programs\Python"
    "%PROGRAMFILES%"
    "C:\"
) do (
    for /d %%p in ("%%~d\Python3*") do (
        if exist "%%p\python.exe" set "PYTHON=%%p\python.exe"
    )
)
if defined PYTHON goto :run

echo.
echo Python still not detected after install.
echo Try restarting your computer and run this script again.
pause
exit /b 1

:run
echo.
echo [OK] Python: !PYTHON!
echo Running build script...
echo.

"!PYTHON!" build_windows.py 2>&1
set "EXITCODE=!errorlevel!"
echo.
echo ----------------------------------------
if !EXITCODE! equ 0 (
    echo Build completed successfully.
) else (
    echo Build script exited with code: !EXITCODE!
    echo If you see errors above, check the messages.
)
echo ----------------------------------------
echo.
pause
exit /b !EXITCODE!
