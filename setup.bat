@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo   Metadata Cleaner - Setup Wizard
echo ============================================
echo.

REM ---- Step 1: Try python command ----
python --version >nul 2>&1
if !errorlevel! equ 0 (
    echo [OK] Python found ^(system PATH^)
    python setup_windows.py
    pause
    exit /b
)

REM ---- Step 2: Try py launcher ----
py --version >nul 2>&1
if !errorlevel! equ 0 (
    echo [OK] Python found ^(py launcher^)
    py setup_windows.py
    pause
    exit /b
)

REM ---- Step 3: Scan registry ----
for /f "tokens=2*" %%a in ('reg query "HKCU\SOFTWARE\Python\PythonCore\3.13\InstallPath" /ve 2^>nul ^| find "REG_SZ"') do set "PYEXE=%%bpython.exe"
if defined PYEXE if exist "!PYEXE!" (
    echo [OK] Python found ^(registry^)
    "!PYEXE!" setup_windows.py
    pause
    exit /b
)

for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Python\PythonCore\3.13\InstallPath" /ve 2^>nul ^| find "REG_SZ"') do set "PYEXE=%%bpython.exe"
if defined PYEXE if exist "!PYEXE!" (
    echo [OK] Python found ^(registry^)
    "!PYEXE!" setup_windows.py
    pause
    exit /b
)

REM ---- Step 4: Scan common folders ----
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%PROGRAMFILES%\Python313"
    "C:\Python313"
) do (
    if exist "%%~d\python.exe" set "PYEXE=%%~d\python.exe"
)
if defined PYEXE (
    echo [OK] Python found ^(file scan^)
    "!PYEXE!" setup_windows.py
    pause
    exit /b
)

REM ---- Not found ----
echo Python 3 not found.
echo.
echo Please install Python 3 first:
echo   https://mirrors.tuna.tsinghua.edu.cn/python/
echo.
echo IMPORTANT: Check [v] Add Python to PATH during install.
echo.
pause
exit /b 1
