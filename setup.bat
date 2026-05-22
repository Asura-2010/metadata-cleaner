@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo   Metadata Cleaner - Setup Wizard
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

echo.
echo Python 3 not found.
echo.
echo Install Python 3 from:
echo   https://mirrors.tuna.tsinghua.edu.cn/python/
echo.
echo Then run this script again.
pause
exit /b 1

:run
echo.
echo [OK] Python: !PYTHON!
echo Running setup...
echo.

"!PYTHON!" "%~dp0setup_windows.py" 2>&1
set "EXITCODE=!errorlevel!"
echo.
echo ----------------------------------------
if !EXITCODE! equ 0 (
    echo Setup completed successfully.
) else (
    echo Setup script exited with code: !EXITCODE!
)
echo ----------------------------------------
echo.
pause
exit /b !EXITCODE!
