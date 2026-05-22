@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

set PYTHON=

REM 1. Try python command
python --version >nul 2>&1
if %errorlevel% equ 0 set PYTHON=python

REM 2. Try py launcher
if not defined PYTHON (
    py --version >nul 2>&1
    if %errorlevel% equ 0 set PYTHON=py
)

REM 3. Scan common install paths
if not defined PYTHON (
    for /d %%d in (
        "%LOCALAPPDATA%\Programs\Python\Python3*"
        "%PROGRAMFILES%\Python3*"
        "%PROGRAMFILES(X86)%\Python3*"
        "C:\Python3*"
    ) do (
        if exist "%%d\python.exe" (
            set PYTHON="%%d\python.exe"
        )
    )
)

REM 4. where command fallback
if not defined PYTHON (
    for /f "delims=" %%i in ('where python 2^>nul') do (
        if not defined PYTHON set PYTHON="%%i"
    )
)

if defined PYTHON (
    echo Python found: %PYTHON%
    %PYTHON% build_windows.py
    pause
    exit
)

echo Python not found.
echo.
echo If Python is installed, please check:
echo   1. Was "Add Python to PATH" checked during install?
echo   2. Try reinstalling from:
echo      https://mirrors.tuna.tsinghua.edu.cn/python/
echo.
pause
