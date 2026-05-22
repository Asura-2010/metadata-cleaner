@echo off
cd /d "%~dp0"

python --version >nul 2>&1
if %errorlevel% equ 0 (
    python setup_windows.py
    pause
    exit
)

py --version >nul 2>&1
if %errorlevel% equ 0 (
    py setup_windows.py
    pause
    exit
)

echo Python not found.
echo Please install Python 3 first:
echo   https://mirrors.tuna.tsinghua.edu.cn/python/
echo.
echo Or download from official site:
echo   https://www.python.org/downloads/
echo.
echo IMPORTANT: Check "Add Python to PATH" during installation.
echo.
pause
