@echo off
REM ============================================================
REM Windows Release Build Script
REM 1. PyInstaller → .exe (--onedir)
REM 2. Inno Setup → setup.exe
REM
REM Prerequisites:
REM   pip install pyinstaller
REM   Inno Setup 6+ (https://jrsoftware.org/isinfo.php)
REM ============================================================
setlocal

cd /d "%~dp0"

for /f "tokens=*" %%i in ('python -c "from metadata_cleaner import __version__; print(__version__)"') do set VERSION=%%i
set APP_NAME=MetadataCleaner
set EXE_NAME=元数据清除工具

echo === Building %APP_NAME% v%VERSION% ===

REM Clean
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Step 1: Build with PyInstaller
echo.
echo --- PyInstaller: building .exe ---
pyinstaller ^
    --windowed ^
    --name "%EXE_NAME%" ^
    --icon icon.ico ^
    --clean ^
    metadata_cleaner.py

echo.
echo --- .exe built: dist\%EXE_NAME%\%EXE_NAME%.exe ---

REM Step 2: Run Inno Setup
echo.
echo --- Inno Setup: creating installer ---
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=%VERSION% installer.iss

echo.
echo === Done ===
echo dist\MetadataCleaner-Setup-%VERSION%.exe
