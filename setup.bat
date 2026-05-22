@echo off
cd /d "%~dp0"

REM Launch the PowerShell setup script
powershell -ExecutionPolicy Bypass -File "%~dp0setup_windows.ps1"

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to launch PowerShell script.
    echo Try right-clicking setup_windows.ps1 and select "Run with PowerShell"
    pause
    exit /b 1
)
