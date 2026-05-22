@echo off
cd /d "%~dp0"

REM Launch the PowerShell build script
powershell -ExecutionPolicy Bypass -File "%~dp0build_windows.ps1"

REM If powershell failed to start at all
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to launch PowerShell script.
    echo Try right-clicking build_windows.ps1 and select "Run with PowerShell"
    pause
    exit /b 1
)
