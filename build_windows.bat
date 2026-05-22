@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

REM ============================================================
REM 元数据清除工具 — Windows 打包脚本 (生成 .exe 安装包)
REM
REM 用途: 在 Windows 上将 Python 脚本打包成可分发的安装包
REM
REM 前提: 已安装 Python 3.8+
REM
REM 步骤:
REM   1. 自动安装 PyInstaller
REM   2. PyInstaller 打包 → dist\元数据清除工具\元数据清除工具.exe
REM   3. 如已安装 Inno Setup 6，自动生成 setup.exe 安装包
REM ============================================================

cd /d "%~dp0"

REM --- 读取版本号 ---
for /f "tokens=*" %%i in ('python -c "from metadata_cleaner import __version__; print(__version__)"') do set VERSION=%%i

set APP_NAME=MetadataCleaner
set EXE_NAME=元数据清除工具

echo.
echo  =========================================
echo   %APP_NAME% v%VERSION%  Windows 打包脚本
echo  =========================================
echo.

REM ============================================================
REM Step 0: Check Python
REM ============================================================
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [错误] 未检测到 Python，请先安装 Python 3.8+
    echo  下载: https://mirrors.tuna.tsinghua.edu.cn/python/
    pause
    exit /b 1
)
python --version
echo.

REM ============================================================
REM Step 1: Install PyInstaller if missing
REM ============================================================
echo  [1/3] 检查 PyInstaller...
pyinstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  正在安装 PyInstaller...
    pip install pyinstaller
    if %errorlevel% neq 0 (
        echo  [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
)
echo  [OK] PyInstaller 就绪
echo.

REM ============================================================
REM Step 2: Clean old builds
REM ============================================================
echo  [2/3] 清理旧构建...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo  [OK] 清理完成
echo.

REM ============================================================
REM Step 3: PyInstaller build
REM ============================================================
echo  [3/3] PyInstaller 打包中 (预计 1-3 分钟) ...
echo.

pyinstaller ^
    --windowed ^
    --name "%EXE_NAME%" ^
    --icon icon.ico ^
    --hidden-import pillow_heif ^
    --hidden-import tkinterdnd2 ^
    --add-data "icon.ico;." ^
    --clean ^
    --noconfirm ^
    metadata_cleaner.py

if %errorlevel% neq 0 (
    echo.
    echo  [错误] PyInstaller 打包失败
    pause
    exit /b 1
)

echo.
echo  =========================================
echo   PyInstaller 打包完成
echo   dist\%EXE_NAME%\%EXE_NAME%.exe
echo  =========================================
echo.

REM ============================================================
REM Step 4 (optional): Inno Setup installer
REM ============================================================
set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if not exist %ISCC% (
    set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
)

if exist %ISCC% (
    echo  [可选] 使用 Inno Setup 生成安装包...
    echo.
    %ISCC% /DMyAppVersion=%VERSION% installer.iss
    if %errorlevel% equ 0 (
        echo.
        echo  =========================================
        echo   安装包已生成
        echo   dist\%APP_NAME%-Setup-%VERSION%.exe
        echo  =========================================
    ) else (
        echo  [警告] Inno Setup 生成失败，但 .exe 文件夹已可用
    )
) else (
    echo  [提示] 未检测到 Inno Setup 6
    echo  如需生成安装包，请安装: https://jrsoftware.org/isinfo.php
    echo.
    echo  当前可直接使用: dist\%EXE_NAME%\%EXE_NAME%.exe
)

echo.
echo  === 打包流程结束 ===
pause
