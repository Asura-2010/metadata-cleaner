@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo  =============================================
echo   元数据清除工具 — 一键打包 (Windows)
echo  =============================================
echo.

REM ============================================================
REM 1. 自动查找 Python
REM ============================================================
echo  [1/5] 正在查找 Python...

set PYTHON=
set PIP=

REM 1a. 先试试直接调用
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    set PIP=pip
    goto :found_python
)

REM 1b. 试试 py 启动器 (Windows 自带)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py
    set PIP=py -m pip
    goto :found_python
)

REM 1c. 扫描常见安装路径
for %%d in (
    "%LOCALAPPDATA%\Programs\Python"
    "%PROGRAMFILES%\Python*"
    "%PROGRAMFILES(X86)%\Python*"
    "C:\Python*"
) do (
    for /d %%p in (%%d) do (
        if exist "%%p\python.exe" (
            set PYTHON="%%p\python.exe"
            set PIP="%%p\python.exe" -m pip
            goto :found_python
        )
    )
)

REM 1d. 用 where 搜索
for /f "delims=" %%i in ('where python 2^>nul') do (
    set PYTHON="%%i"
    set PIP="%%i" -m pip
    goto :found_python
)

echo  [错误] 未找到 Python！
echo.
echo  请先安装 Python 3.8+，安装时务必勾选 "Add Python to PATH"
echo.
echo  清华大学镜像下载 (国内快):
echo  https://mirrors.tuna.tsinghua.edu.cn/python/
echo.
pause
exit /b 1

:found_python
%PYTHON% --version
echo  [OK] Python 就绪
echo.

REM ============================================================
REM 2. 配置 pip 国内镜像
REM ============================================================
echo  [2/5] 配置 pip 清华镜像 (下载更快)...
%PYTHON% -m pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>&1
echo  [OK]
echo.

REM ============================================================
REM 3. 安装依赖
REM ============================================================
echo  [3/5] 安装项目依赖...
echo.
%PYTHON% -m pip install --upgrade pip --quiet
%PYTHON% -m pip install pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller
if %errorlevel% neq 0 (
    echo.
    echo  [警告] 清华镜像失败，尝试阿里云镜像...
    %PYTHON% -m pip install pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller -i https://mirrors.aliyun.com/pypi/simple/
    if %errorlevel% neq 0 (
        echo.
        echo  [错误] 依赖安装失败，请检查网络连接
        pause
        exit /b 1
    )
)
echo.
echo  [OK] 依赖安装完成
echo.

REM ============================================================
REM 4. 读取版本号
REM ============================================================
echo  [4/5] 读取版本号...
for /f "tokens=*" %%i in ('%PYTHON% -c "from metadata_cleaner import __version__; print(__version__)"') do set VERSION=%%i
echo  版本: v%VERSION%
echo.

REM ============================================================
REM 5. PyInstaller 打包
REM ============================================================
echo  [5/5] PyInstaller 打包中 (预计 2-5 分钟，请耐心等待)...
echo.

REM 清理旧构建
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

%PYTHON% -m PyInstaller ^
    --windowed ^
    --name "元数据清除工具" ^
    --icon icon.ico ^
    --hidden-import pillow_heif ^
    --hidden-import tkinterdnd2 ^
    --add-data "icon.ico;." ^
    --clean ^
    --noconfirm ^
    metadata_cleaner.py

if %errorlevel% neq 0 (
    echo.
    echo  [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo  =============================================
echo   打包成功！
echo  =============================================
echo.
echo   输出位置: dist\元数据清除工具\元数据清除工具.exe
echo.
echo   将此文件夹完整复制到其他电脑即可直接运行，
echo   无需安装 Python。
echo.

REM ============================================================
REM 可选: Inno Setup 生成安装包
REM ============================================================
for %%d in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
) do (
    if exist %%d (
        echo  [附加] 检测到 Inno Setup，正在生成安装包...
        %%d /DMyAppVersion=%VERSION% installer.iss
        if %errorlevel% equ 0 (
            echo.
            echo   安装包: dist\MetadataCleaner-Setup-%VERSION%.exe
        )
        goto :done
    )
)

echo  如需 .exe 安装包，可安装 Inno Setup 6 后重新运行本脚本:
echo  https://jrsoftware.org/isinfo.php

:done
echo.
pause
