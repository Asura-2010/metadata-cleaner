@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title 元数据清除工具 — 一键打包

echo.
echo  =============================================
echo   元数据清除工具 — 一键打包 Windows
echo  =============================================
echo.

REM ============================================================
REM 1. 查找或安装 Python
REM ============================================================
echo  [1/5] 检查 Python 环境...
echo.

set PYTHON=
set NEED_INSTALL=0

REM 1a. 直接调用
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    goto :found_python
)

REM 1b. py 启动器
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py
    goto :found_python
)

REM 1c. 扫描常见路径
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\*"
    "%PROGRAMFILES%\Python*"
    "%PROGRAMFILES(X86)%\Python*"
    "C:\Python*"
) do (
    for /d %%p in (%%d) do (
        if exist "%%p\python.exe" (
            set PYTHON="%%p\python.exe"
            goto :found_python
        )
    )
)

REM 1d. where 兜底
for /f "delims=" %%i in ('where python 2^>nul') do (
    set PYTHON="%%i"
    goto :found_python
)

REM ---- Python 未找到，开始自动安装 ----
echo  [X] 未检测到 Python
echo.
echo  +==============================================+
echo  |  即将从国内镜像自动下载并安装 Python 3.13  |
echo  |                                              |
echo  |  !! 安装时请务必勾选：                       |
echo  |  [v] Add Python to PATH，添加到环境变量     |
echo  |                                              |
echo  |  安装完成后本脚本会自动继续                   |
echo  +==============================================+
echo.
choice /c yn /n /m "是否继续安装 Python? [Y]是 [N]否: "
if errorlevel 2 exit /b 1

set PYTHON_URL=https://mirrors.huaweicloud.com/python/3.13.5/python-3.13.5-amd64.exe
set PYTHON_INSTALLER=%TEMP%\python-3.13.5-amd64.exe

echo.
echo  正在下载 Python 3.13，华为云镜像，约 26MB...
echo  如果下载慢请耐心等待...
echo.
powershell -Command "& {$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%'}" 2>&1

if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo  华为云下载失败, 尝试清华镜像...
    set PYTHON_URL=https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/python-3.13.5-amd64.exe
    powershell -Command "& {$ProgressPreference='SilentlyContinue'; Invoke-WebRequest -Uri '!PYTHON_URL!' -OutFile '%PYTHON_INSTALLER%'}" 2>&1
)

if not exist "%PYTHON_INSTALLER%" (
    echo.
    echo  自动下载失败，请手动安装 Python:
    echo  清华镜像: https://mirrors.tuna.tsinghua.edu.cn/python/
    echo  华为云镜像: https://mirrors.huaweicloud.com/python/
    echo.
    echo  安装时务必勾选 "Add Python to PATH"
    echo  装好后重新运行本脚本即可。
    pause
    exit /b 1
)

echo  下载完成，正在启动安装程序...
echo.
echo  +------------------------------------------+
echo  | !! 重要 !! 请勾选:                        |
echo  | [v] Add Python to PATH                   |
echo  |  安装界面底部那个复选框                  |
echo  +------------------------------------------+
echo.

REM /passive = 显示进度条但不弹确认框; PrependPath=1 自动加PATH
"%PYTHON_INSTALLER%" /passive PrependPath=1 InstallAllUsers=0 Include_test=0

echo.
echo  等待 Python 安装完成...
echo  如果安装窗口已关闭，说明安装完毕。
echo.
echo  按任意键继续检查 Python 是否装好...
pause >nul

REM 重新扫描 Python
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    goto :found_python
)

py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py
    goto :found_python
)

REM 再次扫描路径（安装后可能新出现）
for %%d in (
    "%LOCALAPPDATA%\Programs\Python\*"
    "%PROGRAMFILES%\Python*"
    "%PROGRAMFILES(X86)%\Python*"
    "C:\Python*"
) do (
    for /d %%p in (%%d) do (
        if exist "%%p\python.exe" (
            set PYTHON="%%p\python.exe"
            goto :found_python
        )
    )
)

echo.
echo  Python 安装后仍未检测到，可能需要重启电脑。
echo  重启后重新运行本脚本即可。
pause
exit /b 1

:found_python
echo.
%PYTHON% --version
echo  [OK] Python 就绪
echo.

REM ============================================================
REM 2. 安装依赖 (依次尝试多个镜像源)
REM ============================================================
echo  [2/5] 升级 pip...
echo.
%PYTHON% -m pip install --upgrade pip
echo.

echo  [3/5] 安装项目依赖：依次尝试镜像源，请耐心等待...
echo.

echo  尝试镜像 1/4: 清华大学 https
%PYTHON% -m pip install --trusted-host pypi.tuna.tsinghua.edu.cn -i https://pypi.tuna.tsinghua.edu.cn/simple/ pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller
if !errorlevel! equ 0 goto :deps_ok

echo  尝试镜像 2/4: 阿里云 https
%PYTHON% -m pip install --trusted-host mirrors.aliyun.com -i https://mirrors.aliyun.com/pypi/simple/ pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller
if !errorlevel! equ 0 goto :deps_ok

echo  尝试镜像 3/4: 清华大学 http，部分网络屏蔽https
%PYTHON% -m pip install --trusted-host pypi.tuna.tsinghua.edu.cn -i http://pypi.tuna.tsinghua.edu.cn/simple/ pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller
if !errorlevel! equ 0 goto :deps_ok

echo  尝试镜像 4/4: Python 官方源，可能较慢
%PYTHON% -m pip install pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller
if !errorlevel! equ 0 goto :deps_ok

echo.
echo  +==========================================+
echo  |  [X] 所有镜像源均安装失败！               |
echo  |                                          |
echo  |  请检查：                                 |
echo  |  1. 网络连接是否正常                       |
echo  |  2. 是否开启了代理/VPN导致镜像不可达       |
echo  |  3. 公司防火墙是否阻止了 Python 联网       |
echo  |                                          |
echo  |  手动测试: pip install pypdf              |
echo  +==========================================+
pause
exit /b 1

:deps_ok
echo  [OK] 依赖安装成功
echo.


REM ============================================================
echo  [4/5] 读取版本号...
%PYTHON% -c "from metadata_cleaner import __version__; print(__version__)" > "%TEMP%\mc_version.tmp"
set /p VERSION=<"%TEMP%\mc_version.tmp"
del "%TEMP%\mc_version.tmp" 2>nul
echo  版本: v%VERSION%
echo.

REM ============================================================
REM 5. 清理旧构建
REM ============================================================
echo  清理旧构建...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
echo  [OK]
echo.

REM ============================================================
REM 6. PyInstaller 打包
REM ============================================================
echo  [5/5] PyInstaller 打包中，预计 2-5 分钟，请勿关闭此窗口...
echo  看到 "Building EXE" 字样表示正在工作
echo.

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
echo  +==========================================+
echo  |         打 包 成 功 ！                    |
echo  |                                          |
echo  |  dist\元数据清除工具\元数据清除工具.exe    |
echo  |                                          |
echo  |  复制整个"元数据清除工具"文件夹到         |
echo  |  其他电脑即可直接运行，无需Python。       |
echo  +==========================================+
echo.

REM ============================================================
REM 可选: Inno Setup
REM ============================================================
for %%d in (
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    "C:\Program Files\Inno Setup 6\ISCC.exe"
) do (
    if exist %%d (
        echo  检测到 Inno Setup, 正在生成安装包...
        %%d /DMyAppVersion=%VERSION% installer.iss
        if !errorlevel! equ 0 (
            echo  安装包: dist\MetadataCleaner-Setup-%VERSION%.exe
        )
        goto :done
    )
)

echo  如需 .exe 安装包，安装 Inno Setup 6 后重跑本脚本:
echo  https://jrsoftware.org/isinfo.php

:done
echo.
echo  按任意键退出...
pause >nul
