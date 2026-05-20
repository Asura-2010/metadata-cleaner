@echo off
title 元数据清除工具 — 安装向导

REM ===================================================
REM Windows 一键安装脚本 (含国内镜像)
REM ===================================================

REM Force UTF-8 codepage for Chinese text
chcp 65001 >nul 2>&1

:check_python
cls
echo.
echo  =========================================
echo     元数据清除工具 — 安装向导
echo  =========================================
echo.
echo  [检测] 正在检查 Python 环境...

python --version >nul 2>&1
if %errorlevel% equ 0 (
    python --version
    echo.
    echo  Python 已安装，即将开始配置...
    timeout /t 2 /nobreak >nul
    goto :install_deps
)

REM ---- Python NOT found ----
echo.
echo  ╔══════════════════════════════════════════╗
echo  ║  [X] 未检测到 Python 3                  ║
echo  ║                                          ║
echo  ║  Python 是运行本工具所必需的环境。       ║
echo  ║  Python 官网在国内较慢，建议用镜像下载。 ║
echo  ╚══════════════════════════════════════════╝
echo.
echo  ┌──────────────────────────────────────────┐
echo  │  推荐下载源（点击链接即可下载安装包）：  │
echo  └──────────────────────────────────────────┘
echo.
echo    [国内镜像 — 速度快]
echo.
echo    1. 清华大学镜像  (推荐)
echo       https://mirrors.tuna.tsinghua.edu.cn/python/
echo.
echo    2. 华为云镜像  (推荐)
echo       https://mirrors.huaweicloud.com/python/
echo.
echo.
echo    [海外官方源]
echo.
echo    3. Python 官方网站
echo       https://www.python.org/downloads/
echo.

REM Launch download page
powershell -Command "
    Add-Type -AssemblyName System.Windows.Forms
    $result = [System.Windows.Forms.MessageBox]::Show(
        '需要安装 Python 3。' + [char]10 + [char]10 +
        '是否打开清华大学镜像下载页面？' + [char]10 +
        '(国内下载速度更快)',
        '元数据清除工具 — Python 未安装',
        'YesNo',
        'Question'
    )
    if ($result -eq 'Yes') {
        Start-Process 'https://mirrors.tuna.tsinghua.edu.cn/python/'
    }
" >nul 2>&1

echo.
echo  ┌───────────────────────────────────────────┐
echo  │  重要提示：                                │
echo  │                                            │
echo  │  安装 Python 时请务必勾选：                 │
echo  │  「Add Python to PATH」（添加到环境变量）   │
echo  │                                            │
echo  │  安装完成后，请重新运行本脚本继续配置。     │
echo  └───────────────────────────────────────────┘
echo.
pause
exit /b 0

:install_deps
echo.
echo  [安装] 正在配置 pip 国内镜像...
python -m pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>&1
echo  [OK] pip 镜像已切换为清华大学源

echo.
echo  [安装] 正在安装依赖库 (defusedxml, pypdf)...
echo.
python -m pip install --upgrade pip
python -m pip install defusedxml pypdf

if %errorlevel% neq 0 (
    echo.
    echo  [警告] 安装失败，尝试切换到备用镜像...
    python -m pip install defusedxml pypdf -i https://mirrors.aliyun.com/pypi/simple/
)

echo.
echo  =========================================
echo    安装完成！
echo  =========================================
echo.
echo  使用方法：
echo.
echo    1. 双击 metadata_cleaner.py 打开图形界面
echo    2. 将文件拖拽到脚本图标上批量处理
echo    3. 在图形界面点击"添加文件"选择文件
echo.
echo  现在可以关闭此窗口。
echo.
pause
