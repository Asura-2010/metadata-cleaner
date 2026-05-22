#!/bin/bash
# ===================================================
# macOS / Linux 一键安装脚本 (含国内镜像)
# ===================================================

set -e

clear
echo ""
echo "========================================="
echo "  元数据清除工具 — 安装向导"
echo "========================================="
echo ""
echo "[检测] 正在检查 Python 环境..."

# ---- Check Python 3 ----
if command -v python3 &>/dev/null; then
    PYTHON=python3
elif command -v python &>/dev/null; then
    PYTHON=python
else
    echo ""
    echo "╔══════════════════════════════════════════╗"
    echo "║  ✗ 未检测到 Python 3                    ║"
    echo "╚══════════════════════════════════════════╝"
    echo ""
    echo "Python 是运行本工具所必需的环境。"
    echo "Python 官网在国内较慢，建议用镜像下载。"
    echo ""
    echo "┌──────────────────────────────────────────┐"
    echo "│  推荐下载源                              │"
    echo "└──────────────────────────────────────────┘"
    echo ""
    echo "  1. 清华大学镜像 (推荐)"
    echo "     https://mirrors.tuna.tsinghua.edu.cn/python/"
    echo ""
    echo "  2. 华为云镜像"
    echo "     https://mirrors.huaweicloud.com/python/"
    echo ""

    # Show macOS dialog
    osascript -e 'display dialog "需要安装 Python 3。\n\n是否打开下载页面？\n\n推荐使用清华大学镜像，国内下载速度更快。" buttons {"取消", "打开清华镜像", "打开官网"} default button "打开清华镜像" with icon caution with title "元数据清除工具 — Python 未安装"' 2>/dev/null

    # Detect which button was clicked
    BTN=$?
    if [ $BTN -eq 0 ]; then
        # Try to detect the actual button, default to Tsinghua
        open "https://mirrors.tuna.tsinghua.edu.cn/python/" 2>/dev/null || \
        open "https://www.python.org/downloads/" 2>/dev/null
    fi

    echo ""
    echo "┌─────────────────────────────────────────┐"
    echo "│  下载安装 Python 3 后，请重新运行本脚本 │"
    echo "└─────────────────────────────────────────┘"
    exit 0
fi

echo "[OK] $($PYTHON --version)"

# ---- Configure pip mirror ----
echo ""
echo "[配置] 正在设置 pip 国内镜像..."
$PYTHON -m pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple 2>/dev/null || true
echo "[OK] pip 镜像已切换为清华大学源"

# ---- Install dependencies ----
echo ""
echo "[安装] 正在安装依赖库 (defusedxml, pypdf)..."
$PYTHON -m pip install --upgrade pip
$PYTHON -m pip install defusedxml pypdf tkinterdnd2

if [ $? -ne 0 ]; then
    echo ""
    echo "[警告] 默认镜像安装失败，切换备用镜像..."
    $PYTHON -m pip install defusedxml pypdf tkinterdnd2 -i https://mirrors.aliyun.com/pypi/simple/
fi

# ---- Done ----
echo ""
echo "========================================="
echo "  安装完成！"
echo "========================================="
echo ""
echo "使用方法："
echo ""
echo "  1. 双击 metadata_cleaner.py 打开图形界面"
echo "  2. 将文件拖拽到脚本图标上批量处理"
echo "  3. 在终端运行: python3 metadata_cleaner.py"
echo "     (带文件路径可批量处理: python3 metadata_cleaner.py *.docx)"
echo ""
