#!/usr/bin/env python3
"""Windows one-click setup — install Python dependencies for metadata-cleaner."""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent


def install_deps(python_exe="python"):
    """Install required packages with multi-mirror fallback."""
    print()
    print("=" * 45)
    print("   元数据清除工具 — 安装向导")
    print("=" * 45)
    print()

    print("[1/2] 升级 pip...")
    subprocess.run([python_exe, "-m", "pip", "install", "--upgrade", "pip"], check=False)
    print()

    print("[2/2] 安装项目依赖，依次尝试镜像源，请耐心等待...")
    print()

    packages = ["pypdf", "PyPDF2", "Pillow", "pillow-heif", "tkinterdnd2"]

    mirrors = [
        ("清华大学 https", "https://pypi.tuna.tsinghua.edu.cn/simple/", "pypi.tuna.tsinghua.edu.cn"),
        ("阿里云 https", "https://mirrors.aliyun.com/pypi/simple/", "mirrors.aliyun.com"),
        ("清华大学 http", "http://pypi.tuna.tsinghua.edu.cn/simple/", "pypi.tuna.tsinghua.edu.cn"),
        ("Python 官方源", None, None),
    ]

    for i, (name, url, trusted_host) in enumerate(mirrors, 1):
        print(f"  尝试镜像 {i}/4: {name}")
        cmd = [python_exe, "-m", "pip", "install"]
        if url:
            cmd += ["-i", url]
        if trusted_host:
            cmd += ["--trusted-host", trusted_host]
        cmd += packages

        result = subprocess.run(cmd, check=False)
        if result.returncode == 0:
            print()
            print("=" * 45)
            print("   安装完成！")
            print("=" * 45)
            print()
            print("  使用方法：")
            print()
            print("    1. 双击 metadata_cleaner.py 打开图形界面")
            print("    2. 将文件拖拽到脚本图标上批量处理")
            print("    3. 在图形界面点击「添加文件」选择文件")
            print()
            return

    print()
    print("  [X] 所有镜像源均安装失败！")
    print("  请检查网络连接，或手动运行：pip install pypdf")
    print()


def main():
    import os
    os.chdir(str(SCRIPT_DIR))
    install_deps()


if __name__ == "__main__":
    main()
