#!/usr/bin/env python3
"""Windows one-click build script — handles Python detection, dependency install, and PyInstaller packaging."""

import subprocess
import sys
import os
import shutil
import tempfile
import urllib.request
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
OS_NAME = os.name  # 'nt' on Windows

# ── Chinese messages (safe in Python Unicode) ──────────────────────
MSG = {
    "banner": "元数据清除工具 — 一键打包",
    "checking": "[1/5] 检查 Python 环境...",
    "python_ok": "[OK] Python 就绪",
    "upgrade_pip": "[2/5] 升级 pip...",
    "install_deps": "[3/5] 安装项目依赖，依次尝试镜像源，请耐心等待...",
    "try_mirror": "尝试镜像 {n}/4: {name}",
    "deps_ok": "[OK] 依赖安装成功",
    "deps_fail_title": "[X] 所有镜像源均安装失败！",
    "deps_fail_desc": [
        "请检查：",
        "1. 网络连接是否正常",
        "2. 是否开启了代理/VPN 导致镜像不可达",
        "3. 公司防火墙是否阻止了 Python 联网",
        "手动测试: pip install pypdf",
    ],
    "read_version": "[4/5] 读取版本号...",
    "version": "版本: v{ver}",
    "clean": "清理旧构建...",
    "build": "[5/5] PyInstaller 打包中，预计 2-5 分钟，请勿关闭此窗口...",
    "build_ok": "打包成功！",
    "build_fail": "[错误] 打包失败",
    "output_path": r"dist\元数据清除工具\元数据清除工具.exe",
    "output_desc": '复制整个"元数据清除工具"文件夹到其他电脑即可直接运行，无需安装 Python。',
    "download_title": "未检测到 Python，从华为云镜像下载安装...",
    "download_progress": "正在下载 Python 3.13，华为云镜像，约 26MB...",
    "download_fail": "自动下载失败，请手动安装 Python",
    "installer_hint": "安装时请务必勾选 [v] Add Python to PATH",
    "done": "按回车键退出...",
}


def print_box(lines, width=46):
    """Print a box with the given lines."""
    print(f"+{'=' * width}+")
    for line in lines:
        print(f"| {line:<{width}} |")
    print(f"+{'=' * width}+")


def find_python():
    """Find a usable Python interpreter. Returns path string or None."""
    candidates = []

    # Check common python commands
    for cmd in ["python", "python3", "py"]:
        try:
            result = subprocess.run(
                [cmd, "--version"], capture_output=True, timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW if OS_NAME == "nt" else 0
            )
            if result.returncode == 0:
                candidates.append(cmd)
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass

    # Scan common Windows install paths
    if OS_NAME == "nt":
        search_roots = []
        local_appdata = os.environ.get("LOCALAPPDATA", "")
        program_files = os.environ.get("PROGRAMFILES", "")
        program_files_x86 = os.environ.get("PROGRAMFILES(X86)", "")

        if local_appdata:
            py_dir = Path(local_appdata) / "Programs" / "Python"
            if py_dir.exists():
                search_roots.append(py_dir)
        if program_files:
            search_roots.append(Path(program_files))
        if program_files_x86:
            search_roots.append(Path(program_files_x86))
        search_roots.append(Path("C:\\"))

        for root in search_roots:
            try:
                for d in root.glob("Python*"):
                    if d.is_dir():
                        exe = d / "python.exe"
                        if exe.exists():
                            candidates.append(str(exe))
            except PermissionError:
                pass

    return candidates[0] if candidates else None


def install_python_windows():
    """Download and install Python 3.13 on Windows."""
    mirrors = [
        ("华为云", "https://repo.huaweicloud.com/python/3.13.5/python-3.13.5-amd64.exe"),
        ("清华", "https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/python-3.13.5-amd64.exe"),
    ]

    installer_path = Path(tempfile.gettempdir()) / "python-3.13.5-amd64.exe"

    for name, url in mirrors:
        print(f"\n  下载 Python 3.13：{name}镜像...")
        try:
            urllib.request.urlretrieve(url, str(installer_path))
            if installer_path.exists() and installer_path.stat().st_size > 1000000:
                break
        except Exception as e:
            print(f"  {name}镜像下载失败: {e}")
            continue

    if not installer_path.exists() or installer_path.stat().st_size < 1000000:
        print("\n  自动下载失败")
        return False

    print(f"  下载完成，正在启动安装程序...")
    print(f"  {MSG['installer_hint']}")
    print()

    subprocess.run(
        [str(installer_path), "/passive", "PrependPath=1", "InstallAllUsers=0", "Include_test=0"],
        check=False
    )

    # Clean up installer
    try:
        installer_path.unlink()
    except Exception:
        pass

    input("\n  安装完成后按回车继续...")
    return True


def install_dependencies(python_exe):
    """Install required packages with multi-mirror fallback."""
    # Upgrade pip first
    print(f"\n{MSG['upgrade_pip']}")
    subprocess.run(
        [python_exe, "-m", "pip", "install", "--upgrade", "pip"],
        check=False
    )

    print(f"\n{MSG['install_deps']}")

    packages = ["pypdf", "PyPDF2", "Pillow", "pillow-heif", "tkinterdnd2", "pyinstaller"]

    mirrors = [
        ("清华大学 https", f"https://pypi.tuna.tsinghua.edu.cn/simple/", "pypi.tuna.tsinghua.edu.cn"),
        ("阿里云 https", f"https://mirrors.aliyun.com/pypi/simple/", "mirrors.aliyun.com"),
        ("清华大学 http", f"http://pypi.tuna.tsinghua.edu.cn/simple/", "pypi.tuna.tsinghua.edu.cn"),
        ("Python 官方源", None, None),
    ]

    for i, (name, url, trusted_host) in enumerate(mirrors, 1):
        print(f"\n  {MSG['try_mirror'].format(n=i, name=name)}")
        cmd = [python_exe, "-m", "pip", "install"]
        if url:
            cmd += ["-i", url]
        if trusted_host:
            cmd += ["--trusted-host", trusted_host]
        cmd += packages

        result = subprocess.run(cmd, check=False)
        if result.returncode == 0:
            return True

    return False


def read_version(python_exe):
    """Read version from metadata_cleaner module."""
    print(f"\n{MSG['read_version']}")
    try:
        result = subprocess.run(
            [python_exe, "-c", "from metadata_cleaner import __version__; print(__version__)"],
            capture_output=True, text=True, cwd=str(SCRIPT_DIR)
        )
        version = result.stdout.strip()
        if version:
            print(f"  {MSG['version'].format(ver=version)}")
            return version
    except Exception:
        pass
    return "unknown"


def clean_build():
    """Remove old build artifacts."""
    print(f"\n{MSG['clean']}")
    for d in ["build", "dist"]:
        path = SCRIPT_DIR / d
        if path.exists():
            shutil.rmtree(path, ignore_errors=True)


def run_pyinstaller(python_exe):
    """Run PyInstaller to create the executable."""
    print(f"\n{MSG['build']}\n")

    cmd = [
        python_exe, "-m", "PyInstaller",
        "--windowed",
        "--name", "元数据清除工具",
        "--icon", str(SCRIPT_DIR / "icon.ico"),
        "--hidden-import", "pillow_heif",
        "--hidden-import", "tkinterdnd2",
        "--add-data", f"icon.ico;.",
        "--clean",
        "--noconfirm",
        str(SCRIPT_DIR / "metadata_cleaner.py"),
    ]

    result = subprocess.run(cmd, cwd=str(SCRIPT_DIR), check=False)
    return result.returncode == 0


def main():
    os.chdir(str(SCRIPT_DIR))

    print()
    print("=" * 50)
    print(f"  {MSG['banner']}")
    print("=" * 50)
    print()

    # Step 1: Find or install Python
    print(MSG["checking"])
    python_exe = find_python()

    if not python_exe:
        if OS_NAME == "nt":
            print(f"\n  {MSG['download_title']}")
            if install_python_windows():
                # Re-scan for Python
                python_exe = find_python()
        if not python_exe:
            print("\n  Python 未找到，请手动安装后重试")
            input(MSG["done"])
            return

    print(f"\n  {MSG['python_ok']}  ({python_exe})")

    # Step 2 & 3: Install dependencies
    if not install_dependencies(python_exe):
        print()
        print_box(MSG["deps_fail_desc"])
        input()
        return

    print(f"\n  {MSG['deps_ok']}")

    # Step 4: Read version
    version = read_version(python_exe)

    # Step 5: Clean + Build
    clean_build()

    success = run_pyinstaller(python_exe)

    if success:
        output_dir = SCRIPT_DIR / "dist" / "元数据清除工具"
        print()
        print_box([
            MSG["build_ok"],
            "",
            MSG["output_path"],
            "",
            MSG["output_desc"],
        ])
        print(f"\n  输出位置: {output_dir}")
    else:
        print(f"\n  {MSG['build_fail']}")

    input(f"\n{MSG['done']}")


if __name__ == "__main__":
    main()
