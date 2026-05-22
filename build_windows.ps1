# Metadata Cleaner - One-Click Build (PowerShell)
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host ""
Write-Host "============================================"
Write-Host "  Metadata Cleaner - One-Click Build"
Write-Host "============================================"
Write-Host ""
Write-Host "Script dir: $ScriptDir"
Write-Host "Current dir: $(Get-Location)"
Write-Host ""

# Check build_windows.py exists
if (-not (Test-Path "$ScriptDir\build_windows.py")) {
    Write-Host "[ERROR] build_windows.py not found!"
    Write-Host ""
    Write-Host "Files in $ScriptDir :"
    Get-ChildItem $ScriptDir | ForEach-Object { Write-Host "  $($_.Name)" }
    Pause
    exit 1
}

# Find Python
function Find-Python {
    # Method 1: python in PATH
    try {
        $result = & python --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Python found (PATH): python"
            return "python"
        }
    } catch {}

    # Method 2: py launcher
    try {
        $result = & py --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Python found (py launcher): py"
            return "py"
        }
    } catch {}

    # Method 3: Search common install locations
    $searchPaths = @(
        "$env:LOCALAPPDATA\Programs\Python"
        $env:PROGRAMFILES
        "${env:PROGRAMFILES} (x86)"
        "C:\"
    )
    foreach ($root in $searchPaths) {
        if (-not $root) { continue }
        try {
            $dirs = Get-ChildItem -Path $root -Directory -Filter "Python3*" -ErrorAction SilentlyContinue | Sort-Object Name -Descending
            foreach ($dir in $dirs) {
                $exe = Join-Path $dir.FullName "python.exe"
                if (Test-Path $exe) {
                    Write-Host "[OK] Python found: $exe"
                    return $exe
                }
            }
        } catch {}
    }

    # Method 4: Registry
    for ($minor = 13; $minor -ge 7; $minor--) {
        foreach ($hive in @("HKCU", "HKLM")) {
            $regPath = "${hive}:\SOFTWARE\Python\PythonCore\3.$minor\InstallPath"
            try {
                $installPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).'(default)'
                if ($installPath) {
                    $exe = Join-Path $installPath "python.exe"
                    if (Test-Path $exe) {
                        Write-Host "[OK] Python found (registry): $exe"
                        return $exe
                    }
                }
            } catch {}
        }
    }

    return $null
}

function Install-Dependencies($python) {
    Write-Host ""
    Write-Host "[1/3] Upgrading pip..."
    & $python -m pip install --upgrade pip 2>&1 | Out-Null

    Write-Host "[2/3] Installing dependencies (trying mirrors)..."
    $packages = @("pypdf", "PyPDF2", "Pillow", "pillow-heif", "tkinterdnd2", "pyinstaller")

    $mirrors = @(
        @{Name="Tsinghua HTTPS"; Url="https://pypi.tuna.tsinghua.edu.cn/simple/"; Trusted="pypi.tuna.tsinghua.edu.cn"},
        @{Name="Aliyun HTTPS"; Url="https://mirrors.aliyun.com/pypi/simple/"; Trusted="mirrors.aliyun.com"},
        @{Name="Tsinghua HTTP"; Url="http://pypi.tuna.tsinghua.edu.cn/simple/"; Trusted="pypi.tuna.tsinghua.edu.cn"},
        @{Name="PyPI official"; Url=$null; Trusted=$null}
    )

    foreach ($mirror in $mirrors) {
        Write-Host "  Trying: $($mirror.Name)..."
        $args = @("-m", "pip", "install")
        if ($mirror.Url) {
            $args += "-i", $mirror.Url, "--trusted-host", $mirror.Trusted
        }
        $args += $packages

        & $python $args 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[OK] Dependencies installed via $($mirror.Name)"
            return $true
        }
    }

    Write-Host "[X] All mirrors failed!"
    return $false
}

# ---- Main ----
$python = Find-Python

if (-not $python) {
    Write-Host ""
    Write-Host "Python 3 not found."
    Write-Host ""
    $choice = Read-Host "Download Python 3.13 from Tsinghua mirror? [Y]es [N]o"
    if ($choice -notmatch '^[Yy]') { Pause; exit 1 }

    $installerUrl = "https://mirrors.tuna.tsinghua.edu.cn/python/3.13.5/python-3.13.5-amd64.exe"
    $installerPath = "$env:TEMP\python-3.13.5-amd64.exe"

    Write-Host "Downloading Python 3.13..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath

    if (-not (Test-Path $installerPath)) {
        Write-Host "Tsinghua failed, trying Huawei mirror..."
        $installerUrl = "https://repo.huaweicloud.com/python/3.13.5/python-3.13.5-amd64.exe"
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath
    }

    if (-not (Test-Path $installerPath)) {
        Write-Host "Download failed."
        Pause
        exit 1
    }

    Write-Host "Installing Python 3.13..."
    Write-Host "[!!] Make sure to CHECK: [v] Add Python to PATH"
    Start-Process -FilePath $installerPath -ArgumentList "/passive", "PrependPath=1", "InstallAllUsers=0", "Include_test=0" -Wait

    Write-Host "Installation complete. Press any key to continue..."
    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

    # Refresh PATH and try again
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    $python = Find-Python

    if (-not $python) {
        Write-Host "Python still not found after install."
        Pause
        exit 1
    }
}

Write-Host ""
Write-Host "[OK] Using Python: $python"
Write-Host "Test: $(& $python --version 2>&1)"
Write-Host ""

# Install dependencies
if (-not (Install-Dependencies $python)) {
    Write-Host ""
    Write-Host "Dependency install failed. Check:"
    Write-Host "  1. Network connection"
    Write-Host "  2. VPN/proxy settings"
    Write-Host "  Manual test: pip install pypdf"
    Pause
    exit 1
}

# Read version
Write-Host ""
Write-Host "[3/3] Reading version..."
try {
    $version = & $python -c "from metadata_cleaner import __version__; print(__version__)" 2>&1
    Write-Host "  Version: v$version"
} catch {
    Write-Host "  Version: unknown"
}

# Clean old builds
Write-Host ""
Write-Host "Cleaning old builds..."
$toClean = @("$ScriptDir\build", "$ScriptDir\dist")
foreach ($p in $toClean) {
    if (Test-Path $p) {
        Remove-Item -Recurse -Force $p -ErrorAction SilentlyContinue
        Write-Host "  Removed: $p"
    }
}

# Build with PyInstaller
Write-Host ""
Write-Host "Running PyInstaller... (2-5 minutes, do not close this window)"
Write-Host ""

$pyiArgs = @(
    "-m", "PyInstaller",
    "--windowed",
    "--name", "MetadataCleaner",
    "--icon", "$ScriptDir\icon.ico",
    "--hidden-import", "pillow_heif",
    "--hidden-import", "tkinterdnd2",
    "--add-data", "icon.ico;.",
    "--clean",
    "--noconfirm",
    "$ScriptDir\metadata_cleaner.py"
)

& $python $pyiArgs

if ($LASTEXITCODE -eq 0) {
    # Clean up build artifacts
    Write-Host ""
    Write-Host "[OK] PyInstaller finished."
    Write-Host "Current dir: $(Get-Location)"
    Write-Host "Looking for:"
    $cleanup = @("$ScriptDir\build", "$ScriptDir\MetadataCleaner.spec")
    foreach ($p in $cleanup) {
        Write-Host "  Check: $p"
        if (Test-Path $p) {
            Write-Host "  -> exists, removing..."
            Remove-Item -Recurse -Force $p
            Write-Host "  -> Removed: $p"
        } else {
            Write-Host "  -> NOT found"
        }
    }

    Write-Host ""
    Write-Host "Files in $ScriptDir after cleanup:"
    Get-ChildItem $ScriptDir | ForEach-Object { Write-Host "  $($_.Name)" }
    Write-Host ""
    Write-Host "+==============================================+"
    Write-Host "| Build successful!                            |"
    Write-Host "|                                              |"
    Write-Host "| dist\MetadataCleaner\MetadataCleaner.exe      |"
    Write-Host "|                                              |"
    Write-Host "| Copy the MetadataCleaner folder to other PCs |"
    Write-Host "| to run directly, no Python needed.           |"
    Write-Host "+==============================================+"
    Write-Host ""
    Write-Host "Output: $ScriptDir\dist\MetadataCleaner\"
} else {
    Write-Host ""
    Write-Host "[X] Build failed!"
}

Pause
