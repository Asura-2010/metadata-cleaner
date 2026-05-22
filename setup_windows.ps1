# Metadata Cleaner - Setup Wizard (PowerShell)
$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

Write-Host ""
Write-Host "============================================"
Write-Host "  Metadata Cleaner - Setup Wizard"
Write-Host "============================================"
Write-Host ""

# Find Python
function Find-Python {
    try { & python --version 2>&1 | Out-Null; if ($LASTEXITCODE -eq 0) { Write-Host "[OK] Python found: python"; return "python" } } catch {}
    try { & py --version 2>&1 | Out-Null; if ($LASTEXITCODE -eq 0) { Write-Host "[OK] Python found: py"; return "py" } } catch {}

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
                if (Test-Path $exe) { Write-Host "[OK] Python found: $exe"; return $exe }
            }
        } catch {}
    }

    for ($minor = 13; $minor -ge 7; $minor--) {
        foreach ($hive in @("HKCU", "HKLM")) {
            $regPath = "${hive}:\SOFTWARE\Python\PythonCore\3.$minor\InstallPath"
            try {
                $installPath = (Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue).'(default)'
                if ($installPath) {
                    $exe = Join-Path $installPath "python.exe"
                    if (Test-Path $exe) { Write-Host "[OK] Python found (registry): $exe"; return $exe }
                }
            } catch {}
        }
    }

    return $null
}

$python = Find-Python

if (-not $python) {
    Write-Host ""
    Write-Host "Python 3 not found."
    Write-Host ""
    Write-Host "Install Python 3 from:"
    Write-Host "  https://mirrors.tuna.tsinghua.edu.cn/python/"
    Write-Host ""
    Write-Host "Then run this script again."
    Pause
    exit 1
}

Write-Host ""
Write-Host "[1/2] Upgrading pip..."
& $python -m pip install --upgrade pip 2>&1 | Out-Null

Write-Host "[2/2] Installing dependencies (trying mirrors)..."
Write-Host ""

$packages = @("pypdf", "PyPDF2", "Pillow", "pillow-heif", "tkinterdnd2")

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

    & $python $args
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "============================================="
        Write-Host "  Setup completed successfully!"
        Write-Host "============================================="
        Write-Host ""
        Write-Host "Usage:"
        Write-Host "  1. Double-click metadata_cleaner.py for GUI"
        Write-Host "  2. Drag files onto the script icon"
        Write-Host "  3. Click 'Add Files' in the GUI"
        Write-Host ""
        Pause
        exit 0
    }
    Write-Host "  Failed, trying next mirror..."
}

Write-Host ""
Write-Host "[X] All mirrors failed!"
Write-Host "Check your network connection and try again."
Pause
exit 1
