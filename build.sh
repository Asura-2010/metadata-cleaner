#!/bin/bash
# ============================================================
#  Unified Build Script — auto-detects OS, builds accordingly
# ============================================================
set -e

cd "$(dirname "$0")"
PROJECT_DIR="$(pwd)"

# ── Detect OS ──────────────────────────────────────────────
OS_TYPE="unknown"
case "$(uname -s)" in
    Darwin)  OS_TYPE="macos" ;;
    Linux)   OS_TYPE="linux" ;;
    MINGW*|MSYS*|CYGWIN*) OS_TYPE="windows" ;;
esac

echo "=== Metadata Cleaner Build ==="
echo "  OS:      ${OS_TYPE}"
echo "  Project: ${PROJECT_DIR}"

# ── Ensure venv exists ─────────────────────────────────────
if [ ! -f venv/bin/activate ]; then
    echo ""
    echo "--- Creating Python virtual environment ---"
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt pyinstaller
else
    source venv/bin/activate
fi

# ── Read version ───────────────────────────────────────────
VERSION=$(python3 -c "from metadata_cleaner import __version__; print(__version__)")
APP_NAME="MetadataCleaner"
echo "  Version: ${VERSION}"
echo ""

################################################################
#  macOS Build
################################################################
build_macos() {
    echo "--- Building macOS .app ---"
    rm -rf build dist

    pyinstaller \
        --windowed \
        --name "${APP_NAME}" \
        --osx-bundle-identifier "com.asura.metadata-cleaner" \
        --icon icon.icns \
        --clean \
        metadata_cleaner.py

    echo "--- .app built: dist/${APP_NAME}.app ---"

    # Create .dmg
    echo ""
    echo "--- Creating .dmg ---"
    mkdir -p dist/dmg
    cp -R "dist/${APP_NAME}.app" dist/dmg/
    ln -s /Applications dist/dmg/Applications

    hdiutil create \
        -volname "${APP_NAME}" \
        -srcfolder dist/dmg \
        -ov \
        -format UDZO \
        "dist/${APP_NAME}-${VERSION}-macOS.dmg"

    rm -rf dist/dmg build "${APP_NAME}.spec"

    echo ""
    echo "=== macOS build done ==="
    ls -lh dist/
}

################################################################
#  Windows Build (native, when running on Windows)
################################################################
build_windows() {
    echo "--- Building Windows .exe (onefile) ---"
    rm -rf build dist

    pyinstaller \
        --windowed \
        --onefile \
        --name "${APP_NAME}" \
        --icon icon.ico \
        --hidden-import pillow_heif \
        --hidden-import tkinterdnd2 \
        --add-data "icon.ico;." \
        --clean \
        --noconfirm \
        metadata_cleaner.py

    rm -rf build "${APP_NAME}.spec"

    # Rename output
    if [ -f "dist/${APP_NAME}.exe" ]; then
        mv "dist/${APP_NAME}.exe" "dist/${APP_NAME}-${VERSION}-Windows.exe"
    fi

    echo ""
    echo "=== Windows build done ==="
    ls -lh dist/
}

################################################################
#  Windows cross-build via Wine (macOS only)
################################################################
build_windows_via_wine() {
    local WINE_PYTHON="$1"

    echo "--- Cross-building Windows .exe via Wine ---"
    rm -rf build dist

    # Use the venv's pyinstaller but target Windows
    # PyInstaller on macOS can't cross-compile to Windows directly,
    # so we run Windows Python + PyInstaller inside Wine
    local WIN_REQ="pypdf PyPDF2 Pillow pillow-heif tkinterdnd2 pyinstaller"

    echo "  Installing Windows dependencies..."
    ${WINE_PYTHON} -m pip install --quiet ${WIN_REQ} 2>&1 | tail -3

    echo "  Running PyInstaller (Windows)..."
    ${WINE_PYTHON} -m PyInstaller \
        --windowed \
        --onefile \
        --name "${APP_NAME}" \
        --icon icon.ico \
        --hidden-import pillow_heif \
        --hidden-import tkinterdnd2 \
        --add-data "icon.ico;." \
        --clean \
        --noconfirm \
        metadata_cleaner.py

    rm -rf build "${APP_NAME}.spec"

    if [ -f "dist/${APP_NAME}.exe" ]; then
        mv "dist/${APP_NAME}.exe" "dist/${APP_NAME}-${VERSION}-Windows.exe"
    fi

    echo ""
    echo "=== Windows cross-build done ==="
    ls -lh dist/
}

# ── Main ───────────────────────────────────────────────────
case "${OS_TYPE}" in
    macos)
        build_macos

        # Check for Wine-based Windows cross-build
        WINE_PYTHON=""
        for wp in \
            "${HOME}/.wine/drive_c/Python313/python.exe" \
            "${HOME}/.wine/drive_c/Python312/python.exe" \
            "${HOME}/.wine/drive_c/Python311/python.exe" \
            "${HOME}/.wine/drive_c/users/${USER}/AppData/Local/Programs/Python/Python313/python.exe" \
            "${HOME}/.wine/drive_c/users/${USER}/AppData/Local/Programs/Python/Python312/python.exe" \
            "C:/Python313/python.exe" \
            "$(which wine 2>/dev/null && wine cmd /c 'where python' 2>/dev/null | head -1 | tr -d '\r')" \
            ; do
            if [ -n "${wp}" ] && [ -f "${wp}" ]; then
                WINE_PYTHON="wine ${wp}"
                break
            fi
        done

        if command -v wine &>/dev/null && [ -n "${WINE_PYTHON}" ]; then
            echo ""
            echo "--- Wine detected, attempting Windows cross-build ---"
            build_windows_via_wine "${WINE_PYTHON}" || echo "Wine build skipped (error occurred)"
        else
            echo ""
            echo "--- Wine not configured. Skipping Windows build. ---"
            echo "  To enable Windows cross-build from macOS:"
            echo "  1. brew install --cask wine-stable   (requires sudo password)"
            echo "  2. Download Windows Python from https://python.org"
            echo "  3. Install it under Wine: wine python-installer.exe"
            echo "  4. Re-run this script"
        fi
        ;;
    linux)
        echo "--- Building Linux binary (onefile) ---"
        rm -rf build dist
        pyinstaller \
            --windowed \
            --onefile \
            --name "${APP_NAME}" \
            --clean \
            --noconfirm \
            metadata_cleaner.py
        rm -rf build "${APP_NAME}.spec"
        if [ -f "dist/${APP_NAME}" ]; then
            mv "dist/${APP_NAME}" "dist/${APP_NAME}-${VERSION}-Linux"
        fi
        echo "=== Linux build done ==="
        ls -lh dist/
        ;;
    windows)
        build_windows
        ;;
    *)
        echo "Unknown OS. Supported: macOS, Linux, Windows (Git Bash/MSYS2)"
        exit 1
        ;;
esac

echo ""
echo "=== All builds complete ==="
ls -lh dist/ 2>/dev/null
