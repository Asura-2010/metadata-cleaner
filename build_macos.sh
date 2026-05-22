#!/bin/bash
# ============================================================
# macOS Release Build Script
# 1. PyInstaller → .app bundle
# 2. hdiutil   → .dmg image
# ============================================================
set -e

cd "$(dirname "$0")"

# Activate virtual environment
if [ -f venv/bin/activate ]; then
    source venv/bin/activate
else
    echo "ERROR: venv not found. Run: python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt pyinstaller"
    exit 1
fi

VERSION=$(python3 -c "from metadata_cleaner import __version__; print(__version__)")
APP_NAME="MetadataCleaner"
DMG_NAME="MetadataCleaner-${VERSION}"

echo "=== Building ${APP_NAME} v${VERSION} ==="

# Clean previous builds
rm -rf build dist

# Step 1: Build .app with PyInstaller
echo ""
echo "--- PyInstaller: building .app ---"
pyinstaller \
    --windowed \
    --name "${APP_NAME}" \
    --osx-bundle-identifier "com.asura.metadata-cleaner" \
    --icon icon.icns \
    --clean \
    metadata_cleaner.py

echo ""
echo "--- .app built: dist/${APP_NAME}.app ---"

# Step 2: Create .dmg
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
    "dist/${DMG_NAME}.dmg"

# Cleanup staging
rm -rf dist/dmg

# Cleanup build artifacts
rm -rf build "${APP_NAME}.spec"

echo ""
echo "=== Done: dist/${DMG_NAME}.dmg ==="
ls -lh "dist/${DMG_NAME}.dmg"
