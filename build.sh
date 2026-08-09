#!/usr/bin/env bash
#
# OpenOrder — build the macOS app (the cross-platform twin of build.bat)
#
# Builds the frontend, then packages everything into a standalone
# OpenOrder.app with PyInstaller (windowed onedir → .app bundle).
#
# Output: dist/OpenOrder.app   (dist/ is gitignored; regenerable)
# The app stores user data (output/, settings.json, logs/) next to its
# executable inside the bundle by default; point it at a Dropbox folder
# via the in-app data-directory setting (settings.json "data_dir").
#
# Requires: ./setup.sh first (backend/.venv with pyinstaller installed).

set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

VENV_PYTHON="backend/.venv/bin/python"
APP="dist/OpenOrder.app"

if [ ! -x "$VENV_PYTHON" ]; then
  echo "  [X] backend venv not found — run ./setup.sh first."
  exit 1
fi

echo "  Building OpenOrder for macOS"
echo "  ============================"
echo

# --- 1. Frontend ------------------------------------------------------------
echo "  [1/6] Building frontend ..."
( cd frontend && npm run build --silent )
echo "        Done."

# --- 2. PyInstaller ---------------------------------------------------------
echo "  [2/6] Packaging with PyInstaller (this takes a minute) ..."
rm -rf "$APP" dist/OpenOrder build/OpenOrder
"$VENV_PYTHON" -m PyInstaller --noconfirm --onedir --windowed \
    --name "OpenOrder" \
    --icon "resources/images/openorder.icns" \
    --distpath dist \
    --add-data "frontend/dist:frontend/dist" \
    --add-data "resources:resources" \
    --add-data "backend/app:app" \
    --collect-all "fastapi" \
    --collect-all "uvicorn" \
    --collect-all "starlette" \
    --collect-all "pydantic" \
    --collect-all "python_multipart" \
    --collect-all "pythonbible" \
    --collect-all "httpx" \
    --collect-all "docx" \
    --collect-all "pptx" \
    --collect-all "webview" \
    --collect-all "PIL" \
    --hidden-import "app.main" \
    --hidden-import "app.paths" \
    --hidden-import "app.bulletin" \
    --hidden-import "app.slides" \
    --hidden-import "app.hymnal" \
    --hidden-import "app.scripture" \
    --hidden-import "app.models" \
    --hidden-import "app.church_config" \
    --hidden-import "app.calendar_data" \
    --hidden-import "app.themes" \
    --hidden-import "app.themes.oakpark" \
    openorder_app.py > build/pyinstaller.log 2>&1
echo "        Done."

# --- 3. Hymnal data ---------------------------------------------------------
# Ships beside the executable as the DEFAULT hymnal location for a fresh
# install. A user who keeps their hymnal elsewhere points hymnal_dir at it in
# Settings, and this copy is simply never read.
echo "  [3/6] Adding hymnal data ..."
if [ -d hymnal-json ] && [ -f hymnal-json/index.json ]; then
  cp -R hymnal-json "$APP/Contents/MacOS/hymnal-json"
  echo "        Copied."
else
  echo "        No hymnal-json/ found — app will run without hymn search."
fi

# --- 4. Stamp the version and build number ----------------------------------
# macOS reads Info.plist. CFBundleVersion is Apple's build string, so Get Info
# reads "Version 1.1.0 (7)" — semver alone can't tell two builds apart, which
# is exactly what you need when he says "it's doing X" and you must know which
# binary he's running.
echo "  [4/6] Stamping version ..."
VERSION="$(node -p "require('./frontend/package.json').version")"
BUILD_NUMBER=$(( $(cat build-number.txt 2>/dev/null || echo 0) + 1 ))
echo "$BUILD_NUMBER" > build-number.txt
/usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString $VERSION" "$APP/Contents/Info.plist" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :CFBundleShortVersionString string $VERSION" "$APP/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Set :CFBundleVersion $BUILD_NUMBER" "$APP/Contents/Info.plist" 2>/dev/null \
  || /usr/libexec/PlistBuddy -c "Add :CFBundleVersion string $BUILD_NUMBER" "$APP/Contents/Info.plist"
echo "        VERSION: $VERSION"
echo "        BUILD NUMBER: $BUILD_NUMBER"

# --- 5. Re-sign, LAST ------------------------------------------------------
# ORDER IS LOAD-BEARING: signing must follow every edit to the bundle (the
# hymnal copy and the plist stamp above), or the signature won't match the
# contents and macOS may refuse to launch it. Silence from `codesign -v` is
# the receipt; any output is a failure, and a bad signature is invisible from
# the outside — the app still copies and still launches on the machine that
# built it.
echo "  [5/6] Signing ..."
codesign --force --deep --sign - "$APP" 2>/dev/null
if ! codesign -v "$APP" 2>&1 | grep -q .; then
  echo "        Signature verified."
else
  echo "  [X] SIGNATURE INVALID — refusing to install."
  codesign -v "$APP"
  exit 1
fi

# --- 6. Install and verify the installed copy -------------------------------
# Copying is where corruption actually happens, and a stale app, a failed copy
# and a good install all look identical from outside — so read the version
# back out of the installed bundle rather than trusting that cp worked.
echo "  [6/6] Installing to /Applications ..."

# A running OpenOrder is replaced out from under itself by the rm -rf below,
# so quit it first. Jonathan's standing word (2026-08-09): "if the app is
# running you can close it. Better than being stuck when I leave my desk."
#
# GRACEFUL QUIT ONLY, and narrow by name (rules.md §12 — process-stops are
# destructive-tier). `quit app` lets any in-flight write to his data finish;
# a force-kill could truncate a service or the calendar mid-save. If it will
# not quit, ABORT rather than escalate — never -9 an app holding user data.
if pgrep -f "/Applications/OpenOrder.app" > /dev/null 2>&1; then
  echo "        OpenOrder is running — asking it to quit ..."
  osascript -e 'quit app "OpenOrder"' > /dev/null 2>&1 || true
  for _ in $(seq 1 10); do
    pgrep -f "/Applications/OpenOrder.app" > /dev/null 2>&1 || break
    sleep 1
  done
  if pgrep -f "/Applications/OpenOrder.app" > /dev/null 2>&1; then
    echo "  [X] OpenOrder would not quit. Close it by hand and run this again."
    exit 1
  fi
  echo "        Quit."
fi

rm -rf "/Applications/OpenOrder.app"
cp -R "$APP" "/Applications/OpenOrder.app"
INSTALLED_V="$(/usr/libexec/PlistBuddy -c "Print :CFBundleShortVersionString" "/Applications/OpenOrder.app/Contents/Info.plist" 2>/dev/null)"
INSTALLED_B="$(/usr/libexec/PlistBuddy -c "Print :CFBundleVersion" "/Applications/OpenOrder.app/Contents/Info.plist" 2>/dev/null)"
if [ "$INSTALLED_V" != "$VERSION" ] || [ "$INSTALLED_B" != "$BUILD_NUMBER" ]; then
  echo "  [X] INSTALL MISMATCH — /Applications has $INSTALLED_V ($INSTALLED_B),"
  echo "      expected $VERSION ($BUILD_NUMBER). The copy did not land."
  exit 1
fi
echo "        Installed and verified: $INSTALLED_V ($INSTALLED_B)"
# Nudge the icon cache — macOS caches Dock/Finder icons by bundle mtime, so a
# changed icon can otherwise keep showing the previous one.
touch "/Applications/OpenOrder.app"

echo
echo "  ============================"
echo "  Build complete:  /Applications/OpenOrder.app"
echo "  Version $VERSION (build $BUILD_NUMBER)"
echo
