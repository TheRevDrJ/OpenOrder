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
echo "  [1/3] Building frontend ..."
( cd frontend && npm run build --silent )
echo "        Done."

# --- 2. PyInstaller ---------------------------------------------------------
echo "  [2/3] Packaging with PyInstaller (this takes a minute) ..."
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
# The app reads user data from the executable's directory inside the bundle
# (see backend/app/paths.py), so the hymnal rides along there.
echo "  [3/3] Adding hymnal data ..."
if [ -d hymnal-json ] && [ -f hymnal-json/index.json ]; then
  cp -R hymnal-json "$APP/Contents/MacOS/hymnal-json"
  echo "        Copied."
else
  echo "        No hymnal-json/ found — app will run without hymn search."
fi

echo
echo "  ============================"
echo "  Build complete:  $APP"
echo "  Launch it with:  open $APP"
echo
