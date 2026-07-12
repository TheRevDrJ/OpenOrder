#!/usr/bin/env bash
#
# OpenOrder — one-time setup (macOS / Linux)
#
# OpenOrder has two halves:
#   Frontend — a Vite/React app in ./frontend         (needs Node.js + npm)
#   Backend  — a FastAPI app in ./backend             (needs Python 3 + a venv)
#
# This checks the toolchain is present and new enough, then installs both:
#   • creates ./backend/.venv and pip-installs backend/requirements.txt
#   • npm-installs ./frontend
#
# After this:  ./openorder.sh start

set -euo pipefail
cd "$(dirname "${BASH_SOURCE[0]}")"

MIN_NODE_MAJOR=20
MIN_PY_MINOR=10         # Python 3.10+

echo "OpenOrder setup"
echo

# --- Node.js -----------------------------------------------------------------
if ! command -v node >/dev/null 2>&1; then
  echo "  [X] Node.js is not installed."
  echo "      Install it, then run this again:"
  echo "        macOS:  brew install node"
  echo "        or:     https://nodejs.org/"
  exit 1
fi
NODE_VER="$(node -v)"                 # e.g. v22.0.0
NODE_MAJOR="${NODE_VER#v}"; NODE_MAJOR="${NODE_MAJOR%%.*}"
if [ "$NODE_MAJOR" -lt "$MIN_NODE_MAJOR" ]; then
  echo "  [X] Node $NODE_VER found — OpenOrder needs Node ${MIN_NODE_MAJOR}+."
  echo "      Upgrade:  brew upgrade node   (or https://nodejs.org/)"
  exit 1
fi
echo "  [OK] Node $NODE_VER"

# --- npm (ships with Node) ---------------------------------------------------
if ! command -v npm >/dev/null 2>&1; then
  echo "  [X] npm not found — reinstall Node (npm ships with it)."
  exit 1
fi
echo "  [OK] npm $(npm -v)"

# --- Python 3 ----------------------------------------------------------------
# Pick the NEWEST interpreter that satisfies 3.MIN+ — checking versioned names
# first, because a bare `python3` is often an old system build (e.g. macOS ships
# 3.9 via Xcode) while a newer `python3.11` sits alongside it.
PY=""
for cand in python3.13 python3.12 python3.11 python3.10 python3 python; do
  if command -v "$cand" >/dev/null 2>&1 \
     && "$cand" -c "import sys; sys.exit(0 if sys.version_info[:2] >= (3, $MIN_PY_MINOR) else 1)" 2>/dev/null; then
    PY="$cand"; break
  fi
done
if [ -z "$PY" ]; then
  echo "  [X] No Python 3.${MIN_PY_MINOR}+ found (the backend needs 3.${MIN_PY_MINOR}+)."
  if command -v python3 >/dev/null 2>&1; then
    echo "      (Found $(python3 --version 2>&1), which is too old.)"
  fi
  echo "      Install a newer one, then run this again:"
  echo "        macOS:  brew install python@3.12"
  echo "        or:     https://www.python.org/downloads/"
  exit 1
fi
echo "  [OK] $("$PY" --version)  ($PY)"

# --- Backend: venv + dependencies -------------------------------------------
echo
# A venv is platform-specific and fully regenerable, so rebuild it whenever the
# existing one won't do: (a) no bin/python — it was built on Windows
# (Scripts/python.exe) and the repo moved to macOS/Linux; or (b) its Python is
# older than we need. Narrow + safe: the only thing ever removed is this one
# hard-coded, regenerable path.
if [ -x backend/.venv/bin/python ] \
   && ! backend/.venv/bin/python -c "import sys; sys.exit(0 if sys.version_info[:2] >= (3, $MIN_PY_MINOR) else 1)" 2>/dev/null; then
  echo "  Existing backend/.venv uses an older Python — rebuilding it ..."
  rm -rf backend/.venv
elif [ -d backend/.venv ] && [ ! -x backend/.venv/bin/python ]; then
  echo "  Existing backend/.venv is not a macOS/Linux venv — rebuilding it ..."
  rm -rf backend/.venv
fi
if [ ! -x backend/.venv/bin/python ]; then
  echo "  Creating backend venv (backend/.venv) with $("$PY" --version) ..."
  "$PY" -m venv backend/.venv
fi
echo "  Installing backend dependencies (backend/requirements.txt) ..."
backend/.venv/bin/python -m pip install --quiet --upgrade pip
backend/.venv/bin/python -m pip install --quiet -r backend/requirements.txt
echo "  [OK] Backend ready."

# --- Frontend: dependencies --------------------------------------------------
echo
echo "  Installing frontend dependencies (npm install --ignore-scripts) ..."
( cd frontend && npm install --ignore-scripts )
# A file-synced node_modules (Dropbox/OneDrive) carried over from Windows lands
# here without the +x bit on its .bin shims, and npm won't rewrite perms on
# files it considers already installed — so `npm run dev` dies with "vite:
# Permission denied". Restore the executable bit. No-op on a clean install.
chmod +x frontend/node_modules/.bin/* 2>/dev/null || true
echo "  [OK] Frontend ready."

# --- Hymnal data + working directories ----------------------------------------
echo
mkdir -p output scripture-cache
if [ -f hymnal-json/index.json ]; then
  echo "  [OK] Hymnal data found."
else
  echo "  [!] No hymnal data yet (hymnal-json/ is empty)."
  echo "      Copy your hymnal-json folder into the project root when ready —"
  echo "      OpenOrder runs without it, but hymn search will come up empty."
fi

echo
echo "  [OK] Setup complete. Start OpenOrder with:"
echo "        ./openorder.sh start"
