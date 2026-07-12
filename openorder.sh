#!/usr/bin/env bash
#
# OpenOrder — dev server manager (macOS / Linux)
#
# The cross-platform twin of OpenOrder.bat. OpenOrder runs TWO dev servers:
#   Frontend — Vite    :6800  (npm run dev, in ./frontend)
#   Backend  — FastAPI :6801  (uvicorn app.main:app, via ./backend/.venv)
# In production the app runs as a single pywebview window over the backend;
# this two-server split is dev convenience only.
#
# Usage:  ./openorder.sh [start|stop|restart|status|verbose|log|help]
#
# SAFETY (process kills are destructive — kill NARROW, never broad):
#   `stop` targets ONLY the process LISTENING on a hard-coded port (6800 / 6801)
#   via `lsof -sTCP:LISTEN`, validates each target is a numeric PID, and is a
#   no-op when nothing is listening. There is no wildcard / name / broad match
#   anywhere, so an empty target can never degrade into "kill everything."

set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
FRONTEND_PORT=6800
BACKEND_PORT=6801
FRONTEND_DIR="$DIR/frontend"
BACKEND_LOG="$DIR/backend.log"
FRONTEND_LOG="$DIR/frontend.log"
VENV_PYTHON="$DIR/backend/.venv/bin/python"
URL="http://localhost:$FRONTEND_PORT"

# --- narrow helpers ----------------------------------------------------------

# PIDs LISTENING on an exact port (empty if none). Never broad, never by name.
listeners() {
  lsof -ti "tcp:$1" -sTCP:LISTEN 2>/dev/null || true
}

# Stop whatever listens on ONE hard-coded port. Returns 0 if it killed
# something, 1 if nothing was listening. Blank target => no-op (the return 1).
kill_port() {
  local port="$1" label="$2" pids pid
  pids="$(listeners "$port")"
  [ -z "$pids" ] && return 1
  echo "  Stopping $label on $port (pid: $(echo "$pids" | tr '\n' ' ')) ..."
  for pid in $pids; do
    [[ "$pid" =~ ^[0-9]+$ ]] && kill "$pid" 2>/dev/null || true
  done
  sleep 1
  # escalate ONLY against the same narrow port target, if anything survived
  for pid in $(listeners "$port"); do
    [[ "$pid" =~ ^[0-9]+$ ]] && kill -9 "$pid" 2>/dev/null || true
  done
  return 0
}

show_urls() {
  echo "    App:  $URL/"
  echo "    API:  http://localhost:$BACKEND_PORT/api/health"
}

# --- commands ----------------------------------------------------------------

cmd_start() {
  if [ ! -x "$VENV_PYTHON" ]; then
    echo "  [X] backend venv not found at:"
    echo "        $VENV_PYTHON"
    echo "      Run setup first:  ./setup.sh"
    return 1
  fi

  # --- Backend (FastAPI :6801) ---
  if [ -n "$(listeners "$BACKEND_PORT")" ]; then
    echo "  Backend  already running on $BACKEND_PORT."
  else
    echo "  Starting backend  (uvicorn) on $BACKEND_PORT ..."
    ( cd "$DIR" && nohup "$VENV_PYTHON" -m uvicorn app.main:app \
        --host 0.0.0.0 --port "$BACKEND_PORT" --app-dir backend \
        > "$BACKEND_LOG" 2>&1 & )
  fi

  # --- Frontend (Vite :6800) ---
  if [ -n "$(listeners "$FRONTEND_PORT")" ]; then
    echo "  Frontend already running on $FRONTEND_PORT."
  else
    echo "  Starting frontend (Vite) on $FRONTEND_PORT ..."
    ( cd "$FRONTEND_DIR" && nohup npm run dev > "$FRONTEND_LOG" 2>&1 & )
  fi

  # --- Wait for both ports to come up ---
  echo "  Waiting for servers ..."
  local n=0
  while { [ -z "$(listeners "$BACKEND_PORT")" ] || [ -z "$(listeners "$FRONTEND_PORT")" ]; } \
        && [ "$n" -lt 40 ]; do
    sleep 1; n=$((n + 1))
  done

  local back front
  back="$(listeners "$BACKEND_PORT")"
  front="$(listeners "$FRONTEND_PORT")"
  echo
  if [ -n "$back" ] && [ -n "$front" ]; then
    echo "  OpenOrder is running (both servers up in ${n}s)."
    show_urls
    echo "    Logs:  backend.log / frontend.log   (or:  ./openorder.sh log)"
    echo
    echo "  Leave them running; they survive closing this terminal."
    echo "  Use './openorder.sh stop' to shut down."
  else
    echo "  Servers did not both come up within ${n}s."
    [ -z "$back" ]  && echo "    - backend  :$BACKEND_PORT not listening — see backend.log"
    [ -z "$front" ] && echo "    - frontend :$FRONTEND_PORT not listening — see frontend.log"
    echo "  Try './openorder.sh verbose' to watch backend errors live."
    return 1
  fi
}

cmd_stop() {
  local hit=1
  kill_port "$FRONTEND_PORT" "frontend" && hit=0
  kill_port "$BACKEND_PORT"  "backend"  && hit=0
  if [ "$hit" -eq 0 ]; then
    echo "  Stopped."
  else
    echo "  Nothing running on $FRONTEND_PORT or $BACKEND_PORT."
  fi
}

cmd_restart() { cmd_stop; sleep 1; echo; cmd_start; }

cmd_status() {
  local back front
  back="$(listeners "$BACKEND_PORT")"
  front="$(listeners "$FRONTEND_PORT")"
  echo
  echo "  OpenOrder dev servers"
  echo "  ====================="
  if [ -n "$back" ];  then echo "  Backend  :$BACKEND_PORT   RUNNING (pid: $(echo "$back"  | tr '\n' ' '))"; else echo "  Backend  :$BACKEND_PORT   stopped"; fi
  if [ -n "$front" ]; then echo "  Frontend :$FRONTEND_PORT   RUNNING (pid: $(echo "$front" | tr '\n' ' '))"; else echo "  Frontend :$FRONTEND_PORT   stopped"; fi
  echo
  show_urls
  echo
}

# Backend in the FOREGROUND with live output (Ctrl-C to stop); frontend headless.
cmd_verbose() {
  if [ ! -x "$VENV_PYTHON" ]; then
    echo "  [X] backend venv not found — run ./setup.sh first."
    return 1
  fi
  if [ -n "$(listeners "$BACKEND_PORT")" ]; then
    echo "  Backend already running on $BACKEND_PORT — './openorder.sh stop' it first."
    return 1
  fi
  if [ -z "$(listeners "$FRONTEND_PORT")" ]; then
    echo "  Starting frontend (Vite) on $FRONTEND_PORT headless ..."
    ( cd "$FRONTEND_DIR" && nohup npm run dev > "$FRONTEND_LOG" 2>&1 & )
  fi
  echo
  echo "  Backend on $BACKEND_PORT — live logs below. Ctrl-C to stop."
  echo "  ================================================"
  echo
  cd "$DIR"
  exec "$VENV_PYTHON" -m uvicorn app.main:app \
    --host 0.0.0.0 --port "$BACKEND_PORT" --app-dir backend
}

cmd_log() {
  if [ ! -f "$BACKEND_LOG" ] && [ ! -f "$FRONTEND_LOG" ]; then
    echo "  No logs yet. Start the servers first with './openorder.sh start'."
    return 1
  fi
  echo "  Following backend.log + frontend.log. Ctrl-C to stop."
  echo "  ================================================"
  # -F: keep following even as the files are rotated/created
  tail -n 20 -F "$BACKEND_LOG" "$FRONTEND_LOG" 2>/dev/null
}

cmd_help() {
  cat <<EOF

  OpenOrder — dev server manager (macOS / Linux)
  ==============================================

  Usage:  ./openorder.sh [command]

    start     Start both dev servers in the background (frontend $FRONTEND_PORT + backend $BACKEND_PORT)
    stop      Stop both (only ever touches those two ports)
    restart   Stop, then start
    status    Show whether each server is running
    verbose   Run backend in the foreground with live logs (Ctrl-C); frontend stays headless
    log       Follow both server logs live (Ctrl-C)
    help      This help

  Aliases:  up = start      down, drop = stop

    App:  $URL/
    API:  http://localhost:$BACKEND_PORT/api/health

  Servers run in the background and survive closing this terminal. After a
  sleep drops them, just run:  ./openorder.sh start
EOF
}

case "${1:-help}" in
  start|up)        cmd_start ;;
  stop|down|drop)  cmd_stop ;;
  restart)         cmd_restart ;;
  status)          cmd_status ;;
  verbose)         cmd_verbose ;;
  log|logs)        cmd_log ;;
  help|--help|-h)  cmd_help ;;
  *) echo "Unknown command: ${1:-}"; cmd_help; exit 1 ;;
esac
