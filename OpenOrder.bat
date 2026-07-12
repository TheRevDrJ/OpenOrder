@echo off
:: OpenOrder - worship bulletin & slide generator for churches
:: Copyright (c) 2026 Rev. Dr. Jonathan Mellette (TheRevDrJ)
:: Licensed under AGPL-3.0 - see LICENSE file for details
setlocal enabledelayedexpansion

:: ============================================================================
:: OpenOrder Dev Server Manager  (the Windows twin of openorder.sh)
:: Usage: OpenOrder [command]
::
:: Two dev servers, both launched HEADLESS so they survive closing the terminal:
::   Frontend - Vite    :6800  (node, via a hidden-launch .vbs shim - node's
::                              windowless analog to pythonw)
::   Backend  - FastAPI :6801  (python, via pythonw.exe - no console, no window)
:: Production runs as a single pywebview window; this is dev convenience only.
::
:: SAFETY (process-kills are destructive - kill NARROW, never
:: broad): every kill targets a SPECIFIC port's PID (ports are hard-coded
:: literals 6800/6801, never blank) or a narrow command-line match in
:: kill_openorder.ps1. There is NO broad/wildcard process match anywhere, so it
:: cannot sweep ports or kill-all even if a variable were somehow empty. A blank
:: target is a no-op, never a fall-through to broad.
:: ============================================================================

set "SCRIPT_DIR=%~dp0"
set "FRONTEND_PORT=6800"
set "BACKEND_PORT=6801"
set "BACKEND_LOG=%SCRIPT_DIR%backend.log"
set "FRONTEND_LOG=%SCRIPT_DIR%frontend.log"
set "VENV_PYTHON=%SCRIPT_DIR%backend\.venv\Scripts\python.exe"
set "VENV_PYTHONW=%SCRIPT_DIR%backend\.venv\Scripts\pythonw.exe"
set "FRONTEND_VBS=%SCRIPT_DIR%start_frontend_hidden.vbs"
set "KILL_HELPER=%SCRIPT_DIR%kill_openorder.ps1"

if "%~1"=="" goto help
if /i "%~1"=="start" goto start
if /i "%~1"=="up" goto start
if /i "%~1"=="stop" goto stop
if /i "%~1"=="down" goto stop
if /i "%~1"=="drop" goto stop
if /i "%~1"=="restart" goto restart
if /i "%~1"=="status" goto status
if /i "%~1"=="verbose" goto verbose
if /i "%~1"=="log" goto log
if /i "%~1"=="help" goto help
if /i "%~1"=="--help" goto help
if /i "%~1"=="-h" goto help
goto help

:: ============================================================================
:start
::   Launches both servers headless (no console windows). They keep running
::   after you close this terminal. Logs go to backend.log / frontend.log.
:: ============================================================================
if not exist "%VENV_PYTHONW%" (
    echo   ERROR: backend venv not found at %VENV_PYTHONW%
    echo   Run setup first:  setup.bat
    exit /b 1
)

:: --- Backend (FastAPI :6801) ---
call :find_pid %BACKEND_PORT%
if defined RUNNING_PID (
    echo   Backend  already running on %BACKEND_PORT% ^(PID: !RUNNING_PID!^).
) else (
    :: clear any stale OpenOrder backend that failed to bind, then start headless
    powershell -NoProfile -ExecutionPolicy Bypass -File "%KILL_HELPER%" -Quiet > nul 2>&1
    echo   Starting backend  ^(uvicorn^) on %BACKEND_PORT%...
    start "" /b "%VENV_PYTHONW%" -m uvicorn app.main:app --host 0.0.0.0 --port %BACKEND_PORT% --app-dir backend > "%BACKEND_LOG%" 2>&1
)

:: --- Frontend (Vite :6800) ---
call :find_pid %FRONTEND_PORT%
if defined RUNNING_PID (
    echo   Frontend already running on %FRONTEND_PORT% ^(PID: !RUNNING_PID!^).
) else (
    echo   Starting frontend ^(Vite^) on %FRONTEND_PORT%...
    wscript "%FRONTEND_VBS%"
)

:: --- Wait for both ports to come up (fast: ~few seconds each) ---
set /a ELAPSED=0
echo   Waiting for servers...
:start_wait
ping 127.0.0.1 -n 3 > nul
set /a ELAPSED+=2
set "BACK_UP="
set "FRONT_UP="
call :find_pid %BACKEND_PORT%
if defined RUNNING_PID set "BACK_UP=1"
call :find_pid %FRONTEND_PORT%
if defined RUNNING_PID set "FRONT_UP=1"
if defined BACK_UP if defined FRONT_UP goto start_success
if !ELAPSED! geq 40 goto start_timeout
goto start_wait

:start_success
echo.
echo   OpenOrder is running ^(both servers up in !ELAPSED!s^).
echo   App:  http://localhost:%FRONTEND_PORT%/
echo   API:  http://localhost:%BACKEND_PORT%/api/health
echo   Logs: backend.log / frontend.log   ^(or:  OpenOrder log^)
echo.
echo   Use 'OpenOrder stop' to shut down.
echo.
exit /b 0

:start_timeout
echo.
echo   Servers did not both come up within !ELAPSED!s.
if not defined BACK_UP  echo     - backend  :%BACKEND_PORT% not listening - see backend.log
if not defined FRONT_UP echo     - frontend :%FRONTEND_PORT% not listening - see frontend.log
echo   Try 'OpenOrder verbose' to see backend errors live, or 'OpenOrder log'.
echo.
exit /b 1

:: ============================================================================
:stop
::   Stops both servers. Only ever acts on PIDs listening on the two hard-coded
::   ports, plus a narrow command-line match for any stale backend. No broad
::   match - cannot touch any other project's servers.
:: ============================================================================
set "FOUND_SOMETHING=0"

call :find_pid %FRONTEND_PORT%
if defined RUNNING_PID (
    echo   Stopping frontend on %FRONTEND_PORT% ^(PID: !RUNNING_PID!^)...
    taskkill /F /T /PID !RUNNING_PID! > nul 2>&1
    set "FOUND_SOMETHING=1"
)

call :find_pid %BACKEND_PORT%
if defined RUNNING_PID (
    echo   Stopping backend on %BACKEND_PORT% ^(PID: !RUNNING_PID!^)...
    taskkill /F /T /PID !RUNNING_PID! > nul 2>&1
    set "FOUND_SOMETHING=1"
)

:: Narrow stale-backend cleanup (pythonw running OpenOrder's uvicorn, by command
:: line only - kills nothing if there's no match).
powershell -NoProfile -ExecutionPolicy Bypass -File "%KILL_HELPER%"
if errorlevel 1 set "FOUND_SOMETHING=1"

if "!FOUND_SOMETHING!"=="0" (
    echo   Nothing running on %FRONTEND_PORT% or %BACKEND_PORT%.
) else (
    echo   Stopped.
)
exit /b 0

:: ============================================================================
:restart
:: ============================================================================
call :stop
echo.
ping 127.0.0.1 -n 3 > nul
call :start
exit /b 0

:: ============================================================================
:status
:: ============================================================================
echo.
echo   OpenOrder dev servers
echo   =====================
call :find_pid %BACKEND_PORT%
if defined RUNNING_PID ( echo   Backend  :%BACKEND_PORT%   RUNNING ^(PID: !RUNNING_PID!^) ) else ( echo   Backend  :%BACKEND_PORT%   stopped )
call :find_pid %FRONTEND_PORT%
if defined RUNNING_PID ( echo   Frontend :%FRONTEND_PORT%   RUNNING ^(PID: !RUNNING_PID!^) ) else ( echo   Frontend :%FRONTEND_PORT%   stopped )
echo.
echo   App:  http://localhost:%FRONTEND_PORT%/
echo   API:  http://localhost:%BACKEND_PORT%/api/health
echo.
exit /b 0

:: ============================================================================
:verbose
::   Runs the BACKEND in the foreground with live output (Ctrl+C to stop), while
::   the frontend runs headless. Use this to watch Python tracebacks directly.
:: ============================================================================
call :find_pid %BACKEND_PORT%
if defined RUNNING_PID (
    echo   Backend already running headless ^(PID: !RUNNING_PID!^).
    echo   Stop it first with 'OpenOrder stop' before running verbose.
    exit /b 1
)

:: bring the frontend up headless if it isn't already
call :find_pid %FRONTEND_PORT%
if not defined RUNNING_PID (
    echo   Starting frontend ^(Vite^) on %FRONTEND_PORT% headless...
    wscript "%FRONTEND_VBS%"
)

echo.
echo   Starting backend in verbose mode on %BACKEND_PORT%...
echo   Logs appear below. Press Ctrl+C to stop the backend.
echo   ================================================
echo.
"%VENV_PYTHON%" -m uvicorn app.main:app --host 0.0.0.0 --port %BACKEND_PORT% --app-dir backend
exit /b 0

:: ============================================================================
:log
::   Follows both server logs live (recent lines first). Ctrl+C to stop.
:: ============================================================================
if not exist "%BACKEND_LOG%" if not exist "%FRONTEND_LOG%" (
    echo   No logs yet. Start the servers first with 'OpenOrder start'.
    exit /b 1
)
echo.
echo   Following backend.log + frontend.log. Press Ctrl+C to stop.
echo   ================================================
echo.
powershell -NoProfile -Command "Get-Content -Path '%BACKEND_LOG%','%FRONTEND_LOG%' -Tail 20 -Wait"
exit /b 0

:: ============================================================================
:find_pid
::   %1 = port (a hard-coded literal). Sets RUNNING_PID to the PID LISTENING on
::   that exact port, or clears it. Never matches more than one port: the
::   findstr needle is ":<port> " with a trailing space, and the port is always
::   a literal from this script, never user/derived input.
:: ============================================================================
set "RUNNING_PID="
for /f "tokens=5" %%a in ('netstat -ano 2^>nul ^| findstr "LISTENING" ^| findstr ":%~1 "') do (
    set "RUNNING_PID=%%a"
)
exit /b 0

:: ============================================================================
:help
:: ============================================================================
echo.
echo   OpenOrder Dev Server Manager
echo   ============================
echo.
echo   Usage: OpenOrder [command]
echo.
echo     start     Start both dev servers headless (frontend %FRONTEND_PORT% + backend %BACKEND_PORT%)
echo     stop      Stop both (only ever touches those two ports)
echo     restart   Stop, then start
echo     status    Show whether each server is running
echo     verbose   Run backend in foreground with live logs (Ctrl+C); frontend stays headless
echo     log       Follow both server logs live (Ctrl+C)
echo     help      This help
echo.
echo   Aliases:  up = start      down, drop = stop
echo.
echo   Servers run HEADLESS and survive closing this terminal. After a PC sleep
echo   drops them, just run:  OpenOrder start
echo.
exit /b 0
