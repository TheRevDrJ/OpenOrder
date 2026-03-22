@echo off
setlocal enabledelayedexpansion

:: OpenOrder --Worship bulletin & slide generator
:: Usage: OpenOrder [start|stop|restart|status|help] [-v]

set "ROOT=%~dp0"
set "BACKEND_PORT=8316"
set "FRONTEND_PORT=5173"
set "VERBOSE=0"

:: Add Node.js to PATH
set "PATH=C:\Program Files\nodejs;%PATH%"

:: Parse arguments
set "CMD=%~1"
set "FLAG=%~2"

if "%CMD%"=="" set "CMD=start"
if /I "%FLAG%"=="-v" set "VERBOSE=1"

:: Route to command
if /I "%CMD%"=="start" goto :start
if /I "%CMD%"=="stop" goto :stop
if /I "%CMD%"=="restart" goto :restart
if /I "%CMD%"=="status" goto :status
if /I "%CMD%"=="help" goto :help
if /I "%CMD%"=="-h" goto :help
if /I "%CMD%"=="--help" goto :help
if /I "%CMD%"=="/?" goto :help

echo Unknown command: %CMD%
echo Run "OpenOrder help" for usage.
exit /b 1

:: ============================================================
:start
:: ============================================================
echo.
echo   ___                   ___          _
echo  / _ \ _ __   ___ _ __ / _ \ _ __ __^| ^| ___ _ __
echo ^| ^| ^| ^| '_ \ / _ \ '_ \ ^| ^| ^| '__/ _` ^|/ _ \ '__^|
echo ^| ^|_^| ^| ^|_) ^|  __/ ^| ^| ^| ^|_^| ^| ^| ^(_^| ^|  __/ ^|
echo  \___/^| .__/ \___^|_^| ^|_^|\___/^|_^|  \__,_^|\___^|_^|
echo       ^|_^|
echo.

:: Check if already running
call :check_port %BACKEND_PORT% backend_running
call :check_port %FRONTEND_PORT% frontend_running

if "!backend_running!"=="1" if "!frontend_running!"=="1" (
    echo   Already running!
    echo   Backend:  http://localhost:%BACKEND_PORT%
    echo   Frontend: http://localhost:%FRONTEND_PORT%
    echo.
    exit /b 0
)

:: Start backend
if "!backend_running!"=="1" (
    echo   Backend already running on port %BACKEND_PORT%
) else (
    echo   Starting backend on port %BACKEND_PORT%...
    if "%VERBOSE%"=="1" (
        start "OpenOrder Backend" cmd /k "cd /d "%ROOT%backend" && python -m uvicorn app.main:app --host 0.0.0.0 --port %BACKEND_PORT% --reload"
    ) else (
        start "OpenOrder Backend" /min cmd /c "cd /d "%ROOT%backend" && python -m uvicorn app.main:app --host 0.0.0.0 --port %BACKEND_PORT% --reload"
    )
)

:: Start frontend
if "!frontend_running!"=="1" (
    echo   Frontend already running on port %FRONTEND_PORT%
) else (
    echo   Starting frontend on port %FRONTEND_PORT%...
    if "%VERBOSE%"=="1" (
        start "OpenOrder Frontend" cmd /k "cd /d "%ROOT%frontend" && node node_modules\vite\bin\vite.js"
    ) else (
        start "OpenOrder Frontend" /min cmd /c "cd /d "%ROOT%frontend" && node node_modules\vite\bin\vite.js"
    )
)

:: Wait a moment for servers to start
timeout /t 2 /nobreak >nul

echo.
echo   Backend:  http://localhost:%BACKEND_PORT%  [API]
echo   Frontend: http://localhost:%FRONTEND_PORT%  [App]
echo.
echo   Use "OpenOrder stop" to shut down.
echo.
exit /b 0

:: ============================================================
:stop
:: ============================================================
echo.
echo   Stopping OpenOrder...

set "stopped=0"

:: Kill backend
call :kill_port %BACKEND_PORT% "Backend"
if "!killed!"=="1" set "stopped=1"

:: Kill frontend
call :kill_port %FRONTEND_PORT% "Frontend"
if "!killed!"=="1" set "stopped=1"

if "!stopped!"=="0" (
    echo   Nothing was running.
) else (
    echo.
    echo   Stopped.
)
echo.
exit /b 0

:: ============================================================
:restart
:: ============================================================
call :stop
timeout /t 1 /nobreak >nul
call :start
exit /b 0

:: ============================================================
:status
:: ============================================================
echo.
echo   OpenOrder Status
echo   ----------------

call :check_port %BACKEND_PORT% backend_running
call :check_port %FRONTEND_PORT% frontend_running

if "!backend_running!"=="1" (
    echo   Backend:  RUNNING  [port %BACKEND_PORT%]
) else (
    echo   Backend:  STOPPED
)

if "!frontend_running!"=="1" (
    echo   Frontend: RUNNING  [port %FRONTEND_PORT%]
) else (
    echo   Frontend: STOPPED
)

echo.
exit /b 0

:: ============================================================
:help
:: ============================================================
echo.
echo   OpenOrder --Worship bulletin ^& slide generator
echo.
echo   USAGE:
echo     OpenOrder [command] [flags]
echo.
echo   COMMANDS:
echo     start      Start both servers [default if no command given]
echo     stop       Stop both servers
echo     restart    Stop then start
echo     status     Show whether servers are running
echo     help       Show this message
echo.
echo   FLAGS:
echo     -v         Verbose mode --show server output in console windows
echo.
echo   PORTS:
echo     Backend:   %BACKEND_PORT%  [John 3:16]
echo     Frontend:  %FRONTEND_PORT%
echo.
echo   EXAMPLES:
echo     OpenOrder              Start servers [minimized]
echo     OpenOrder start -v     Start with visible server output
echo     OpenOrder stop         Shut everything down
echo     OpenOrder status       Check if servers are running
echo.
exit /b 0

:: ============================================================
:: Utility: check if a port is in use
:: Usage: call :check_port PORT RESULT_VAR
:: ============================================================
:check_port
set "%~2=0"
netstat -ano 2>nul | findstr /R "LISTENING" | findstr ":%~1 " >nul 2>&1
if !errorlevel! equ 0 set "%~2=1"
exit /b 0

:: ============================================================
:: Utility: kill process on a port
:: Usage: call :kill_port PORT LABEL
:: ============================================================
:kill_port
set "killed=0"
for /f "tokens=5" %%p in ('netstat -ano 2^>nul ^| findstr /R "LISTENING" ^| findstr ":%~1 "') do (
    if not "%%p"=="0" (
        taskkill /F /PID %%p >nul 2>&1
        if !errorlevel! equ 0 (
            echo   %~2 stopped (PID %%p)
            set "killed=1"
        )
    )
)
exit /b 0
