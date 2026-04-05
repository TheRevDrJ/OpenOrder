@echo off
setlocal enabledelayedexpansion

:: OpenOrder -- Worship bulletin & slide generator
:: Usage: OpenOrder [start|stop|restart|status|help] [-v]

set "ROOT=%~dp0"
set "PORT=8316"
set "VERBOSE=0"

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
call :check_port %PORT% server_running

if "!server_running!"=="1" (
    echo   Already running!
    echo   App: http://localhost:%PORT%
    echo.
    exit /b 0
)

:: Start server
echo   Starting OpenOrder on port %PORT%...
if "%VERBOSE%"=="1" (
    start "OpenOrder" cmd /k "cd /d "%ROOT%backend" && python -m uvicorn app.main:app --host 0.0.0.0 --port %PORT%"
) else (
    start "OpenOrder" /min cmd /c "cd /d "%ROOT%backend" && python -m uvicorn app.main:app --host 0.0.0.0 --port %PORT%"
)

:: Wait for server to start
timeout /t 2 /nobreak >nul

echo.
echo   App: http://localhost:%PORT%
echo.
echo   Use "OpenOrder stop" to shut down.
echo.
exit /b 0

:: ============================================================
:stop
:: ============================================================
echo.
echo   Stopping OpenOrder...

call :kill_port %PORT% "Server"
if "!killed!"=="0" (
    echo   Nothing was running.
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

call :check_port %PORT% server_running

if "!server_running!"=="1" (
    echo   Server:  RUNNING  [port %PORT%]
) else (
    echo   Server:  STOPPED
)

echo.
exit /b 0

:: ============================================================
:help
:: ============================================================
echo.
echo   OpenOrder -- Worship bulletin ^& slide generator
echo.
echo   USAGE:
echo     OpenOrder [command] [flags]
echo.
echo   COMMANDS:
echo     start      Start server [default if no command given]
echo     stop       Stop server
echo     restart    Stop then start
echo     status     Show whether server is running
echo     help       Show this message
echo.
echo   FLAGS:
echo     -v         Verbose mode -- show server output in console window
echo.
echo   PORT:
echo     %PORT%  [John 3:16]
echo.
echo   EXAMPLES:
echo     OpenOrder              Start server [minimized]
echo     OpenOrder start -v     Start with visible server output
echo     OpenOrder stop         Shut everything down
echo     OpenOrder status       Check if server is running
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
            echo   %~2 stopped [PID %%p]
            set "killed=1"
        )
    )
)
exit /b 0
