@echo off
:: OpenOrder - worship bulletin & slide generator for churches
:: Copyright (c) 2026 Rev. Dr. Jonathan Mellette (TheRevDrJ)
:: Licensed under AGPL-3.0 - see LICENSE file for details
setlocal enabledelayedexpansion

:: ============================================================
::  OpenOrder Setup  (the Windows twin of setup.sh)
::  One-time setup after cloning the repo.
::  Run this from the project root: setup.bat
::
::  Checks the toolchain (Node 20+, Python 3.10+), then:
::    - creates backend\.venv and pip-installs backend\requirements.txt
::    - npm-installs .\frontend
::    - offers to copy in hymnal data
::
::  After this:  OpenOrder start
:: ============================================================

echo.
echo   ==============================
echo    OpenOrder Setup
echo   ==============================
echo.

:: --------------------------------------------------
::  1. Check for Python 3.10+
:: --------------------------------------------------
echo   [1/5] Checking for Python...
where python >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ERROR: Python not found in PATH.
    echo   Install Python 3.10+ from https://www.python.org/downloads/
    echo   Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)
python -c "import sys; sys.exit(0 if sys.version_info[:2] >= (3, 10) else 1)" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ERROR: Python is too old - OpenOrder needs Python 3.10+.
    for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do echo   Found Python %%v
    echo   Install a newer one from https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo          Found Python %PYVER%

:: --------------------------------------------------
::  2. Check for Node.js 20+
:: --------------------------------------------------
echo   [2/5] Checking for Node.js...
where node >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ERROR: Node.js not found in PATH.
    echo   Install Node.js 20+ from https://nodejs.org/
    echo.
    pause
    exit /b 1
)
for /f "tokens=1 delims=." %%v in ('node --version 2^>^&1') do set NODEMAJOR=%%v
set "NODEMAJOR=%NODEMAJOR:v=%"
if %NODEMAJOR% lss 20 (
    echo.
    echo   ERROR: Node.js %NODEMAJOR% found - OpenOrder needs Node 20+.
    echo   Upgrade from https://nodejs.org/
    echo.
    pause
    exit /b 1
)
for /f "tokens=1 delims= " %%v in ('node --version 2^>^&1') do set NODEVER=%%v
echo          Found Node.js %NODEVER%

:: --------------------------------------------------
::  3. Backend: venv + dependencies
::     A venv is platform-specific and fully regenerable, so rebuild it if the
::     existing one won't do: no Scripts\python.exe means it was built on
::     macOS/Linux (bin/python) and the repo moved back to Windows. Narrow +
::     safe: the only thing ever removed is this one hard-coded path.
:: --------------------------------------------------
echo   [3/5] Setting up the backend venv...
if exist "backend\.venv" if not exist "backend\.venv\Scripts\python.exe" (
    echo          Existing backend\.venv is not a Windows venv - rebuilding it...
    rmdir /s /q "backend\.venv"
)
if not exist "backend\.venv\Scripts\python.exe" (
    echo          Creating backend venv ^(backend\.venv^) with Python %PYVER%...
    python -m venv backend\.venv
)
echo          Installing backend dependencies ^(backend\requirements.txt^)...
backend\.venv\Scripts\python.exe -m pip install --quiet --upgrade pip --disable-pip-version-check
backend\.venv\Scripts\python.exe -m pip install --quiet -r backend\requirements.txt --disable-pip-version-check
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ERROR: pip install failed - see output above.
    pause
    exit /b 1
)
echo          Done.

:: --------------------------------------------------
::  4. Frontend: dependencies
:: --------------------------------------------------
echo   [4/5] Installing frontend packages...
cd frontend
call npm install --ignore-scripts --silent 2>nul
cd ..
echo          Done.

:: --------------------------------------------------
::  5. Hymnal data
:: --------------------------------------------------
echo   [5/5] Checking for hymnal data...

if exist "hymnal-json\index.json" (
    echo          Hymnal data found.
    goto :HYMNAL_DONE
)

echo.
echo   Hymnal data not found. You have two options:
echo.
echo     1. Copy from a flash drive or local folder
echo     2. Skip for now (you can add it later^)
echo.
set /p HYMNAL_CHOICE="   Choice (1 or 2): "

if "%HYMNAL_CHOICE%"=="1" (
    echo.
    echo   Enter the path to your hymnal-json folder.
    echo   Example: E:\hymnal-json  or  D:\Backup\hymnal-json
    echo.
    set /p HYMNAL_SRC="   Path: "

    if not exist "!HYMNAL_SRC!\index.json" (
        echo.
        echo   WARNING: No index.json found at !HYMNAL_SRC!
        echo   Make sure you're pointing to the hymnal-json folder itself.
        echo   Skipping for now -- you can copy it manually later.
        goto :HYMNAL_DONE
    )

    echo   Copying hymnal data...
    xcopy "!HYMNAL_SRC!" "hymnal-json\" /E /I /Q /Y >nul
    echo          Copied successfully.
) else (
    echo          Skipped. Add your hymnal-json folder later.
)

:HYMNAL_DONE

:: --------------------------------------------------
::  Create directories
:: --------------------------------------------------
if not exist "output" mkdir output
if not exist "scripture-cache" mkdir scripture-cache

:: --------------------------------------------------
::  Summary
:: --------------------------------------------------
echo.
echo   ==============================
echo    Setup Complete
echo   ==============================
echo.
echo   To start OpenOrder:
echo     OpenOrder start
echo.
echo   Then open http://localhost:6800
echo.

:: --------------------------------------------------
::  Offer to start
:: --------------------------------------------------
set /p START_NOW="   Start OpenOrder now? (Y/N): "
if /i "%START_NOW%"=="Y" (
    call OpenOrder.bat start
)

endlocal
