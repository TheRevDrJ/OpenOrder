@echo off
setlocal enabledelayedexpansion

:: ============================================================
::  OpenOrder Setup
::  One-time setup after cloning the repo.
::  Run this from the project root: setup.bat
:: ============================================================

echo.
echo   ==============================
echo    OpenOrder Setup
echo   ==============================
echo.

:: --------------------------------------------------
::  1. Check for Python
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
for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo          Found Python %PYVER%

:: --------------------------------------------------
::  2. Check for Node.js
:: --------------------------------------------------
echo   [2/5] Checking for Node.js...
where node >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo   ERROR: Node.js not found in PATH.
    echo   Install Node.js 18+ from https://nodejs.org/
    echo.
    pause
    exit /b 1
)
for /f "tokens=1 delims= " %%v in ('node --version 2^>^&1') do set NODEVER=%%v
echo          Found Node.js %NODEVER%

:: --------------------------------------------------
::  3. Install Python dependencies
:: --------------------------------------------------
echo   [3/5] Installing Python packages...
pip install -r backend\requirements.txt --quiet --disable-pip-version-check
if %ERRORLEVEL% neq 0 (
    echo.
    echo   WARNING: pip install had issues. Trying with --user flag...
    pip install -r backend\requirements.txt --user --quiet --disable-pip-version-check
)
echo          Done.

:: --------------------------------------------------
::  4. Install Node dependencies
:: --------------------------------------------------
echo   [4/5] Installing frontend packages...
cd frontend
call npm install --silent 2>nul
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
echo   Then open http://localhost:5173
echo.

:: --------------------------------------------------
::  Offer to start
:: --------------------------------------------------
set /p START_NOW="   Start OpenOrder now? (Y/N): "
if /i "%START_NOW%"=="Y" (
    call OpenOrder.bat start
)

endlocal
