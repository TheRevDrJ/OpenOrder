@echo off
setlocal

:: OpenOrder Build Script
:: Builds the frontend and packages everything into a standalone exe

set "ROOT=%~dp0"
set "PATH=C:\Program Files\nodejs;%PATH%"

echo.
echo   Building OpenOrder...
echo   =====================
echo.

:: Step 1: Build frontend
echo   [1/3] Building frontend...
cd /d "%ROOT%frontend"
call npm run build >nul 2>&1
if errorlevel 1 (
    echo   FAILED: npm run build
    exit /b 1
)
echo         Done.

:: Step 2: Run PyInstaller
echo   [2/3] Packaging with PyInstaller...
cd /d "%ROOT%"
python -m PyInstaller --noconfirm --onedir --windowed --name "OpenOrder" --icon "resources\images\openorder.ico" --distpath "D:\claude" ^
    --add-data "frontend/dist;frontend/dist" ^
    --add-data "resources;resources" ^
    --add-data "backend/app;app" ^
    --collect-all "fastapi" ^
    --collect-all "uvicorn" ^
    --collect-all "starlette" ^
    --collect-all "pydantic" ^
    --collect-all "python_multipart" ^
    --collect-all "pythonbible" ^
    --collect-all "httpx" ^
    --collect-all "docx" ^
    --collect-all "pptx" ^
    --collect-all "webview" ^
    --hidden-import "app.main" ^
    --hidden-import "app.paths" ^
    --hidden-import "app.bulletin" ^
    --hidden-import "app.slides" ^
    --hidden-import "app.hymnal" ^
    --hidden-import "app.scripture" ^
    --hidden-import "app.models" ^
    --hidden-import "app.church_config" ^
    --hidden-import "app.themes" ^
    --hidden-import "app.themes.oakpark" ^
    openorder_app.py >nul 2>&1
if errorlevel 1 (
    echo   FAILED: PyInstaller
    exit /b 1
)
echo         Done.

:: Step 3: Create output directory in dist
echo   [3/3] Preparing dist...
mkdir "%ROOT%dist\OpenOrder\output" 2>nul

echo.
echo   =====================
echo   Build complete!
echo.
echo   Output: %ROOT%dist\OpenOrder\
echo.
echo   To deploy, copy the dist\OpenOrder folder and add:
echo     - hymnal-json\   (your hymn data)
echo     - settings.json  (optional, for custom data directory)
echo.
echo   Then just double-click OpenOrder.exe
echo.
