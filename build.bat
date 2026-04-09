@echo off
setlocal

:: OpenOrder Build Script
:: Builds the frontend and packages everything into a standalone exe
:: Preserves user data (output, settings, scripture-cache) across rebuilds

set "ROOT=%~dp0"
set "DIST=D:\claude"
set "APP=%DIST%\OpenOrder"
set "BACKUP=%DIST%\_backup"
set "PATH=C:\Program Files\nodejs;%PATH%"

echo.
echo   Building OpenOrder...
echo   =====================
echo.

:: Step 1: Build frontend
echo   [1/4] Building frontend...
cd /d "%ROOT%frontend"
call npm run build >nul 2>&1
if errorlevel 1 (
    echo   FAILED: npm run build
    exit /b 1
)
echo         Done.

:: Step 2: Preserve user data
echo   [2/4] Preserving user data...
if exist "%APP%\output" (
    mkdir "%BACKUP%" 2>nul
    if exist "%APP%\output" xcopy /E /I /Q /Y "%APP%\output" "%BACKUP%\output" >nul 2>&1
    if exist "%APP%\scripture-cache" xcopy /E /I /Q /Y "%APP%\scripture-cache" "%BACKUP%\scripture-cache" >nul 2>&1
    if exist "%APP%\settings.json" copy /Y "%APP%\settings.json" "%BACKUP%\settings.json" >nul 2>&1
    if exist "%APP%\logs" xcopy /E /I /Q /Y "%APP%\logs" "%BACKUP%\logs" >nul 2>&1
    echo         Backed up output, cache, settings, and logs.
) else (
    echo         No existing data to preserve.
)

:: Step 3: Run PyInstaller
echo   [3/4] Packaging with PyInstaller...
cd /d "%ROOT%"
if exist "%APP%" rmdir /S /Q "%APP%" 2>nul
python -m PyInstaller --noconfirm --onedir --windowed --name "OpenOrder" --icon "resources\images\openorder.ico" --distpath "%DIST%" ^
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
    --collect-all "PIL" ^
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

:: Step 4: Restore user data and add hymnal
echo   [4/4] Restoring data...
mkdir "%APP%\output" 2>nul
xcopy /E /I /Q /Y "%ROOT%hymnal-json" "%APP%\hymnal-json" >nul 2>&1

if exist "%BACKUP%\output" xcopy /E /I /Q /Y "%BACKUP%\output" "%APP%\output" >nul 2>&1
if exist "%BACKUP%\scripture-cache" xcopy /E /I /Q /Y "%BACKUP%\scripture-cache" "%APP%\scripture-cache" >nul 2>&1
if exist "%BACKUP%\settings.json" copy /Y "%BACKUP%\settings.json" "%APP%\settings.json" >nul 2>&1
if exist "%BACKUP%\logs" xcopy /E /I /Q /Y "%BACKUP%\logs" "%APP%\logs" >nul 2>&1

:: Clean up backup
if exist "%BACKUP%" rmdir /S /Q "%BACKUP%" 2>nul

echo.
echo   =====================
echo   Build complete!
echo.
echo   Output: %APP%\
echo   User data preserved across rebuild.
echo.
echo   Double-click OpenOrder.exe to launch.
echo.
