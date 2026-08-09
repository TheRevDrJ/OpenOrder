# The WINDOWS build - build.sh's twin, run ON ZORA (the build machine) from
# the d:\claude\OpenOrder checkout. PowerShell, never .bat (rules.md sec 12).
#
# ENCODING LAW (bled for 2026-07-22 on FirstLight): this file must stay pure
# ASCII. Windows PowerShell 5.1 - the ssh default shell - reads a BOM-less
# UTF-8 .ps1 as ANSI, so em-dash bytes decode as SMART QUOTES, which PS treats
# as real string delimiters, and the parse dies with a phantom
# "Missing closing '}'" far from the true cause. No em-dashes, no curly
# quotes, no arrows in this file.
#
# ONE NUMBER PER EVENT: build.sh (the mac leg, first in the ritual) owns the
# single bump; the driver scp's build-number.txt here, and this script READS
# it and never bumps. Both artifacts of an event wear the same number.
#
# THE HYMNAL IS DELIBERATELY NOT BUNDLED. It is copyrighted lyric data, it is
# gitignored, and this clone therefore does not have it. The app resolves it
# through the hymnal_dir setting instead, which on the test rig points at
# D:\dropbox\Code\Claude\OrderOfWorship\hymnal-json. Keeping it out of the
# artifact is the point, not a limitation.
#
# Output: d:\claude\openorder-stage\OpenOrder\       (the app folder)
#         d:\claude\openorder-stage\OpenOrder-win.zip (what travels)

$ErrorActionPreference = 'Stop'

$Dist  = 'd:\claude\openorder-dist'
$Work  = 'd:\claude\openorder-build'
$Stage = 'd:\claude\openorder-stage'
$App   = "$Stage\OpenOrder"
$Zip   = "$Stage\OpenOrder-win.zip"
$Py    = 'backend\.venv\Scripts\python.exe'

if (-not (Test-Path $Py)) {
  throw "ABORT: no backend venv at $Py. Create it: python -m venv backend\.venv; then pip install -r backend\requirements.txt"
}

# READ the event's number (bumped once by the mac leg, scp'd here) - never bump
$N = (Get-Content build-number.txt).Trim()
$Ver = (Get-Content frontend\package.json | ConvertFrom-Json).version
Write-Output "BUILD NUMBER: $N (windows - shared with the event's other artifacts)"

# ---- 1. frontend ----------------------------------------------------------
Write-Output '[1/4] Building frontend ...'
Push-Location frontend
npm run build --silent
if ($LASTEXITCODE -ne 0) { Pop-Location; throw "ABORT: npm run build failed ($LASTEXITCODE)" }
Pop-Location

# ---- 2. PyInstaller -------------------------------------------------------
# NOTE the --add-data separator: Windows uses ';', macOS/Linux ':'. Getting
# this wrong silently produces a bundle with no frontend and no resources.
Write-Output '[2/4] Packaging with PyInstaller ...'
Remove-Item -Recurse -Force "$Dist\OpenOrder" -ErrorAction SilentlyContinue
& $Py -m PyInstaller --noconfirm --onedir --windowed `
  --name 'OpenOrder' `
  --icon 'resources\images\openorder.ico' `
  --distpath $Dist `
  --workpath $Work `
  --add-data 'frontend/dist;frontend/dist' `
  --add-data 'resources;resources' `
  --add-data 'backend/app;app' `
  --collect-all 'fastapi' `
  --collect-all 'uvicorn' `
  --collect-all 'starlette' `
  --collect-all 'pydantic' `
  --collect-all 'python_multipart' `
  --collect-all 'pythonbible' `
  --collect-all 'httpx' `
  --collect-all 'docx' `
  --collect-all 'pptx' `
  --collect-all 'webview' `
  --collect-all 'PIL' `
  --hidden-import 'app.main' `
  --hidden-import 'app.paths' `
  --hidden-import 'app.bulletin' `
  --hidden-import 'app.slides' `
  --hidden-import 'app.hymnal' `
  --hidden-import 'app.scripture' `
  --hidden-import 'app.models' `
  --hidden-import 'app.church_config' `
  --hidden-import 'app.calendar_data' `
  --hidden-import 'app.themes' `
  --hidden-import 'app.themes.oakpark' `
  openorder_app.py
if ($LASTEXITCODE -ne 0) { throw "ABORT: PyInstaller failed ($LASTEXITCODE)" }

$Exe = "$Dist\OpenOrder\OpenOrder.exe"
if (-not (Test-Path $Exe)) { throw "ABORT: PyInstaller reported success but $Exe is not there." }

# ---- 3. stage -------------------------------------------------------------
Write-Output '[3/4] Staging ...'
Remove-Item -Recurse -Force $App -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force $Stage | Out-Null
Copy-Item -Recurse $Dist\OpenOrder $App

# ---- 4. zip - ONE file to copy, and ONE hash to verify on the far side -----
# PyInstaller onedir is a folder of hundreds of files; hashing each across the
# wire is impractical, and a partial directory copy looks exactly like a code
# bug when he runs it. One archive means one SHA256 answers "did it arrive
# intact" for the whole artifact.
Write-Output '[4/4] Zipping ...'
Remove-Item -Force $Zip -ErrorAction SilentlyContinue
Compress-Archive -Path "$App\*" -DestinationPath $Zip -CompressionLevel Optimal
if (-not (Test-Path $Zip)) { throw "ABORT: the zip was not produced." }
$ZipMB = [math]::Round((Get-Item $Zip).Length / 1MB, 1)

# ONE CONSOLIDATED RECEIPT, LAST LINE, EVERY BUILD. It carries the version as
# well as the number, in the same greppable shape the mac leg uses, so
# `grep RECEIPT:` finds it in either - and reading only the tail of a long
# build log can never miss it.
Write-Output "BUILT: $Zip ($ZipMB MB)"
Write-Output "RECEIPT: windows v$Ver build $N"
