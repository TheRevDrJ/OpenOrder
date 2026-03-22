# Convert all .ppt files to .pptx using PowerPoint COM automation
# Skips Dropbox "conflicted copy" files
# Saves .pptx alongside the original .ppt files

param(
    [string]$RootDir = (Join-Path $PSScriptRoot "..\Hymnals"),
    [switch]$DryRun
)

$RootDir = (Resolve-Path $RootDir).Path

# Folders to convert
$folders = @(
    "Hymnal\Hymns by Number",
    "Hymnal\Other General Services and Acts of Worship",
    "Hymnal\General Services",
    "Hymnal\Psalter",
    "The Faith We Sing\Number Index (Cross & Flame background)",
    "The Faith We Sing\Title Index (Photographic background)"
)

# Gather all .ppt files, excluding conflicted copies and WS_FTP logs
$allFiles = @()
foreach ($folder in $folders) {
    $fullPath = Join-Path $RootDir $folder
    if (Test-Path $fullPath) {
        $files = Get-ChildItem -Path $fullPath -Filter "*.ppt" |
            Where-Object { $_.Name -notmatch "conflicted copy" -and $_.Extension -eq ".ppt" }
        $allFiles += $files
    } else {
        Write-Warning "Folder not found: $fullPath"
    }
}

$total = $allFiles.Count
Write-Host "Found $total .ppt files to convert (excluding conflicted copies)" -ForegroundColor Cyan

if ($DryRun) {
    Write-Host "DRY RUN - would convert:"
    foreach ($f in $allFiles) { Write-Host "  $($f.FullName)" }
    exit
}

if ($total -eq 0) {
    Write-Host "Nothing to convert."
    exit
}

# Start PowerPoint
Write-Host "Starting PowerPoint..." -ForegroundColor Yellow
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = [Microsoft.Office.Interop.PowerPoint.MsoTriState]::msoTrue

$converted = 0
$skipped = 0
$failed = 0

foreach ($file in $allFiles) {
    $pptxPath = [System.IO.Path]::ChangeExtension($file.FullName, ".pptx")

    # Skip if .pptx already exists
    if (Test-Path $pptxPath) {
        $skipped++
        continue
    }

    try {
        $presentation = $ppt.Presentations.Open($file.FullName, $true, $false, $false)  # ReadOnly, Untitled, WithWindow
        # 24 = ppSaveAsOpenXMLPresentation (.pptx)
        $presentation.SaveAs($pptxPath, 24)
        $presentation.Close()
        $converted++

        $pct = [math]::Round(($converted + $skipped + $failed) / $total * 100)
        Write-Host "[$pct%] Converted: $($file.Name)" -ForegroundColor Green
    }
    catch {
        $failed++
        Write-Warning "FAILED: $($file.Name) - $($_.Exception.Message)"
    }
}

# Clean up
$ppt.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host "  Converted: $converted"
Write-Host "  Skipped (already existed): $skipped"
Write-Host "  Failed: $failed"
Write-Host "  Total: $total"
