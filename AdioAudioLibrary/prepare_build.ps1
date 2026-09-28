# Stop bij fouten
$ErrorActionPreference = "Stop"

# Configuration
$PROJECT_DIR = "./Source/"
$COPY_LIST   = "../Resources/build-files.txt"

Set-Location $PROJECT_DIR

# Pre-Build
Write-Host "=== Prepare-Build ==="

if (-not (Test-Path $COPY_LIST)) {
    Write-Error "$COPY_LIST not found"
    exit 1
}

Get-Content $COPY_LIST | ForEach-Object {
    if ([string]::IsNullOrWhiteSpace($_)) { return }

    # Split on ;
    $parts = $_ -split ";", 2
    if ($parts.Count -ne 2) { return }

    $SRC = $parts[0].Trim()
    $DST = $parts[1].Trim()

    Write-Host "Copying $SRC"

    if (-not (Test-Path $DST)) {
        New-Item -ItemType Directory -Path $DST -Force | Out-Null
    }

    Copy-Item `
        -Path $SRC `
        -Destination $DST `
        -Recurse `
        -Force `
        -ErrorAction Stop
}

Write-Host "=== Prepare-Build Completed ==="
exit 0