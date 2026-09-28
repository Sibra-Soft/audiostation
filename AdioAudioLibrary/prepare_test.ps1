# Stop bij fouten
$ErrorActionPreference = "Stop"

# Configuration
$PROJECT_DIR = "./Tests/"
$COPY_LIST   = "../Resources/build-files.txt"

function Register($file)
{
	$Regsvr32_32 = Join-Path $env:WINDIR "SysWOW64\regsvr32.exe"
	$Regsvr32_64 = Join-Path $env:WINDIR "System32\regsvr32.exe"

    $useRegsvr32 = $Regsvr32_64
    if ($filePath -match '\\SysWOW64\\') {
        $useRegsvr32 = $Regsvr32_32
    }
	
	Write-Host "Dependency: $file"
    $p = Start-Process -FilePath $useRegsvr32 -ArgumentList "/s", "`"$file`"" -Wait -PassThru

    if ($p.ExitCode -ne 0) {
        Write-Host "  [ERROR] regsvr32 exitcode = $($p.ExitCode)" -ForegroundColor Red
    }
    else {
        Write-Host "  [OK]" -ForegroundColor Green
    }
}

Set-Location $PROJECT_DIR
Write-Host "=== Prepare-Test ==="

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

# Register the compiled library for testing
Register("../Build/AdioLibrary.ocx")

Write-Host "=== Prepare-Test Completed ==="
exit 0