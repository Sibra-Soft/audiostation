param(
    [Parameter(Mandatory = $true)]
    [string]$VbpFile
)

if (!(Test-Path $VbpFile)) {
    Write-Host "FOUT: VBP bestand niet gevonden: $VbpFile" -ForegroundColor Red
    exit 1
}

Write-Host "======================================" 
Write-Host "VB6 Dependency Register (PowerShell)"
Write-Host "Project: $VbpFile"
Write-Host "======================================"
Write-Host ""

$Regsvr32_32 = Join-Path $env:WINDIR "SysWOW64\regsvr32.exe"
$Regsvr32_64 = Join-Path $env:WINDIR "System32\regsvr32.exe"

$errors = 0

$lines = Get-Content -LiteralPath $VbpFile

foreach ($line in $lines) {
    if ($line -notmatch '\.(ocx|dll)\b') { continue }

    $filePath = $null

	if ($line -match ';') {
		$filePath = ($line -split ';')[-1].Trim()
	}
	elseif ($line -match '\|') {
		$filePath = ($line -split '\|')[-1].Trim()
	}

    if ([string]::IsNullOrWhiteSpace($filePath)) {
        continue
    }

    $filePath = $filePath.Trim('"').Trim()

    Write-Host "Dependency: $filePath"

    if (!(Test-Path -LiteralPath $filePath)) {
        Write-Host "  Status: Not Found" -ForegroundColor Red
        $errors++
        continue
    }

    $useRegsvr32 = $Regsvr32_64
    if ($filePath -match '\\SysWOW64\\') {
        $useRegsvr32 = $Regsvr32_32
    }


    $p = Start-Process -FilePath $useRegsvr32 -ArgumentList "/s", "`"$filePath`"" -Wait -PassThru

    if ($p.ExitCode -ne 0) {
        Write-Host "  [ERROR] regsvr32 exitcode = $($p.ExitCode)" -ForegroundColor Red
        $errors++
    }
    else {
        Write-Host "  [OK]" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host "======================================"
if ($errors -eq 0) {
    Write-Host "Alle dependencies succesvol verwerkt." -ForegroundColor Green
    exit 0
}
else {
    Write-Host "Er zijn fouten opgetreden: $errors" -ForegroundColor Red
    exit 1
}