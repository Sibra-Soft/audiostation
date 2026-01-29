param(
    [Parameter(Mandatory = $true)]
    [string]$VbpFile
)

function Register($file)
{
	Write-Host "Dependency: $file"
    $p = Start-Process -FilePath $useRegsvr32 -ArgumentList "/s", "`"$file`"" -Wait -PassThru

    if ($p.ExitCode -ne 0) {
        Write-Host "  [ERROR] regsvr32 exitcode = $($p.ExitCode)" -ForegroundColor Red
    }
    else {
        Write-Host "  [OK]" -ForegroundColor Green
    }
}

if (!(Test-Path $VbpFile)) {
    Write-Host "FOUT: VBP bestand niet gevonden: $VbpFile" -ForegroundColor Red
    exit 1
}

Write-Host "======================================" 
Write-Host "VB6 Dependency Register"
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

    if (!(Test-Path -LiteralPath $filePath)) {
        Write-Host "  Status: Not Found" -ForegroundColor Red
        $errors++
        continue
    }

    $useRegsvr32 = $Regsvr32_64
    if ($filePath -match '\\SysWOW64\\') {
        $useRegsvr32 = $Regsvr32_32
    }
	
	if(Register($filePath) == 0) {
		$errors++
	}
}

Register("midifl2k.ocx")
Register("midifl32.ocx")
Register("midiio2k.ocx")
Register("midiio32.ocx")

Write-Host ""
Write-Host "======================================"