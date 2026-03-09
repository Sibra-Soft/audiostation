param(
    [string]$PackagesPath = ".\packages"
)

$ErrorActionPreference = "Stop"

function Test-Admin {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-RegSvr32Path {
    param(
        [string]$OcxPath
    )

    $is64BitOs = [Environment]::Is64BitOperatingSystem
    $fileName = [System.IO.Path]::GetFileName($OcxPath).ToLower()

    # Probeer de bitness van de OCX te bepalen via PE-header
    try {
        $fs = [System.IO.File]::OpenRead($OcxPath)
        $br = New-Object System.IO.BinaryReader($fs)

        $fs.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peOffset = $br.ReadInt32()

        $fs.Seek($peOffset + 4, [System.IO.SeekOrigin]::Begin) | Out-Null
        $machine = $br.ReadUInt16()

        $br.Close()
        $fs.Close()

        switch ($machine) {
            0x8664 { 
                return "$env:windir\System32\regsvr32.exe"   # 64-bit OCX
            }
            0x14C { 
                if ($is64BitOs) {
                    return "$env:windir\SysWOW64\regsvr32.exe" # 32-bit OCX op 64-bit OS
                } else {
                    return "$env:windir\System32\regsvr32.exe" # 32-bit OS
                }
            }
            default {
                # Fallback
                if ($is64BitOs) {
                    return "$env:windir\SysWOW64\regsvr32.exe"
                } else {
                    return "$env:windir\System32\regsvr32.exe"
                }
            }
        }
    }
    catch {
        # Fallback als detectie mislukt
        if ($is64BitOs) {
            return "$env:windir\SysWOW64\regsvr32.exe"
        } else {
            return "$env:windir\System32\regsvr32.exe"
        }
    }
}

if (-not (Test-Admin)) {
    Write-Error "Dit script moet als Administrator worden uitgevoerd, anders kan regsvr32 mislukken."
    exit 1
}

$fullPackagesPath = Resolve-Path -Path $PackagesPath -ErrorAction SilentlyContinue

if (-not $fullPackagesPath) {
    Write-Error "De packages directory bestaat niet: $PackagesPath"
    exit 1
}

Write-Host "Zoeken naar OCX-bestanden in: $($fullPackagesPath.Path)" -ForegroundColor Cyan

$ocxFiles = Get-ChildItem -Path $fullPackagesPath.Path -Recurse -Filter "*.ocx" -File

if (-not $ocxFiles -or $ocxFiles.Count -eq 0) {
    Write-Host "Geen OCX-bestanden gevonden." -ForegroundColor Yellow
    exit 0
}

$failed = @()

foreach ($ocx in $ocxFiles) {
    try {
        $regsvr32 = Get-RegSvr32Path -OcxPath $ocx.FullName
        Write-Host "Registreren: $($ocx.FullName)" -ForegroundColor Green
        Write-Host "Gebruikt:    $regsvr32" -ForegroundColor DarkGray

        $process = Start-Process -FilePath $regsvr32 `
                                 -ArgumentList "/s", "`"$($ocx.FullName)`"" `
                                 -Wait `
                                 -PassThru `
                                 -NoNewWindow

        if ($process.ExitCode -ne 0) {
            Write-Warning "Registratie mislukt voor: $($ocx.FullName) (ExitCode: $($process.ExitCode))"
            $failed += $ocx.FullName
        }
    }
    catch {
        Write-Warning "Fout bij registreren van $($ocx.FullName): $($_.Exception.Message)"
        $failed += $ocx.FullName
    }
}

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "De volgende OCX-bestanden konden niet geregistreerd worden:" -ForegroundColor Red
    $failed | ForEach-Object { Write-Host " - $_" -ForegroundColor Red }
    exit 1
}

Write-Host ""
Write-Host "Alle gevonden OCX-bestanden zijn geregistreerd." -ForegroundColor Cyan
exit 0