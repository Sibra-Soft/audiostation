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
                return "$env:windir\System32\regsvr32.exe" # 64-bit
            }
            0x14C { 
                if ($is64BitOs) {
                    return "$env:windir\SysWOW64\regsvr32.exe" # 32-bit
                } else {
                    return "$env:windir\System32\regsvr32.exe" # 32-bit
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

function PrepareBuild(){
	$PROJECT_DIR = "./"
	$COPY_LIST   = "./Resources/build-files.txt"

	Set-Location $PROJECT_DIR

	# Pre-Build
	Write-Host "=== Prepare-Build ==="

	if (-not (Test-Path $COPY_LIST)) {
		Write-Error "$COPY_LIST not found"
		exit 1
	}

	Get-Content $COPY_LIST | ForEach-Object {
		if ([string]::IsNullOrWhiteSpace($_)) { return }

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
}

function CompileProject(){
	Write-Host "Start Compiling" -ForegroundColor Cyan
	
	if (-not (Test-Path "./Build")) {
		New-Item -ItemType Directory -Path "./Build" | Out-Null
	}

	$vb6Exe = "C:\Program Files\Develop\Visual Basic 6\VB6.exe"

	Start-Process `
		-FilePath $vb6Exe `
		-ArgumentList '/MAKE', '".\source\AdioAudioLibrary.vbp"', '/outdir', '"./Build/"', '/out', '"build.log"' `
		-Wait
	
	Write-Host "Compile Complete" -ForegroundColor Green
	Write-Host "Verify Compilation" -ForegroundColor Cyan
	
	Start-Sleep -Seconds 5
	
	$logFile = ".\build.log"
	if (-not (Test-Path $logFile)) {
		Write-Host "build.log not found" -ForegroundColor Red
		exit 1
	}
	
	Get-Content $logFile
	$logContent = Get-Content $logFile -Raw
	
	if ($logContent -match '(?i)succeeded') {
		Write-Host "Build succeeded" -ForegroundColor Green
	}
	else {
		Write-Host "Build failed" -ForegroundColor Red
		exit 1
	}
}

# ==================================

if (-not (Test-Admin)) {
    Write-Error "This script must be run with administrator privileges"
    exit 1
}

PrepareBuild
CompileProject

exit 0