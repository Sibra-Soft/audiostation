@echo off

setlocal EnableExtensions EnableDelayedExpansion

REM Compiling
echo === Compiling ===
cd ./project
mkdir Build
start /wait "" "C:\Program Files\Develop\Visual Basic 6\VB6.exe" /MAKE ".\source\Audiostation.vbp" /outdir "Build/" /out "build.log"
echo === Compiling Completed ===

CALL :sleep 5

REM Verify Compilation
echo === Verify Compilation ===
type build.log

IF NOT EXIST build.log (
	echo build.log not found
	EXIT /B 1
)

findstr /I "succeeded" build.log >nul
IF ERRORLEVEL 1 (
	echo Build failed
	EXIT /B 1
) ELSE (
	echo Build succeeded
	EXIT /B 0
)

REM Helpers		  
:sleep
ping 127.0.0.1 -n 2 -w 1000 > NUL
ping 127.0.0.1 -n %1 -w 1000 > NUL