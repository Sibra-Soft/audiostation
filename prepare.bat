@echo off

setlocal EnableExtensions EnableDelayedExpansion

REM Configuration
set PROJECT_DIR=%~dp0
set COPY_LIST=./resources/build-files.txt

cd /d "%PROJECT_DIR%"

REM Pre-Build
echo === Pre-Build: copying required files ===

IF NOT EXIST "%COPY_LIST%" (
  echo %COPY_LIST% not found
  exit /b 1
)

for /f "usebackq tokens=1,2 delims=;" %%A in ("%COPY_LIST%") do (
  set SRC=%%A
  set DST=%%B

  set SRC=!SRC:~0,-1!
  set DST=!DST:~1!

  echo Copying !SRC!

  IF NOT EXIST "!DST!" mkdir "!DST!"

  xcopy "!SRC!" "!DST!" /E /Y /I >nul
  IF ERRORLEVEL 1 (
    echo ERROR
    exit /b 1
  )
)

echo === Pre-build Completed ===
exit /b 0