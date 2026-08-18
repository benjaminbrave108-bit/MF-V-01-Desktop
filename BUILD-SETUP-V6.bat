@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title MF-V-01 1.0.8 Setup Builder V6

set "NPM_CONFIG_REGISTRY=https://registry.npmjs.org/"
set "npm_config_registry=https://registry.npmjs.org/"
set "npm_config_fetch_retries=5"
set "npm_config_fetch_retry_factor=2"
set "npm_config_fetch_retry_mintimeout=20000"
set "npm_config_fetch_retry_maxtimeout=120000"
set "npm_config_fetch_timeout=300000"
set "ELECTRON_BUILDER_ALLOW_UNRESOLVED_DEPENDENCIES=false"

 echo ============================================================
 echo   MF-V-01 1.0.8 - Windows Setup Builder V6
 echo ============================================================
 echo.

rem ---- Node.js check -------------------------------------------------------
where node >nul 2>nul
if errorlevel 1 (
  echo [INFO] Node.js not found.
  where winget >nul 2>nul
  if errorlevel 1 goto :node_missing
  echo [INFO] Installing Node.js LTS with winget...
  winget install --id OpenJS.NodeJS.LTS -e --accept-package-agreements --accept-source-agreements
  if errorlevel 1 goto :node_missing
  set "PATH=%ProgramFiles%\nodejs;%PATH%"
)

for /f "tokens=*" %%V in ('node -p "process.versions.node" 2^>nul') do set "NODEVER=%%V"
echo [OK] Node.js !NODEVER!
node -e "const [a,b]=process.versions.node.split('.').map(Number); process.exit(a>22||(a===22&&b>=13)?0:1)"
if errorlevel 1 (
  echo [ERROR] Node.js 22.13.0 or newer is required.
  goto :fail
)

where npm >nul 2>nul
if errorlevel 1 goto :node_missing

rem ---- Git Bash check (do NOT accept Windows WSL bash.exe) ----------------
set "GIT_BASH="
if exist "%ProgramFiles%\Git\bin\bash.exe" set "GIT_BASH=%ProgramFiles%\Git\bin\bash.exe"
if not defined GIT_BASH if exist "%ProgramFiles%\Git\usr\bin\bash.exe" set "GIT_BASH=%ProgramFiles%\Git\usr\bin\bash.exe"
if not defined GIT_BASH if exist "%LocalAppData%\Programs\Git\bin\bash.exe" set "GIT_BASH=%LocalAppData%\Programs\Git\bin\bash.exe"
if not defined GIT_BASH if exist "%LocalAppData%\Programs\Git\usr\bin\bash.exe" set "GIT_BASH=%LocalAppData%\Programs\Git\usr\bin\bash.exe"

if not defined GIT_BASH (
  echo [INFO] Git Bash not found. Installing Git for Windows...
  where winget >nul 2>nul
  if errorlevel 1 goto :bash_missing
  winget install --id Git.Git -e --accept-package-agreements --accept-source-agreements
  if errorlevel 1 goto :bash_missing
  if exist "%ProgramFiles%\Git\bin\bash.exe" set "GIT_BASH=%ProgramFiles%\Git\bin\bash.exe"
  if not defined GIT_BASH if exist "%LocalAppData%\Programs\Git\bin\bash.exe" set "GIT_BASH=%LocalAppData%\Programs\Git\bin\bash.exe"
)

if not defined GIT_BASH goto :bash_missing
for %%G in ("!GIT_BASH!") do set "GIT_BIN=%%~dpG"
set "PATH=!GIT_BIN!;%ProgramFiles%\Git\usr\bin;%LocalAppData%\Programs\Git\usr\bin;%PATH%"
echo [OK] Git Bash: !GIT_BASH!
"!GIT_BASH!" --version
if errorlevel 1 goto :bash_missing

rem ---- Release locked files ------------------------------------------------
echo [INFO] Closing stale Node/Electron processes that can lock node_modules...
taskkill /F /IM node.exe >nul 2>nul
taskkill /F /IM electron.exe >nul 2>nul
timeout /t 2 /nobreak >nul

rem ---- Clean partial install robustly --------------------------------------
if exist "node_modules" (
  echo [INFO] Cleaning previous partial node_modules...
  attrib -R -S -H "node_modules\*" /S /D >nul 2>nul
  rmdir /s /q "node_modules" >nul 2>nul
  if exist "node_modules" (
    echo [INFO] Standard cleanup was blocked. Using robocopy cleanup method...
    mkdir "%TEMP%\mfv-empty" >nul 2>nul
    robocopy "%TEMP%\mfv-empty" "node_modules" /MIR >nul 2>nul
    rmdir /s /q "node_modules" >nul 2>nul
  )
  if exist "node_modules" (
    echo [ERROR] Windows is still locking node_modules.
    echo Close File Explorer windows, VS Code and antivirus scan for this folder, then run this file again.
    goto :fail
  )
)

rem ---- Install npm packages without Electron binary ------------------------
echo.
echo [1/5] Installing npm packages (compatibility mode)...
set "ELECTRON_SKIP_BINARY_DOWNLOAD=1"
call npm install --include=dev --ignore-scripts --registry=https://registry.npmjs.org/ --no-audit --no-fund
if errorlevel 1 goto :npm_fail
set "ELECTRON_SKIP_BINARY_DOWNLOAD="

rem ---- Download Electron binary separately with retries --------------------
echo.
echo [2/5] Downloading Electron runtime...
call :electron_official
if exist "node_modules\electron\dist\electron.exe" goto :electron_ok

echo [WARN] Official Electron download failed. Trying alternative mirror...
call :electron_mirror
if exist "node_modules\electron\dist\electron.exe" goto :electron_ok

echo [ERROR] Electron runtime could not be downloaded from official or mirror source.
goto :electron_fail

:electron_ok
echo [OK] Electron runtime is ready.

rem ---- Build app ------------------------------------------------------------
echo.
echo [3/5] Building application...
"!GIT_BASH!" scripts/build-verified.sh
if errorlevel 1 goto :build_fail

rem ---- Package with official source first ----------------------------------
echo.
echo [4/5] Creating Windows NSIS package...
set "ELECTRON_MIRROR="
set "ELECTRON_BUILDER_BINARIES_MIRROR="
call npx electron-builder --win nsis --x64
if not errorlevel 1 goto :check_setup

echo [WARN] Packaging download failed. Retrying with alternative binary mirror...
set "ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/"
set "ELECTRON_BUILDER_BINARIES_MIRROR=https://npmmirror.com/mirrors/electron-builder-binaries/"
call npx electron-builder --win nsis --x64
if errorlevel 1 goto :builder_fail

:check_setup
echo.
echo [5/5] Verifying Setup file...
if exist "release\MF-V-01-Setup-1.0.8.exe" (
  for %%F in ("release\MF-V-01-Setup-1.0.8.exe") do set "SETSIZE=%%~zF"
  echo.
  echo ============================================================
  echo [SUCCESS] Setup created successfully.
  echo File: %CD%\release\MF-V-01-Setup-1.0.8.exe
  echo Size: !SETSIZE! bytes
  echo ============================================================
  echo.
  start "" explorer.exe /select,"%CD%\release\MF-V-01-Setup-1.0.8.exe"
  pause
  exit /b 0
)

echo [ERROR] Packaging ended but expected Setup file was not found.
goto :fail

:electron_official
set "ELECTRON_MIRROR="
for /L %%R in (1,1,3) do (
  echo [INFO] Official Electron attempt %%R of 3...
  rmdir /s /q "node_modules\electron\dist" >nul 2>nul
  del /q "node_modules\electron\path.txt" >nul 2>nul
  node "node_modules\electron\install.js"
  if exist "node_modules\electron\dist\electron.exe" exit /b 0
  timeout /t 5 /nobreak >nul
)
exit /b 1

:electron_mirror
set "ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/"
for /L %%R in (1,1,3) do (
  echo [INFO] Mirror Electron attempt %%R of 3...
  rmdir /s /q "node_modules\electron\dist" >nul 2>nul
  del /q "node_modules\electron\path.txt" >nul 2>nul
  node "node_modules\electron\install.js"
  if exist "node_modules\electron\dist\electron.exe" exit /b 0
  timeout /t 5 /nobreak >nul
)
exit /b 1

:node_missing
echo [ERROR] Node.js/npm could not be found or installed automatically.
goto :fail

:bash_missing
echo [ERROR] Git Bash could not be found or installed automatically.
goto :fail

:npm_fail
echo.
echo [ERROR] npm package installation failed.
echo If this happens again, send the last 20 lines shown above.
goto :fail

:electron_fail
echo.
echo [ERROR] Electron binary download failed on both sources.
echo This is usually a network/proxy/firewall interruption.
goto :fail

:build_fail
echo.
echo [ERROR] Application build failed. Send the last 20 lines shown above.
goto :fail

:builder_fail
echo.
echo [ERROR] Electron/NSIS packaging failed after both download routes.
echo Send the last 20 lines shown above.
goto :fail

:fail
echo.
echo Setup was NOT created.
pause
exit /b 1
