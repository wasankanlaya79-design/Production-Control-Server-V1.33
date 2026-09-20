@echo off
setlocal
cd /d "%~dp0"
title Production Control - Install and Run

echo ===============================================
echo Production Control Server - First Installation
echo ===============================================

echo [1/4] Checking Node.js...
where node >nul 2>&1
if errorlevel 1 (
  echo.
  echo ERROR: Node.js not found.
  echo Please install Node.js 20 or newer first.
  pause
  exit /b 1
)
node -v

echo [2/4] Installing required packages...
if not exist node_modules\exceljs (
  call npm.cmd install
  if errorlevel 1 (
    echo ERROR: npm install failed.
    pause
    exit /b 1
  )
) else (
  echo node_modules already installed.
)

echo [3/4] Building modular server...
node tools\build-server.js
if errorlevel 1 (
  echo ERROR: server build failed.
  pause
  exit /b 1
)

echo [4/4] Starting server...
start "Production Control Server" cmd /k "cd /d ""%~dp0"" && node server.js"
exit /b 0
