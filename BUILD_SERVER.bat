@echo off
setlocal
cd /d "%~dp0"
where node >nul 2>&1
if errorlevel 1 (
  echo Node.js not found. Please install Node.js 20+ first.
  pause
  exit /b 1
)
node tools\build-server.js
if errorlevel 1 (
  echo BUILD FAILED.
  pause
  exit /b 1
)
echo.
echo server.js rebuilt successfully from server_src modules.
pause
