@echo off
setlocal
cd /d "%~dp0"
title Production Control Server

where node >nul 2>&1
if errorlevel 1 (
  echo Node.js not found. Run INSTALL_AND_RUN.bat after installing Node.js 20+.
  pause
  exit /b 1
)

if not exist node_modules\exceljs (
  echo Required packages are not installed.
  echo Please run INSTALL_AND_RUN.bat once first.
  pause
  exit /b 1
)

node server.js
pause
