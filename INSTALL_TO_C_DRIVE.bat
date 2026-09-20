@echo off
setlocal EnableExtensions
title Production Control V1.33 - Install to C Drive

set "DEST=C:\ProductionControlServer"
set "DATA=C:\ProductionControlData"

echo ==================================================
echo  Production Control V1.33 - C Drive Installation
echo ==================================================
echo.
echo Program : %DEST%
echo Data    : %DATA%
echo Web     : http://localhost:3000
echo.

if not exist "%DEST%" mkdir "%DEST%"
if not exist "%DATA%" mkdir "%DATA%"

echo [1/3] Copying program to C drive...
robocopy "%~dp0" "%DEST%" /E /XD node_modules /XF server.log >nul
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
  echo ERROR: Could not copy program files.
  pause
  exit /b 1
)

echo [2/3] Program copied successfully.
echo [3/3] Starting first-time installation...
echo.

cd /d "%DEST%"
call INSTALL_AND_RUN.bat
exit /b 0
