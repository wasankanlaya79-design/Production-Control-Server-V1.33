@echo off
setlocal
cd /d "%~dp0"
set "TASKNAME=ProductionControlServer"
set "RUNFILE=%~dp0START_PRODUCTION_AUTO.bat"

echo Creating Windows logon task: %TASKNAME%
schtasks /Create /F /SC ONLOGON /TN "%TASKNAME%" /TR "\"%RUNFILE%\"" >nul 2>&1
if errorlevel 1 (
  echo Could not create task automatically.
  echo Please right-click this file and choose Run as administrator.
  pause
  exit /b 1
)
echo Auto-start configured successfully.
pause
