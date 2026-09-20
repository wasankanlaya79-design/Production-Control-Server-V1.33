@echo off
net session >nul 2>&1
if errorlevel 1 (
  echo Please right-click this file and choose Run as administrator.
  pause
  exit /b 1
)
netsh advfirewall firewall add rule name="Production Control 3000" dir=in action=allow protocol=TCP localport=3000
ipconfig
pause
