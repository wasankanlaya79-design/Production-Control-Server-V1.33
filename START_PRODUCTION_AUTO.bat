@echo off
cd /d "%~dp0"
start "Production Control Server" /min cmd /c "cd /d ""%~dp0"" && node server.js >> server.log 2>&1"
exit
