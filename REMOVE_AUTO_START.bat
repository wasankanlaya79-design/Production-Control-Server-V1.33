@echo off
schtasks /Delete /F /TN "ProductionControlServer"
echo Auto-start task removed.
pause
