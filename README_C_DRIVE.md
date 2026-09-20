# Production Control Server V1.33 — C Drive Edition

## Standard installation paths
- Program: `C:\ProductionControlServer`
- Production data: `C:\ProductionControlData`
- Web: `http://localhost:3000`
- Port: `3000`

## First installation
1. Extract this ZIP.
2. Double-click `INSTALL_TO_C_DRIVE.bat`.
3. Wait for npm packages to install and the server to start.
4. Open `http://localhost:3000`.

## Normal daily start
Double-click:
`C:\ProductionControlServer\START_PRODUCTION_CONTROL_WEB.bat`

## Auto start with Windows
Right-click `INSTALL_AUTO_START.bat` and choose **Run as administrator**.

## Allow other PCs on the same LAN
Right-click `OPEN_FIREWALL_3000_ADMIN.bat` and choose **Run as administrator**.
Then open from another PC using:
`http://SERVER-IP:3000`

## OneDrive
OneDrive mirroring is disabled by default in this C-drive package.
Set `ONEDRIVE_SYNC_FOLDER` in `config.json` to the actual OneDrive folder when needed.
