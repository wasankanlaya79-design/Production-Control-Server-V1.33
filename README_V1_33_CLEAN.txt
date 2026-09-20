Production Control Server V1.33 - CLEAN / LOGO FIXED

Baseline: V1.33 IMPORT ALL AS TEXT + EFF AUTO FROM PLAN
Changes in this package:
1) Embedded TNI / Faculty / Logistics logo directly inside index.html so the logo does not depend on an /assets path.
2) Removed old QA reports, old version READMEs, update notes, rebuild/verify utility files, and unused image assets.
3) Retained only files needed for normal installation, run, firewall, and Windows auto-start.

Main files:
- server.js
- config.json
- package.json
- public/index.html
- public/executive_dashboard_18_lines_realtime.html
- INSTALL_AND_RUN.bat
- RUN_WINDOWS.bat
- START_PRODUCTION_AUTO.bat
- SETUP_AUTO_START.bat
- REMOVE_AUTO_START.bat
- OPEN_FIREWALL_3000_ADMIN.bat

First use: run INSTALL_AND_RUN.bat
Normal use: run RUN_WINDOWS.bat
