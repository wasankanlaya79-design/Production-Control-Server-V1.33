# Production Control Server V1.33 — MODULAR EDITION

เวอร์ชันนี้จัดโครงสร้างใหม่เพื่อให้แก้ไขโปรแกรมง่ายขึ้น โดยคง Logic เดิมของ V1.33 CLEAN LOGO FIXED เป็นฐาน

## จุดสำคัญ
- หน้า `public/index.html` ถูกตัด CSS และ JavaScript ออกเป็นไฟล์ย่อยตามหน้าที่
- `server.js` ถูกแยก Source เป็น `server_src/01-...07-...` และสร้างกลับอัตโนมัติก่อนรัน
- การแก้ Server ให้แก้ที่ `server_src/` ไม่ควรแก้ `server.js` โดยตรง เพราะจะถูก Build ทับ
- `RUN_WINDOWS.bat`, `INSTALL_AND_RUN.bat` และ `START_PRODUCTION_AUTO.bat` จะ Build `server.js` จาก Source ก่อนเริ่มระบบ

## แผนที่ไฟล์ Front-end
- `public/css/01-base-layout.css` — พื้นฐาน Theme/Layout
- `public/css/02-components.css` — Component/UI
- `public/css/03-production.css` — Production/Planner/Report UI
- `public/css/04-responsive-overrides.css` — ส่วนท้าย/Responsive/Override
- `public/js/01-input-storage.js` — Input Mode, Voice, Local Queue, Workspace
- `public/js/02-logs-reports.js` — Production Log, CSV/Excel/Print, Daily/History
- `public/js/03-production-control.js` — Production Status, Save/Close/Edit FO
- `public/js/04-planner-assignment.js` — Planner Import, Plan Table, FO Assignment
- `public/js/05-overview-events.js` — LINE01–LINE18 Overview/Login และ Event Binding
- `public/js/06-planning-engine.js` — Shift/OT/Target/Render Planning Engine
- `public/js/07-server-bridge.js` — เชื่อม Browser กับ Node API
- `public/js/08-executive-dashboard.js` — Executive Dashboard

## แผนที่ไฟล์ Server
- `server_src/01-config-utils.js` — Config/Path/Utility/File I/O/Lock
- `server_src/02-workbook-sync.js` — Excel Workbook และ Sync
- `server_src/03-event-journal.js` — Event Journal / Pending Queue
- `server_src/04-workbook-reports.js` — Daily Summary / Report / Audit / Workbook Update
- `server_src/05-processing-verify-fo.js` — Process/Verify/FO Reserve/Status
- `server_src/06-api-routes.js` — API Routes
- `server_src/07-worker-startup.js` — Background Worker / Start Server

## วิธีแก้โปรแกรม
1. ถ้าแก้หน้าจอ/ปุ่ม/สูตรฝั่ง Browser ให้แก้ไฟล์ใน `public/js/` หรือ `public/css/`
2. ถ้าแก้ API/Excel/การบันทึกข้อมูล ให้แก้ไฟล์ใน `server_src/`
3. ดับเบิลคลิก `RUN_WINDOWS.bat` ระบบจะ Build Server และรันให้อัตโนมัติ
4. ถ้าต้องการ Build Server อย่างเดียว ใช้ `BUILD_SERVER.bat`

## Master Runtime
- `server.js` = ไฟล์ Runtime ที่สร้างจาก `server_src/`
- `public/index.html` = หน้า Production Control ที่อ้าง CSS/JS แบบแยกไฟล์แล้ว
