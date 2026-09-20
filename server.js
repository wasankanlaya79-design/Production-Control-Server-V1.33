"use strict";

const express = require("express");
const path = require("path");
const fs = require("fs");
const fsp = fs.promises;
const ExcelJS = require("exceljs");
const { exec } = require("child_process");

const ROOT = __dirname;
const cfg = JSON.parse(fs.readFileSync(path.join(ROOT, "config.json"), "utf8"));
const PORT = Number(cfg.PORT || 3000);
const DATA_ROOT = path.resolve(String(cfg.LOCAL_DATA_ROOT || "E:\\ProductionControlData"));
const ONEDRIVE_ROOT = String(cfg.ONEDRIVE_SYNC_FOLDER || "").trim();
const ADMIN_PIN = String(cfg.ADMIN_PIN || "12345");
const LINE_HEAD_PIN = String(cfg.LINE_HEAD_PIN || ADMIN_PIN);
const PLANNER_PIN = String(cfg.PLANNER_PIN || ADMIN_PIN);
const LINE_COUNT = Math.max(1, Math.min(99, Number(cfg.LINES || 18)));
const RETRY_MS = Math.max(1000, Number(cfg.RETRY_SECONDS || 5) * 1000);
const LINE_CODES = Array.from({length: LINE_COUNT}, (_, i) => `LINE${String(i+1).padStart(2,"0")}`);

const app = express();
app.use(express.json({limit:"20mb"}));
app.use(express.static(path.join(ROOT, "public")));

const lineLocks = new Map();
const journalIdCache = new Map();
const appliedIdCache = new Map();
const lastExcelError = new Map();
const lastMirrorError = new Map();
const lastExcelSavedAt = new Map();
const lastMirrorAt = new Map();
let centralLock = Promise.resolve();

const SHIFT = [
  {label:"08:00-09:00", h:1, type:"WORK"},
  {label:"09:00-10:00", h:1, type:"WORK"},
  {label:"10:00-11:00", h:1, type:"WORK"},
  {label:"11:00-12:00", h:1, type:"WORK"},
  {label:"12:00-13:00", h:1, type:"BREAK"},
  {label:"13:00-14:00", h:1, type:"WORK"},
  {label:"14:00-15:00", h:1, type:"WORK"},
  {label:"15:00-16:00", h:1, type:"WORK"},
  {label:"16:00-17:00", h:1, type:"WORK"},
  {label:"17:00-17:30", h:.5, type:"BREAK"},
  {label:"17:30-18:30", h:1, type:"OT", otNo:1},
  {label:"18:30-19:30", h:1, type:"OT", otNo:2},
  {label:"19:30-20:30", h:1, type:"OT", otNo:3}
];

function nowBkk(){
  const p = Object.fromEntries(new Intl.DateTimeFormat("sv-SE",{
    timeZone:"Asia/Bangkok", year:"numeric", month:"2-digit", day:"2-digit",
    hour:"2-digit", minute:"2-digit", second:"2-digit", hour12:false
  }).formatToParts(new Date()).map(x=>[x.type,x.value]));
  return `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}`;
}
function isoTodayBkk(){ return nowBkk().slice(0,10); }
function validLine(line){ return LINE_CODES.includes(String(line||"").toUpperCase()); }
function lineCode(line){
  const x = String(line||"").trim().toUpperCase();
  if(!validLine(x)) throw Object.assign(new Error("Invalid Line"), {status:400});
  return x;
}
function safeSheetName(v, max=31){
  return String(v||"").replace(/[:\\/?*\[\]]/g,"-").replace(/\s+/g,"_").slice(0,max) || "NA";
}
function samMinutes(raw){
  const s = String(raw??"").trim();
  if(!s) return 0;
  const [m0, sec0=""] = s.split(".");
  const min = Number(m0||0);
  const sec = Number((sec0+"00").slice(0,2)||0);
  if(!Number.isFinite(min)||!Number.isFinite(sec)||min<0||sec<0||sec>59) return 0;
  return min + sec/60;
}
function effDecimal(raw){
  const s=String(raw??"").trim();
  if(s==="") return 0; // Eff ต้องมาจากแผนการผลิต ไม่มีค่า default
  const n=Number(s);
  if(!Number.isFinite(n) || n<0) return 0;
  return n>1 ? n/100 : n;
}
function targetHrFromInputs(rawSam, manpowerRaw=1, effRaw=100){
  const sm = samMinutes(rawSam);
  const manpower = Number(manpowerRaw||1);
  const eff = effDecimal(effRaw);
  if(sm<=0 || !Number.isFinite(manpower) || manpower<=0 || eff<=0) return 0;

  // Target/Hr = (Manpower × 60 ÷ SAM) × Eff
  const raw = (manpower * 60 / sm) * eff;
  return Math.floor(raw + 0.5);
}
function targetHrFromForm(form){
  const f=form||{};
  return targetHrFromInputs(f.sam, f.manpower||1, f.eff);
}
function targetHrFromWorkspace(workspace){
  return targetHrFromForm(workspace?.form||{});
}
function targetHrFromSam(raw){
  // Backward compatible old usage only.
  return targetHrFromInputs(raw,1,100);
}
function ensureArray(v){ return Array.isArray(v)?v:[]; }
function deepClone(v){ return JSON.parse(JSON.stringify(v ?? null)); }
function normalizeFo(v){ return String(v||"").trim().toUpperCase(); }
function timeToMinutes(t){
  const [h,m]=String(t||"00:00").split(":").map(Number);
  return (Number(h)||0)*60+(Number(m)||0);
}
function slotStartMinutes(label){
  return timeToMinutes(String(label||"").split("-")[0]);
}
function isBeforeFirstDayStartServer(index, slotLabel, workspace){
  if(Number(index)!==0) return false;
  const startTime=String(workspace?.form?.startTime||"08:00");
  return slotStartMinutes(slotLabel) < timeToMinutes(startTime);
}
function activeTargetSlotServer(workspace,date,index,slot,ot){
  if(isBeforeFirstDayStartServer(index,slot.label,workspace)) return false;
  if(slot.type==="WORK") return true;
  if(slot.type==="OT" && slot.otNo<=ot) return true;
  return false;
}
function capacityForDay(workspace,date,index){
  const targetHr=targetHrFromWorkspace(workspace);
  const ot=Math.max(0,Math.min(3,otHoursForDay(workspace,date,index)));
  let cap=0;
  for(const slot of SHIFT){
    if(activeTargetSlotServer(workspace,date,index,slot,ot)){
      cap += targetHr*slot.h;
    }
  }
  return cap;
}

function lineDir(line){ return path.join(DATA_ROOT, line); }
function lineWorkbook(line){ return path.join(lineDir(line), `Production_Control_${line}.xlsx`); }
function lineStateFile(line){ return path.join(lineDir(line), "state.json"); }
function lineStateAppliedFile(line){ return path.join(lineDir(line), "state_applied.json"); }
function lineJournalFile(line){ return path.join(lineDir(line), "event_journal.jsonl"); }
function lineDataLogWorkbook(line){ return path.join(lineDir(line), `DATA_LOG_${line}.xlsx`); }
function mirrorLineDataLogWorkbook(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),`DATA_LOG_${line}.xlsx`) : ""; }
function lineDataLogSyncStateFile(line){ return path.join(lineDir(line), "_data_log_sync_state.json"); }

function lineFoDailySummaryWorkbook(line){ return path.join(lineDir(line), `FO_HISTORY_DAILY_SUMMARY_${line}.xlsx`); }
function mirrorLineFoDailySummaryWorkbook(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),`FO_HISTORY_DAILY_SUMMARY_${line}.xlsx`) : ""; }
function allLinesFoDailySummaryWorkbook(){ return path.join(DATA_ROOT, "FO_HISTORY_DAILY_SUMMARY_ALL_LINES.xlsx"); }
function mirrorAllLinesFoDailySummaryWorkbook(){ return ONEDRIVE_ROOT ? path.join(ONEDRIVE_ROOT,"FO_HISTORY_DAILY_SUMMARY_ALL_LINES.xlsx") : ""; }

function allLinesDataLogWorkbook(){ return path.join(DATA_ROOT, "DATA_LOG_ALL_LINES.xlsx"); }
function mirrorAllLinesDataLogWorkbook(){ return ONEDRIVE_ROOT ? path.join(ONEDRIVE_ROOT, "DATA_LOG_ALL_LINES.xlsx") : ""; }
function allLinesSyncStateFile(){ return path.join(DATA_ROOT, "_all_lines_data_log_sync_state.json"); }

function mirrorLineJournalFile(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),"event_journal.jsonl") : ""; }
function mirrorAppliedEventIdsFile(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),"applied_event_ids.json") : ""; }
function mirrorLineStateFile(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),"_state.json") : ""; }
function mirrorDataLogSyncStateFile(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),"_data_log_sync_state.json") : ""; }


function lineAppliedFile(line){ return path.join(lineDir(line), "applied_event_ids.json"); }
function lineBackupDir(line){ return path.join(lineDir(line), "Backup"); }
function assignmentJson(){ return path.join(DATA_ROOT, "FO_ASSIGNMENT.json"); }
function assignmentXlsx(){ return path.join(DATA_ROOT, "FO_ASSIGNMENT.xlsx"); }
function planJson(){ return path.join(DATA_ROOT, "PRODUCTION_PLAN_MASTER.json"); }
function planXlsx(){ return path.join(DATA_ROOT, "PRODUCTION_PLAN_MASTER.xlsx"); }
function summaryXlsx(){ return path.join(DATA_ROOT, "Production_Summary.xlsx"); }
function mirrorLineDir(line){ return ONEDRIVE_ROOT ? path.join(ONEDRIVE_ROOT,line) : ""; }
function mirrorLineWorkbook(line){ return ONEDRIVE_ROOT ? path.join(mirrorLineDir(line),`Production_Control_${line}.xlsx`) : ""; }
function mirrorCentral(file){ return ONEDRIVE_ROOT ? path.join(ONEDRIVE_ROOT,path.basename(file)) : ""; }

async function exists(p){ try{ await fsp.access(p); return true; }catch{return false;} }
async function ensureDir(p){ await fsp.mkdir(p,{recursive:true}); }
async function atomicWriteText(file, text){
  await ensureDir(path.dirname(file));
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  await fsp.writeFile(tmp, text, "utf8");
  await fsp.rename(tmp,file).catch(async()=>{
    await fsp.copyFile(tmp,file);
    await fsp.unlink(tmp).catch(()=>{});
  });
}
async function atomicWriteJson(file, value){ return atomicWriteText(file, JSON.stringify(value,null,2)); }
async function readJson(file, fallback){
  try{return JSON.parse(await fsp.readFile(file,"utf8"));}catch{return deepClone(fallback);}
}
async function appendJsonl(file, value){
  await ensureDir(path.dirname(file));
  await fsp.appendFile(file, JSON.stringify(value)+"\n", "utf8");
}
async function readJsonl(file){
  try{
    const text = await fsp.readFile(file,"utf8");
    return text.split(/\r?\n/).filter(Boolean).map(x=>{try{return JSON.parse(x);}catch{return null;}}).filter(Boolean);
  }catch{return [];}
}
async function saveWorkbookAtomic(wb,file){
  await ensureDir(path.dirname(file));
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  try{
    await wb.xlsx.writeFile(tmp);
    try{ await fsp.rename(tmp,file); }
    catch(e){ await fsp.copyFile(tmp,file); await fsp.unlink(tmp).catch(()=>{}); }
  }catch(e){
    await fsp.unlink(tmp).catch(()=>{});
    const err = new Error(`Excel save failed: ${e.message}`);
    err.code="EXCEL_LOCKED_OR_SAVE_FAILED";
    throw err;
  }
}
async function withLineLock(line, fn){
  const key = lineCode(line);
  const prev = lineLocks.get(key) || Promise.resolve();
  let release;
  const gate = new Promise(r=>release=r);
  const chain = prev.catch(()=>{}).then(()=>gate);
  lineLocks.set(key,chain);
  await prev.catch(()=>{});
  try{return await fn();}
  finally{ release(); if(lineLocks.get(key)===chain) lineLocks.delete(key); }
}
async function withCentralLock(fn){
  const prev = centralLock;
  let release;
  centralLock = new Promise(r=>release=r);
  await prev.catch(()=>{});
  try{return await fn();} finally{release();}
}

function styleHeader(ws,row,cols){
  for(let c=1;c<=cols;c++){
    const cell=ws.getCell(row,c);
    cell.font={bold:true,color:{argb:"FFFFFFFF"}};
    cell.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF0F172A"}};
    cell.alignment={horizontal:"center",vertical:"middle"};
    cell.border={bottom:{style:"thin",color:{argb:"FFCBD5E1"}}};
  }
}
function setWidths(ws,widths){ widths.forEach((w,i)=>ws.getColumn(i+1).width=w); }

async function createLineWorkbook(line,file){
  const wb = new ExcelJS.Workbook();
  wb.creator="Production Control";
  wb.created=new Date();

  let ws=wb.addWorksheet("INFO");
  ws.addRow(["Key","Value"]); styleHeader(ws,1,2);
  ws.addRows([
    ["Line",line],["Created At",nowBkk()],["Storage","Local Master"],["Workbook",path.basename(file)],
    ["Rule","DATA_LOG continuous across FO; one RPT_<FO> sheet per FO"]
  ]); setWidths(ws,[28,70]);

  ws=wb.addWorksheet("FO_INDEX");
  ws.addRow(["FO","Style Name","Color","Order Qty","Cutting Qty","Previous Cumulative","SAM","Eff (%)","Manpower","Target/Hr","Start Date","Start Time","Status","OK","Repair","NG","Production OK+NG","Good Cumulative","FO Cumulative","ยอดงานที่ต้องผลิต","Report Sheet","Updated At"]);
  styleHeader(ws,1,22); setWidths(ws,[18,28,18,12,12,18,12,10,10,12,14,12,14,10,10,10,18,18,18,18,20,22]);

  ws=wb.addWorksheet("DATA_LOG");
  ws.addRow(["Event ID","Date Time","Production Date","Line","FO","Style","Color","Time Slot","Type","Qty","Source","Status"]);
  styleHeader(ws,1,12); setWidths(ws,[34,22,16,12,18,26,18,18,12,8,14,14]);

  ws=wb.addWorksheet("DAILY_SUMMARY");
  ws.addRow(["Production Date","Line","FO","Target Today","OK","Repair","NG","Production OK+NG","Achievement %","FO Cumulative","Good Cumulative","Remaining Good","OT Hours","Status","Updated At"]);
  styleHeader(ws,1,15); setWidths(ws,[16,12,18,14,10,10,10,18,16,18,18,18,12,14,22]);

  ws=wb.addWorksheet("AUDIT_TRAIL");
  ws.addRow(["Date Time","Line","FO","Action","Details"]); styleHeader(ws,1,5); setWidths(ws,[22,12,18,24,70]);

  ws=wb.addWorksheet("PROBLEM_LOG");
  ws.addRow(["Date Time","Production Date","Line","FO","Time Slot","Reason","Note","User"]); styleHeader(ws,1,8); setWidths(ws,[22,16,12,18,18,22,60,18]);

  await saveWorkbookAtomic(wb,file);
}

async function ensureLineWorkbook(line){
  const file=lineWorkbook(line);
  await ensureDir(lineDir(line));
  await ensureDir(lineBackupDir(line));
  if(!(await exists(file))) await createLineWorkbook(line,file);
  if(!(await exists(lineAppliedFile(line)))) await atomicWriteJson(lineAppliedFile(line),[]);
  if(!(await exists(lineStateFile(line)))) await atomicWriteJson(lineStateFile(line),null);
  return file;
}


async function syncStandaloneDataLog(line, force=false){
  const outFile = lineDataLogWorkbook(line);
  const journalFile = lineJournalFile(line);

  await ensureDir(lineDir(line));

  const journalStat = await fsp.stat(journalFile).catch(()=>null);
  const syncState = await readJson(lineDataLogSyncStateFile(line), {journalMtimeMs:0});

  // If there is no journal yet, still create an empty dashboard-friendly file once.
  if(!force && await exists(outFile)){
    if(!journalStat) return {synced:true,unchanged:true,path:outFile};
    if(Number(syncState?.journalMtimeMs||0) >= journalStat.mtimeMs){
      return {synced:true,unchanged:true,path:outFile};
    }
  }

  const events = await readJsonl(journalFile);
  const wb = new ExcelJS.Workbook();
  wb.creator = "Production Control";
  wb.created = new Date();

  const ws = wb.addWorksheet("DATA_LOG");
  ws.addRow([
    "Event ID","Date Time","Production Date","Line","FO","Style","Color",
    "Time Slot","Type","Qty","Source","Status"
  ]);
  styleHeader(ws,1,12);
  setWidths(ws,[34,22,16,12,18,26,18,18,12,8,14,14]);

  for(const e of events){
    ws.addRow([
      e.eventId||"",
      e.dateTime||"",
      e.productionDate||"",
      line,
      e.fo||"",
      e.style||"",
      e.color||"",
      e.timeSlot||"",
      String(e.type||"").toUpperCase(),
      Number(e.qty||1),
      e.source||"",
      e.status||"RUNNING"
    ]);
  }

  try{
    await saveWorkbookAtomic(wb,outFile);
    await atomicWriteJson(lineDataLogSyncStateFile(line),{
      journalMtimeMs: journalStat?.mtimeMs || 0,
      rows: events.length,
      syncedAt: nowBkk()
    });

    // Mirror the standalone DATA_LOG workbook independently from the main workbook.
    if(ONEDRIVE_ROOT){
      await mirrorFile(outFile, mirrorLineDataLogWorkbook(line)).catch(()=>{});
    }

    return {synced:true,rows:events.length,path:outFile};
  }catch(e){
    // Derived file only: failure must never stop production.
    return {queued:true,error:e.message,path:outFile};
  }
}



function thaiDateFromIso(iso){
  const m=String(iso||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(!m) return String(iso||"");
  return `${m[3]}/${m[2]}/${Number(m[1])+543}`;
}

function foDailySummaryHeaders(){
  return [
    "วันที่",
    "Line",
    "FO",
    "Target วันนี้",
    "OK วันนี้",
    "Repair วันนี้",
    "NG วันนี้",
    "ผลิตวันนี้ (OK+NG)",
    "Achievement",
    "FO Cum (OK+NG)",
    "Good Cum (OK)",
    "Remaining (งานดี)",
    "OT",
    "สถานะวัน"
  ];
}


function buildFoDailySummaryRowsFromWorkspace(workspace){
  const f=workspace?.form||{};
  const line=String(workspace?.line||"");
  const fo=String(f.fo||"").trim();
  if(!fo) return [];

  const cutting=Number(f.qty||0);
  const prev=Number(f.previousCum||0);
  const tHr=targetHrFromForm(f);
  const today=isoTodayBkk();
  const daily=dailyActuals(workspace);

  const planned=[];
  let remainingPlan=Math.max(0,cutting-prev);
  let date=f.start||isoTodayBkk();
  let idx=0;

  // Create every planned day immediately, even before production starts.
  while((remainingPlan>0 || planned.length===0) && idx<120){
    const target=targetForDay(workspace,date,idx,remainingPlan);
    planned.push({date,index:idx,target});
    remainingPlan=Math.max(0,remainingPlan-target);
    date=nextDateIso(date);
    idx++;
    if(tHr<=0) break;
  }

  // Keep actual dates even if they fall outside the current recalculated plan.
  for(const actualDate of Object.keys(daily).sort()){
    if(!planned.some(x=>x.date===actualDate)){
      planned.push({date:actualDate,index:planned.length,target:0});
    }
  }
  planned.sort((a,b)=>a.date.localeCompare(b.date));
  planned.forEach((d,i)=>{ d.index=i; });

  let cumOk=0;
  let cumNg=0;
  const rows=[];

  for(let i=0;i<planned.length;i++){
    const p=planned[i];
    const a=daily[p.date]||{ok:0,repair:0,ng:0};
    cumOk+=Number(a.ok||0);
    cumNg+=Number(a.ng||0);

    const production=Number(a.ok||0)+Number(a.ng||0);
    const achievement=p.target>0?production/p.target:0;
    const foCum=prev+cumOk+cumNg;
    const goodCum=prev+cumOk;
    const remaining=Math.max(0,cutting-foCum);

    let dayStatus="PLANNED";
    if(p.date<today) dayStatus="LOCKED";
    else if(p.date===today) dayStatus="CURRENT";

    if(String(workspace.prodStatus||"").toUpperCase()==="CLOSED" && p.date<=today){
      dayStatus="LOCKED";
    }

    rows.push({
      productionDate:p.date,
      thaiDate:thaiDateFromIso(p.date),
      line,
      fo,
      target:Number(p.target||0),
      ok:Number(a.ok||0),
      repair:Number(a.repair||0),
      ng:Number(a.ng||0),
      production,
      achievement,
      foCum,
      goodCum,
      remaining,
      ot:Number(otHoursForDay(workspace,p.date,i)||0),
      dayStatus
    });
  }

  return rows;
}

async function readFoDailySummaryRowsFromState(line){
  const workspace=await readJson(lineStateFile(line),null);
  if(!workspace) return [];
  workspace.line=line;
  return buildFoDailySummaryRowsFromWorkspace(workspace);
}

async function readDailySummaryRowsFromMainWorkbook(line){
  const file=lineWorkbook(line);
  if(!(await exists(file))) return [];

  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws=wb.getWorksheet("DAILY_SUMMARY");
  if(!ws) return [];

  const rows=[];
  for(let r=2;r<=ws.rowCount;r++){
    const date=String(ws.getCell(r,1).value||"").trim();
    const fo=String(ws.getCell(r,3).value||"").trim();
    if(!date || !fo) continue;

    rows.push({
      productionDate:date,
      thaiDate:thaiDateFromIso(date),
      line:String(ws.getCell(r,2).value||line),
      fo,
      target:Number(ws.getCell(r,4).value||0),
      ok:Number(ws.getCell(r,5).value||0),
      repair:Number(ws.getCell(r,6).value||0),
      ng:Number(ws.getCell(r,7).value||0),
      production:Number(ws.getCell(r,8).value||0),
      achievement:Number(ws.getCell(r,9).value||0),
      foCum:Number(ws.getCell(r,10).value||0),
      goodCum:Number(ws.getCell(r,11).value||0),
      remaining:Number(ws.getCell(r,12).value||0),
      ot:Number(ws.getCell(r,13).value||0),
      dayStatus:String(ws.getCell(r,14).value||"")
    });
  }

  rows.sort((a,b)=>{
    const d=a.productionDate.localeCompare(b.productionDate);
    if(d) return d;
    return a.fo.localeCompare(b.fo);
  });
  return rows;
}

async function writeFoDailySummaryWorkbook(rows,outFile,sheetName){
  const wb=new ExcelJS.Workbook();
  wb.creator="Production Control";
  wb.created=new Date();

  const ws=wb.addWorksheet(sheetName);
  ws.views=[{state:"frozen",ySplit:1}];
  ws.autoFilter={from:"A1",to:"N1"};

  ws.addRow(foDailySummaryHeaders());
  styleHeader(ws,1,14);
  setWidths(ws,[16,12,18,14,12,14,12,20,14,18,18,18,12,14]);

  for(const x of rows){
    ws.addRow([
      x.thaiDate,
      x.line,
      x.fo,
      x.target,
      x.ok,
      x.repair,
      x.ng,
      x.production,
      x.achievement,
      x.foCum,
      x.goodCum,
      x.remaining,
      x.ot,
      x.dayStatus
    ]);
  }

  ws.getColumn(9).numFmt="0.0%";

  for(let r=2;r<=ws.rowCount;r++){
    const status=String(ws.getCell(r,14).value||"");
    if(status==="CURRENT"){
      ws.getCell(r,14).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFE0F2FE"}};
      ws.getCell(r,14).font={bold:true,color:{argb:"FF075985"}};
    }else if(status==="LOCKED"){
      ws.getCell(r,14).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF1F5F9"}};
      ws.getCell(r,14).font={color:{argb:"FF475569"}};
    }else if(status==="PLANNED"){
      ws.getCell(r,14).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFECFDF5"}};
      ws.getCell(r,14).font={color:{argb:"FF047857"}};
    }
  }

  await saveWorkbookAtomic(wb,outFile);
  return {rows:rows.length,path:outFile};
}



async function ensureStandaloneFoDailySummaryWorkbook(line){
  const outFile=lineFoDailySummaryWorkbook(line);
  if(await exists(outFile)) return {exists:true,path:outFile};

  // Create the file immediately even when the Line has no FO yet.
  // This guarantees LINE01-LINE18 always have their own analytics workbook.
  const result=await writeFoDailySummaryWorkbook([],outFile,"FO_HISTORY_DAILY_SUMMARY");

  if(ONEDRIVE_ROOT){
    await mirrorFile(outFile,mirrorLineFoDailySummaryWorkbook(line)).catch(()=>{});
  }
  return {created:true,...result};
}

async function ensureAllLinesFoDailySummaryWorkbook(){
  const outFile=allLinesFoDailySummaryWorkbook();
  if(await exists(outFile)) return {exists:true,path:outFile};

  const result=await writeFoDailySummaryWorkbook([],outFile,"FO_HISTORY_ALL_LINES");
  if(ONEDRIVE_ROOT){
    await mirrorFile(outFile,mirrorAllLinesFoDailySummaryWorkbook()).catch(()=>{});
  }
  return {created:true,...result};
}

async function ensureFoHistoryDailySummaryNow(line){
  // Called immediately after a Line workspace/FO is saved.
  // Keeps 1 Line = 1 continuous FO_HISTORY_DAILY_SUMMARY_LINEXX.xlsx
  const perLine=await syncStandaloneFoDailySummary(line);
  const allLines=await syncAllLinesFoDailySummary().catch(e=>({queued:true,error:e.message}));
  return {perLine,allLines};
}

async function syncStandaloneFoDailySummary(line){
  const outFile=lineFoDailySummaryWorkbook(line);
  try{
    await ensureStandaloneFoDailySummaryWorkbook(line);

    // V1.20: build directly from state.json / current workspace.
    // Do not depend on DAILY_SUMMARY inside the main workbook.
    const rows=await readFoDailySummaryRowsFromState(line);
    const result=await writeFoDailySummaryWorkbook(rows,outFile,"FO_HISTORY_DAILY_SUMMARY");

    if(ONEDRIVE_ROOT){
      await mirrorFile(outFile,mirrorLineFoDailySummaryWorkbook(line)).catch(()=>{});
    }
    return result;
  }catch(e){
    return {queued:true,error:e.message,path:outFile};
  }
}

async function syncAllLinesFoDailySummary(){
  const outFile=allLinesFoDailySummaryWorkbook();
  try{
    await ensureAllLinesFoDailySummaryWorkbook();

    const all=[];
    for(const line of LINE_CODES){
      const rows=await readFoDailySummaryRowsFromState(line).catch(()=>[]);
      all.push(...rows);
    }

    all.sort((a,b)=>{
      const d=a.productionDate.localeCompare(b.productionDate);
      if(d) return d;
      const l=a.line.localeCompare(b.line);
      if(l) return l;
      return a.fo.localeCompare(b.fo);
    });

    const result=await writeFoDailySummaryWorkbook(all,outFile,"FO_HISTORY_ALL_LINES");

    if(ONEDRIVE_ROOT){
      await mirrorFile(outFile,mirrorAllLinesFoDailySummaryWorkbook()).catch(()=>{});
    }
    return result;
  }catch(e){
    return {queued:true,error:e.message,path:outFile};
  }
}


async function latestJournalMtimeAcrossLines(){
  let maxMtime=0;
  for(const line of LINE_CODES){
    const st=await fsp.stat(lineJournalFile(line)).catch(()=>null);
    if(st && st.mtimeMs>maxMtime) maxMtime=st.mtimeMs;
  }
  return maxMtime;
}

async function syncAllLinesDataLog(force=false){
  const outFile=allLinesDataLogWorkbook();
  const newestJournalMtime=await latestJournalMtimeAcrossLines();
  const syncState=await readJson(allLinesSyncStateFile(),{journalMtimeMs:0});

  if(!force && await exists(outFile)){
    if(Number(syncState?.journalMtimeMs||0)>=newestJournalMtime){
      // Even when unchanged locally, make sure OneDrive has the current combined file.
      if(ONEDRIVE_ROOT){
        await mirrorFile(outFile,mirrorAllLinesDataLogWorkbook()).catch(()=>{});
      }
      return {synced:true,unchanged:true,path:outFile};
    }
  }

  const all=[];
  for(const line of LINE_CODES){
    const rows=await readJsonl(lineJournalFile(line));
    for(const e of rows){
      all.push({
        eventId:e.eventId||"",
        dateTime:e.dateTime||"",
        productionDate:e.productionDate||"",
        line,
        fo:e.fo||"",
        style:e.style||"",
        color:e.color||"",
        timeSlot:e.timeSlot||"",
        type:String(e.type||"").toUpperCase(),
        qty:Number(e.qty||1),
        source:e.source||"",
        status:e.status||"RUNNING"
      });
    }
  }

  // Stable chronological order; Event ID is final tie-breaker.
  all.sort((a,b)=>{
    const da=String(a.dateTime||""), db=String(b.dateTime||"");
    if(da<db) return -1;
    if(da>db) return 1;
    return String(a.eventId).localeCompare(String(b.eventId));
  });

  const wb=new ExcelJS.Workbook();
  wb.creator="Production Control";
  wb.created=new Date();

  const ws=wb.addWorksheet("DATA_LOG_ALL_LINES");
  ws.addRow([
    "Event ID","Date Time","Production Date","Line","FO","Style","Color",
    "Time Slot","Type","Qty","Source","Status"
  ]);
  styleHeader(ws,1,12);
  setWidths(ws,[34,22,16,12,18,26,18,18,12,8,14,14]);
  ws.views=[{state:"frozen",ySplit:1}];

  for(const e of all){
    ws.addRow([
      e.eventId,e.dateTime,e.productionDate,e.line,e.fo,e.style,e.color,
      e.timeSlot,e.type,e.qty,e.source,e.status
    ]);
  }

  try{
    await saveWorkbookAtomic(wb,outFile);
    await atomicWriteJson(allLinesSyncStateFile(),{
      journalMtimeMs:newestJournalMtime,
      rows:all.length,
      syncedAt:nowBkk()
    });

    if(ONEDRIVE_ROOT){
      await mirrorFile(outFile,mirrorAllLinesDataLogWorkbook());
      await mirrorFile(allLinesSyncStateFile(),path.join(ONEDRIVE_ROOT,"_all_lines_data_log_sync_state.json")).catch(()=>{});
    }
    return {synced:true,rows:all.length,path:outFile};
  }catch(e){
    // This is a derived dashboard file. Failure must never stop production.
    return {queued:true,error:e.message,path:outFile};
  }
}

async function mirrorLineSupportFiles(line){
  if(!ONEDRIVE_ROOT) return {disabled:true};
  await ensureDir(mirrorLineDir(line));
  const jobs=[
    [lineJournalFile(line),mirrorLineJournalFile(line)],
    [lineAppliedIdsFile(line),mirrorAppliedEventIdsFile(line)],
    [lineStateFile(line),mirrorLineStateFile(line)],
    [lineDataLogSyncStateFile(line),mirrorDataLogSyncStateFile(line)]
  ];
  const result=[];
  for(const [src,dst] of jobs){
    if(await exists(src)){
      try{ await mirrorFile(src,dst); result.push({src,ok:true}); }
      catch(e){ result.push({src,ok:false,error:e.message}); }
    }
  }
  return {mirrored:result};
}

async function ensureCentralFiles(){
  await ensureDir(DATA_ROOT);
  if(!(await exists(assignmentJson()))) await atomicWriteJson(assignmentJson(),{});
  if(!(await exists(planJson()))) await atomicWriteJson(planJson(),[]);
  await syncAssignmentsWorkbook().catch(()=>{});
  await syncPlanWorkbook().catch(()=>{});
  if(!(await exists(summaryXlsx()))){
    const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet("LINE_STATUS");
    ws.addRow(["Line","FO","Style","Color","Status","OK","Repair","NG","Good Cumulative","FO Cumulative","ยอดงานที่ต้องผลิต","Updated At"]);
    styleHeader(ws,1,12); setWidths(ws,[12,18,26,18,14,10,10,10,18,18,18,22]);
    await saveWorkbookAtomic(wb,summaryXlsx()).catch(()=>{});
  }
}

async function bootstrap(){
  await ensureDir(DATA_ROOT);
  if(ONEDRIVE_ROOT) await ensureDir(ONEDRIVE_ROOT).catch(()=>{});

  // Create core workbooks/directories first.
  for(const line of LINE_CODES){
    await ensureLineWorkbook(line);
    await ensureStandaloneFoDailySummaryWorkbook(line).catch(()=>{});
  }
  await ensureCentralFiles();
  await ensureAllLinesFoDailySummaryWorkbook().catch(()=>{});

  // V1.31 startup recovery sequence:
  // Journal is source of truth. Reconcile first, then rebuild Excel/report.
  for(const line of LINE_CODES){
    try{
      await reconcileLineStateFromJournal(line);
      await processLine(line);
    }catch(e){
      console.error("bootstrap recovery",line,e.message);
    }
  }

  // Derived all-line outputs.
  await syncAllLinesDataLog(true).catch(()=>{});
  await syncAllLinesFoDailySummary().catch(()=>{});
  await syncAssignmentsWorkbook().catch(()=>{});
  await syncPlanWorkbook().catch(()=>{});
}

async function loadAppliedSet(line){
  if(appliedIdCache.has(line)) return appliedIdCache.get(line);
  const arr=ensureArray(await readJson(lineAppliedFile(line),[]));
  const set=new Set(arr.map(String)); appliedIdCache.set(line,set); return set;
}
async function saveAppliedSet(line,set){
  appliedIdCache.set(line,set);
  await atomicWriteJson(lineAppliedFile(line),[...set]);
}
async function loadJournalIds(line){
  if(journalIdCache.has(line)) return journalIdCache.get(line);
  const events=await readJsonl(lineJournalFile(line));
  const set=new Set(events.map(x=>String(x.eventId||""))); journalIdCache.set(line,set); return set;
}
async function journalEvent(line,event){
  const ids=await loadJournalIds(line);
  const id=String(event.eventId||"");
  if(!id) throw Object.assign(new Error("eventId required"),{status:400});
  if(ids.has(id)) return false;
  await appendJsonl(lineJournalFile(line),event);
  ids.add(id); return true;
}


function preserveProblemMap(prodRows){
  const out={};
  for(const [key,v] of Object.entries(prodRows||{})){
    if(String(v?.problem||"").trim()){
      out[key]=String(v.problem);
    }
  }
  return out;
}

async function rebuildProdRowsFromJournal(line,workspace){
  if(!workspace) return workspace;
  const fo=normalizeFo(workspace?.form?.fo);
  if(!fo) return workspace;

  const problems=preserveProblemMap(workspace.prodRows||{});
  const rows={};

  // Keep problem-only rows even before an event exists in that slot.
  for(const [key,problem] of Object.entries(problems)){
    rows[key]={ok:0,repair:0,ng:0,problem};
  }

  const events=await readJsonl(lineJournalFile(line));
  for(const e of events){
    if(normalizeFo(e.fo)!==fo) continue;

    const date=String(e.productionDate||"").trim();
    const slot=String(e.timeSlot||"").trim();
    const type=String(e.type||"").toUpperCase();
    if(!date || !slot || !["OK","REPAIR","NG"].includes(type)) continue;

    const key=`${date}|${slot}`;
    if(!rows[key]) rows[key]={ok:0,repair:0,ng:0,problem:problems[key]||""};

    const qty=Math.max(0,Number(e.qty||1));
    if(type==="OK") rows[key].ok+=qty;
    else if(type==="REPAIR") rows[key].repair+=qty;
    else if(type==="NG") rows[key].ng+=qty;
  }

  workspace.prodRows=rows;
  return workspace;
}

async function reconcileLineStateFromJournal(line){
  const file=lineStateFile(line);
  const workspace=await readJson(file,null);
  if(!workspace) return null;

  const before=JSON.stringify(workspace.prodRows||{});
  workspace.line=line;
  await rebuildProdRowsFromJournal(line,workspace);
  const after=JSON.stringify(workspace.prodRows||{});

  if(before!==after){
    await atomicWriteJson(file,workspace);
  }
  return workspace;
}

async function pendingEvents(line){
  const events=await readJsonl(lineJournalFile(line));
  const applied=await loadAppliedSet(line);
  return events.filter(e=>e.eventId && !applied.has(String(e.eventId)));
}

async function workbookHasEvent(ws,id){
  const col=1;
  for(let r=Math.max(2,ws.rowCount-2000); r<=ws.rowCount; r++){
    if(String(ws.getCell(r,col).value||"")===String(id)) return true;
  }
  return false;
}

async function applyPendingEvents(line){
  const pending=await pendingEvents(line);
  if(!pending.length){ lastExcelError.delete(line); return {saved:0,pending:0}; }
  const file=await ensureLineWorkbook(line);
  const wb=new ExcelJS.Workbook();
  try{
    await wb.xlsx.readFile(file);
    ensureCurrentWorkbookHeaders(wb);
    const ws=wb.getWorksheet("DATA_LOG") || wb.addWorksheet("DATA_LOG");
    const newly=[];
    for(const e of pending){
      if(!(await workbookHasEvent(ws,e.eventId))){
        ws.addRow([
          e.eventId||"", e.dateTime||nowBkk(), e.productionDate||"", line, e.fo||"", e.style||"", e.color||"",
          e.timeSlot||"", String(e.type||"").toUpperCase(), Number(e.qty||1), e.source||"", e.status||"RUNNING"
        ]);
      }
      newly.push(String(e.eventId));
    }
    await saveWorkbookAtomic(wb,file);
    const applied=await loadAppliedSet(line);
    newly.forEach(id=>applied.add(id));
    await saveAppliedSet(line,applied);
    lastExcelError.delete(line); lastExcelSavedAt.set(line,nowBkk());
    return {saved:newly.length,pending:(await pendingEvents(line)).length};
  }catch(e){
    lastExcelError.set(line,e.message);
    return {saved:0,pending:pending.length,error:e.message};
  }
}

function workspaceTotals(workspace){
  let ok=0,repair=0,ng=0;
  const rows=workspace?.prodRows || {};
  for(const v of Object.values(rows)){
    ok+=Number(v?.ok||0); repair+=Number(v?.repair||0); ng+=Number(v?.ng||0);
  }
  return {ok,repair,ng,production:ok+ng};
}
function dailyActuals(workspace){
  const out={};
  for(const [key,v] of Object.entries(workspace?.prodRows||{})){
    const date=String(key).split("|")[0];
    if(!out[date]) out[date]={ok:0,repair:0,ng:0};
    out[date].ok+=Number(v?.ok||0); out[date].repair+=Number(v?.repair||0); out[date].ng+=Number(v?.ng||0);
  }
  return out;
}
function otHoursForDay(workspace,date,index){
  const map=workspace?.otByDay||{};
  const exact=`DAY_${index+1}_${date}`;
  if(Object.prototype.hasOwnProperty.call(map,exact)) return Number(map[exact]||0);
  const k=Object.keys(map).find(x=>x.endsWith(`_${date}`));
  return k?Number(map[k]||0):0;
}
function nextDateIso(iso){
  const d=new Date(`${iso}T12:00:00Z`); d.setUTCDate(d.getUTCDate()+1); return d.toISOString().slice(0,10);
}
function planDatesFromWorkspace(workspace,maxDays=120){
  const start=workspace?.form?.start || isoTodayBkk();
  const dates=[]; let d=start;
  for(let i=0;i<maxDays;i++){ dates.push(d); d=nextDateIso(d); }
  return dates;
}
function targetForDay(workspace,date,index,remainingBefore){
  const capacity=capacityForDay(workspace,date,index);
  return Math.max(0,Math.min(capacity,remainingBefore));
}


function setRowValues(ws,rowNo,values){
  values.forEach((v,i)=>{ ws.getCell(rowNo,i+1).value=v; });
}

function ensureCurrentWorkbookHeaders(wb){
  const foIndex=wb.getWorksheet("FO_INDEX");
  if(foIndex){
    setRowValues(foIndex,1,[
      "FO","Style Name","Color","Order Qty","Cutting Qty","Previous Cumulative","SAM","Eff (%)","Manpower","Target/Hr",
      "Start Date","Start Time","Status","OK","Repair","NG","Production OK+NG","Good Cumulative","FO Cumulative",
      "ยอดงานที่ต้องผลิต","Report Sheet","Updated At"
    ]);
    styleHeader(foIndex,1,22);
  }

  const daily=wb.getWorksheet("DAILY_SUMMARY");
  if(daily){
    setRowValues(daily,1,[
      "Production Date","Line","FO","Target Today","OK","Repair","NG","Production OK+NG",
      "Achievement %","FO Cumulative","Good Cumulative","ยอดงานที่ต้องผลิต","OT Hours","Status","Updated At"
    ]);
    styleHeader(daily,1,15);
  }

  const log=wb.getWorksheet("DATA_LOG");
  if(log){
    setRowValues(log,1,[
      "Event ID","Date Time","Production Date","Line","FO","Style","Color","Time Slot",
      "Type","Qty","Source","Status"
    ]);
    styleHeader(log,1,12);
  }
}

async function upsertFoIndex(ws,workspace){
  const f=workspace?.form||{}; const fo=String(f.fo||"").trim(); if(!fo) return;
  const totals=workspaceTotals(workspace);
  const cutting=Number(f.qty||0), prev=Number(f.previousCum||0);
  const goodCum=prev+totals.ok, foCum=prev+totals.production, rem=Math.max(0,cutting-foCum);
  let row=-1;
  for(let r=2;r<=ws.rowCount;r++) if(normalizeFo(ws.getCell(r,1).value)===normalizeFo(fo)){row=r;break;}
  if(row<0) row=ws.rowCount+1;
  const report=`RPT_${safeSheetName(fo,25)}`;
  const values=[fo,f.style||"",f.color||"",Number(f.orderQty||0),cutting,prev,f.sam||"",f.eff===""||f.eff==null?0:Number(f.eff||0),Number(f.manpower||0),targetHrFromForm(f),f.start||"",f.startTime||"08:00",workspace.prodStatus||"READY",totals.ok,totals.repair,totals.ng,totals.production,goodCum,foCum,rem,report,nowBkk()];
  values.forEach((v,i)=>ws.getCell(row,i+1).value=v);
}

async function replaceRowsForFo(ws,fo,newRows,foColIndex){
  const keep=[];
  for(let r=2;r<=ws.rowCount;r++){
    const vals=[]; for(let c=1;c<=ws.columnCount;c++) vals.push(ws.getCell(r,c).value);
    if(normalizeFo(vals[foColIndex-1])!==normalizeFo(fo)) keep.push(vals);
  }
  if(ws.rowCount>1) ws.spliceRows(2,ws.rowCount-1);
  keep.forEach(r=>ws.addRow(r)); newRows.forEach(r=>ws.addRow(r));
}

async function updateDailySummary(ws,workspace){
  const f=workspace?.form||{};
  const fo=String(f.fo||"").trim();
  if(!fo) return;

  const daily=dailyActuals(workspace);
  const actualDates=Object.keys(daily).sort();
  const cutting=Number(f.qty||0);
  const prev=Number(f.previousCum||0);
  const tHr=targetHrFromForm(f);
  const today=isoTodayBkk();

  // Build the complete planned date range at once.
  const planned=[];
  let remainingPlan=Math.max(0,cutting-prev);
  let date=f.start||isoTodayBkk();
  let idx=0;

  while((remainingPlan>0 || planned.length===0) && idx<120){
    const target=targetForDay(workspace,date,idx,remainingPlan);
    planned.push({date,index:idx,target});
    remainingPlan=Math.max(0,remainingPlan-target);
    date=nextDateIso(date);
    idx++;
    if(tHr<=0) break;
  }

  // Preserve any actual dates outside the recalculated plan.
  for(const d of actualDates){
    if(!planned.some(x=>x.date===d)){
      planned.push({date:d,index:planned.length,target:0});
    }
  }
  planned.sort((a,b)=>a.date.localeCompare(b.date));
  planned.forEach((d,i)=>{ d.index=i; });

  let cumOk=0,cumNg=0;
  const rows=[];

  for(let i=0;i<planned.length;i++){
    const p=planned[i];
    const a=daily[p.date]||{ok:0,repair:0,ng:0};
    cumOk+=a.ok;
    cumNg+=a.ng;

    const prod=a.ok+a.ng;
    const ach=p.target>0?prod/p.target:0;
    const goodCum=prev+cumOk;
    const foCum=prev+cumOk+cumNg;

    // Current business rule:
    // ยอดงานที่ต้องผลิต = Cutting Qty - ยอดผลิตสะสม (OK+NG)
    const remaining=Math.max(0,cutting-foCum);

    let dayStatus="PLANNED";
    if(p.date<today) dayStatus="LOCKED";
    else if(p.date===today) dayStatus="CURRENT";
    if(String(workspace.prodStatus||"").toUpperCase()==="CLOSED" && p.date<=today){
      dayStatus="LOCKED";
    }

    rows.push([
      p.date,
      workspace.line||"",
      fo,
      p.target,
      a.ok,
      a.repair,
      a.ng,
      prod,
      ach,
      foCum,
      goodCum,
      remaining,
      otHoursForDay(workspace,p.date,i),
      dayStatus,
      nowBkk()
    ]);
  }

  await replaceRowsForFo(ws,fo,rows,3);
  ws.getColumn(9).numFmt="0.0%";
}

async function updateReportSheet(wb,workspace){
  const f=workspace?.form||{};
  const fo=String(f.fo||"").trim();
  if(!fo) return {skipped:true};

  // 1 FO = 1 Sheet
  const name=`RPT_${safeSheetName(fo,25)}`;
  let ws=wb.getWorksheet(name);
  if(ws) wb.removeWorksheet(ws.id);
  ws=wb.addWorksheet(name);

  ws.pageSetup={orientation:"landscape",fitToPage:true,fitToWidth:1,fitToHeight:0,paperSize:9};
  ws.views=[{state:"frozen",ySplit:1}];

  const cutting=Number(f.qty||0);
  const orderQty=Number(f.orderQty||0);
  const prev=Number(f.previousCum||0);
  const targetHr=targetHrFromForm(f);
  const totals=workspaceTotals(workspace);

  // Actual rows come from the same workspace used by the web page.
  const prodRows=workspace?.prodRows||{};
  const actualByDate={};
  for(const [key,v] of Object.entries(prodRows)){
    const [date,slot]=String(key).split("|");
    if(!date || !slot) continue;
    if(!actualByDate[date]) actualByDate[date]={};
    actualByDate[date][slot]=v||{};
  }

  // ------------------------------------------------------------
  // CREATE ALL PLANNED DAYS AT ONCE UNTIL FO IS FINISHED
  // ------------------------------------------------------------
  const plannedDays=[];
  let planRemaining=Math.max(0,cutting-prev);
  let date=f.start||isoTodayBkk();
  let dayIndex=0;

  while((planRemaining>0 || plannedDays.length===0) && dayIndex<120){
    const ot=Math.max(0,Math.min(3,otHoursForDay(workspace,date,dayIndex)));
    const capacity=capacityForDay(workspace,date,dayIndex);
    const dayTarget=Math.max(0,Math.min(capacity,planRemaining));
    plannedDays.push({date,index:dayIndex,ot,dayTarget});
    planRemaining=Math.max(0,planRemaining-dayTarget);
    date=nextDateIso(date);
    dayIndex++;
    if(targetHr<=0) break;
  }

  // Preserve real production dates outside the recalculated plan.
  for(const actualDate of Object.keys(actualByDate).sort()){
    if(!plannedDays.some(x=>x.date===actualDate)){
      plannedDays.push({
        date:actualDate,
        index:plannedDays.length,
        ot:Math.max(0,Math.min(3,otHoursForDay(workspace,actualDate,plannedDays.length))),
        dayTarget:0
      });
    }
  }
  plannedDays.sort((a,b)=>a.date.localeCompare(b.date));
  plannedDays.forEach((d,i)=>{ d.index=i; });

  // Overall FO information — old report style
  ws.mergeCells("A1:L1");
  ws.getCell("A1").value=`Production Control Report — ${workspace.line||""} — ${fo}`;
  ws.getCell("A1").font={bold:true,size:16,color:{argb:"FF0F172A"}};
  ws.getCell("A1").alignment={horizontal:"center"};

  const goodCum=prev+totals.ok;
  const productionCum=prev+totals.production;
  const remainingNow=Math.max(0,cutting-productionCum);

  const metaRows=[
    ["Production Control",workspace.line||"","FO",fo,"Style",f.style||"","Color",f.color||"","","","",""],
    ["Status",workspace.prodStatus||"READY","Order Qty",orderQty,"Cutting Qty",cutting,"Previous Cumulative",prev,"","","",""],
    ["SAM",f.sam||"","Eff (%)",f.eff===""||f.eff==null?0:Number(f.eff||0),"Manpower",Number(f.manpower||0),"Start Time",f.startTime||"08:00","Target/Hr",targetHr,"",""],
    ["Plan Days",plannedDays.length,"Good Cum (OK)",goodCum,"Production Cum (OK+NG)",productionCum,"ยอดงานที่ต้องผลิต",remainingNow,"Updated At",nowBkk(),"",""]
  ];

  let row=3;
  for(const r of metaRows){
    ws.getRow(row).values=r;
    row++;
  }

  // Summary plan table first
  row++;
  ws.mergeCells(row,1,row,12);
  ws.getCell(row,1).value=`แผนการผลิตทั้ง FO — ${plannedDays.length} วัน`;
  ws.getCell(row,1).font={bold:true,size:13,color:{argb:"FFFFFFFF"}};
  ws.getCell(row,1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1F4E78"}};
  row++;

  ws.getRow(row).values=[
    "Day","Production Date","Target/Day","OT (ชม.)","Plan Cum",
    "OK","Repair","NG","Actual (OK+NG)","Actual Cum","Remaining","Status"
  ];
  styleHeader(ws,row,12);
  row++;

  let planCum=0;
  let actualFoCumForPlan=prev;
  const today=isoTodayBkk();

  for(const d of plannedDays){
    const slots=actualByDate[d.date]||{};
    let ok=0,repair=0,ng=0;
    for(const v of Object.values(slots)){
      ok+=Number(v?.ok||0);
      repair+=Number(v?.repair||0);
      ng+=Number(v?.ng||0);
    }
    const prod=ok+ng;
    planCum+=d.dayTarget;
    actualFoCumForPlan+=prod;

    let dayStatus="PLANNED";
    if(d.date<today) dayStatus="LOCKED";
    else if(d.date===today) dayStatus="CURRENT";
    if(String(workspace.prodStatus||"").toUpperCase()==="CLOSED" && d.date<=today) dayStatus="LOCKED";

    ws.getRow(row).values=[
      d.index+1,d.date,d.dayTarget,d.ot,planCum,
      ok,repair,ng,prod,actualFoCumForPlan,
      Math.max(0,cutting-actualFoCumForPlan),dayStatus
    ];
    row++;
  }

  // ------------------------------------------------------------
  // OLD-STYLE DAILY TABLES, CREATED FOR EVERY PLANNED DAY UP FRONT
  // On future days actuals are 0; when the date arrives and web data
  // is recorded, this same sheet is rebuilt and actual values fill in.
  // ------------------------------------------------------------
  let foCumBeforeDay=prev;
  let targetCumBeforeDay=0;

  for(const d of plannedDays){
    const slots=actualByDate[d.date]||{};

    let dayOk=0,dayRepair=0,dayNg=0;
    for(const v of Object.values(slots)){
      dayOk+=Number(v?.ok||0);
      dayRepair+=Number(v?.repair||0);
      dayNg+=Number(v?.ng||0);
    }
    const dayProd=dayOk+dayNg;

    row++;
    ws.mergeCells(row,1,row,12);
    ws.getCell(row,1).value=`Production Date: ${d.date}  (Day ${d.index+1}/${plannedDays.length})`;
    ws.getCell(row,1).font={bold:true,size:13,color:{argb:"FFFFFFFF"}};
    ws.getCell(row,1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF17365D"}};
    row++;

    // Header block like old sheet
    const remainingBefore=Math.max(0,cutting-foCumBeforeDay);
    ws.getRow(row).values=[
      "Production Date",d.date,
      "Status",d.date<today?"LOCKED":(d.date===today?"CURRENT":"PLANNED"),
      "Target Today",d.dayTarget,
      "OT",`${d.ot} ชม.`,
      "FO Cum Before",foCumBeforeDay,
      "ยอดงานที่ต้องผลิตก่อนเริ่มวัน",remainingBefore
    ];
    row++;

    ws.getRow(row).values=[
      "Order Qty",orderQty,
      "Cutting Qty",cutting,
      "Previous Cumulative",prev,
      "SAM",f.sam||"",
      "Manpower",Number(f.manpower||0),
      "Target/Hr",targetHr
    ];
    row++;

    ws.getRow(row).values=[
      "เวลา","Target/Hr","Target Cum","OK","Repair","NG",
      "Actual/Hr","Actual Cum","Diff","Achievement %","Incentive","ปัญหา"
    ];
    styleHeader(ws,row,12);
    row++;

    // Match the web Production Board:
    // cumulative Target/Actual continue across the FO, not reset at each day.
    let targetCum=targetCumBeforeDay;
    let actualCum=foCumBeforeDay;
    let targetRemaining=d.dayTarget;

    for(const slot of SHIFT){
      const isBreak=slot.type==="BREAK";
      const beforeStart=isBeforeFirstDayStartServer(d.index,slot.label,workspace);
      const activeWork=!beforeStart && slot.type==="WORK";
      const activeOt=!beforeStart && slot.type==="OT" && slot.otNo<=d.ot;
      const active=activeWork || activeOt;

      let targetThisSlot=0;
      if(active && !isBreak && targetRemaining>0){
        const nominal=targetHr*slot.h;
        targetThisSlot=Math.min(nominal,targetRemaining);
        targetRemaining=Math.max(0,targetRemaining-targetThisSlot);
      }
      targetCum+=targetThisSlot;

      const a=slots[slot.label]||{};
      const ok=Number(a.ok||0);
      const repair=Number(a.repair||0);
      const ng=Number(a.ng||0);
      const actualHr=ok+ng;
      actualCum+=actualHr;

      const diff=actualCum-targetCum;
      const achievement=targetCum>0?actualCum/targetCum:0;

      let problem=String(a.problem||"");
      if(beforeStart) problem="ก่อนเริ่มผลิต";
      else if(isBreak) problem=slot.label==="12:00-13:00"?"พักเที่ยง":"พักก่อน OT";
      else if(slot.type==="OT" && !activeOt) problem="OT (ไม่ได้วางแผน)";
      else if(slot.type==="OT" && activeOt && !problem) problem="OT";

      ws.getRow(row).values=[
        slot.label,
        targetThisSlot,
        targetCum,
        ok,
        repair,
        ng,
        actualHr,
        actualCum,
        diff,
        achievement,
        0,
        problem
      ];
      ws.getCell(row,10).numFmt="0.0%";
      row++;
    }

    const remainingAfter=Math.max(0,cutting-(foCumBeforeDay+dayProd));

    row++;
    ws.mergeCells(row,1,row,12);
    ws.getCell(row,1).value="Daily Summary";
    ws.getCell(row,1).font={bold:true,color:{argb:"FFFFFFFF"}};
    ws.getCell(row,1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF24557F"}};
    row++;

    ws.getRow(row).values=[
      "OK Total",dayOk,
      "Repair Total",dayRepair,
      "NG Total",dayNg,
      "Production",dayProd,
      "FO Cum",foCumBeforeDay+dayProd,
      "ยอดงานที่ต้องผลิต",remainingAfter
    ];

    foCumBeforeDay+=dayProd;
    targetCumBeforeDay+=d.dayTarget;
  }

  // Final FO summary
  row+=2;
  ws.mergeCells(row,1,row,12);
  ws.getCell(row,1).value="FO Summary";
  ws.getCell(row,1).font={bold:true,size:13,color:{argb:"FFFFFFFF"}};
  ws.getCell(row,1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1F4E78"}};
  row++;

  ws.getRow(row).values=[
    "OK Total",totals.ok,
    "Repair Total",totals.repair,
    "NG Total",totals.ng,
    "Good Cum",goodCum,
    "Production Cum",productionCum,
    "ยอดงานที่ต้องผลิต",remainingNow
  ];

  setWidths(ws,[20,15,15,12,12,12,12,14,14,16,12,30]);

  return {
    sheet:name,
    days:plannedDays.length,
    oldStyle:true,
    precreated:true
  };
}

async function updateAuditTrail(ws,workspace){
  const existing=new Set();
  for(let r=2;r<=ws.rowCount;r++) existing.add(`${ws.getCell(r,1).value}|${ws.getCell(r,3).value}|${ws.getCell(r,4).value}`);
  for(const a of ensureArray(workspace?.auditLog)){
    const key=`${a.at||""}|${a.fo||""}|${a.action||""}`;
    if(existing.has(key)) continue;
    ws.addRow([a.at||nowBkk(),a.line||workspace.line||"",a.fo||workspace?.form?.fo||"",a.action||"",a.detail||""]); existing.add(key);
  }
}


async function ensureFoReportSheetNow(line,workspace){
  const fo=String(workspace?.form?.fo||"").trim();
  if(!fo) return {skipped:true,reason:"no fo"};

  const file=await ensureLineWorkbook(line);
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);

  const reportName=`RPT_${safeSheetName(fo,25)}`;
  const existed=!!wb.getWorksheet(reportName);

  // V1.22 critical fix:
  // ALWAYS rebuild RPT_<FO> from the current workspace.
  // Do not return early just because an older report sheet already exists.
  const report=await updateReportSheet(wb,workspace);
  await saveWorkbookAtomic(wb,file);

  if(ONEDRIVE_ROOT){
    await mirrorFile(file,mirrorLineWorkbook(line)).catch(()=>{});
  }

  return {
    rebuilt:existed,
    created:!existed,
    sheet:report?.sheet||reportName,
    days:report?.days||0
  };
}

async function syncLineWorkbookFromState(line){
  const stateStat=await fsp.stat(lineStateFile(line)).catch(()=>null);
  if(!stateStat) return {skipped:true};

  const workspace=await readJson(lineStateFile(line),null);
  if(!workspace) return {skipped:true};

  const fo=String(workspace?.form?.fo||"").trim();
  const reportName=fo?`RPT_${safeSheetName(fo,25)}`:"";
  const file=await ensureLineWorkbook(line);

  const applied=await readJson(lineStateAppliedFile(line),{mtimeMs:0});
  let reportMissing=false;
  let reportNeedsUpgrade=false;

  if(reportName){
    try{
      const probe=new ExcelJS.Workbook();
      await probe.xlsx.readFile(file);
      const rpt=probe.getWorksheet(reportName);
      reportMissing=!rpt;

      if(rpt){
        // Old report sheets do not contain this full-plan title.
        let markerFound=false;
        const maxRow=Math.min(rpt.rowCount,20);
        for(let r=1;r<=maxRow;r++){
          for(let c=1;c<=Math.min(rpt.columnCount||12,12);c++){
            const val=String(rpt.getCell(r,c).value||"");
            if(val.includes("แผนการผลิตทั้ง FO")){
              markerFound=true;
              break;
            }
          }
          if(markerFound) break;
        }
        reportNeedsUpgrade=!markerFound;
      }
    }catch{
      reportMissing=true;
      reportNeedsUpgrade=true;
    }
  }

  if(Number(applied?.mtimeMs||0)>=stateStat.mtimeMs && !reportMissing && !reportNeedsUpgrade){
    return {skipped:true};
  }

  const wb=new ExcelJS.Workbook();
  try{
    await wb.xlsx.readFile(file);
    ensureCurrentWorkbookHeaders(wb);
    await upsertFoIndex(wb.getWorksheet("FO_INDEX"),workspace);
    await updateDailySummary(wb.getWorksheet("DAILY_SUMMARY"),workspace);
    await updateAuditTrail(wb.getWorksheet("AUDIT_TRAIL"),workspace);
    const report=await updateReportSheet(wb,workspace);
    await saveWorkbookAtomic(wb,file);

    await atomicWriteJson(lineStateAppliedFile(line),{
      mtimeMs:stateStat.mtimeMs,
      at:nowBkk(),
      report:report?.sheet||reportName,
      reportVersion:"V1.24-JOURNAL-SYNC-CUMULATIVE"
    });

    lastExcelError.delete(line);
    lastExcelSavedAt.set(line,nowBkk());
    await updateSummaryFromWorkspace(workspace).catch(()=>{});

    if(ONEDRIVE_ROOT){
      await mirrorFile(file,mirrorLineWorkbook(line)).catch(()=>{});
    }

    return {
      saved:true,
      report,
      forced:reportMissing||reportNeedsUpgrade,
      upgraded:reportNeedsUpgrade
    };
  }catch(e){
    lastExcelError.set(line,e.message);
    return {queued:true,error:e.message};
  }
}

async function updateSummaryFromWorkspace(workspace){
  const line=workspace?.line; if(!validLine(line)) return;
  const file=summaryXlsx(); const wb=new ExcelJS.Workbook(); await wb.xlsx.readFile(file);
  const ws=wb.getWorksheet("LINE_STATUS"); const f=workspace.form||{}, t=workspaceTotals(workspace), prev=Number(f.previousCum||0), cutting=Number(f.qty||0);
  let row=-1; for(let r=2;r<=ws.rowCount;r++) if(String(ws.getCell(r,1).value||"")===line){row=r;break;} if(row<0) row=ws.rowCount+1;
  const vals=[line,f.fo||"",f.style||"",f.color||"",workspace.prodStatus||"READY",t.ok,t.repair,t.ng,prev+t.ok,prev+t.production,Math.max(0,cutting-(prev+t.production)),nowBkk()];
  vals.forEach((v,i)=>ws.getCell(row,i+1).value=v); await saveWorkbookAtomic(wb,file); await mirrorFile(file,mirrorCentral(file)).catch(()=>{});
}

async function mirrorFile(src,dst){
  if(!dst) return {disabled:true};
  try{
    await ensureDir(path.dirname(dst));
    await fsp.copyFile(src,dst);
    return {synced:true};
  }catch(e){ throw e; }
}
async function mirrorLine(line){
  if(!ONEDRIVE_ROOT) return {disabled:true};
  const src=lineWorkbook(line), dst=mirrorLineWorkbook(line);
  if(!(await exists(src))) return {missing:true};
  try{
    const ss=await fsp.stat(src); const ds=await fsp.stat(dst).catch(()=>null);
    if(ds && ds.mtimeMs>=ss.mtimeMs){ lastMirrorError.delete(line); return {synced:true,unchanged:true}; }
    await mirrorFile(src,dst); lastMirrorError.delete(line); lastMirrorAt.set(line,nowBkk()); return {synced:true};
  }catch(e){ lastMirrorError.set(line,e.message); return {queued:true,error:e.message}; }
}

async function processLine(line){
  return withLineLock(line,async()=>{
    // V1.24: journal is the actual-production source of truth.
    // Heal state.json before every Excel/report rebuild.
    await reconcileLineStateFromJournal(line).catch(()=>{});

    const ev=await applyPendingEvents(line);
    const st=await syncLineWorkbookFromState(line);

    // Dashboard/Power BI friendly per-Line standalone files.
    const dl=await syncStandaloneDataLog(line).catch(e=>({queued:true,error:e.message}));
    const dailySummary=await syncStandaloneFoDailySummary(line).catch(e=>({queued:true,error:e.message}));

    // Mirror main workbook and operational support data to OneDrive.
    const mi=await mirrorLine(line);
    const support=await mirrorLineSupportFiles(line).catch(e=>({queued:true,error:e.message}));

    // Combined 18-Line dashboard file. Derived output; never blocks production.
    const allLines=await syncAllLinesDataLog().catch(e=>({queued:true,error:e.message}));
    const allDailySummary=await syncAllLinesFoDailySummary().catch(e=>({queued:true,error:e.message}));

    return {events:ev,state:st,dataLog:dl,dailySummary,allLines,allDailySummary,mirror:mi,support};
  });
}


function totalsFromEvents(events,fo){
  const want=normalizeFo(fo);
  let ok=0,repair=0,ng=0;
  for(const e of events||[]){
    if(want && normalizeFo(e.fo)!==want) continue;
    const qty=Math.max(0,Number(e.qty||1));
    const type=String(e.type||"").toUpperCase();
    if(type==="OK") ok+=qty;
    else if(type==="REPAIR") repair+=qty;
    else if(type==="NG") ng+=qty;
  }
  return {ok,repair,ng,production:ok+ng};
}

function sameTotals(a,b){
  return ["ok","repair","ng","production"].every(k=>Number(a?.[k]||0)===Number(b?.[k]||0));
}

async function workbookDataLogTotals(file,fo){
  if(!(await exists(file))) return {exists:false,ok:0,repair:0,ng:0,production:0};
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws=wb.getWorksheet("DATA_LOG");
  if(!ws) return {exists:true,sheet:false,ok:0,repair:0,ng:0,production:0};

  let ok=0,repair=0,ng=0;
  const want=normalizeFo(fo);
  for(let r=2;r<=ws.rowCount;r++){
    if(want && normalizeFo(ws.getCell(r,5).value)!==want) continue;
    const type=String(ws.getCell(r,9).value||"").toUpperCase();
    const qty=Math.max(0,Number(ws.getCell(r,10).value||1));
    if(type==="OK") ok+=qty;
    else if(type==="REPAIR") repair+=qty;
    else if(type==="NG") ng+=qty;
  }
  return {exists:true,sheet:true,ok,repair,ng,production:ok+ng};
}

async function workbookReportTotals(file,fo){
  if(!(await exists(file))) return {exists:false};
  const wb=new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const name=`RPT_${safeSheetName(fo,25)}`;
  const ws=wb.getWorksheet(name);
  if(!ws) return {exists:true,sheet:false,name};

  // Read the last "FO Summary" block generated by updateReportSheet().
  for(let r=ws.rowCount;r>=1;r--){
    if(String(ws.getCell(r,1).value||"").trim()==="FO Summary"){
      const vr=r+1;
      const ok=Number(ws.getCell(vr,2).value||0);
      const repair=Number(ws.getCell(vr,4).value||0);
      const ng=Number(ws.getCell(vr,6).value||0);
      return {exists:true,sheet:true,name,ok,repair,ng,production:ok+ng,row:vr};
    }
  }
  return {exists:true,sheet:true,name,summary:false};
}

async function statInfo(file){
  try{
    const s=await fsp.stat(file);
    return {exists:true,size:s.size,mtimeMs:s.mtimeMs,mtime:new Date(s.mtimeMs).toISOString()};
  }catch{
    return {exists:false,size:0,mtimeMs:0,mtime:""};
  }
}

async function verifyLineConsistency(line){
  line=lineCode(line);
  const workspace=await reconcileLineStateFromJournal(line);
  const fo=String(workspace?.form?.fo||"").trim();
  const events=await readJsonl(lineJournalFile(line));

  const journal=totalsFromEvents(events,fo);
  const state=workspace ? workspaceTotals(workspace) : {ok:0,repair:0,ng:0,production:0};

  const localFile=lineWorkbook(line);
  const localDataLog=await workbookDataLogTotals(localFile,fo).catch(e=>({error:e.message}));
  const localReport=await workbookReportTotals(localFile,fo).catch(e=>({error:e.message}));

  let mirrorDataLog={disabled:true};
  let mirrorReport={disabled:true};
  const mirrorFilePath=mirrorLineWorkbook(line);
  if(ONEDRIVE_ROOT){
    mirrorDataLog=await workbookDataLogTotals(mirrorFilePath,fo).catch(e=>({error:e.message}));
    mirrorReport=await workbookReportTotals(mirrorFilePath,fo).catch(e=>({error:e.message}));
  }

  const pending=(await pendingEvents(line)).length;
  const localStat=await statInfo(localFile);
  const mirrorStat=ONEDRIVE_ROOT ? await statInfo(mirrorFilePath) : {exists:false};

  const checks={
    stateMatchesJournal:sameTotals(state,journal),
    dataLogMatchesJournal:sameTotals(localDataLog,journal),
    reportMatchesJournal:sameTotals(localReport,journal),
    mirrorDataLogMatchesJournal:ONEDRIVE_ROOT ? sameTotals(mirrorDataLog,journal) : true,
    mirrorReportMatchesJournal:ONEDRIVE_ROOT ? sameTotals(mirrorReport,journal) : true,
    noPendingEvents:pending===0
  };

  return {
    line,fo,
    journal,state,
    local:{dataLog:localDataLog,report:localReport,file:localStat},
    oneDrive:{enabled:!!ONEDRIVE_ROOT,dataLog:mirrorDataLog,report:mirrorReport,file:mirrorStat},
    pendingEvents:pending,
    checks,
    allOk:Object.values(checks).every(Boolean),
    checkedAt:nowBkk()
  };
}


async function syncAssignmentsWorkbook(){
  const data=await readJson(assignmentJson(),{});
  const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet("FO_ASSIGNMENT");
  ws.addRow(["FO","Style","Color","Assigned Line","Status","Assigned At","Closed At","Updated At"]); styleHeader(ws,1,8); setWidths(ws,[18,28,18,16,14,22,22,22]);
  Object.values(data).sort((a,b)=>String(a.fo).localeCompare(String(b.fo))).forEach(a=>ws.addRow([a.fo||"",a.style||"",a.color||"",a.assignedLine||"",a.status||"",a.assignedAt||"",a.closedAt||"",a.updatedAt||""]));
  await saveWorkbookAtomic(wb,assignmentXlsx()); await mirrorFile(assignmentXlsx(),mirrorCentral(assignmentXlsx())).catch(()=>{});
}
async function syncPlanWorkbook(){
  const rows=ensureArray(await readJson(planJson(),[])); const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet("PLAN_MASTER");
  ws.addRow(["ลำดับ","ไลน์","FO","Style Name","สี","SAM เย็บ","Eff","Order Qty"]); styleHeader(ws,1,8); setWidths(ws,[10,14,18,30,18,14,10,14]);
  rows.forEach((r,i)=>ws.addRow([i+1,r.line||"",r.fo||"",r.style||"",r.color||"",r.sam||"",r.eff??"",r.orderQty??""]));
  await saveWorkbookAtomic(wb,planXlsx()); await mirrorFile(planXlsx(),mirrorCentral(planXlsx())).catch(()=>{});
}

async function reserveFo(body){
  const line=lineCode(body.line), fo=normalizeFo(body.fo); if(!fo) throw Object.assign(new Error("FO required"),{status:400});
  return withCentralLock(async()=>{
    const all=await readJson(assignmentJson(),{}); const ex=all[fo];
    if(ex && ["READY","RUNNING","PAUSED"].includes(ex.status) && ex.assignedLine!==line){
      const err=new Error(`FO ${fo} กำลังใช้งานที่ ${ex.assignedLine} (${ex.status})`); err.status=409; err.assignment=ex; throw err;
    }
    const now=nowBkk(); all[fo]={fo,style:body.style||ex?.style||"",color:body.color||ex?.color||"",assignedLine:line,status:body.status||ex?.status||"READY",assignedAt:ex?.assignedAt||now,closedAt:null,updatedAt:now};
    await atomicWriteJson(assignmentJson(),all); await syncAssignmentsWorkbook().catch(()=>{}); return all[fo];
  });
}
async function setFoStatus(body){
  const line=lineCode(body.line), fo=normalizeFo(body.fo);
  if(!fo) throw Object.assign(new Error("FO required"),{status:400});

  const requestedStatus=String(body.status||"READY").toUpperCase();

  // CLOSED is privileged:
  // 1) AUTO only when GOOD cumulative >= Cutting Qty
  // 2) SUPERVISOR requires Line Head PIN + mandatory reason
  if(requestedStatus==="CLOSED"){
    const mode=String(body.closeMode||"").toUpperCase();
    if(mode==="SUPERVISOR"){
      if(String(body.pin||"")!==LINE_HEAD_PIN){
        throw Object.assign(new Error("รหัสหัวหน้า Line ไม่ถูกต้อง"),{status:403});
      }
      if(!String(body.reason||"").trim()){
        throw Object.assign(new Error("ต้องระบุเหตุผลในการปิด FO"),{status:400});
      }
    }else if(mode==="AUTO"){
      const good=Number(body.goodCumulative||0);
      const cutting=Number(body.cuttingQty||0);
      if(!(cutting>0 && good>=cutting)){
        throw Object.assign(new Error("AUTO CLOSE ไม่ผ่าน: ยอดงานดียังไม่ครบ Cutting Qty"),{status:403});
      }
    }else{
      throw Object.assign(new Error("การปิด FO ต้องเป็น AUTO หรือ SUPERVISOR ที่ได้รับอนุญาต"),{status:403});
    }
  }

  return withCentralLock(async()=>{
    const all=await readJson(assignmentJson(),{});
    let ex=all[fo];

    if(!ex){
      const now=nowBkk();
      ex={
        fo,style:body.style||"",color:body.color||"",
        assignedLine:line,status:requestedStatus,
        assignedAt:now,closedAt:requestedStatus==="CLOSED"?now:null,
        updatedAt:now,
        closeReason:requestedStatus==="CLOSED"?String(body.reason||""):""
      };
      all[fo]=ex;
      await atomicWriteJson(assignmentJson(),all);
      await syncAssignmentsWorkbook().catch(()=>{});
      return ex;
    }

    if(ex.assignedLine!==line && ["READY","RUNNING","PAUSED"].includes(ex.status)){
      throw Object.assign(new Error(`FO ${fo} belongs to ${ex.assignedLine}`),{status:409});
    }

    ex.status=requestedStatus||ex.status;
    ex.updatedAt=nowBkk();
    if(ex.status==="CLOSED"){
      ex.closedAt=nowBkk();
      ex.closeReason=String(body.reason||"");
      ex.closeMode=String(body.closeMode||"").toUpperCase();
    }

    all[fo]=ex;
    await atomicWriteJson(assignmentJson(),all);
    await syncAssignmentsWorkbook().catch(()=>{});
    return ex;
  });
}

async function storageStatus(line){
  const pending=(await pendingEvents(line)).length;
  const stateStat=await fsp.stat(lineStateFile(line)).catch(()=>null); const applied=await readJson(lineStateAppliedFile(line),{mtimeMs:0});
  const statePending=!!(stateStat && Number(applied?.mtimeMs||0)<stateStat.mtimeMs);
  let mirrorPending=false;
  if(ONEDRIVE_ROOT){
    const ss=await fsp.stat(lineWorkbook(line)).catch(()=>null), ds=await fsp.stat(mirrorLineWorkbook(line)).catch(()=>null);
    mirrorPending=!!(ss && (!ds || ds.mtimeMs<ss.mtimeMs));
  }
  return {
    line,
    localQueue:{status:"SAVED",pendingEvents:pending,statePending},
    excelMaster:{status:lastExcelError.has(line)?"WAITING":"SAVED",error:lastExcelError.get(line)||null,lastSavedAt:lastExcelSavedAt.get(line)||null,path:lineWorkbook(line)},
    dataLogExport:{
      status:(await exists(lineDataLogWorkbook(line)))?"READY":"WAITING",
      path:lineDataLogWorkbook(line),
      mirrorPath:mirrorLineDataLogWorkbook(line)
    },
    allLinesDataLog:{
      status:(await exists(allLinesDataLogWorkbook()))?"READY":"WAITING",
      path:allLinesDataLogWorkbook(),
      mirrorPath:mirrorAllLinesDataLogWorkbook()
    },
    oneDrive:{status:!ONEDRIVE_ROOT?"DISABLED":lastMirrorError.has(line)?"WAITING":mirrorPending?"SYNCING":"SYNCED",error:lastMirrorError.get(line)||null,lastSyncedAt:lastMirrorAt.get(line)||null,path:mirrorLineWorkbook(line)}
  };
}

// ---------------- API ----------------
app.get("/api/health",(req,res)=>res.json({ok:true,version:"SERVER-V1.33-IMPORT-ALL-AS-TEXT",time:nowBkk(),dataRoot:DATA_ROOT,oneDriveRoot:ONEDRIVE_ROOT,lines:LINE_CODES}));

app.get("/api/plan",async(req,res)=>res.json({ok:true,rows:ensureArray(await readJson(planJson(),[]))}));
app.post("/api/plan",async(req,res,next)=>{
  try{
    if(String(req.body?.pin||"")!==PLANNER_PIN){
      throw Object.assign(new Error("Invalid planner PIN"),{status:403});
    }

    // V1.33: load and save every row as text.
    // Do not block duplicate FO, blank FO, invalid SAM, or non-numeric quantity.
    const input=ensureArray(req.body?.rows);
    const rows=input
      .map(r=>({
        line:String(r?.line??"").trim(),
        fo:String(r?.fo??"").trim(),
        style:String(r?.style??"").trim(),
        color:String(r?.color??"").trim(),
        sam:String(r?.sam??"").trim(),
        eff:String(r?.eff??"").trim(),
        orderQty:String(r?.orderQty??"").trim()
      }))
      .filter(r=>[r.line,r.fo,r.style,r.color,r.sam,r.eff,r.orderQty].some(v=>String(v).trim()!==""));

    await atomicWriteJson(planJson(),rows);
    await syncPlanWorkbook().catch(()=>{});

    res.json({ok:true,count:rows.length,mode:"IMPORT_ALL_AS_TEXT"});
  }catch(e){next(e);}
});

app.get("/api/assignments",async(req,res)=>res.json({ok:true,assignments:await readJson(assignmentJson(),{})}));
app.post("/api/fo/reserve",async(req,res,next)=>{try{res.json({ok:true,assignment:await reserveFo(req.body||{})});}catch(e){next(e);}});
app.post("/api/fo/status",async(req,res,next)=>{try{res.json({ok:true,assignment:await setFoStatus(req.body||{})});}catch(e){next(e);}});

app.get("/api/line/state",async(req,res,next)=>{
  try{
    const line=lineCode(req.query.line);
    await ensureLineWorkbook(line);

    const workspace=await reconcileLineStateFromJournal(line);
    const logs=await readJsonl(lineJournalFile(line));

    if(workspace) workspace.productionLog=logs;
    res.json({ok:true,line,workspace,logs});
  }catch(e){next(e);}
});
app.post("/api/line/state",async(req,res,next)=>{
  try{
    const line=lineCode(req.body?.line);
    const workspace=deepClone(req.body?.workspace||null);
    if(!workspace) throw Object.assign(new Error("workspace required"),{status:400});

    workspace.line=line;
    workspace.productionLog=[]; // DATA_LOG is journal/Excel, not duplicated in state.json

    // V1.24: never let browser state overwrite actual production totals.
    // Rebuild OK/Repair/NG from the accepted event journal first.
    await rebuildProdRowsFromJournal(line,workspace);
    await atomicWriteJson(lineStateFile(line),workspace);

    // Critical V1.10 rule:
    // As soon as a NEW FO/workspace is saved, create RPT_<FO> immediately.
    const report=await ensureFoReportSheetNow(line,workspace).catch(e=>({
      queued:true,error:e.message
    }));

    const result=await processLine(line);

    // Every FO save/update refreshes the Line's standalone history file immediately.
    const foHistory=await ensureFoHistoryDailySummaryNow(line).catch(e=>({
      queued:true,error:e.message
    }));

    res.json({ok:true,result,report,foHistory});
  }catch(e){next(e);}
});

app.post("/api/event",async(req,res,next)=>{
  try{
    const e=deepClone(req.body?.event||{});
    const line=lineCode(e.line||req.body?.line);
    e.line=line;

    if(!e.eventId) throw Object.assign(new Error("eventId required"),{status:400});
    if(!["OK","REPAIR","NG"].includes(String(e.type||"").toUpperCase())){
      throw Object.assign(new Error("Invalid event type"),{status:400});
    }

    e.type=String(e.type).toUpperCase();
    e.qty=Number(e.qty||1);
    e.dateTime=e.dateTime||nowBkk();

    const isNew=await journalEvent(line,e);

    // Critical V1.24 fix:
    // Reconcile state immediately after the journal accepts the event.
    // This removes the browser-state/event race that caused Excel RPT to lag behind the web UI.
    await reconcileLineStateFromJournal(line);

    const result=await processLine(line);
    const pending=(await pendingEvents(line)).length;

    // Compact post-event consistency snapshot.
    const currentState=await readJson(lineStateFile(line),null);
    const currentEvents=await readJsonl(lineJournalFile(line));
    const currentFo=String(currentState?.form?.fo||e.fo||"").trim();
    const journalTotals=totalsFromEvents(currentEvents,currentFo);
    const stateTotals=currentState?workspaceTotals(currentState):{ok:0,repair:0,ng:0,production:0};

    res.json({
      ok:true,
      accepted:true,
      newEvent:isNew,
      excelSaved:pending===0 && !result?.state?.queued,
      pending,
      consistency:{
        fo:currentFo,
        journal:journalTotals,
        state:stateTotals,
        stateMatchesJournal:sameTotals(journalTotals,stateTotals)
      },
      result
    });
  }catch(e){next(e);}
});


app.get("/api/line/verify",async(req,res,next)=>{
  try{
    const line=lineCode(req.query.line);
    res.json({ok:true,...await verifyLineConsistency(line)});
  }catch(e){next(e);}
});

app.post("/api/line/rebuild",async(req,res,next)=>{
  try{
    if(String(req.body?.pin||"")!==ADMIN_PIN){
      throw Object.assign(new Error("Invalid admin PIN"),{status:403});
    }
    const line=lineCode(req.body?.line);
    await reconcileLineStateFromJournal(line);
    const result=await processLine(line);
    const verify=await verifyLineConsistency(line);
    res.json({ok:true,line,result,verify});
  }catch(e){next(e);}
});

app.get("/api/storage/status",async(req,res,next)=>{try{const line=lineCode(req.query.line);res.json({ok:true,...await storageStatus(line)});}catch(e){next(e);}});
app.get("/api/storage/pending",async(req,res,next)=>{
  try{
    const line=lineCode(req.query.line);
    const events=await readJsonl(lineJournalFile(line));
    const applied=await loadAppliedSet(line);
    const st=await storageStatus(line);
    const rows=events.slice(-20).reverse().map(e=>({
      eventId:e.eventId||"", line, fo:e.fo||"", type:e.type||"", qty:Number(e.qty||1),
      queueStatus:"SAVED", excelStatus:applied.has(String(e.eventId))?"SAVED":"WAITING",
      oneDriveStatus:st.oneDrive.status==="SYNCED"?"SYNCED":"WAITING"
    }));
    res.json({ok:true,rows});
  }catch(e){next(e);}
});

app.post("/api/admin/bootstrap",async(req,res,next)=>{
  try{ if(String(req.body?.pin||"")!==ADMIN_PIN) throw Object.assign(new Error("Invalid admin PIN"),{status:403}); await bootstrap(); res.json({ok:true,lines:LINE_CODES}); }catch(e){next(e);}
});

app.get("/line/:line",(req,res)=>res.sendFile(path.join(ROOT,"public","index.html")));
app.get("*",(req,res,next)=>{ if(req.path.startsWith("/api/")) return next(); res.sendFile(path.join(ROOT,"public","index.html")); });

app.use((err,req,res,next)=>{
  console.error(nowBkk(),err);
  res.status(Number(err.status||500)).json({ok:false,message:err.message||"Server error",assignment:err.assignment||null});
});

async function backgroundWorker(){
  for(const line of LINE_CODES){ try{ await processLine(line); }catch(e){ console.error("worker",line,e.message); } }
  try{ await syncAllLinesDataLog(); }catch(e){ console.error("worker all-lines",e.message); }
  try{ await syncAllLinesFoDailySummary(); }catch(e){ console.error("worker all-daily-summary",e.message); }
  try{ await syncAssignmentsWorkbook(); }catch{}
  try{ await syncPlanWorkbook(); }catch{}
}

(async()=>{
  try{
    await bootstrap();
    app.listen(PORT,"0.0.0.0",()=>{
      console.log(`[${nowBkk()}] Production Control Server V1.33 + Import All As Text + ALL LINES + OneDrive`);
      console.log(`Local: http://localhost:${PORT}`);
      console.log(`Data: ${DATA_ROOT}`);
      console.log(`OneDrive: ${ONEDRIVE_ROOT || "DISABLED"}`);
      if(cfg.OPEN_BROWSER_ON_START && process.platform==="win32"){
        setTimeout(()=>exec(`start "" "http://localhost:${PORT}"`),1200);
      }
    });
    setInterval(backgroundWorker,RETRY_MS);
  }catch(e){ console.error("STARTUP FAILED",e); process.exitCode=1; }
})();
