const express = require('express');
const cors    = require('cors');
const XLSX    = require('xlsx');
const path    = require('path');
const multer  = require('multer');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, 'SAT Progress.xlsx');
const TMP_EXCEL  = '/tmp/hq_latest.xlsx';
// TOTAL, PROJ_START/END คำนวณจาก Excel ใน parseData()

let cache = null, cacheTime = 0;

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') {
    // Buddhist calendar serial (> 200000) ต้องหัก offset 198347
    const serial = v > 200000 ? v - 198347 : v;
    return new Date((serial - 25569) * 86400000);
  }
  if (typeof v === 'string') {
    const s = v.trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      let y = parseInt(m[3]); if (y < 100) y += 2000;
      return new Date(y, parseInt(m[2])-1, parseInt(m[1]));
    }
    const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return new Date(parseInt(iso[1]), parseInt(iso[2])-1, parseInt(iso[3]));
    return null;
  }
  return null;
}

function isoDate(v) {
  const d = toDate(v);
  return d ? d.toISOString().slice(0,10) : null;
}

function fmtLbl(d) {
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
}

const thMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
function fmtLbl2(isoStr) {
  if (!isoStr) return '–';
  const d = new Date(isoStr+'T00:00:00');
  return `${d.getDate()} ${thMonths[d.getMonth()]} ${(d.getFullYear()+543)%100}`;
}

function parseData() {
  const excelPath = EXCEL_PATH; // ใช้ Excel จาก repo เสมอ
  console.log('Reading Excel:', excelPath);
  const wb = XLSX.readFile(excelPath);

  const hqRows = XLSX.utils.sheet_to_json(wb.Sheets['HQ'], { header:1, defval:null });

  // หา proj_start/end จาก col T(19)=เริ่ม, col V(21)=สิ้นสุด (Helper)
  // fallback: col H(7) ถ้า col T ว่างทั้งหมด
  let PROJ_START = null, PROJ_END = null;
  for (let i=2; i<hqRows.length; i++) {
    const r=hqRows[i]; if(!r) continue;
    const dT = toDate(r[19]); // col T = วันที่เริ่ม Helper
    const dV = toDate(r[21]); // col V = วันที่สิ้นสุด Helper
    const dH = toDate(r[7]);  // col H = Migration Plan เริ่ม (fallback)
    const dI = toDate(r[8]);  // col I = Migration Plan สิ้นสุด (fallback)
    // ใช้ T ถ้ามี ไม่งั้น H
    const dStart = (dT && dT.getFullYear() > 2000) ? dT : (dH && dH.getFullYear() > 2000 ? dH : null);
    const dEnd   = (dV && dV.getFullYear() > 2000) ? dV : (dI && dI.getFullYear() > 2000 ? dI : null);
    if (dStart) {
      if (!PROJ_START || dStart < PROJ_START) PROJ_START = dStart;
      if (!PROJ_END   || dStart > PROJ_END)   PROJ_END   = dStart;
    }
    if (dEnd) {
      if (!PROJ_END   || dEnd > PROJ_END)     PROJ_END   = dEnd;
    }
  }
  if (!PROJ_START) PROJ_START = new Date('2026-02-02');
  if (!PROJ_END)   PROJ_END   = new Date('2026-04-30');
  PROJ_START.setHours(0,0,0,0); PROJ_END.setHours(0,0,0,0);

  // คำนวณ TOTAL จาก col G(6) จำนวนใหม่ ทุก row ที่มี Category
  let TOTAL = 0;
  for (let i=2; i<hqRows.length; i++) {
    const r=hqRows[i]; if(!r) continue;
    const qty = typeof r[6]==='number' ? r[6] : 0;
    const cat = r[18] ? String(r[18]).trim() : '';
    if (qty>0 && ['Switch','Infra'].includes(cat)) TOTAL += qty;
  }
  // เพิ่ม AP total จาก HQ-WL (คำนวณทีหลัง จะ += apTotal)
  // TOTAL จะถูก update หลัง parse HQ-WL
  console.log('PROJ_START:', PROJ_START.toISOString().slice(0,10), 'PROJ_END:', PROJ_END.toISOString().slice(0,10));

  let installed=0, inProgress=0, notStarted=0, hold=0, overdue=0;
  let _curSite = null;
  let onTimeQty=0, earlyQty=0, lateQty=0;
  let instSW=0, instAP=0, instInf=0;
  let lastInstallDate = null;
  let curSite = null;

  const siteMap={}, typeMap={}, dayActMap={}, dayPlanMap={}, dayActBySite={}, dayPlanBySite={}, daySwActMap={}, dayApActMap={};
  const devices=[];
  const today = new Date(); today.setHours(0,0,0,0);

  for (let i = 2; i < hqRows.length; i++) {
    const r = hqRows[i];
    if (!r || !r.length) continue;
    if (r[0]) curSite = String(r[0]).trim();
    if (curSite && curSite.startsWith('%')) { curSite = null; continue; } // skip summary rows

    const device  = r[3] ? String(r[3]).trim() : null;
    const qty     = typeof r[6]==='number' ? r[6] : 0;
    const status  = r[11] ? String(r[11]).trim() : '';
    // ใช้ helper dates (col T=19 plan, col U=20 actual) แทน col J/H
    const instDt  = toDate(r[20]) || toDate(r[19]);  // col U ก่อน fallback col T
    const schedDt = toDate(r[19]);                    // col T = plan date
    const instStr = isoDate(instDt);
    const schedStr= isoDate(schedDt);

    // Category จาก column S (index 18) — Switch/AP/Infra
    let cat = r[18] ? String(r[18]).trim() : 'Infra';
    if (!['Switch','AP','Infra'].includes(cat)) cat = 'Infra';

    if (!device || !curSite || qty <= 0) continue;

    const site = curSite; // ไม่ truncate
    if (!siteMap[site]) siteMap[site] = {total:0,done:0,inp:0,start:null,end:null};
    // track start/end date per site จาก col T + V
    const _hStr = r[19] ? toDate(r[19])?.toISOString().slice(0,10) : null;
    const _eStr = r[21] ? toDate(r[21])?.toISOString().slice(0,10) : null;
    if (_hStr) {
      if (!siteMap[site].start || _hStr < siteMap[site].start) siteMap[site].start = _hStr;
      if (!siteMap[site].end   || _hStr > siteMap[site].end)   siteMap[site].end   = _hStr;
    }
    if (_eStr) {
      if (!siteMap[site].end   || _eStr > siteMap[site].end)   siteMap[site].end   = _eStr;
    }
    siteMap[site].total += qty;

    const dev = device.length>60 ? device.slice(0,60)+'…' : device;
    if (!typeMap[dev]) typeMap[dev] = {plan:0,done:0,cat};
    typeMap[dev].plan += qty;

    // dayPlanMap ใช้ col T(19) วันที่เริ่ม Helper — clamp ไม่ต่ำกว่า PROJ_START
    const helperDt = toDate(r[19]);
    let helperStr = helperDt ? helperDt.toISOString().slice(0,10) : schedStr;
    if (helperStr && helperStr < PROJ_START.toISOString().slice(0,10)) helperStr = PROJ_START.toISOString().slice(0,10);
    if (helperStr && cat !== 'AP') {
      dayPlanMap[helperStr] = (dayPlanMap[helperStr]||0) + qty;
      // per-site plan tracking
      if (site) {
        if (!dayPlanBySite[site]) dayPlanBySite[site]={total:0,byDate:{}};
        dayPlanBySite[site].total += qty;
        dayPlanBySite[site].byDate[helperStr]=(dayPlanBySite[site].byDate[helperStr]||0)+qty;
      }
    }

    // นับจาก Migration column (index 15) — เฉพาะ SW และ Infra
    const migration = typeof r[15]==='number' ? Math.round(r[15]) : 0;
    if (migration > 0 && cat !== 'AP') {
      installed += migration;
      siteMap[site].done += migration;
      typeMap[dev].done += migration;
      if (cat==='Switch') instSW+=migration;
      else if (cat==='Infra') instInf+=migration;
      // ใช้ col U(20) วันที่ติดตั้ง Helper — fallback col T(19) ถ้าไม่มี
      const instDt2 = toDate(r[20]) || toDate(r[19]);
      let instStr2 = instDt2 ? instDt2.toISOString().slice(0,10) : null;
      if (instStr2) {
        if (instStr2 < PROJ_START.toISOString().slice(0,10)) instStr2 = PROJ_START.toISOString().slice(0,10);
        if (!lastInstallDate||instStr2>lastInstallDate) lastInstallDate=instStr2;
        dayActMap[instStr2]=(dayActMap[instStr2]||0)+migration;
        // per-category tracking
        if (cat==='Switch') daySwActMap[instStr2]=(daySwActMap[instStr2]||0)+migration;
        else daySwActMap[instStr2]=(daySwActMap[instStr2]||0); // Infra นับรวม SW
        // per-site actual tracking
        if (site) {
          if (!dayActBySite[site]) dayActBySite[site]={};
          dayActBySite[site][instStr2]=(dayActBySite[site][instStr2]||0)+migration;
        }
      }
      // on-time check ใช้ ISO string เปรียบเทียบตรงๆ
      if (instStr2 && helperStr) {
        if (instStr2 <= helperStr){onTimeQty+=migration;if(instStr2<helperStr)earlyQty+=migration;}
        else lateQty+=migration;
      }
    } else if (status.includes('Progress')) {
      inProgress+=qty; siteMap[site].inp+=qty;
      if (schedDt&&schedDt<today) overdue+=qty;
    } else if (status==='Hold') {
      hold+=qty;
    } else {
      notStarted+=qty;
    }

    devices.push({site:curSite, device, qty, status, instDt:instStr, schedDt:schedStr, cat});
  }

  const _swInfActSum = Object.values(dayActMap).reduce((a,v)=>a+v,0);
  console.log('[SW/INF ACT] dayActMap sum before AP:', _swInfActSum, 'instSW:', instSW, 'instInf:', instInf);
  const onTimePct = installed>0 ? Math.round(onTimeQty/installed*1000)/10 : 0;

  // Weekly
  const wkSheet = wb.Sheets['HQ-กราฟรายสัปดาห์'];
  const wkRows  = wkSheet ? XLSX.utils.sheet_to_json(wkSheet,{header:1,defval:null}) : [];
  let wkLabels=[], planPct=[], actPct=[];
  for (const row of wkRows) {
    if (!row||!row[1]) continue;
    if (!wkLabels.length && String(row[1]).startsWith('W.')) {
      for (let i=1;i<row.length;i++) {
        if (row[i]&&String(row[i]).startsWith('W.')) wkLabels.push(String(row[i]).split('\n')[0]);
      }
    }
    if (String(row[0]).includes('% ความก้าวหน้าโดยรวม'))
      planPct = wkLabels.map((_,i)=>Math.round((row[i+1]||0)*100));
    if (String(row[0]).includes('จำนวนที่เสร็จสิ้นสะสม'))
      actPct  = wkLabels.map((_,i)=>Math.round((row[i+1]||0)*100));
  }
  if (!wkLabels.length) { wkLabels=Array.from({length:14},(_,i)=>`W.${i+1}`); planPct=wkLabels.map(()=>0); actPct=planPct.slice(); }

  const bdPlan = planPct.map(p=>Math.round((1-p/100)*TOTAL));
  let bdCum=TOTAL; const bdAct=actPct.map((v,i)=>{ if(v>0) bdCum=Math.round((1-v/100)*TOTAL); return i<=9?bdCum:null; });

  // Insight
  // DEBUG HQ plan
  const _hqPlanSum = Object.values(dayPlanMap).reduce((a,v)=>a+v,0);

  // sw_inf_sites — แยก SW/Infra per site จาก HQ sheet
  const swInfSiteMap = {};
  for (let i=2; i<hqRows.length; i++) {
    const r=hqRows[i]; if(!r||!r.length) continue;
    if (r[0]) _curSite = String(r[0]).trim();
    if (_curSite && _curSite.startsWith('%')) { _curSite = null; continue; }
    const qty = typeof r[6]==='number' ? r[6] : 0;
    const mig = typeof r[15]==='number' ? Math.round(r[15]) : 0;
    const cat = r[18] ? String(r[18]).trim() : '';
    if (qty<=0 || !_curSite || !['Switch','Infra'].includes(cat)) continue;
    const site = _curSite; // ไม่ truncate เพื่อให้ตรงกับ fab key
    if (!swInfSiteMap[site]) swInfSiteMap[site]={sw_t:0,sw_d:0,inf_t:0,inf_d:0};
    if (cat==='Switch') { swInfSiteMap[site].sw_t+=qty; swInfSiteMap[site].sw_d+=mig; }
    else               { swInfSiteMap[site].inf_t+=qty; swInfSiteMap[site].inf_d+=mig; }
  }

  const COLORS=['#4361ee','#2bc48a','#ff9f43','#a855f7','#22b8cf','#f76707','#74c0fc'];
  const fabrics = Object.entries(siteMap)
    .filter(([k])=>k&&!k.match(/^\d/)&&!k.startsWith('%'))
    .map(([name,v],i)=>({
      n:name, t:v.total, d:v.done,
      p:v.total>0?Math.round(v.done/v.total*100):0,
      h:0, r:v.total-v.done, c:COLORS[i%7],
      s: v.start ? fmtLbl2(v.start) : '–', e: v.end ? fmtLbl2(v.end) : '–',
      sw:{t:swInfSiteMap[name]?.sw_t||0, d:swInfSiteMap[name]?.sw_d||0},
      ap:{t:0, d:0},
      inf:{t:swInfSiteMap[name]?.inf_t||0, d:swInfSiteMap[name]?.inf_d||0}, weekly:null
    }))
    .sort((a,b)=>b.t-a.t);

  // ── AP จาก HQ-WL sheet ──────────────────────────────────────────────
  const wlRows = XLSX.utils.sheet_to_json(wb.Sheets['HQ-WL'], { header:1, defval:null });
  let apTotal=0, apDone=0;
  // ข้าม row สุดท้าย (summary row) — วน wlRows[1] ถึง length-2
  // หา index ของ summary row จริง (row ที่ col C = 'Summary :')
  // หา wlEndIdx — ข้าม summary rows ทั้งหมดท้าย sheet
  // summary row = col C มี 'Summary' หรือ col D > 100 (ผลรวม)
  // หา wlEndIdx — หา last data row ที่ D (col3) > 0 และ <= 100
  let wlEndIdx = 1;
  for (let i=1; i<wlRows.length; i++) {
    const r = wlRows[i]; if (!r||!r.length) continue;
    const d3 = typeof r[3]==='number' ? r[3] : 0;
    if (d3 > 0 && d3 <= 100) wlEndIdx = i + 1; // +1 เพราะ loop ใช้ i < wlEndIdx
  }
  const apSiteMap = {};
  let _apCurSite = null;
  for (let i=1; i<wlEndIdx; i++) {
    const r = wlRows[i]; if (!r||!r.length) continue;
    if (r[0] && typeof r[0]==='string' && r[0].trim().length>3)
      _apCurSite = r[0].trim();
    const qty = typeof r[3]==='number' ? r[3] : 0;
    const mig = typeof r[16]==='number' ? Math.round(r[16]) : 0;
    if (qty<=0) continue;
    apTotal += qty; apDone += mig;
    if (_apCurSite) {
      if (!apSiteMap[_apCurSite]) apSiteMap[_apCurSite]={total:0,done:0};
      apSiteMap[_apCurSite].total += qty;
      apSiteMap[_apCurSite].done  += mig;
    }
  }
  // AP plan จาก HQ-WL col T(19) — sheet มีข้อมูล 2 ชุดซ้ำ หาร 2
  const _projStartStr = PROJ_START.toISOString().slice(0,10);
  const apPlanByDate = {};
  for (let i=1; i<wlEndIdx; i++) {
    const r = wlRows[i]; if (!r||!r.length) continue;
    const qty = typeof r[3]==='number' ? r[3] : 0;
    const planDt = toDate(r[19]);
    let planStr = planDt ? planDt.toISOString().slice(0,10) : null;
    if (!planStr || qty<=0) continue;
    if (planStr < _projStartStr) planStr = _projStartStr;
    apPlanByDate[planStr] = (apPlanByDate[planStr]||0) + qty;
  }
  const _apPlanSum = Object.values(apPlanByDate).reduce((a,v)=>a+v,0);

  Object.entries(apPlanByDate).forEach(([d,q])=>{
    dayPlanMap[d] = (dayPlanMap[d]||0) + q;
  });
  const _totalPlanSum = Object.values(dayPlanMap).reduce((a,v)=>a+v,0);

  // AP actual — col Q(16) migration + col T(19) date — หาร 2 เพราะข้อมูลซ้ำ
  const apActByDate = {};
  for (let i=1; i<wlEndIdx; i++) {
    const r = wlRows[i]; if (!r||!r.length) continue;
    const mig = typeof r[16]==='number' ? Math.round(r[16]) : 0;
    const helperDt = toDate(r[19]);
    let helperStr = helperDt ? helperDt.toISOString().slice(0,10) : null;
    if (mig<=0 || !helperStr) continue;
    if (helperStr < _projStartStr) helperStr = _projStartStr;
    apActByDate[helperStr] = (apActByDate[helperStr]||0) + mig;
  }
  const _apActSum = Object.values(apActByDate).reduce((a,v)=>a+v,0);
  console.log('[AP ACT] apActByDate sum:', _apActSum, 'entries:', Object.keys(apActByDate).length);
  Object.entries(apActByDate).forEach(([d,q])=>{
    dayActMap[d] = (dayActMap[d]||0) + q;
    dayApActMap[d] = (dayApActMap[d]||0) + q;
    if (!lastInstallDate||d>lastInstallDate) lastInstallDate=d;
  });
  const _totalActMap = Object.values(dayActMap).reduce((a,v)=>a+v,0);
  console.log('[TOTAL ACT] dayActMap after AP merge:', _totalActMap);

  // AP installed จาก HQ-WL
  instAP = apDone;
  installed += apDone;
  TOTAL += apTotal; // รวม AP total เข้า TOTAL

  // คำนวณ swTotal/infTotal สำหรับใช้ใน daily loop
  let swTotal=0, infTotal=0;
  for (let i=2; i<hqRows.length; i++) {
    const r=hqRows[i]; if(!r||!r.length) continue;
    const qty = typeof r[6]==='number' ? r[6] : 0;
    const cat = r[18] ? String(r[18]).trim() : '';
    if (cat==='Switch') swTotal+=qty;
    else if (cat==='Infra') infTotal+=qty;
  }

  // Daily cumulative — ต้องอยู่หลัง AP loops ทั้งหมด
  const PROJ_START_D = new Date(PROJ_START); PROJ_START_D.setHours(0,0,0,0);
  const PROJ_END_D   = new Date(PROJ_END);   PROJ_END_D.setHours(0,0,0,0);
  const _chartEnd = new Date(Math.max(PROJ_END_D.getTime(), today.getTime()+7*86400000));
  const lastActDt    = lastInstallDate ? new Date(lastInstallDate+'T00:00:00') : null;
  console.log('[DAILY] lastInstallDate:', lastInstallDate, 'dayActMap on that date:', dayActMap[lastInstallDate]||0);
  // sum all dayActMap up to lastInstallDate
  let _cumCheck=0;
  Object.entries(dayActMap).forEach(([d,v])=>{ if(d<=lastInstallDate) _cumCheck+=v; });
  console.log('[DAILY] cumAct up to lastInstallDate:', _cumCheck, 'of TOTAL:', TOTAL);

  const dailyLabels=[],dailyActCum=[],dailyPlanCum=[],dailyBdPlan=[],dailyBdAct=[],dailySwActCum=[],dailyApActCum=[];
  let cumAct=0, cumPlan=0, cumSwAct=0, cumApAct=0;
  const cur2 = new Date(PROJ_START_D);
  while (cur2 <= _chartEnd) {
    const k   = cur2.toISOString().slice(0,10);
    const lbl = fmtLbl(cur2);
    // burndown — push ก่อน accumulate ทำให้วันแรก = TOTAL
    const inAct = lastActDt && cur2 <= lastActDt;
    dailyBdPlan.push(TOTAL - Math.round(Math.min(cumPlan/TOTAL,1)*TOTAL));
    dailyBdAct.push(inAct ? TOTAL - Math.round(cumAct) : null);
    // accumulate ก่อน push % สะสม ทำให้วันสุดท้ายได้ครบ
    cumSwAct += daySwActMap[k]||0;
    cumApAct += dayApActMap[k]||0;
    cumAct   += dayActMap[k]||0;
    cumPlan  += dayPlanMap[k]||0;
    dailyLabels.push(lbl);
    dailyPlanCum.push(Math.round(Math.min(cumPlan/TOTAL,1)*10000)/100);
    dailyActCum.push(inAct ? Math.round(cumAct/TOTAL*10000)/100 : null);
    dailySwActCum.push(inAct ? Math.round(cumSwAct/Math.max(swTotal,1)*10000)/100 : null);
    dailyApActCum.push(inAct ? Math.round(cumApAct/Math.max(apTotal,1)*10000)/100 : null);
    cur2.setDate(cur2.getDate()+1);
  }

  const elapsed   = Math.max(1,Math.round((today-PROJ_START)/86400000));
  const projDays  = Math.round((PROJ_END-PROJ_START)/86400000);
  const daysLeft  = Math.max(0,Math.round((PROJ_END-today)/86400000));
  const remaining = TOTAL-installed;
  const dailyRate = Math.round(installed/elapsed*10)/10;
  // ถ้าเลยกำหนดแล้ว ใช้ finish_date คำนวณ req_rate แทน
  const _daysToFinish = dailyRate>0 ? Math.ceil(remaining/dailyRate) : 0;
  const reqRate   = daysLeft>0 ? Math.ceil(remaining/daysLeft) : (remaining>0 ? Math.ceil(remaining/Math.max(_daysToFinish,1)) : 0);
  const needMore  = Math.round((reqRate-dailyRate)*10)/10;
  const gaugePct  = reqRate>0 ? Math.min(150,Math.round(dailyRate/reqRate*100)) : 100;
  const pctDone   = Math.round(installed/TOTAL*100);
  const todayWk   = Math.floor((today-PROJ_START)/(7*86400000));

  let finishDate = null, daysLate=0, daysEarly=0;
  if (dailyRate>0) {
    const fd=new Date(today); fd.setDate(fd.getDate()+Math.ceil(remaining/dailyRate));
    fd.setHours(0,0,0,0);
    finishDate = fd.toISOString().slice(0,10);
    const diffDays = Math.round((PROJ_END-fd)/86400000);
    if (diffDays < 0) daysLate  = Math.abs(diffDays);
    else              daysEarly = diffDays;
  }




  const types = Object.entries(typeMap)
    .map(([n,v])=>({n,plan:v.plan,done:v.done}))
    .sort((a,b)=>b.plan-a.plan).slice(0,20);

  const holdItems = devices.filter(d=>d.status==='Hold')
    .map(d=>({fab:d.site.slice(0,30),loc:'',dev:d.device.slice(0,40),qty:d.qty,done:0,rem:d.qty}));

  return {
    meta:{
      total:TOTAL, installed, in_progress:inProgress, not_started:notStarted,
      remaining, pct_done:pctDone, hold, overdue,
      installed_sw:instSW, installed_ap:instAP, installed_inf:instInf,
      sw_total:swTotal, ap_total:apTotal, inf_total:infTotal,
      on_time_qty:onTimeQty, on_time_pct:onTimePct,
      on_time_early:earlyQty, on_time_late:lateQty,
      proj_start:PROJ_START.toISOString().slice(0,10),
      proj_end:PROJ_END.toISOString().slice(0,10),
      proj_days:projDays, days_left:daysLeft,
    },
    insight:{
      daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
      gauge_pct:gaugePct, elapsed, remaining,
      days_left:daysLeft, days_late:daysLate, days_early:daysEarly,
      finish_date:finishDate,
      pct_more:dailyRate>0?Math.round((reqRate/dailyRate-1)*100):0,
    },
    wk:wkLabels,
    weekly:{
      labels:wkLabels, plan_all:planPct, act_all:actPct,
      plan_sw:planPct, act_sw:actPct, plan_ap:planPct, act_ap:actPct,
      bd_plan:bdPlan, bd_act:bdAct,
    },
    daily:{labels:dailyLabels,sw:[],ap:[],inf:[],plan:[],cum_d:[],cum_sw:[],cum_ap:[],cum_inf:[]},
    daily_progress:{
      labels:dailyLabels, plan_cum:dailyPlanCum, act_cum:dailyActCum,
      sw_plan:dailyPlanCum, sw_act:dailySwActCum,
      ap_plan:dailyPlanCum, ap_act:dailyApActCum,
      bd_plan:dailyBdPlan,
      bd_act: dailyBdAct,
      fab: (() => {
        // per-site daily plan/actual — คำนวณจาก dayPlanBySite/dayActBySite
        const fab = {};
        const allSites = Object.keys(swInfSiteMap);
        allSites.forEach(site => {
          const actDateMap  = dayActBySite[site]  || {};
          const planDateMap = (dayPlanBySite[site]||{}).byDate || {};
          const siteTotal   = swInfSiteMap[site].sw_t + swInfSiteMap[site].inf_t || 1;
          let cumAct2=0, cumPlan2=0;
          const sw_plan=[], sw_act=[];
          // หา lastActDt สำหรับ site นี้
          const siteActDates = Object.keys(actDateMap).sort();
          const siteLastActStr = siteActDates.length ? siteActDates[siteActDates.length-1] : null;
          const siteLastActDt = siteLastActStr ? new Date(siteLastActStr+'T00:00:00') : null;
          // ถ้า site ไม่มี actual ใช้ lastActDt global
          const _siteActDt = siteLastActDt || lastActDt;
          dailyLabels.forEach((lbl) => {
            const parts = lbl.split('/');
            const k2 = `2026-${parts[1]}-${parts[0]}`;
            cumPlan2 += planDateMap[k2]||0;
            cumAct2  += actDateMap[k2]||0;
            const inAct2 = _siteActDt && new Date(k2+'T00:00:00') <= _siteActDt;
            sw_plan.push(Math.round(Math.min(cumPlan2/siteTotal,1)*10000)/100);
            sw_act.push(inAct2 ? Math.round(cumAct2/siteTotal*10000)/100 : null);
          });
          fab[site] = { sw_plan, sw_act, ap_plan:[], ap_act:[] };
        });
        return fab;
      })(),
    },
    fab_colors:{}, fab_plan_totals:{}, fab_totals:{},
    fab_weekly:{},
    fab_daily_plan: (()=>{
      const fdp={};
      Object.keys(swInfSiteMap).forEach(site=>{
        fdp[site]={};
        const planByDate=(dayPlanBySite[site]||{}).byDate||{};
        dailyLabels.forEach(lbl=>{
          const[dd,mm]=lbl.split('/'); const k=`2026-${mm}-${dd}`;
          if(planByDate[k]) fdp[site][lbl]=planByDate[k];
        });
      });
      return fdp;
    })(),
    fab_daily: (()=>{
      const fd={};
      Object.keys(swInfSiteMap).forEach(site=>{
        const actByDate=dayActBySite[site]||{};
        const sw=[], ap=[], inf=[];
        dailyLabels.forEach(lbl=>{
          const[dd,mm]=lbl.split('/'); const k=`2026-${mm}-${dd}`;
          sw.push(actByDate[k]||0); ap.push(0); inf.push(0);
        });
        fd[site]={sw,ap,inf};
      });
      return fd;
    })(),
    daily:{
      labels:dailyLabels,
      sw:dailyLabels.map(lbl=>{const[dd,mm]=lbl.split('/');return daySwActMap[`2026-${mm}-${dd}`]||0;}),
      ap:dailyLabels.map(lbl=>{const[dd,mm]=lbl.split('/');return dayApActMap[`2026-${mm}-${dd}`]||0;}),
      inf:dailyLabels.map(()=>0), plan:dailyLabels.map(lbl=>{const[dd,mm]=lbl.split('/');return dayPlanMap[`2026-${mm}-${dd}`]||0;}),
      cum_d:[],cum_sw:[],cum_ap:[],cum_inf:[]
    },
    locations: (()=>{
      // Site → Room จาก HQ sheet เท่านั้น (SW+Infra) col A=site, col B=room, col G=qty, col P=mig
      const locMap = {};
      let locSite = null;
      for (let i=2; i<hqRows.length; i++) {
        const r=hqRows[i]; if(!r) continue;
        if (r[0]) locSite=String(r[0]).trim();
        if (!locSite) continue;
        if (locSite.startsWith('%') || locSite.match(/^[0-9]/)) { locSite=null; continue; } // skip summary rows
        const qty  = typeof r[6]==='number' ? r[6] : 0;
        if (qty<=0 || qty > 500) continue; // skip summary/aggregate rows
        if (typeof r[1] !== 'string') continue; // skip rows ที่ B ไม่ใช่ string (เช่น ตัวเลข)
        const room = r[1].trim() || '(ไม่ระบุห้อง)';
        const mig  = typeof r[15]==='number' ? Math.round(r[15]) : 0;
        if (!locMap[locSite]) locMap[locSite]={};
        if (!locMap[locSite][room]) locMap[locSite][room]={t:0,d:0};
        locMap[locSite][room].t += qty;
        locMap[locSite][room].d += mig;
      }
      return Object.fromEntries(Object.entries(locMap).map(([site,rooms])=>[
        site, Object.entries(rooms).map(([room,v])=>({l:room,t:v.t,d:v.d,p:v.t>0?Math.round(v.d/v.t*100):0}))
      ]));
    })(),
    types, hold_items:holdItems, fabrics,
    today_wk:todayWk, last_install_date:lastInstallDate,
    upcoming: (()=>{
      const up={};
      const todayStr=today.toISOString().slice(0,10);
      const end14=new Date(today.getTime()+14*86400000).toISOString().slice(0,10);
      let upSite=null;
      for (let i=2; i<hqRows.length; i++) {
        const r=hqRows[i]; if(!r) continue;
        if (r[0]) upSite=String(r[0]).trim();
        if (!upSite) continue;
        const qty=typeof r[6]==='number'?r[6]:0; if(qty<=0) continue;
        const cat=r[18]?String(r[18]).trim():''; if(cat==='AP') continue;
        const mig=typeof r[15]==='number'?Math.round(r[15]):0;
        const hDt=toDate(r[19]); if(!hDt) continue;
        const hStr=hDt.toISOString().slice(0,10);
        if(hStr<todayStr||hStr>end14) continue;
        const dev=r[3]?String(r[3]).slice(0,50):'อุปกรณ์';
        if(!up[hStr]) up[hStr]={};
        if(!up[hStr][upSite]) up[hStr][upSite]={qty:0,rem:0,cats:[],types:[],locs:new Set()};
        up[hStr][upSite].qty+=qty;
        up[hStr][upSite].rem+=Math.max(0,qty-mig);
        if(!up[hStr][upSite].cats.includes(cat)) up[hStr][upSite].cats.push(cat);
        if(!up[hStr][upSite].types.includes(dev)) up[hStr][upSite].types.push(dev);
        if(r[1]) up[hStr][upSite].locs.add(String(r[1]).trim());
      }
      Object.values(up).forEach(day=>Object.values(day).forEach(v=>{v.locs=[...v.locs];}));
      return up;
    })(),
    sites:fabrics.filter(f=>f.t>0).map(f=>({name:f.n,total:f.t,done:f.d,inp:siteMap[f.n]?.inp||0,pct:f.p})),
    sw_inf_sites: Object.entries(swInfSiteMap).map(([name,v])=>({
      name, sw_t:v.sw_t, sw_d:v.sw_d, inf_t:v.inf_t, inf_d:v.inf_d,
      total:v.sw_t+v.inf_t, done:v.sw_d+v.inf_d,
      pct: (v.sw_t+v.inf_t)>0 ? Math.round((v.sw_d+v.inf_d)/(v.sw_t+v.inf_t)*100) : 0
    })).sort((a,b)=>b.total-a.total),
    ap_sites: Object.entries(apSiteMap)
      .filter(([,v])=>v.total>0)
      .map(([name,v])=>({
        name, total:v.total, done:v.done,
        pct: v.total>0?Math.round(v.done/v.total*100):0
      })).sort((a,b)=>b.total-a.total),
  };
}

const upload = multer({dest:'/tmp/'});
app.post('/api/upload-excel', upload.single('excel'), (req,res)=>{
  if (!req.file) return res.status(400).json({error:'No file'});
  fs.renameSync(req.file.path, TMP_EXCEL);
  cache=null;
  res.json({ok:true,filename:req.file.originalname});
});

app.get('/api/dashboard', (req,res)=>{
  const now=Date.now();
  if (!cache||now-cacheTime>300000) {
    try { cache=parseData(); cacheTime=now; }
    catch(e) { console.error(e); return res.status(500).json({error:String(e)}); }
  }
  res.json(cache);
});

app.post('/api/cache/refresh', (req,res)=>{
  cache=null;
  try { res.json({success:true,data:parseData()}); }
  catch(e) { res.status(500).json({error:String(e)}); }
});

app.post('/api/clear-upload', (req,res)=>{
  try {
    if(fs.existsSync(TMP_EXCEL)) fs.unlinkSync(TMP_EXCEL);
    cache=null;
    res.json({success:true, cleared:true, using:'repo Excel'});
  } catch(e) { res.status(500).json({error:String(e)}); }
});

const PORT=process.env.PORT||3000;
app.listen(PORT,()=>console.log(`HQ Dashboard running on port ${PORT}`));
