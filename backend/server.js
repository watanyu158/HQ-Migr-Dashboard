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

function parseData() {
  const excelPath = fs.existsSync(TMP_EXCEL) ? TMP_EXCEL : EXCEL_PATH;
  console.log('Reading Excel:', excelPath);
  const wb = XLSX.readFile(excelPath);

  const hqRows = XLSX.utils.sheet_to_json(wb.Sheets['HQ'], { header:1, defval:null });

  // หา proj_start/end จาก Migration Plan column H(7)=เริ่ม, I(8)=สิ้นสุด
  let planDates = [];
  let startDates=[], endDates=[];
  for (let i=2; i<hqRows.length; i++) {
    const r=hqRows[i]; if(!r) continue;
    // col H(7) = เริ่ม, col I(8) = สิ้นสุด
    [r[7]].forEach(v => {
      const d = typeof v==='number'&&v>40000 ? new Date((v-25569)*86400000) : toDate(v);
      if (d&&!isNaN(d.getTime())) startDates.push(d);
    });
    [r[8]].forEach(v => {
      const d = typeof v==='number'&&v>40000 ? new Date((v-25569)*86400000) : toDate(v);
      if (d&&!isNaN(d.getTime())) endDates.push(d);
    });
  }
  const PROJ_START = startDates.length ? new Date(Math.min(...startDates)) : new Date('2026-02-02');
  const PROJ_END   = endDates.length   ? new Date(Math.max(...endDates))   : new Date('2026-04-30');
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

  const siteMap={}, typeMap={}, dayActMap={}, dayPlanMap={}, dayActBySite={}, dayPlanBySite={};
  const devices=[];
  const today = new Date(); today.setHours(0,0,0,0);

  for (let i = 2; i < hqRows.length; i++) {
    const r = hqRows[i];
    if (!r || !r.length) continue;
    if (r[0]) curSite = String(r[0]).trim();

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

    const site = curSite.length>50 ? curSite.slice(0,50)+'…' : curSite;
    if (!siteMap[site]) siteMap[site] = {total:0,done:0,inp:0};
    siteMap[site].total += qty;

    const dev = device.length>60 ? device.slice(0,60)+'…' : device;
    if (!typeMap[dev]) typeMap[dev] = {plan:0,done:0,cat};
    typeMap[dev].plan += qty;

    // dayPlanMap ใช้ col T(19) วันที่เริ่ม Helper — clamp ไม่ต่ำกว่า PROJ_START
    const helperDt = toDate(r[19]);
    let helperStr = helperDt ? helperDt.toISOString().slice(0,10) : schedStr;
    if (helperStr && helperStr < PROJ_START.toISOString().slice(0,10)) helperStr = PROJ_START.toISOString().slice(0,10);
    if (helperStr && cat !== 'AP') dayPlanMap[helperStr] = (dayPlanMap[helperStr]||0) + qty;

    // นับจาก Migration column (index 15) — เฉพาะ SW และ Infra
    const migration = typeof r[15]==='number' ? r[15] : 0;
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
        // per-site tracking
        if (site) {
          if (!dayActBySite[site]) dayActBySite[site]={};
          dayActBySite[site][instStr2]=(dayActBySite[site][instStr2]||0)+migration;
          if (!dayPlanBySite[site]) dayPlanBySite[site]={total:0};
          dayPlanBySite[site].total += qty;
        }
      }
      // on-time check ใช้ instDt2 (helper) vs schedDt
      if (instDt2 && schedDt) {
        const id=new Date(instDt2+'T00:00:00'); id.setHours(0,0,0,0);
        const sd=new Date(schedDt); sd.setHours(0,0,0,0);
        if (id<=sd){onTimeQty+=migration;if(id<sd)earlyQty+=migration;}
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
    const qty = typeof r[6]==='number' ? r[6] : 0;
    const mig = typeof r[15]==='number' ? r[15] : 0;
    const cat = r[18] ? String(r[18]).trim() : '';
    if (qty<=0 || !_curSite || !['Switch','Infra'].includes(cat)) continue;
    const site = _curSite.length>50?_curSite.slice(0,50)+'…':_curSite;
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
      s:'–', e:'–',
      sw:{t:0,d:0}, ap:{t:0,d:0}, inf:{t:0,d:0}, weekly:null
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
    const mig = typeof r[16]==='number' ? r[16] : 0;
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
    const mig = typeof r[16]==='number' ? r[16] : 0;
    const helperDt = toDate(r[19]);
    let helperStr = helperDt ? helperDt.toISOString().slice(0,10) : null;
    if (mig<=0 || !helperStr) continue;
    if (helperStr < _projStartStr) helperStr = _projStartStr;
    apActByDate[helperStr] = (apActByDate[helperStr]||0) + mig;
  }
  Object.entries(apActByDate).forEach(([d,q])=>{
    dayActMap[d] = (dayActMap[d]||0) + q;
    if (!lastInstallDate||d>lastInstallDate) lastInstallDate=d;
  });

  // AP installed จาก HQ-WL
  instAP = apDone;
  installed += apDone;
  TOTAL += apTotal; // รวม AP total เข้า TOTAL

  // Daily cumulative — ต้องอยู่หลัง AP loops ทั้งหมด
  const PROJ_START_D = new Date(PROJ_START); PROJ_START_D.setHours(0,0,0,0);
  const PROJ_END_D   = new Date(PROJ_END);   PROJ_END_D.setHours(0,0,0,0);
  const _chartEnd = new Date(Math.max(PROJ_END_D.getTime(), today.getTime()+7*86400000));
  const lastActDt    = lastInstallDate ? new Date(lastInstallDate+'T00:00:00') : null;

  const dailyLabels=[],dailyActCum=[],dailyPlanCum=[],dailyBdPlan=[],dailyBdAct=[];
  let cumAct=0, cumPlan=0;
  const cur2 = new Date(PROJ_START_D);
  while (cur2 <= _chartEnd) {
    const k   = cur2.toISOString().slice(0,10);
    const lbl = fmtLbl(cur2);
    const inAct = lastActDt && cur2 <= lastActDt;
    dailyLabels.push(lbl);
    // % สะสม (plan_cum/act_cum) — สำหรับกราฟความก้าวหน้า
    dailyPlanCum.push(Math.round(Math.min(cumPlan/TOTAL,1)*10000)/100);
    dailyActCum.push(inAct ? Math.round(cumAct/TOTAL*10000)/100 : null);
    // burndown — สำหรับ burn-down chart (push ก่อน accumulate = วันแรก TOTAL)
    dailyBdPlan.push(TOTAL - Math.round(Math.min(cumPlan/TOTAL,1)*TOTAL));
    dailyBdAct.push(inAct ? TOTAL - Math.round(cumAct) : null);
    cumAct  += dayActMap[k]||0;
    cumPlan += dayPlanMap[k]||0;
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


  // นับ SW/Infra total จาก HQ sheet
  let swTotal=0, infTotal=0;
  for (let i=2; i<hqRows.length; i++) {
    const r=hqRows[i]; if(!r||!r.length) continue;
    const qty = typeof r[6]==='number' ? r[6] : 0;
    const cat = r[18] ? String(r[18]).trim() : '';
    if (qty<=0) continue;
    if (cat==='Switch') swTotal+=qty;
    else if (cat==='Infra') infTotal+=qty;
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
      sw_plan:dailyPlanCum, sw_act:dailyActCum,
      ap_plan:dailyPlanCum, ap_act:dailyActCum,
      bd_plan:dailyBdPlan,
      bd_act: dailyBdAct,
      fab: (() => {
        // per-site daily act จาก dayActBySite
        const fab = {};
        Object.entries(dayActBySite||{}).forEach(([site, dateMap]) => {
          let cum = 0;
          const act_cum = dailyLabels.map((lbl, i) => {
            const k = cur2 ? null : null; // ใช้ key จาก dailyLabels
            // map lbl DD/MM กลับเป็น YYYY-MM-DD
            const parts = lbl.split('/');
            const k2 = `2026-${parts[1]}-${parts[0]}`;
            cum += (dateMap[k2]||0);
            const inAct = lastActDt && new Date(k2+'T00:00:00') <= lastActDt;
            return inAct ? Math.round(cum/(dayPlanBySite[site]?.total||1)*10000)/100 : null;
          });
          fab[site] = { sw_plan: dailyPlanCum, sw_act: act_cum, ap_plan: dailyPlanCum, ap_act: [] };
        });
        return fab;
      })(),
    },
    fab_colors:{}, fab_plan_totals:{}, fab_totals:{}, fab_weekly:{}, fab_daily:{}, fab_daily_plan:{},
    locations:{},
    types, hold_items:holdItems, fabrics,
    today_wk:todayWk, last_install_date:lastInstallDate, upcoming:{},
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

const PORT=process.env.PORT||3000;
app.listen(PORT,()=>console.log(`HQ Dashboard running on port ${PORT}`));
