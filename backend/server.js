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
const PROJ_START = new Date('2026-02-02T00:00:00');
const PROJ_END   = new Date('2026-04-30T00:00:00');
const TOTAL      = 193;

let cache = null, cacheTime = 0;

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v === 'number') return new Date((v - 25569) * 86400000);
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

  let installed=0, inProgress=0, notStarted=0, hold=0, overdue=0;
  let onTimeQty=0, earlyQty=0, lateQty=0;
  let instSW=0, instAP=0, instInf=0;
  let lastInstallDate = null;
  let curSite = null;

  const siteMap={}, typeMap={}, dayActMap={}, dayPlanMap={};
  const devices=[];
  const today = new Date(); today.setHours(0,0,0,0);

  for (let i = 2; i < hqRows.length; i++) {
    const r = hqRows[i];
    if (!r || !r.length) continue;
    if (r[0]) curSite = String(r[0]).trim();

    const device  = r[3] ? String(r[3]).trim() : null;
    const qty     = typeof r[6]==='number' ? r[6] : 0;
    const status  = r[11] ? String(r[11]).trim() : '';
    const instDt  = toDate(r[9]);
    const schedDt = toDate(r[7]);
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

    if (schedStr) dayPlanMap[schedStr] = (dayPlanMap[schedStr]||0) + qty;

    if (status === 'Complete') {
      installed += qty;
      siteMap[site].done += qty;
      typeMap[dev].done += qty;
      if (cat==='Switch') instSW+=qty;
      else if (cat==='AP') instAP+=qty;
      else instInf+=qty;
      if (instStr) {
        if (!lastInstallDate||instStr>lastInstallDate) lastInstallDate=instStr;
        dayActMap[instStr]=(dayActMap[instStr]||0)+qty;
      }
      if (instDt && schedDt) {
        const id=new Date(instDt); id.setHours(0,0,0,0);
        const sd=new Date(schedDt); sd.setHours(0,0,0,0);
        if (id<=sd){onTimeQty+=qty;if(id<sd)earlyQty+=qty;}
        else lateQty+=qty;
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

  // Daily cumulative
  const PROJ_START_D = new Date(PROJ_START); PROJ_START_D.setHours(0,0,0,0);
  const PROJ_END_D   = new Date(PROJ_END);   PROJ_END_D.setHours(0,0,0,0);
  const lastActDt    = lastInstallDate ? new Date(lastInstallDate+'T00:00:00') : null;

  const dailyLabels=[],dailyActCum=[],dailyPlanCum=[];
  let cumAct=0, cumPlan=0;
  const cur = new Date(PROJ_START_D);
  while (cur <= PROJ_END_D) {
    const k   = cur.toISOString().slice(0,10);
    const lbl = fmtLbl(cur);
    cumAct  += dayActMap[k]||0;
    cumPlan += dayPlanMap[k]||0;
    const inAct = lastActDt && cur <= lastActDt;
    dailyLabels.push(lbl);
    dailyActCum.push(inAct ? Math.round(cumAct/TOTAL*10000)/100 : null);
    dailyPlanCum.push(Math.round(cumPlan/TOTAL*10000)/100);
    cur.setDate(cur.getDate()+1);
  }

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
  const elapsed   = Math.max(1,Math.round((today-PROJ_START)/86400000));
  const projDays  = Math.round((PROJ_END-PROJ_START)/86400000);
  const daysLeft  = Math.max(0,Math.round((PROJ_END-today)/86400000));
  const remaining = TOTAL-installed;
  const dailyRate = Math.round(installed/elapsed*10)/10;
  const reqRate   = daysLeft>0 ? Math.ceil(remaining/daysLeft) : 0;
  const needMore  = Math.round((reqRate-dailyRate)*10)/10;
  const gaugePct  = reqRate>0 ? Math.min(150,Math.round(dailyRate/reqRate*100)) : 100;
  const pctDone   = Math.round(installed/TOTAL*100);
  const todayWk   = Math.floor((today-PROJ_START)/(7*86400000));

  let finishDate = null;
  if (dailyRate>0) {
    const fd=new Date(today); fd.setDate(fd.getDate()+Math.ceil(remaining/dailyRate));
    finishDate = fd.toISOString().slice(0,10);
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
  // ข้าม row สุดท้าย (summary) — วน wlRows[1] ถึง length-2
  for (let i=1; i<wlRows.length-1; i++) {
    const r = wlRows[i]; if (!r||!r.length) continue;
    const qty = typeof r[3]==='number' ? r[3] : 0;
    const mig = typeof r[6]==='number' ? r[6] : 0;
    if (qty<=0) continue;
    apTotal += qty; apDone += mig;
  }
  // update instAP
  instAP = apDone;

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
      ap_total:apTotal,
      on_time_qty:onTimeQty, on_time_pct:onTimePct,
      on_time_early:earlyQty, on_time_late:lateQty,
      proj_start:PROJ_START.toISOString().slice(0,10),
      proj_end:PROJ_END.toISOString().slice(0,10),
      proj_days:projDays, days_left:daysLeft,
    },
    insight:{
      daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
      gauge_pct:gaugePct, elapsed, remaining,
      days_left:daysLeft, days_late:0, days_early:0,
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
      fab:{},
    },
    fab_colors:{}, fab_plan_totals:{}, fab_totals:{}, fab_weekly:{}, fab_daily:{}, fab_daily_plan:{},
    locations:{},
    types, hold_items:holdItems, fabrics,
    today_wk:todayWk, last_install_date:lastInstallDate, upcoming:{},
    sites:fabrics.map(f=>({name:f.n,total:f.t,done:f.d,inp:siteMap[f.n]?.inp||0,pct:f.p})),
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
