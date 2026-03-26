const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const LOCAL_EXCEL = path.join(__dirname, 'SAT_Progress.xlsx');
const TOTAL = 193;
const PROJ_START = new Date('2026-02-03');
const PROJ_END   = new Date('2026-04-03');

let cacheTime = 0, cachedData = null;
const CACHE_TTL = 5 * 60 * 1000;

function calcDashboard(wb) {
  const wsWk  = wb.Sheets['HQ-กราฟรายสัปดาห์'];
  const wsDay = wb.Sheets['HQ-กราฟรายวัน'];
  const wsHQ  = wb.Sheets['HQ'];

  const wkRows  = XLSX.utils.sheet_to_json(wsWk,  {header:1, defval:null});
  const dayRows = XLSX.utils.sheet_to_json(wsDay, {header:1, defval:null});
  const hqRaw   = XLSX.utils.sheet_to_json(wsHQ,  {header:1, defval:null});

  const N_WK = 9;
  const num = v => typeof v === 'number' ? v : null;
  const pct = v => typeof v === 'number' ? Math.round(v * 10000) / 100 : null;

  // Weekly
  const plan_wk      = wkRows[4].slice(1, N_WK+1).map(v => num(v) || 0);
  const cfg_wk       = wkRows[5].slice(1, N_WK+1).map(num);
  const mig_wk       = wkRows[7].slice(1, N_WK+1).map(num);
  const plan_cum_pct = wkRows[12].slice(1, N_WK+1).map(pct);
  const act_cum_pct  = wkRows[13].slice(1, N_WK+1).map(pct);

  const wk = wkRows[3].slice(1, N_WK+1).map(v =>
    v ? String(v).split('\n')[0].replace(/\r/g,'') : ''
  );
  const wk_dates = wkRows[3].slice(1, N_WK+1).map(v => {
    if (!v) return '';
    const parts = String(v).split('\n');
    return parts[1] ? parts[1].replace(/[()]/g,'').replace(/\r/g,'') : '';
  });

  const mig_total = mig_wk.reduce((s,v) => s + (v||0), 0);
  const cfg_total = cfg_wk.reduce((s,v) => s + (v||0), 0);

  // Burndown
  let s = 0;
  const bd_plan = plan_wk.map(v => TOTAL - (s += v));
  s = 0; let last = null;
  const bd_act = mig_wk.map((v,i) => {
    if (v !== null) { s += v; last = TOTAL - s; }
    return i <= 7 ? last : null;
  });

  // Insight
  const today = new Date(); today.setHours(0,0,0,0);
  const elapsed   = Math.max(1, Math.floor((today - PROJ_START) / 86400000) + 1);
  const daysLeft  = Math.max(1, Math.floor((PROJ_END - today) / 86400000) + 1);
  const remaining = TOTAL - mig_total;
  const dailyRate = Math.round(mig_total / elapsed * 100) / 100;
  const reqRate   = Math.ceil(remaining / daysLeft);
  const needMore  = Math.round((reqRate - dailyRate) * 100) / 100;
  const pctMore   = dailyRate > 0 ? Math.round((reqRate / dailyRate - 1) * 100) : 0;
  const daysNeeded = dailyRate > 0 ? Math.ceil(remaining / dailyRate) : 9999;
  const finishDt  = new Date(today); finishDt.setDate(today.getDate() + daysNeeded);
  const daysLate  = Math.max(0, Math.floor((finishDt - PROJ_END) / 86400000));
  const gaugePct  = reqRate > 0 ? Math.min(150, Math.round(dailyRate / reqRate * 100)) : 100;
  const todayWk   = Math.max(0, Math.min(N_WK-1, Math.floor((elapsed-1) / 7)));

  // Daily
  const dayLabels=[], planDay=[], migDay=[], planCumD=[], actCumD=[];
  for (let i = 1; i < (dayRows[3]||[]).length; i++) {
    const raw = dayRows[3][i];
    if (raw === null) continue;
    dayLabels.push(String(raw).replace(/\r/g,'').replace('\n','/'));
    planDay.push(typeof dayRows[4][i]==='number' ? dayRows[4][i] : 0);
    migDay.push(typeof dayRows[5][i]==='number' ? dayRows[5][i] : 0);
    planCumD.push(pct(dayRows[8][i]) || 0);
    actCumD.push(pct(dayRows[9][i]));
  }

  // Last install date
  let lastInstall = null;
  hqRaw.slice(2).forEach(r => {
    const d = r[9];
    if (typeof d !== 'number') return;
    const dt = new Date((d - 25569) * 86400000);
    dt.setHours(0,0,0,0);
    if (dt <= today && (!lastInstall || dt > lastInstall)) lastInstall = dt;
  });
  const lastInstallDate = lastInstall ? lastInstall.toISOString().slice(0,10) : null;

  // Locations — col index: 0=สถานที่, 3=อุปกรณ์, 5=TOR, 13=Cfg, 14=Inst, 15=Mig
  const locMap = {};
  let curLoc = null;
  hqRaw.slice(2).forEach(r => {
    if (r[0]) curLoc = String(r[0]).trim();
    if (!curLoc || curLoc === 'เพิ่มเติม') return;
    const device = r[3];
    const tor    = r[5];
    if (!device || typeof device !== 'string') return;
    if (typeof tor !== 'number' || !Number.isInteger(tor) || tor <= 0) return;
    const mig  = typeof r[15] === 'number' ? Math.round(r[15]) : 0;
    const cfg  = typeof r[13] === 'number' ? Math.round(r[13]) : 0;
    const inst = typeof r[14] === 'number' ? Math.round(r[14]) : 0;
    if (!locMap[curLoc]) locMap[curLoc] = {tor:0, cfg:0, inst:0, mig:0};
    locMap[curLoc].tor  += tor;
    locMap[curLoc].cfg  += cfg;
    locMap[curLoc].inst += inst;
    locMap[curLoc].mig  += mig;
  });

  const locations = Object.entries(locMap)
    .filter(([,v]) => v.tor > 0)
    .map(([n,v]) => ({
      n, tor:v.tor, cfg:v.cfg, inst:v.inst, mig:v.mig,
      pct: Math.round(v.mig / v.tor * 100)
    }))
    .sort((a,b) => b.pct - a.pct || b.tor - a.tor);

  console.log(`✓ Locations: ${locations.length} | ${locations.map(l=>l.n.slice(0,10)+':'+l.mig).join(', ')}`);

  return {
    wk, wk_dates, today_wk: todayWk,
    last_install_date: lastInstallDate,
    meta: {total:TOTAL, mig:mig_total, cfg:cfg_total, remaining, hold:0},
    insight: {
      daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
      pct_more:pctMore, days_late:daysLate, gauge_pct:gaugePct,
      finish_date:finishDt.toISOString().slice(0,10), days_left:daysLeft
    },
    weekly: {plan:plan_wk, cfg:cfg_wk, mig:mig_wk, bd_plan, bd_act, plan_cum_pct, act_cum_pct},
    daily:  {labels:dayLabels, plan:planDay, mig:migDay, plan_cum_pct:planCumD, act_cum_pct:actCumD},
    locations
  };
}

async function getDashboard(force=false) {
  const now = Date.now();
  if (!force && cachedData && (now - cacheTime) < CACHE_TTL) return cachedData;
  if (!fs.existsSync(LOCAL_EXCEL)) throw new Error('SAT_Progress.xlsx not found');
  console.log('Reading Excel...');
  const wb = XLSX.readFile(LOCAL_EXCEL);
  cachedData = calcDashboard(wb);
  cacheTime = now;
  console.log(`✓ mig=${cachedData.meta.mig}/${TOTAL} rate=${cachedData.insight.daily_rate}`);
  return cachedData;
}

app.get('/api/dashboard', async (req,res) => {
  try { res.json(await getDashboard()); }
  catch(e) { res.status(500).json({error:e.message}); }
});

app.post('/api/cache/refresh', async (req,res) => {
  try { res.json({success:true, data: await getDashboard(true)}); }
  catch(e) { res.status(500).json({error:e.message}); }
});

app.get('/health', (req,res) => res.json({status:'ok', cached_at: cacheTime ? new Date(cacheTime).toISOString() : null}));

app.use(express.static(path.join(__dirname,'../frontend')));
app.get('*', (req,res) => res.sendFile(path.join(__dirname,'../frontend/index.html')));

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`HQ Dashboard on port ${PORT}`));
