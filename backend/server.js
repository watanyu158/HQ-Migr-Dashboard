const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');
const https   = require('https');
const http    = require('http');

const app = express();
app.use(cors());
app.use(express.json());

// ── Config ────────────────────────────────────────────────────────────────────
const SHAREPOINT_URL = process.env.SHAREPOINT_URL || '';
const LOCAL_EXCEL    = path.join(__dirname, 'SDA_Installation_Plan_V2.xlsx');
const CACHE_PATH     = path.join(__dirname, 'sda_cache.xlsx');
const CACHE_TTL      = 5 * 60 * 1000; // 5 min

const TOTAL = 1592, TOTAL_SW = 1121, TOTAL_AP = 445, TOTAL_INF = 26;
const PROJ_START = new Date('2026-02-09');
const PROJ_END   = new Date('2026-06-23');
const FABRICS    = ['D1-041','CFZ','T1-015','D1-091','RFF','AMF','PPW'];
const FAB_COLORS = { 'D1-041':'#4361ee','CFZ':'#2bc48a','T1-015':'#ff9f43',
                     'D1-091':'#a855f7','RFF':'#22b8cf','AMF':'#f76707','PPW':'#74c0fc' };
const FAB_PLAN_T = { 'D1-041':710,'CFZ':187,'T1-015':327,'D1-091':147,'RFF':99,'AMF':114,'PPW':8 };

// Week bounds W7–W26 (20 weeks)
const WK_BOUNDS = [
  ['2026-02-09','2026-02-15','09 Feb','W7'],  ['2026-02-16','2026-02-22','16 Feb','W8'],
  ['2026-02-23','2026-03-01','23 Feb','W9'],  ['2026-03-02','2026-03-08','02 Mar','W10'],
  ['2026-03-09','2026-03-15','09 Mar','W11'], ['2026-03-16','2026-03-22','16 Mar','W12'],
  ['2026-03-23','2026-03-29','23 Mar','W13'], ['2026-03-30','2026-04-05','30 Mar','W14'],
  ['2026-04-06','2026-04-12','06 Apr','W15'], ['2026-04-13','2026-04-19','13 Apr','W16'],
  ['2026-04-20','2026-04-26','20 Apr','W17'], ['2026-04-27','2026-05-03','27 Apr','W18'],
  ['2026-05-04','2026-05-10','04 May','W19'], ['2026-05-11','2026-05-17','11 May','W20'],
  ['2026-05-18','2026-05-24','18 May','W21'], ['2026-05-25','2026-05-31','25 May','W22'],
  ['2026-06-01','2026-06-08','01 Jun','W23'], ['2026-06-09','2026-06-15','09 Jun','W24'],
  ['2026-06-16','2026-06-22','16 Jun','W25'], ['2026-06-23','2026-06-29','23 Jun','W26'],
].map(([s,e,label,name]) => ({ s:new Date(s), e:new Date(e), label, name }));
const N_WK = WK_BOUNDS.length;

let cacheTime = 0, cachedData = null;

// ── File helpers ──────────────────────────────────────────────────────────────
function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith('https') ? https : http;
    proto.get(url, { headers:{'User-Agent':'Mozilla/5.0'} }, res => {
      if ([301,302,303,307,308].includes(res.statusCode))
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      if (res.statusCode !== 200) return reject(new Error(`HTTP ${res.statusCode}`));
      const f = fs.createWriteStream(dest);
      res.pipe(f);
      f.on('finish', () => f.close(resolve));
      f.on('error', reject);
    }).on('error', reject);
  });
}

async function getWorkbook() {
  if (SHAREPOINT_URL) {
    try {
      console.log('Fetching Excel from SharePoint...');
      await downloadFile(SHAREPOINT_URL, CACHE_PATH);
      console.log('SharePoint OK');
      return XLSX.readFile(CACHE_PATH);
    } catch(e) {
      console.warn('SharePoint failed:', e.message);
    }
  }
  // fallback: local file (uploaded to GitHub)
  if (fs.existsSync(LOCAL_EXCEL)) {
    console.log('Using local Excel from GitHub');
    return XLSX.readFile(LOCAL_EXCEL);
  }
  throw new Error('No Excel source available');
}

// ── Week index helper ─────────────────────────────────────────────────────────
function wkIdx(dt) {
  if (!dt) return -1;
  const d = new Date(dt); d.setHours(0,0,0,0);
  return WK_BOUNDS.findIndex(w => d >= w.s && d <= w.e);
}

function fmtDate(dt) {
  const d = new Date(dt);
  return `${d.getDate()}/${String(d.getMonth()+1).padStart(2,'0')}`;
}

function cumPct(arr, total) {
  let c = 0;
  return arr.map(v => Math.round((c += v) / total * 10000) / 100);
}

function cumActNull(arr, total, upto) {
  let c = 0, last = null;
  return arr.map((v, i) => {
    if (v > 0) { c += v; last = Math.round(c / total * 10000) / 100; }
    return i <= upto ? last : null;
  });
}

// ── Main calculation ──────────────────────────────────────────────────────────
function calcDashboard(wb) {
  const wsD  = wb.Sheets['Dashboard'];
  const wsA  = wb.Sheets['All_Detail'];
  const dRows = XLSX.utils.sheet_to_json(wsD, { header:1, defval:null });
  const aRows = XLSX.utils.sheet_to_json(wsA, { defval:null });

  // Overview
  // installed = R7[3] Actual Installed
  const installed = dRows[6][3] || 0;

  // hold = นับจำนวน rows ที่ Status='Hold' (ไม่ใช่ qty)
  const hold = aRows.filter(r => r['Status'] === 'Hold').length;

  // overdue = นับ rows ที่ Days Until Due < 0 และยังไม่เสร็จ (ok < qty)
  const overdue = aRows.filter(r => {
    const days = r['Days Until Due'];
    const qty  = r['Qty'] || 0;
    const ok   = r['Qty. Success'] || 0;
    return typeof days === 'number' && days < 0 && ok < qty;
  }).length;

  const remaining = TOTAL - installed;

  // TODAY & insight
  const today = new Date(); today.setHours(0,0,0,0);
  // elapsed: นับรวมวันแรก (proj_start) ด้วย → +1
  const elapsed   = Math.max(1, Math.floor((today - PROJ_START) / 86400000) + 1);
  // daysLeft: นับรวมวันนี้ด้วย → +1
  const daysLeft  = Math.max(1, Math.floor((PROJ_END - today) / 86400000) + 1);
  const dailyRate = Math.round(installed / elapsed * 10) / 10;
  // reqRate ปัดขึ้นเต็มจำนวน
  const reqRate   = Math.ceil(remaining / daysLeft);
  const needMore  = Math.round((reqRate - dailyRate) * 10) / 10;
  const pctMore   = dailyRate > 0 ? Math.round((reqRate / dailyRate - 1) * 100) : 0;
  const daysNeeded= dailyRate > 0 ? Math.ceil(remaining / dailyRate) : 9999;
  const finishDt  = new Date(today); finishDt.setDate(today.getDate() + daysNeeded);
  const daysLate  = Math.max(0, Math.floor((finishDt - PROJ_END) / 86400000));
  const gaugePct  = reqRate > 0 ? Math.min(150, Math.round(dailyRate / reqRate * 100)) : 100;
  const todayWk   = Math.max(0, Math.min(N_WK - 1, Math.floor((elapsed-1) / 7)));

  // Arrays
  const planWk=new Array(N_WK).fill(0); const actWk =new Array(N_WK).fill(0);
  const swPlan=new Array(N_WK).fill(0); const swAct =new Array(N_WK).fill(0);
  const apPlan=new Array(N_WK).fill(0); const apAct =new Array(N_WK).fill(0);
  const infAct=new Array(N_WK).fill(0);

  const fabSwPlan={}, fabApPlan={}, fabSwAct={}, fabApAct={};
  FABRICS.forEach(f => {
    fabSwPlan[f]=new Array(N_WK).fill(0); fabApPlan[f]=new Array(N_WK).fill(0);
    fabSwAct[f] =new Array(N_WK).fill(0); fabApAct[f] =new Array(N_WK).fill(0);
  });

  const dailyMap = {}; // dk → {sw,ap,inf,plan}
  const fabDailyAct={}, fabDailyPlan={};
  FABRICS.forEach(f => { fabDailyAct[f]={}; fabDailyPlan[f]={}; });

  const typeMap = {};
  const locMap  = {}; // fab → loc → {t,d}

  aRows.forEach(r => {
    const fab  = r['Fabric'];
    const cat  = r['Category'];
    const qty  = r['Qty']           || 0;
    const ok   = r['Qty. Success']  || 0;
    const dt   = r['Device Type'];
    const loc  = r['Location'];

    // Device types (all rows)
    if (dt) {
      if (!typeMap[dt]) typeMap[dt] = {plan:0, done:0};
      typeMap[dt].plan += qty;
      typeMap[dt].done += ok;
    }

    if (!FABRICS.includes(fab)) return;

    // Location map
    if (loc) {
      if (!locMap[fab]) locMap[fab] = {};
      if (!locMap[fab][loc]) locMap[fab][loc] = {t:0, d:0};
      locMap[fab][loc].t += qty;
      locMap[fab][loc].d += ok;
    }

    // Scheduled date → plan
    let schedDt = r['Scheduled Date'];
    if (typeof schedDt === 'number') schedDt = new Date((schedDt - 25569) * 86400000);
    if (schedDt && qty > 0) {
      const wi = wkIdx(schedDt); const dk = fmtDate(schedDt);
      if (wi >= 0) {
        planWk[wi] += qty;
        if (cat === 'Switch') { swPlan[wi]+=qty; fabSwPlan[fab][wi]+=qty; }
        else if (cat === 'AP') { apPlan[wi]+=qty; fabApPlan[fab][wi]+=qty; }
      }
      if (!dailyMap[dk]) dailyMap[dk]={sw:0,ap:0,inf:0,plan:0};
      dailyMap[dk].plan += qty;
      fabDailyPlan[fab][dk] = (fabDailyPlan[fab][dk]||0) + qty;
    }

    // Install date → actual
    let instDt = r['Install Date'];
    if (typeof instDt === 'number') instDt = new Date((instDt - 25569) * 86400000);
    if (instDt && ok > 0) {
      const wi = wkIdx(instDt); const dk = fmtDate(instDt);
      if (wi >= 0) {
        actWk[wi] += ok;
        if (cat === 'Switch') { swAct[wi]+=ok; fabSwAct[fab][wi]+=ok; }
        else if (cat === 'AP') { apAct[wi]+=ok; fabApAct[fab][wi]+=ok; }
        else infAct[wi] += ok;
      }
      if (!dailyMap[dk]) dailyMap[dk]={sw:0,ap:0,inf:0,plan:0};
      if (cat === 'Switch') dailyMap[dk].sw += ok;
      else if (cat === 'AP') dailyMap[dk].ap += ok;
      else dailyMap[dk].inf += ok;
      fabDailyAct[fab][dk] = (fabDailyAct[fab][dk]||0) + ok;
    }
  });

  // Cumulative weekly %
  const PLAN_ALL = cumPct(planWk, TOTAL);
  const ACT_ALL  = cumActNull(actWk, TOTAL, todayWk);
  const PLAN_SW  = cumPct(swPlan, TOTAL_SW);
  const ACT_SW   = cumActNull(swAct, TOTAL_SW, todayWk);
  const PLAN_AP  = cumPct(apPlan, TOTAL_AP);
  const ACT_AP   = cumActNull(apAct, TOTAL_AP, todayWk);

  // Burndown
  let s = 0;
  const BD_PLAN = planWk.map(v => TOTAL - (s += v));
  s = 0; let last = null;
  const BD_ACT = actWk.map((v,i) => {
    if (v > 0) { s += v; last = TOTAL - s; }
    return i <= todayWk ? last : null;
  });

  // FAB_WEEKLY
  const fabWeekly = {};
  FABRICS.forEach(f => {
    const swT = fabSwPlan[f].reduce((a,v)=>a+v, 0);
    const apT = fabApPlan[f].reduce((a,v)=>a+v, 0);
    fabWeekly[f] = {
      sw_plan: cumPct(fabSwPlan[f], swT || 1),
      sw_act:  cumActNull(fabSwAct[f], swT || 1, todayWk),
      ap_plan: cumPct(fabApPlan[f], apT || 1),
      ap_act:  cumActNull(fabApAct[f], apT || 1, todayWk),
    };
  });

  // Daily sorted
  const sortedDates = Object.keys(dailyMap).sort((a,b) => {
    const [da,ma] = a.split('/').map(Number);
    const [db,mb] = b.split('/').map(Number);
    return ma !== mb ? ma - mb : da - db;
  });

  let cSW=0, cAP=0, cIN=0;
  const daily = {
    labels:    sortedDates,
    sw:        sortedDates.map(d => dailyMap[d].sw),
    ap:        sortedDates.map(d => dailyMap[d].ap),
    inf:       sortedDates.map(d => dailyMap[d].inf),
    plan:      sortedDates.map(d => dailyMap[d].plan),
    cum_sw:    sortedDates.map(d => (cSW += dailyMap[d].sw)),
    cum_ap:    sortedDates.map(d => (cAP += dailyMap[d].ap)),
    cum_inf:   sortedDates.map(d => (cIN += dailyMap[d].inf)),
  };
  daily.cum_d = daily.cum_sw.map((v,i) => v + daily.cum_ap[i] + daily.cum_inf[i]);

  // FAB_DAILY
  const fabDaily = {}, fabDailyPlanOut = {};
  FABRICS.forEach(f => {
    fabDaily[f] = sortedDates.map(d => fabDailyAct[f][d] || 0);
    fabDailyPlanOut[f] = sortedDates.map(d => fabDailyPlan[f][d] || 0);
  });

  // Fabrics from Dashboard rows 27-33
  const fabRows = dRows.slice(27, 34);
  const fabrics = FABRICS.map(fn => {
    const fr = fabRows.find(r => r[0] === fn) || [];
    const swT=fr[1]||0, swD=fr[2]||0, apT=fr[4]||0, apD=fr[5]||0, infT=fr[7]||0, infD=fr[8]||0;
    const tot=swT+apT+infT, done=swD+apD+infD;
    return {
      n:fn, t:tot, d:done, p:tot>0?Math.round(done/tot*10000)/100:0,
      h:fr[3]||0, r:tot-done, c:FAB_COLORS[fn],
      sw:{t:swT,d:swD}, ap:{t:apT,d:apD}, inf:{t:infT,d:infD},
      weekly: fabWeekly[fn],
    };
  });

  const fabTotals = {};
  fabrics.forEach(f => { fabTotals[f.n] = {sw:f.sw.d, ap:f.ap.d, inf:f.inf.d}; });

  // Types sorted by plan desc
  const types = Object.entries(typeMap)
    .map(([n,d]) => ({n, plan:d.plan, done:d.done}))
    .sort((a,b) => b.plan - a.plan);

  // Locations
  const locations = {};
  Object.entries(locMap).forEach(([fab, locs]) => {
    locations[fab] = Object.entries(locs)
      .map(([l,v]) => ({l, t:v.t, d:v.d, p:Math.round(v.d/Math.max(v.t,1)*100)}))
      .sort((a,b) => b.t - a.t);
  });

  // on-time = installed ที่ Install Date <= Scheduled Date
  let onTimeQty = 0, earlyQty = 0, lateQty = 0;
  aRows.forEach(r => {
    const ok    = r['Qty. Success'] || 0;
    let instDt  = r['Install Date'];
    let schedDt = r['Scheduled Date'];
    if (typeof instDt  === 'number') instDt  = new Date((instDt  - 25569) * 86400000);
    if (typeof schedDt === 'number') schedDt = new Date((schedDt - 25569) * 86400000);
    if (ok > 0 && instDt instanceof Date && schedDt instanceof Date && !isNaN(instDt) && !isNaN(schedDt)) {
      instDt.setHours(0,0,0,0); schedDt.setHours(0,0,0,0);
      if (instDt <= schedDt) {
        onTimeQty += ok;
        const diff = Math.floor((schedDt - instDt) / 86400000);
        if (diff > 0) earlyQty += ok; // ติดตั้งก่อน scheduled date
      } else {
        lateQty += ok;
      }
    }
  });
  const onTimePct = installed > 0 ? Math.round(onTimeQty / installed * 1000) / 10 : 0;

  // hold items list
  const holdItems = aRows
    .filter(r => r['Status']==='Hold')
    .map(r => ({
      fab:  r['Fabric']||'',
      loc:  r['Location']||'',
      dev:  r['Device Type']||'',
      qty:  r['Qty']||0,
      done: r['Qty. Success']||0,
      rem:  r['Qty. Remaining']||0,
    }));

  // Last install date
  let lastInstallDate = null;
  aRows.forEach(r => {
    let d = r['Install Date'];
    if (typeof d === 'number') d = new Date((d - 25569) * 86400000);
    if (d instanceof Date && !isNaN(d)) {
      const ds = d.toISOString().slice(0,10);
      if (!lastInstallDate || ds > lastInstallDate) lastInstallDate = ds;
    }
  });

  return {
    wk:        WK_BOUNDS.map(w => w.label),
    today_wk:  todayWk,
    last_install_date: lastInstallDate,
    meta:      { total:TOTAL, installed, remaining, hold, overdue, on_time_qty:onTimeQty, on_time_pct:onTimePct, on_time_early:earlyQty, on_time_late:lateQty },
    hold_items: holdItems,
    insight:   { daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
                 pct_more:pctMore, days_late:daysLate, gauge_pct:gaugePct,
                 finish_date:finishDt.toISOString().slice(0,10), days_left:daysLeft },
    weekly:    { plan_all:PLAN_ALL, act_all:ACT_ALL, plan_sw:PLAN_SW, act_sw:ACT_SW,
                 plan_ap:PLAN_AP, act_ap:ACT_AP, bd_plan:BD_PLAN, bd_act:BD_ACT },
    fab_weekly: fabWeekly,
    fab_daily:  fabDaily,
    fab_daily_plan: fabDailyPlanOut,
    daily,
    fabrics,
    fab_totals:      fabTotals,
    fab_plan_totals: FAB_PLAN_T,
    fab_colors:      FAB_COLORS,
    types,
    locations,
  };
}

// ── Cache wrapper ─────────────────────────────────────────────────────────────
async function getDashboard(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedData && (now - cacheTime) < CACHE_TTL) return cachedData;
  const wb = await getWorkbook();
  cachedData = calcDashboard(wb);
  cacheTime  = now;
  return cachedData;
}

// ── Routes ────────────────────────────────────────────────────────────────────
app.get('/api/dashboard', async (req, res) => {
  try   { res.json(await getDashboard()); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/cache/refresh', async (req, res) => {
  try   { res.json({ success:true, data: await getDashboard(true) }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/health', (req, res) => res.json({
  status: 'ok',
  source: SHAREPOINT_URL ? 'sharepoint' : 'local_excel',
  cached_at: cacheTime ? new Date(cacheTime).toISOString() : null,
  cache_age_s: cacheTime ? Math.round((Date.now() - cacheTime) / 1000) : null,
}));

app.use(express.static(path.join(__dirname, '../frontend')));

app.get('/api/devices', async (req, res) => {
  try {
    const d = await getDashboard();
    const { limit=20 } = req.query;
    res.json({ data: (d.hold_items || []).slice(0, parseInt(limit)) });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('*', (req, res) => res.sendFile(path.join(__dirname, '../frontend/index.html')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SVB Dashboard running on port ${PORT}`));
