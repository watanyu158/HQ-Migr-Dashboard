const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, 'SAT Progress.xlsx');
const TOTAL = 193;

// Project dates
const PROJ_START = new Date('2026-02-02');
const PROJ_END   = new Date('2026-04-30');

let cache = null;
let cacheTime = 0;

function parseData() {
  const wb = XLSX.readFile(EXCEL_PATH);

  // ── Sheet: HQ (device list) ─────────────────────────────────────────
  const hqSheet = wb.Sheets['HQ'];
  const hqRows  = XLSX.utils.sheet_to_json(hqSheet, { header: 1, defval: null });

  let installed = 0, inProgress = 0, notStarted = 0;
  let lastInstallDate = null;
  let curSite = null, curRoom = null, curFloor = null;

  const siteMap = {};
  const devices = [];

  for (let i = 2; i < hqRows.length; i++) {
    const r = hqRows[i];
    if (!r || !r.length) continue;

    if (r[0]) curSite  = String(r[0]).trim();
    if (r[1]) curRoom  = String(r[1]).trim();
    if (r[2]) curFloor = r[2];

    const device   = r[3] ? String(r[3]).trim() : null;
    const qty      = typeof r[6] === 'number' ? r[6] : 0;
    const status   = r[11] ? String(r[11]).trim() : '';
    const instDt   = typeof r[9] === 'number'
      ? new Date(Date.UTC(1899,11,30) + r[9]*86400000).toISOString().slice(0,10)
      : null;
    const compDt   = typeof r[10] === 'number'
      ? new Date(Date.UTC(1899,11,30) + r[10]*86400000).toISOString().slice(0,10)
      : null;
    const planStart = typeof r[7] === 'number'
      ? new Date(Date.UTC(1899,11,30) + r[7]*86400000).toISOString().slice(0,10)
      : null;
    const planEnd   = typeof r[8] === 'number'
      ? new Date(Date.UTC(1899,11,30) + r[8]*86400000).toISOString().slice(0,10)
      : null;

    if (!device || !curSite || qty <= 0) continue;

    if (status === 'Complete') {
      installed += qty;
      if (instDt && (!lastInstallDate || instDt > lastInstallDate)) lastInstallDate = instDt;
    } else if (status.includes('Progress')) {
      inProgress += qty;
    } else {
      notStarted += qty;
    }

    // สรุปรายสถานที่
    const siteName = curSite.length > 50
      ? curSite.slice(0,50) + '…'
      : curSite;
    if (!siteMap[siteName]) siteMap[siteName] = { total:0, done:0, inp:0 };
    siteMap[siteName].total += qty;
    if (status === 'Complete')         siteMap[siteName].done += qty;
    else if (status.includes('Progress')) siteMap[siteName].inp += qty;

    devices.push({ site: curSite, room: curRoom, device, qty, status, instDt, compDt, planStart, planEnd });
  }

  // ── Sheet: HQ-กราฟรายสัปดาห์ (weekly plan vs actual) ─────────────
  const wkSheet = wb.Sheets['HQ-กราฟรายสัปดาห์'];
  const wkRows  = XLSX.utils.sheet_to_json(wkSheet, { header: 1, defval: null });

  // หา row ที่มี W.1, W.2, ...
  let wkLabels = [], planPct = [], actPct = [];
  for (const row of wkRows) {
    if (!row[1]) continue;
    const lbl = String(row[1]);
    if (lbl.startsWith('W.')) {
      // นี่คือ header row
      for (let i = 1; i < row.length; i++) {
        if (row[i] && String(row[i]).startsWith('W.')) {
          wkLabels.push(String(row[i]).split('\n')[0]);
        }
      }
    }
    if (String(row[0]).includes('% ความก้าวหน้าโดยรวม')) {
      planPct = wkLabels.map((_,i) => Math.round((row[i+1]||0)*100));
    }
    if (String(row[0]).includes('จำนวนที่เสร็จสิ้นสะสม')) {
      actPct = wkLabels.map((_,i) => Math.round((row[i+1]||0)*100));
    }
  }

  // fallback weekly labels
  if (!wkLabels.length) {
    wkLabels = ['W.1','W.2','W.3','W.4','W.5','W.6','W.7','W.8','W.9','W.10'];
  }

  // ── คำนวณ insight ──────────────────────────────────────────────────
  const today    = new Date();
  today.setHours(0,0,0,0);
  const elapsed  = Math.max(1, Math.round((today - PROJ_START) / 86400000));
  const projDays = Math.round((PROJ_END - PROJ_START) / 86400000);
  const daysLeft = Math.max(0, Math.round((PROJ_END - today) / 86400000));
  const remaining = TOTAL - installed;
  const dailyRate = Math.round(installed / elapsed * 10) / 10;
  const reqRate   = daysLeft > 0 ? Math.ceil(remaining / daysLeft) : 0;
  const needMore  = Math.round((reqRate - dailyRate) * 10) / 10;
  const gaugePct  = reqRate > 0 ? Math.min(150, Math.round(dailyRate / reqRate * 100)) : 100;
  const pctDone   = Math.round(installed / TOTAL * 100);

  // today_wk
  const msPerWk = 7*86400000;
  const todayWk = Math.floor((today - PROJ_START) / msPerWk);

  return {
    meta: {
      total: TOTAL, installed, in_progress: inProgress, not_started: notStarted,
      remaining, pct_done: pctDone,
      proj_start: PROJ_START.toISOString().slice(0,10),
      proj_end:   PROJ_END.toISOString().slice(0,10),
      proj_days: projDays, days_left: daysLeft,
    },
    insight: {
      daily_rate: dailyRate, req_rate: reqRate, need_more: needMore,
      gauge_pct: gaugePct, elapsed, remaining,
      days_late:  daysLeft < 0 ? Math.abs(daysLeft) : 0,
      days_early: 0,
    },
    weekly: {
      labels: wkLabels,
      plan_pct: planPct,
      act_pct:  actPct,
    },
    today_wk: todayWk,
    last_install_date: lastInstallDate,
    sites: Object.entries(siteMap)
      .filter(([k]) => k && !k.match(/^[\d\.]+$/) && !k.startsWith('%'))
      .map(([name,v]) => ({ name, total:v.total, done:v.done, inp:v.inp,
        pct: v.total>0 ? Math.round(v.done/v.total*100) : 0 }))
      .sort((a,b) => b.total - a.total),
    devices: devices.slice(0, 200),
  };
}

app.get('/api/dashboard', (req, res) => {
  const now = Date.now();
  if (!cache || now - cacheTime > 60000) {
    try { cache = parseData(); cacheTime = now; }
    catch(e) { console.error(e); return res.status(500).json({error:String(e)}); }
  }
  res.json(cache);
});

app.post('/api/cache/refresh', (req, res) => {
  cache = null; res.json({ok:true});
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`HQ Dashboard API running on port ${PORT}`));
