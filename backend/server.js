const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const https   = require('https');
const http    = require('http');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

// ── SharePoint URL (อัพเดทเป็น link ใหม่เมื่อ share) ─────────────────────────
const SHAREPOINT_URL = process.env.SHAREPOINT_URL || '';
const CACHE_PATH     = path.join(__dirname, 'hq_cache.xlsx');
const LOCAL_PATH     = path.join(__dirname, 'SAT_Progress.xlsx');
const CACHE_TTL_MS   = 5 * 60 * 1000;

let cacheTime = 0;
let cachedWb  = null;

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

async function readWorkbook() {
  const now = Date.now();
  if (cachedWb && (now - cacheTime) < CACHE_TTL_MS) return cachedWb;
  if (SHAREPOINT_URL) {
    try {
      await downloadFile(SHAREPOINT_URL, CACHE_PATH);
      cachedWb = XLSX.readFile(CACHE_PATH);
      cacheTime = Date.now();
      return cachedWb;
    } catch(e) { console.error('SharePoint err:', e.message); }
  }
  // Fallback: local file
  const src = fs.existsSync(CACHE_PATH) ? CACHE_PATH : LOCAL_PATH;
  cachedWb  = XLSX.readFile(src);
  cacheTime = Date.now();
  return cachedWb;
}

// ── Helper: parse Thai number safely ─────────────────────────────────────────
const num = v => (typeof v === 'number' ? v : 0);
const pct = v => (typeof v === 'number' ? Math.round(v * 10000) / 100 : null);

// ── GET /api/summary ─────────────────────────────────────────────────────────
app.get('/api/summary', async (req, res) => {
  try {
    const wb = await readWorkbook();

    // ── HQ main sheet: count items ─────────────────────────────────────────
    const ws  = wb.Sheets['HQ'];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });

    let total_tor=0, total_new=0, complete=0, in_progress=0, not_started=0;
    let total_config=0, total_install=0, total_migrate=0;
    const locs = {};
    let curLoc = null;

    for (let i = 2; i < raw.length; i++) {
      const r = raw[i];
      if (r[0]) curLoc = String(r[0]);
      if (!curLoc) continue;
      if (!locs[curLoc]) locs[curLoc] = {tor:0,new:0,cfg:0,inst:0,mig:0};
      locs[curLoc].tor  += num(r[5]);
      locs[curLoc].new  += num(r[6]);
      locs[curLoc].cfg  += num(r[13]);
      locs[curLoc].inst += num(r[14]);
      locs[curLoc].mig  += num(r[15]);
      total_tor     += num(r[5]);
      total_new     += num(r[6]);
      total_config  += num(r[13]);
      total_install += num(r[14]);
      total_migrate += num(r[15]);
      const st = r[11];
      if (st === 'Complete')    complete++;
      else if (st === 'In Progress') in_progress++;
      else if (st === '-' || st === null) not_started++;
    }

    // ── Weekly graph sheet ─────────────────────────────────────────────────
    const wsW = wb.Sheets['HQ-กราฟรายสัปดาห์'];
    const rawW = XLSX.utils.sheet_to_json(wsW, { header:1, defval:null });
    const wk_labels  = rawW[3].slice(1,10).filter(Boolean).map(v=>String(v).split('\n')[0]);
    const wk_plan    = rawW[4].slice(1,10).map(v=>num(v));
    const wk_cfg     = rawW[5].slice(1,10).map(v=>typeof v==='number'?v:null);
    const wk_inst    = rawW[6].slice(1,10).map(v=>typeof v==='number'?v:null);
    const wk_mig     = rawW[7].slice(1,10).map(v=>typeof v==='number'?v:null);
    const wk_cum_plan= rawW[12].slice(1,10).map(v=>pct(v));
    const wk_cum_act = rawW[13].slice(1,10).map(v=>pct(v));

    // ── Daily graph sheet ──────────────────────────────────────────────────
    const wsD  = wb.Sheets['HQ-กราฟรายวัน'];
    const rawD = XLSX.utils.sheet_to_json(wsD, { header:1, defval:null });
    const day_labels   = rawD[3].slice(1).filter(Boolean).map(v=>String(v).replace('\n',''));
    const day_plan     = rawD[4].slice(1, day_labels.length+1).map(v=>num(v));
    const day_act      = rawD[5].slice(1, day_labels.length+1).map(v=>typeof v==='number'?v:null);
    const day_cum_plan = rawD[8].slice(1, day_labels.length+1).map(v=>pct(v));
    const day_cum_act  = rawD[9].slice(1, day_labels.length+1).map(v=>pct(v));

    // ── Overall stats ──────────────────────────────────────────────────────
    const total_plan  = wk_plan.reduce((a,b)=>a+b,0);  // 193
    const done_mig    = Math.round(total_migrate);
    const pct_done    = total_plan > 0 ? Math.round(done_mig/total_plan*1000)/10 : 0;
    const cur_cum_plan = wk_cum_plan.find((v,i) => wk_cum_act[i]===null && v!==null) || wk_cum_plan[wk_cum_plan.length-1];
    const cur_cum_act  = [...wk_cum_act].reverse().find(v=>v!==null) || 0;
    const on_schedule  = cur_cum_act >= (cur_cum_plan||0);

    res.json({
      overall: {
        total_plan, done_mig, pct_done,
        total_config: Math.round(total_config),
        total_install: Math.round(total_install),
        total_migrate: Math.round(total_migrate),
        complete, in_progress,
        cur_cum_plan, cur_cum_act, on_schedule,
      },
      locations: Object.entries(locs).map(([name, d]) => ({
        name,
        tor: Math.round(d.tor),
        new: Math.round(d.new),
        cfg: Math.round(d.cfg),
        inst: Math.round(d.inst),
        mig: Math.round(d.mig),
        pct: d.tor > 0 ? Math.round(d.mig/d.tor*100) : 0,
      })).filter(l=>l.tor>0),
      weekly: { labels: wk_labels, plan: wk_plan, cfg: wk_cfg, inst: wk_inst, mig: wk_mig, cum_plan: wk_cum_plan, cum_act: wk_cum_act },
      daily:  { labels: day_labels, plan: day_plan, act: day_act, cum_plan: day_cum_plan, cum_act: day_cum_act },
      cached_at: new Date(cacheTime).toISOString(),
    });
  } catch(err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/devices ─────────────────────────────────────────────────────────
app.get('/api/devices', async (req, res) => {
  try {
    const wb  = await readWorkbook();
    const ws  = wb.Sheets['HQ'];
    const raw = XLSX.utils.sheet_to_json(ws, { header:1, defval:null });
    let curLoc = null;
    const rows = [];
    for (let i = 2; i < raw.length; i++) {
      const r = raw[i];
      if (r[0]) curLoc = String(r[0]);
      if (!curLoc || !r[3]) continue;
      rows.push({
        location: curLoc,
        room:     r[1] || '',
        floor:    r[2] || '',
        device:   r[3],
        brand:    r[4] || '',
        tor:      num(r[5]),
        new_qty:  num(r[6]),
        plan_start: r[7] ? (r[7] instanceof Date ? r[7].toISOString().slice(0,10) : r[7]) : null,
        plan_end:   r[8] ? (r[8] instanceof Date ? r[8].toISOString().slice(0,10) : r[8]) : null,
        install_date: r[9] ? (r[9] instanceof Date ? r[9].toISOString().slice(0,10) : r[9]) : null,
        status:   r[11] || '-',
        remark:   r[12] || '',
        config:   num(r[13]),
        install:  num(r[14]),
        migrate:  num(r[15]),
      });
    }
    const { location, status } = req.query;
    let out = rows;
    if (location) out = out.filter(r => r.location.includes(location));
    if (status)   out = out.filter(r => r.status === status);
    res.json({ total: out.length, data: out });
  } catch(err) {
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/cache/refresh ───────────────────────────────────────────────────
app.post('/api/cache/refresh', async (req, res) => {
  cacheTime = 0; cachedWb = null;
  try { await readWorkbook(); res.json({ success:true, cached_at: new Date(cacheTime).toISOString() }); }
  catch(e) { res.status(500).json({ error:e.message }); }
});

app.get('/health', (req, res) => res.json({ status:'ok', cached_at: cacheTime ? new Date(cacheTime).toISOString() : null }));

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => console.log(`HQ API running on port ${PORT}`));
