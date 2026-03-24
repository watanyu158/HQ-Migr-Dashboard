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

// ── Excel source: SharePoint direct download ──────────────────────────────────
const SHAREPOINT_URL = 'https://aitcoth-my.sharepoint.com/:x:/g/personal/suttipong_s_ait_co_th/IQB4depTDLOdRbI2UEHtAB7RAbaE9Ybz60zc_CjOHPUMkmI?e=WqeQTd&download=1';
const CACHE_PATH     = path.join(__dirname, 'sda_cache.xlsx');
const CACHE_TTL_MS   = 5 * 60 * 1000; // 5 minutes

let cacheTime = 0;
let cachedWb  = null;

// Download file following redirects (SharePoint redirects multiple times)
function downloadFile(url, dest) {
  return new Promise((resolve, reject) => {
    const proto = url.startsWith('https') ? https : http;
    proto.get(url, { headers: { 'User-Agent': 'Mozilla/5.0' } }, res => {
      // Follow redirect
      if (res.statusCode === 301 || res.statusCode === 302 || res.statusCode === 303 || res.statusCode === 307 || res.statusCode === 308) {
        return downloadFile(res.headers.location, dest).then(resolve).catch(reject);
      }
      if (res.statusCode !== 200) {
        return reject(new Error(`HTTP ${res.statusCode} from SharePoint`));
      }
      const file = fs.createWriteStream(dest);
      res.pipe(file);
      file.on('finish', () => file.close(resolve));
      file.on('error', reject);
    }).on('error', reject);
  });
}

async function readWorkbook() {
  const now = Date.now();

  // Return cache if fresh
  if (cachedWb && (now - cacheTime) < CACHE_TTL_MS) {
    return cachedWb;
  }

  // Try SharePoint download
  try {
    console.log('Fetching Excel from SharePoint...');
    await downloadFile(SHAREPOINT_URL, CACHE_PATH);
    cachedWb  = XLSX.readFile(CACHE_PATH);
    cacheTime = Date.now();
    console.log('Excel loaded from SharePoint OK');
    return cachedWb;
  } catch (err) {
    console.error('SharePoint fetch failed:', err.message);
    // Fallback: use local file if exists
    const localPath = path.join(__dirname, 'SDA_Installation_Plan_V2.xlsx');
    if (fs.existsSync(localPath)) {
      console.log('Using local Excel fallback');
      cachedWb  = XLSX.readFile(localPath);
      cacheTime = Date.now();
      return cachedWb;
    }
    // Last resort: use stale cache
    if (cachedWb) {
      console.log('Using stale cache');
      return cachedWb;
    }
    throw new Error('No Excel source available: ' + err.message);
  }
}

// ── GET /api/summary ─────────────────────────────────────────────────────────
app.get('/api/summary', async (req, res) => {
  try {
    const wb  = await readWorkbook();
    const ws  = wb.Sheets['Dashboard'];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });

    const overall = {
      total_devices:    rows[5][1],
      completed:        rows[5][3],
      on_plan:          rows[5][5],
      hold:             rows[5][7],
      pending:          rows[5][9],
      progress_pct:     Math.round(rows[6][1] * 10000) / 100,
      actual_installed: rows[6][3],
      overdue_items:    rows[6][8],
    };

    const fabrics = rows.slice(10, 17).map(r => ({
      fabric:     r[0],
      total:      r[1],
      done:       r[2],
      pct_done:   Math.round((r[3] || 0) * 10000) / 100,
      hold:       r[4],
      remaining:  r[5],
      start_date: r[6] ? new Date((r[6] - 25569) * 86400000).toISOString().slice(0,10) : null,
      end_date:   r[7] ? new Date((r[7] - 25569) * 86400000).toISOString().slice(0,10) : null,
      on_plan:    r[8],
      overdue:    r[9],
    }));

    res.json({ overall, fabrics, cached_at: new Date(cacheTime).toISOString() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/devices ──────────────────────────────────────────────────────────
app.get('/api/devices', async (req, res) => {
  try {
    const { fabric, status, location, limit = 100, offset = 0 } = req.query;
    const wb = await readWorkbook();
    const ws = wb.Sheets['All_Detail'];
    let rows = XLSX.utils.sheet_to_json(ws, { defval: null });

    rows = rows.map(r => ({
      ...r,
      'Install Date':   r['Install Date']   ? new Date((r['Install Date']   - 25569) * 86400000).toISOString().slice(0,10) : null,
      'Scheduled Date': r['Scheduled Date'] ? new Date((r['Scheduled Date'] - 25569) * 86400000).toISOString().slice(0,10) : null,
    }));

    if (fabric)   rows = rows.filter(r => r['Fabric']   === fabric);
    if (status)   rows = rows.filter(r => r['Status']   === status);
    if (location) rows = rows.filter(r => r['Location'] && r['Location'].includes(location));

    const total = rows.length;
    const data  = rows.slice(Number(offset), Number(offset) + Number(limit));
    res.json({ total, offset: Number(offset), limit: Number(limit), data });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── GET /api/filters ─────────────────────────────────────────────────────────
app.get('/api/filters', async (req, res) => {
  try {
    const wb   = await readWorkbook();
    const ws   = wb.Sheets['All_Detail'];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null });

    res.json({
      fabrics:      [...new Set(rows.map(r => r['Fabric']).filter(Boolean))].sort(),
      statuses:     [...new Set(rows.map(r => r['Status']).filter(Boolean))].sort(),
      locations:    [...new Set(rows.map(r => r['Location']).filter(Boolean))].sort(),
      device_types: [...new Set(rows.map(r => r['Device Type']).filter(Boolean))].sort(),
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── POST /api/cache/refresh ───────────────────────────────────────────────────
// Force refresh cache from SharePoint
app.post('/api/cache/refresh', async (req, res) => {
  cacheTime = 0; // expire cache
  cachedWb  = null;
  try {
    await readWorkbook();
    res.json({ success: true, cached_at: new Date(cacheTime).toISOString() });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ── GET /health ───────────────────────────────────────────────────────────────
app.get('/health', (req, res) => res.json({
  status: 'ok',
  source: 'sharepoint',
  cached_at: cacheTime ? new Date(cacheTime).toISOString() : null,
  cache_age_s: cacheTime ? Math.round((Date.now()-cacheTime)/1000) : null,
}));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SDA API (SharePoint mode) running on port ${PORT}`));
