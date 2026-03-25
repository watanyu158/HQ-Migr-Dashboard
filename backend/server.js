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

const SHAREPOINT_URL = process.env.SHAREPOINT_URL || '';
const LOCAL_EXCEL    = path.join(__dirname, 'SAT_Progress.xlsx');
const CACHE_PATH     = path.join(__dirname, 'hq_cache.xlsx');
const CACHE_TTL      = 5 * 60 * 1000;

const TOTAL     = 193;
const PROJ_START = new Date('2026-02-03');
const PROJ_END   = new Date('2026-04-03');

const WK_BOUNDS = [
  ['2026-02-02','2026-02-08'],['2026-02-09','2026-02-15'],
  ['2026-02-16','2026-02-22'],['2026-02-23','2026-03-01'],
  ['2026-03-02','2026-03-08'],['2026-03-09','2026-03-15'],
  ['2026-03-16','2026-03-22'],['2026-03-23','2026-03-29'],
  ['2026-03-30','2026-04-05'],
].map(([s,e])=>({s:new Date(s),e:new Date(e)}));
const N_WK = WK_BOUNDS.length;

let cacheTime=0, cachedData=null;

function downloadFile(url, dest) {
  return new Promise((resolve,reject)=>{
    const proto=url.startsWith('https')?https:http;
    proto.get(url,{headers:{'User-Agent':'Mozilla/5.0'}},res=>{
      if([301,302,303,307,308].includes(res.statusCode))
        return downloadFile(res.headers.location,dest).then(resolve).catch(reject);
      if(res.statusCode!==200) return reject(new Error(`HTTP ${res.statusCode}`));
      const f=fs.createWriteStream(dest);
      res.pipe(f); f.on('finish',()=>f.close(resolve)); f.on('error',reject);
    }).on('error',reject);
  });
}

async function getWorkbook() {
  if(SHAREPOINT_URL){
    try{
      console.log('Fetching from SharePoint...');
      await downloadFile(SHAREPOINT_URL,CACHE_PATH);
      console.log('SharePoint OK');
      return XLSX.readFile(CACHE_PATH);
    }catch(e){ console.warn('SharePoint failed:',e.message); }
  }
  if(fs.existsSync(LOCAL_EXCEL)){ console.log('Using local Excel'); return XLSX.readFile(LOCAL_EXCEL); }
  throw new Error('No Excel source');
}

function calcDashboard(wb) {
  const wsWk  = wb.Sheets['HQ-กราฟรายสัปดาห์'];
  const wsDay = wb.Sheets['HQ-กราฟรายวัน'];
  const wsHQ  = wb.Sheets['HQ'];
  const wkRows  = XLSX.utils.sheet_to_json(wsWk,  {header:1, defval:null});
  const dayRows = XLSX.utils.sheet_to_json(wsDay, {header:1, defval:null});
  const hqRows  = XLSX.utils.sheet_to_json(wsHQ,  {header:1, defval:null});

  // Weekly (R5-R18)
  const plan_wk  = wkRows[4].slice(1,N_WK+1).map(v=>typeof v==='number'?v:0);
  const cfg_wk   = wkRows[5].slice(1,N_WK+1).map(v=>typeof v==='number'?v:null);
  const inst_wk  = wkRows[6].slice(1,N_WK+1).map(v=>typeof v==='number'?v:null);
  const mig_wk   = wkRows[7].slice(1,N_WK+1).map(v=>typeof v==='number'?v:null);
  const mig_cum  = wkRows[8].slice(1,N_WK+1).map(v=>typeof v==='number'?v:null);
  const plan_cum_pct = wkRows[12].slice(1,N_WK+1).map(v=>typeof v==='number'?Math.round(v*10000)/100:0);
  const act_cum_pct  = wkRows[13].slice(1,N_WK+1).map(v=>typeof v==='number'?Math.round(v*10000)/100:null);
  const cfg_pct_wk   = wkRows[14].slice(1,N_WK+1).map(v=>typeof v==='number'?Math.round(v*10000)/100:null);
  const mig_pct_wk   = wkRows[16].slice(1,N_WK+1).map(v=>typeof v==='number'?Math.round(v*10000)/100:null);

  // WK labels from R4
  const wk_labels = wkRows[3].slice(1,N_WK+1).map(v=>{
    if(!v) return '';
    const s=String(v); const parts=s.split('\n');
    return parts[0]||'';
  });
  const wk_dates = wkRows[3].slice(1,N_WK+1).map(v=>{
    if(!v) return '';
    const s=String(v); const parts=s.split('\n');
    return parts[1]?parts[1].replace(/[()]/g,''):'';
  });

  // Daily
  const dayLabels=[], planDay=[], migDay=[], planCumDay=[], actCumDay=[];
  const thM=['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  for(let i=1; i<dayRows[3].length; i++){
    const raw=dayRows[3][i]; if(raw===null) continue;
    dayLabels.push(String(raw).replace('\n','/'));
    planDay.push(typeof dayRows[4][i]==='number'?dayRows[4][i]:0);
    migDay.push(typeof dayRows[5][i]==='number'?dayRows[5][i]:0);
    planCumDay.push(typeof dayRows[8][i]==='number'?Math.round(dayRows[8][i]*10000)/100:0);
    actCumDay.push(typeof dayRows[9][i]==='number'?Math.round(dayRows[9][i]*10000)/100:null);
  }

  // Summary from weekly totals
  const cfg_total  = cfg_wk.reduce((s,v)=>s+(v||0),0);
  const inst_total = inst_wk.reduce((s,v)=>s+(v||0),0);
  const mig_total  = mig_wk.reduce((s,v)=>s+(v||0),0);

  // TODAY & insight
  const today=new Date(); today.setHours(0,0,0,0);
  const elapsed  =Math.max(1,Math.floor((today-PROJ_START)/86400000)+1);
  const daysLeft =Math.max(1,Math.floor((PROJ_END-today)/86400000)+1);
  const remaining=TOTAL-mig_total;
  const dailyRate=Math.round(mig_total/elapsed*100)/100;
  const reqRate  =Math.ceil(remaining/daysLeft);
  const needMore =Math.round((reqRate-dailyRate)*100)/100;
  const pctMore  =dailyRate>0?Math.round((reqRate/dailyRate-1)*100):0;
  const daysNeeded=dailyRate>0?Math.ceil(remaining/dailyRate):9999;
  const finishDt =new Date(today); finishDt.setDate(today.getDate()+daysNeeded);
  const daysLate =Math.max(0,Math.floor((finishDt-PROJ_END)/86400000));
  const gaugePct =reqRate>0?Math.min(150,Math.round(dailyRate/reqRate*100)):100;
  const todayWk  =Math.max(0,Math.min(N_WK-1,Math.floor((elapsed-1)/7)));
  const holdCount=hqRows.slice(2).filter(r=>r[11]==='Hold').length;
  const overdueCount=hqRows.slice(2).filter(r=>{
    const inst=r[9]; return inst&&inst!=='-'&&inst!==null&&(r[11]===null||r[11]==='');
  }).length;

  // Burndown
  let s=0;
  const bd_plan=plan_wk.map(v=>TOTAL-(s+=v));
  s=0; let last=null;
  const bd_act=mig_wk.map((v,i)=>{
    if(v!==null) {s+=v; last=TOTAL-s;}
    return i<=todayWk?last:null;
  });

  // Location summary
  const locMap={};
  let curLoc=null;
  hqRows.slice(2).forEach(r=>{
    if(r[0]) curLoc=r[0];
    if(!curLoc) return;
    const qty=typeof r[5]==='number'?r[5]:0;
    if(!locMap[curLoc]) locMap[curLoc]={tor:0,mig:0};
    locMap[curLoc].tor+=qty;
    if(r[11]&&r[11]!=='-'&&typeof r[11]==='string'&&r[11].trim()!=='')
      locMap[curLoc].mig+=qty;
  });
  const locations=Object.entries(locMap)
    .filter(([,v])=>v.tor>0)
    .map(([n,v])=>({n,tor:v.tor,mig:v.mig,pct:Math.round(v.mig/v.tor*100)}))
    .sort((a,b)=>b.pct-a.pct||b.tor-a.tor);

  return {
    wk: wk_labels,
    wk_dates,
    today_wk: todayWk,
    meta: {total:TOTAL, cfg:cfg_total, inst:inst_total, mig:mig_total,
           hold:holdCount, overdue:remaining, remaining},
    insight: {daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
              pct_more:pctMore, days_late:daysLate, gauge_pct:gaugePct,
              finish_date:finishDt.toISOString().slice(0,10), days_left:daysLeft},
    weekly: {
      plan:plan_wk, cfg:cfg_wk, inst:inst_wk, mig:mig_wk, mig_cum,
      plan_cum_pct, act_cum_pct, cfg_pct_wk, mig_pct_wk,
      bd_plan, bd_act,
    },
    daily: {labels:dayLabels, plan:planDay, mig:migDay,
            plan_cum_pct:planCumDay, act_cum_pct:actCumDay},
    locations,
  };
}

async function getDashboard(force=false){
  const now=Date.now();
  if(!force&&cachedData&&(now-cacheTime)<CACHE_TTL) return cachedData;
  const wb=await getWorkbook();
  cachedData=calcDashboard(wb); cacheTime=now;
  return cachedData;
}

app.get('/api/dashboard', async(req,res)=>{
  try{ res.json(await getDashboard()); }
  catch(e){ res.status(500).json({error:e.message}); }
});

app.post('/api/cache/refresh', async(req,res)=>{
  try{ res.json({success:true, data:await getDashboard(true)}); }
  catch(e){ res.status(500).json({error:e.message}); }
});

app.get('/health', (req,res)=>res.json({
  status:'ok',
  source:SHAREPOINT_URL?'sharepoint':'local_excel',
  cached_at:cacheTime?new Date(cacheTime).toISOString():null,
}));

app.use(express.static(path.join(__dirname,'../frontend')));
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'../frontend/index.html')));

const PORT=process.env.PORT||3001;
app.listen(PORT,()=>console.log(`HQ Dashboard API running on port ${PORT}`));
