const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

const EXCEL_PATH = path.join(__dirname, 'SAT Progress.xlsx');
const CACHE_TTL  = 5 * 60 * 1000;
let cacheTime = 0, cachedData = null;

function toDate(v) {
  if (typeof v !== 'number' || v <= 0) return null;
  const serial = v > 200000 ? v - 198327 : v;
  const d = new Date((serial - 25569) * 86400000);
  return isNaN(d) ? null : d;
}
function isoDate(d) { return d ? d.toISOString().slice(0,10) : null; }
function fmtLbl(d) {
  return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}`;
}

function parseData() {
  const wb = XLSX.readFile(EXCEL_PATH);

  // ── HKT sheet (SW + Infra) ──
  const wsHKT = wb.Sheets['HKT'];
  if (!wsHKT) throw new Error('HKT sheet not found');
  const hktRows = XLSX.utils.sheet_to_json(wsHKT, { header:1, defval:null });

  // ── HKT-WL sheet (AP) ──
  const wsWL = wb.Sheets['HKT-WL'];
  if (!wsWL) throw new Error('HKT-WL sheet not found');
  const wlRows = XLSX.utils.sheet_to_json(wsWL, { header:1, defval:null });

  // ── หา last valid row ของ WL — หยุดก่อน Summary/Progress row ──
  let wlEndIdx = wlRows.length - 1;
  for (let i = 1; i < wlRows.length; i++) {
    const c2 = wlRows[i][2] ? String(wlRows[i][2]).trim() : '';
    if (c2.includes('Summary') || c2.includes('Progress') || c2.includes('%')) {
      wlEndIdx = i - 1; break;
    }
  }

  // ── คำนวณ PROJ_START / PROJ_END จาก col T(19) ──
  let PROJ_START = null, PROJ_END = null;
  for (let i = 2; i < hktRows.length; i++) {
    const r = hktRows[i]; if (!r) continue;
    const d = toDate(r[19]);
    if (!d) continue;
    if (!PROJ_START || d < PROJ_START) PROJ_START = d;
    if (!PROJ_END   || d > PROJ_END)   PROJ_END   = d;
  }
  if (!PROJ_START) PROJ_START = new Date();
  if (!PROJ_END)   PROJ_END   = new Date(PROJ_START.getTime() + 90*86400000);

  const projStartStr = PROJ_START.toISOString().slice(0,10);
  const projEndStr   = PROJ_END.toISOString().slice(0,10);
  const projDays     = Math.round((PROJ_END - PROJ_START) / 86400000);

  // ── today ──
  const today = new Date(); today.setHours(0,0,0,0);
  const todayStr = today.toISOString().slice(0,10);
  const daysLeft = Math.max(0, Math.ceil((PROJ_END - today) / 86400000));
  const elapsed  = Math.round((today - PROJ_START) / 86400000);

  // ── daily labels (PROJ_START → PROJ_END) ──
  const dailyLabels = [];
  for (let d = new Date(PROJ_START); d <= PROJ_END; d.setDate(d.getDate()+1)) {
    dailyLabels.push(fmtLbl(d));
  }

  // ── parse HKT rows ──
  let TOTAL_SW = 0, TOTAL_INF = 0;
  let installed = 0, inProgress = 0, notStarted = 0, hold = 0;
  let instSW = 0, instInf = 0;
  let lastInstallDate = null;
  let onTimeQty = 0, earlyQty = 0, lateQty = 0;
  let curSite = null;

  const siteMap = {}, dayActMap = {}, dayPlanMap = {};
  const dayActBySite = {}, dayPlanBySite = {};
  const daySwActMap = {};
  const swInfSiteMap = {};
  const devices = [], typeMap = {};

  for (let i = 2; i < hktRows.length; i++) {
    const r = hktRows[i]; if (!r || !r.length) continue;
    if (r[0]) curSite = String(r[0]).trim();
    const device = r[3] ? String(r[3]).trim() : null;
    const qty    = typeof r[6] === 'number' ? r[6] : 0;
    const status = r[11] ? String(r[11]).trim() : '';
    let cat = r[18] ? String(r[18]).trim() : 'Infra';
    if (!['Switch','AP','Infra'].includes(cat)) cat = 'Infra';

    const helperDt = toDate(r[19]);
    let helperStr  = helperDt ? isoDate(helperDt) : null;
    if (helperStr && helperStr < projStartStr) helperStr = projStartStr;

    const instDt2 = toDate(r[20]) || toDate(r[19]);
    let instStr2  = instDt2 ? isoDate(instDt2) : null;
    if (instStr2 && instStr2 < projStartStr) instStr2 = projStartStr;

    if (!device || !curSite || qty <= 0) continue;
    const site = curSite;

    if (!siteMap[site]) siteMap[site] = {total:0, done:0, inp:0};
    if (!swInfSiteMap[site]) swInfSiteMap[site] = {sw_t:0, sw_d:0, inf_t:0, inf_d:0};

    if (cat !== 'AP') {
      if (cat === 'Switch') { TOTAL_SW += qty; swInfSiteMap[site].sw_t += qty; }
      else                  { TOTAL_INF += qty; swInfSiteMap[site].inf_t += qty; }
      siteMap[site].total += qty;

      const dev = device.length > 60 ? device.slice(0,60)+'…' : device;
      if (!typeMap[dev]) typeMap[dev] = {plan:0, done:0, cat};
      typeMap[dev].plan += qty;

      if (helperStr) {
        dayPlanMap[helperStr] = (dayPlanMap[helperStr]||0) + qty;
        if (!dayPlanBySite[site]) dayPlanBySite[site] = {total:0, byDate:{}};
        dayPlanBySite[site].total += qty;
        dayPlanBySite[site].byDate[helperStr] = (dayPlanBySite[site].byDate[helperStr]||0) + qty;
      }

      const migration = typeof r[15] === 'number' ? r[15] : 0;
      if (migration > 0) {
        installed += migration; siteMap[site].done += migration; typeMap[dev].done += migration;
        if (cat==='Switch') { instSW+=migration; swInfSiteMap[site].sw_d+=migration; }
        else                { instInf+=migration; swInfSiteMap[site].inf_d+=migration; }
        if (instStr2) {
          if (!lastInstallDate || instStr2 > lastInstallDate) lastInstallDate = instStr2;
          dayActMap[instStr2] = (dayActMap[instStr2]||0) + migration;
          daySwActMap[instStr2] = (daySwActMap[instStr2]||0) + (cat==='Switch'?migration:0);
          if (!dayActBySite[site]) dayActBySite[site] = {};
          dayActBySite[site][instStr2] = (dayActBySite[site][instStr2]||0) + migration;
        }
        // on-time
        if (instStr2 && helperStr) {
          if (instStr2 <= helperStr) { onTimeQty+=migration; if(instStr2<helperStr) earlyQty+=migration; }
          else lateQty += migration;
        }
      } else if (status.includes('Progress')) {
        inProgress += qty; siteMap[site].inp += qty;
      } else if (status.includes('Hold')) {
        hold += qty;
        devices.push({site, device:dev, qty, status:'Hold'});
      } else {
        notStarted += qty;
      }
    }
  }

  // ── AP (HKT-WL) ──
  let TOTAL_AP = 0, instAP = 0;
  const dayApActMap = {};
  let apCurSite = null;
  const apSiteMap = {};

  for (let i = 1; i <= wlEndIdx; i++) {
    const r = wlRows[i]; if (!r) continue;
    if (r[0] && typeof r[0] === 'string') apCurSite = r[0].trim();
    const qty = typeof r[3] === 'number' ? r[3] : 0;
    const mig = typeof r[9] === 'number' ? r[9] : 0;
    if (qty <= 0) continue;
    TOTAL_AP += qty;
    if (!apSiteMap[apCurSite||'']) apSiteMap[apCurSite||''] = {total:0,done:0};
    apSiteMap[apCurSite||''].total += qty;
    if (mig > 0) {
      instAP += mig;
      apSiteMap[apCurSite||''].done += mig;
    }
  }

  const TOTAL = TOTAL_SW + TOTAL_INF + TOTAL_AP;
  const swTotal = TOTAL_SW + TOTAL_INF;
  const totalInstalled = installed + instAP;
  const remaining = TOTAL - totalInstalled;
  const pctDone = TOTAL > 0 ? Math.round(totalInstalled/TOTAL*1000)/10 : 0;
  const onTimePct = installed > 0 ? Math.round(onTimeQty/installed*1000)/10 : 0;
  const overdue = 0;

  // ── daily cumulative ──
  const lastActDt = lastInstallDate ? new Date(lastInstallDate+'T00:00:00') : null;
  let cumPlan = 0, cumAct = 0;
  const dailyPlanCum = [], dailyActCum = [], bdPlan = [], bdAct = [];
  const dailySwActCum = [], dailyApActCum = [];
  let cumSwAct = 0, cumApAct = 0;

  // distribute AP plan proportional กับ SW+Infra plan ต่อวัน
  const _swInfPlanSum = Object.values(dayPlanMap).reduce((a,v)=>a+v,0)||1;
  const dayPlanFull = {}; // dayPlanMap + AP distributed
  dailyLabels.forEach(lbl=>{
    const[dd,mm]=lbl.split('/'); const k=`2026-${mm}-${dd}`;
    const swInfQty = dayPlanMap[k]||0;
    const apShare  = TOTAL_AP * (swInfQty/_swInfPlanSum);
    dayPlanFull[k] = swInfQty + apShare;
  });

  dailyLabels.forEach((lbl, di) => {
    const [dd,mm] = lbl.split('/');
    const k = `2026-${mm}-${dd}`;
    bdPlan.push(Math.round(TOTAL - cumPlan));
    bdAct.push(lastActDt && new Date(k+'T00:00:00') <= lastActDt ? TOTAL - cumAct : null);
    cumPlan += dayPlanFull[k]||0;
    cumAct  += dayActMap[k]||0;
    cumSwAct += daySwActMap[k]||0;
    cumApAct += dayApActMap[k]||0;
    dailyPlanCum.push(Math.round(Math.min(cumPlan/TOTAL,1)*10000)/100);
    const inAct = lastActDt && new Date(k+'T00:00:00') <= lastActDt;
    dailyActCum.push(inAct ? Math.round(cumAct/TOTAL*10000)/100 : null);
    dailySwActCum.push(inAct ? Math.round(cumSwAct/TOTAL*10000)/100 : null);
    dailyApActCum.push(inAct ? Math.round(cumApAct/TOTAL*10000)/100 : null);
  });

  // ── weekly ──
  const WK_MS = 7*86400000;
  const nWk = Math.ceil((PROJ_END - PROJ_START) / WK_MS) + 1;
  const swInfTotal = TOTAL_SW + TOTAL_INF;

  // คำนวณ SW+Infra plan per week ก่อน แล้วเอา AP มาเฉลี่ยตาม proportion
  const wkSwInfPlan = [];
  for (let w = 0; w < nWk; w++) {
    const ws = new Date(PROJ_START.getTime() + w*WK_MS);
    const we = new Date(ws.getTime() + 6*86400000);
    let wp = 0;
    for (let d=new Date(ws); d<=we; d.setDate(d.getDate()+1)) {
      wp += dayPlanMap[isoDate(d)]||0;
    }
    wkSwInfPlan.push(wp);
  }
  const totalSwInfPlanSum = wkSwInfPlan.reduce((a,v)=>a+v, 0) || 1;

  const wkLabels = [], planPct = [], actPct = [];
  const wkBdPlan = [], wkBdAct = [];
  let wkCumPlan = 0, wkCumAct = 0;
  for (let w = 0; w < nWk; w++) {
    const ws = new Date(PROJ_START.getTime() + w*WK_MS);
    const we = new Date(ws.getTime() + 6*86400000);
    wkLabels.push(`${fmtLbl(ws)}-${fmtLbl(we)}`);
    let wa = 0;
    for (let d=new Date(ws); d<=we; d.setDate(d.getDate()+1)) {
      wa += dayActMap[isoDate(d)]||0;
    }
    // distribute AP plan proportional กับ SW+Infra plan
    const swInfWp = wkSwInfPlan[w];
    const apWp = TOTAL_AP * (swInfWp / totalSwInfPlanSum);
    const wp = swInfWp + apWp;

    wkBdPlan.push(Math.round(TOTAL - wkCumPlan));
    wkCumPlan += wp; wkCumAct += wa;
    planPct.push(Math.round(Math.min(wkCumPlan/TOTAL,1)*10000)/100);
    const inWkAct = lastActDt && we <= lastActDt;
    actPct.push(inWkAct ? Math.round(wkCumAct/TOTAL*10000)/100 : null);
    wkBdAct.push(inWkAct ? TOTAL - wkCumAct : null);
  }
  // สร้าง bd_plan ใหม่ linear จาก TOTAL → 0
  for (let i = 0; i < wkBdPlan.length; i++) {
    wkBdPlan[i] = Math.round(TOTAL * (1 - i/(wkBdPlan.length-1)));
  }

  // ── insight ──
  const daysToFinish = elapsed > 0 ? elapsed : 1;
  const dailyRate   = Math.round(totalInstalled/daysToFinish*10)/10;
  const reqRate     = daysLeft > 0 ? Math.round(remaining/daysLeft*10)/10 : remaining;
  const needMore    = Math.round((reqRate - dailyRate)*10)/10;
  const gaugePct    = pctDone;
  const daysLate    = daysLeft < 0 ? Math.abs(daysLeft) : 0;
  const finishDateObj = dailyRate > 0 ? new Date(today.getTime() + Math.ceil(remaining/dailyRate)*86400000) : null;
  const finishDate  = finishDateObj ? finishDateObj.toISOString().slice(0,10) : null;
  // daysEarly: บวก = ก่อนกำหนด, ลบ = ช้ากว่ากำหนด
  const daysEarly   = finishDateObj ? Math.round((PROJ_END - finishDateObj) / 86400000) : 0;
  const daysLateAdj = daysEarly < 0 ? Math.abs(daysEarly) : 0;

  // ── fabrics / sites ──
  const COLORS = ['#f97316','#0ea5e9','#10b981','#a855f7','#f43f5e','#eab308','#06b6d4'];
  const fabrics = Object.entries(siteMap)
    .filter(([,v]) => v.total > 0)
    .sort((a,b) => (b[1].done/b[1].total||0) - (a[1].done/a[1].total||0))
    .map(([n,v],i) => ({
      n, t:v.total, d:v.done, p:Math.round(v.done/v.total*1000)/10,
      h:0, r:v.total-v.done, c:COLORS[i%COLORS.length], s:'–', e:'–',
      sw:{t:swInfSiteMap[n]?.sw_t||0, d:swInfSiteMap[n]?.sw_d||0},
      ap:{t:0, d:0},
      inf:{t:swInfSiteMap[n]?.inf_t||0, d:swInfSiteMap[n]?.inf_d||0},
    }));

  // per-site fab daily
  const fabDailyProg = {};
  Object.keys(swInfSiteMap).forEach(site => {
    const planByDate = (dayPlanBySite[site]||{}).byDate || {};
    const actByDate  = dayActBySite[site] || {};
    const siteTotal  = (swInfSiteMap[site].sw_t + swInfSiteMap[site].inf_t) || 1;
    const sActDates  = Object.keys(actByDate).sort();
    const sActDt     = sActDates.length ? new Date(sActDates[sActDates.length-1]+'T00:00:00') : lastActDt;
    let cPlan=0, cAct=0;
    const sw_plan=[], sw_act=[];
    dailyLabels.forEach(lbl => {
      const [dd,mm] = lbl.split('/');
      const k = `2026-${mm}-${dd}`;
      cPlan += planByDate[k]||0;
      cAct  += actByDate[k]||0;
      const inA = sActDt && new Date(k+'T00:00:00') <= sActDt;
      sw_plan.push(Math.round(Math.min(cPlan/siteTotal,1)*10000)/100);
      sw_act.push(inA ? Math.round(cAct/siteTotal*10000)/100 : null);
    });
    fabDailyProg[site] = { sw_plan, sw_act, ap_plan:[], ap_act:[] };
  });

  // ap sites
  const apSites = Object.entries(apSiteMap)
    .filter(([n,v]) => n && v.total > 0)
    .map(([n,v]) => ({ name:n, total:v.total, done:v.done, pct:Math.round(v.done/v.total*100) }));

  // sw_inf_sites
  const swInfSites = Object.entries(swInfSiteMap)
    .filter(([,v]) => v.sw_t+v.inf_t > 0)
    .map(([name,v]) => ({
      name,
      sw_t:v.sw_t, sw_d:v.sw_d,
      inf_t:v.inf_t, inf_d:v.inf_d,
      total: v.sw_t+v.inf_t,
      done:  v.sw_d+v.inf_d,
    }));

  // types
  const types = Object.entries(typeMap)
    .map(([n,v])=>({n, plan:v.plan, done:v.done}))
    .sort((a,b)=>b.plan-a.plan).slice(0,20);

  // today_wk = weeks elapsed since PROJ_START
  const todayWk = Math.max(0, Math.floor((today - PROJ_START) / (7*86400000)));
  const lastInstallDate2 = lastInstallDate;

  return {
    today_wk: todayWk,
    last_install_date: lastInstallDate2,
    meta:{
      total:TOTAL, installed:totalInstalled, in_progress:inProgress,
      not_started:notStarted, remaining, pct_done:pctDone, hold, overdue,
      installed_sw:instSW, installed_ap:instAP, installed_inf:instInf,
      sw_total:TOTAL_SW, ap_total:TOTAL_AP, inf_total:TOTAL_INF, sw_inf_total:swTotal,
      on_time_qty:onTimeQty, on_time_pct:onTimePct,
      on_time_early:earlyQty, on_time_late:lateQty,
      proj_start:projStartStr, proj_end:projEndStr,
      proj_days:projDays, days_left:daysLeft,
    },
    insight:{
      daily_rate:dailyRate, req_rate:reqRate, need_more:needMore,
      gauge_pct:gaugePct, elapsed, remaining,
      days_left:daysLeft, days_late:daysLateAdj, days_early:Math.max(0,daysEarly),
      finish_date:finishDate,
      pct_more:dailyRate>0?Math.round((reqRate/dailyRate-1)*100):0,
    },
    daily:(()=>{
      const sw=[],ap=[],inf=[],plan=[],cum_d=[],cum_sw=[],cum_ap=[],cum_inf=[];
      let cSW=0,cAP=0,cINF=0,cD=0;
      dailyLabels.forEach(lbl=>{
        const[dd,mm]=lbl.split('/'); const k=`2026-${mm}-${dd}`;
        const swV=daySwActMap[k]||0;
        const apV=dayApActMap[k]||0;
        const infV=(dayActMap[k]||0)-swV-apV;
        const planV=dayPlanMap[k]||0;
        sw.push(swV); ap.push(apV); inf.push(Math.max(0,infV)); plan.push(planV);
        cSW+=swV; cAP+=apV; cINF+=Math.max(0,infV); cD+=swV+apV+Math.max(0,infV);
        cum_sw.push(cSW); cum_ap.push(cAP); cum_inf.push(cINF); cum_d.push(cD);
      });
      return {labels:dailyLabels,sw,ap,inf,plan,cum_d,cum_sw,cum_ap,cum_inf};
    })(),
    weekly:{
      labels:wkLabels, plan_all:planPct, act_all:actPct,
      plan_sw:planPct, act_sw:actPct, plan_ap:planPct, act_ap:actPct,
      bd_plan:wkBdPlan, bd_act:wkBdAct,
    },
    daily_progress:{
      labels:dailyLabels, plan_cum:dailyPlanCum, act_cum:dailyActCum,
      sw_plan:dailyPlanCum, sw_act:dailySwActCum,
      ap_plan:dailyPlanCum, ap_act:dailyApActCum,
      bd_plan:(()=>{
        // rescale ให้จบที่ 0 พอดี
        if(!bdPlan.length) return bdPlan;
        const lastVal = bdPlan[bdPlan.length-1];
        if(lastVal===0) return bdPlan;
        // กระจาย remainder ให้ลดลงจาก lastVal→0 ใน 1 step
        bdPlan[bdPlan.length-1]=0;
        return bdPlan;
      })(),
      bd_act:bdAct,
      fab: (()=>{
        const fab={};
        Object.keys(swInfSiteMap).forEach(site=>{
          const planByDate=(dayPlanBySite[site]||{}).byDate||{};
          const actByDate=dayActBySite[site]||{};
          const siteTotal=(swInfSiteMap[site].sw_t+swInfSiteMap[site].inf_t)||1;
          const sActDates=Object.keys(actByDate).sort();
          const sActDt=sActDates.length?new Date(sActDates[sActDates.length-1]+'T00:00:00'):lastActDt;
          let cP=0,cA=0; const sp=[],sa=[];
          dailyLabels.forEach(lbl=>{
            const[dd,mm]=lbl.split('/'); const k=`2026-${mm}-${dd}`;
            cP+=planByDate[k]||0; cA+=actByDate[k]||0;
            const inA=sActDt&&new Date(k+'T00:00:00')<=sActDt;
            sp.push(Math.round(Math.min(cP/siteTotal,1)*10000)/100);
            sa.push(inA?Math.round(cA/siteTotal*10000)/100:null);
          });
          fab[site]={sw_plan:sp,sw_act:sa,ap_plan:[],ap_act:[]};
        });
        return fab;
      })(),
    },
    fab_colors:{},
    fab_plan_totals:{},
    fab_totals: Object.fromEntries(
      Object.entries(swInfSiteMap).map(([k,v])=>[k, v.sw_t+v.inf_t])
    ),
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
      // format: {siteName: {sw:[], ap:[], inf:[]}} indexed by dailyLabels
      const fd={};
      Object.keys(swInfSiteMap).forEach(site=>{
        const actByDate=dayActBySite[site]||{};
        const sw=[], ap=[], inf=[];
        dailyLabels.forEach(lbl=>{
          const[dd,mm]=lbl.split('/'); const k=`2026-${mm}-${dd}`;
          sw.push(actByDate[k]||0);
          ap.push(0);
          inf.push(0);
        });
        fd[site]={sw,ap,inf};
      });
      return fd;
    })(),
    fabrics, sites:fabrics, sw_inf_sites:swInfSites, ap_sites:apSites,
    types, hold_items:[],
  };
}

async function getData() {
  const now = Date.now();
  if (cachedData && (now-cacheTime) < CACHE_TTL) return cachedData;
  cachedData = parseData();
  cacheTime = now;
  return cachedData;
}

app.get('/api/dashboard', async (req,res) => {
  try { res.json(await getData()); }
  catch(e) { res.status(500).json({error:e.message}); }
});
app.get('/api/health', (req,res) => res.json({status:'ok'}));
app.post('/api/cache/refresh', (req,res) => {
  cacheTime=0; cachedData=null;
  try { res.json({success:true, data:parseData()}); }
  catch(e) { res.json({success:false, error:e.message}); }
});

app.use(express.static(path.join(__dirname,'../frontend')));
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'../frontend/index.html')));

const PORT = process.env.PORT||3001;
app.listen(PORT,()=>console.log(`HKT Dashboard on port ${PORT}`));
