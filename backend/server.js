const express = require('express');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');
const fs      = require('fs');

const app = express();
app.use(cors());
app.use(express.json({limit:'50mb'}));
app.use(express.raw({limit:'50mb', type:'application/octet-stream'}));

const EXCEL_PATH = path.join(__dirname, 'SAT Progress.xlsx');
const CACHE_TTL  = 5 * 60 * 1000;
let cacheTime = 0, cachedData = null;
let excelUpdatedAt = null; // timestamp ล่าสุดที่ Make.com ส่ง Excel มา
let serverStartedAt = Date.now(); // เวลา server start

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

const thMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
function fmtLbl2(isoStr) {
  if (!isoStr) return '–';
  const d = new Date(isoStr+'T00:00:00');
  return `${d.getDate()} ${thMonths[d.getMonth()]} ${(d.getFullYear()+543)%100}`;
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

  // ── คำนวณ PROJ_START / PROJ_END จาก col T(19)=เริ่ม, col V(21)=สิ้นสุด ──
  let PROJ_START = null, PROJ_END = null;
  for (let i = 2; i < hktRows.length; i++) {
    const r = hktRows[i]; if (!r) continue;
    const dStart = toDate(r[19]); // col T = วันที่เริ่ม
    const dEnd   = toDate(r[21]); // col V = วันที่สิ้นสุด
    if (dStart) {
      if (!PROJ_START || dStart < PROJ_START) PROJ_START = dStart;
      if (!PROJ_END   || dStart > PROJ_END)   PROJ_END   = dStart;
    }
    if (dEnd) {
      if (!PROJ_END   || dEnd > PROJ_END)     PROJ_END   = dEnd;
    }
  }
  if (!PROJ_START) PROJ_START = new Date();
  if (!PROJ_END)   PROJ_END   = new Date(PROJ_START.getTime() + 90*86400000);

  const projStartStr = PROJ_START.toISOString().slice(0,10);
  const projEndStr   = PROJ_END.toISOString().slice(0,10);
  const projDays     = Math.round((PROJ_END - PROJ_START) / 86400000);

  // ── today ──
  // ── pre-scan หา lastInstallDate เพื่อ freeze today ถ้างานเสร็จ ──
  let _preLastInstall = null;
  let _preTotal = 0, _preInstalled = 0;
  for (let i = 2; i < hktRows.length; i++) {
    const r = hktRows[i]; if (!r) continue;
    const _device = r[3] ? String(r[3]).trim() : null;
    const _qty = typeof r[6]==='number' ? r[6] : 0;
    const _mig = typeof r[15]==='number' ? r[15] : 0;
    if (!_device || _qty <= 0) continue;
    _preTotal += _qty; _preInstalled += _mig;
    const _instStr = r[10] ? toDate(r[10])?.toISOString().slice(0,10) : (r[20] ? toDate(r[20])?.toISOString().slice(0,10) : null);
    if (_instStr && (!_preLastInstall || _instStr > _preLastInstall)) _preLastInstall = _instStr;
  }
  const _isDoneEarly = _preTotal > 0 && _preInstalled >= _preTotal;

  // ── today (freeze ถ้างานเสร็จแล้ว) ──
  const _realToday = new Date(); _realToday.setHours(0,0,0,0);
  const today = _isDoneEarly && _preLastInstall ? new Date(_preLastInstall+'T00:00:00') : _realToday;
  today.setHours(0,0,0,0);
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

    const helperDt    = toDate(r[19]); // col T = วันที่เริ่ม
    const helperEndDt = toDate(r[21]); // col V = วันที่สิ้นสุด
    let helperStr  = helperDt ? isoDate(helperDt) : null;
    let helperEndStr = helperEndDt ? isoDate(helperEndDt) : helperStr;
    if (helperStr && helperStr < projStartStr) helperStr = projStartStr;

    const instDt2 = toDate(r[20]) || toDate(r[19]);
    let instStr2  = instDt2 ? isoDate(instDt2) : null;
    if (instStr2 && instStr2 < projStartStr) instStr2 = projStartStr;

    if (!device || !curSite || qty <= 0) continue;
    const site = curSite;

    if (!siteMap[site]) siteMap[site] = {total:0, done:0, inp:0, start:null, end:null};
    if (!swInfSiteMap[site]) swInfSiteMap[site] = {sw_t:0, sw_d:0, inf_t:0, inf_d:0};

    // track start/end date per site
    if (helperStr) {
      if (!siteMap[site].start || helperStr < siteMap[site].start) siteMap[site].start = helperStr;
      if (!siteMap[site].end   || helperStr > siteMap[site].end)   siteMap[site].end   = helperStr;
    }
    if (helperEndStr) {
      if (!siteMap[site].end   || helperEndStr > siteMap[site].end) siteMap[site].end = helperEndStr;
    }

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
        // on-time: ตรงเวลาถ้าทำเสร็จภายใน start→end window
        if (instStr2 && helperEndStr) {
          if (instStr2 <= helperEndStr) { onTimeQty+=migration; if(instStr2<helperStr) earlyQty+=migration; }
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
      // populate dayApActMap — ใช้ lastInstallDate เป็น proxy วันติดตั้ง AP
      if (lastInstallDate) {
        dayApActMap[lastInstallDate] = (dayApActMap[lastInstallDate]||0) + mig;
        dayActMap[lastInstallDate]   = (dayActMap[lastInstallDate]||0) + mig;
      }
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
    cumPlan += dayPlanFull[k]||0;
    cumAct  += dayActMap[k]||0;
    cumSwAct += daySwActMap[k]||0;
    cumApAct += dayApActMap[k]||0;
    bdAct.push(lastActDt && new Date(k+'T00:00:00') <= lastActDt ? TOTAL - cumAct : null);
    dailyPlanCum.push(Math.round(Math.min(cumPlan/TOTAL,1)*10000)/100);
    const inAct = lastActDt && new Date(k+'T00:00:00') <= lastActDt;
    dailyActCum.push(inAct ? Math.round(cumAct/TOTAL*10000)/100 : null);
    dailySwActCum.push(inAct ? Math.round(cumSwAct/TOTAL_SW*10000)/100 : null);
    dailyApActCum.push(inAct ? Math.round(cumApAct/(TOTAL_AP||1)*10000)/100 : null);
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
  // หา week ที่ lastInstallDate ตกอยู่ แล้วใส่ TOTAL-totalInstalled ที่ week นั้น
  if (lastActDt) {
    let targetIdx = -1;
    for (let w = 0; w < nWk; w++) {
      const ws2 = new Date(PROJ_START.getTime() + w*WK_MS);
      const we2 = new Date(ws2.getTime() + 6*86400000);
      if (lastActDt >= ws2 && lastActDt <= we2) { targetIdx = w; break; }
    }
    // ถ้าไม่เจอ (เกิน week สุดท้าย) ใช้ week สุดท้ายที่มีข้อมูล
    if (targetIdx < 0) targetIdx = wkBdAct.reduce((a,v,i)=>v!=null?i:a,-1);
    if (targetIdx >= 0) {
      wkBdAct[targetIdx] = TOTAL - totalInstalled;
      // week ก่อนหน้าที่ยังเป็น null ให้ carry forward ค่าก่อนหน้า
      for (let w = targetIdx-1; w >= 0; w--) {
        if (wkBdAct[w] === null) wkBdAct[w] = wkBdAct[w+1];
        else break;
      }
    }
  }
  // สร้าง bd_plan ใหม่ linear จาก TOTAL → 0
  for (let i = 0; i < wkBdPlan.length; i++) {
    wkBdPlan[i] = Math.round(TOTAL * (1 - i/(wkBdPlan.length-1)));
  }

  // ── insight ──
  const _isDone = remaining <= 0;
  const daysToFinish = elapsed > 0 ? elapsed : 1;
  const dailyRate   = Math.round(totalInstalled/daysToFinish*10)/10;
  const reqRate     = _isDone ? 0 : (daysLeft > 0 ? Math.round(remaining/daysLeft*10)/10 : remaining);
  const needMore    = Math.round((reqRate - dailyRate)*10)/10;
  const gaugePct    = _isDone ? 100 : (reqRate > 0 ? Math.min(150, Math.round(dailyRate / reqRate * 100)) : (totalInstalled > 0 ? 100 : 0));
  const daysLate    = _isDone ? 0 : (daysLeft < 0 ? Math.abs(daysLeft) : 0);
  const finishDateObj = _isDone
    ? (lastInstallDate ? new Date(lastInstallDate+'T00:00:00') : today)
    : (dailyRate > 0 ? new Date(today.getTime() + Math.ceil(remaining/dailyRate)*86400000) : null);
  const finishDate  = finishDateObj ? finishDateObj.toISOString().slice(0,10) : null;
  const daysEarly   = finishDateObj ? Math.round((PROJ_END - finishDateObj) / 86400000) : 0;
  const daysLateAdj = _isDone ? 0 : (daysEarly < 0 ? Math.abs(daysEarly) : 0);

  // ── locations: site → room ──
  const locationMap = {};
  let locSite = null;
  for (let i = 2; i < hktRows.length; i++) {
    const r = hktRows[i];
    if (!r) continue;
    if (r[0]) locSite = String(r[0]).trim();
    if (!locSite || locSite.startsWith('%')) { locSite = null; continue; }
    if (typeof r[1] !== 'string') continue; // skip rows ที่ col B ไม่ใช่ string
    const room = r[1].trim() || '(ไม่ระบุห้อง)';
    const qty  = typeof r[6]==='number' ? r[6] : 0;
    const mig  = typeof r[15]==='number' ? r[15] : 0;
    if (qty <= 0) continue;
    if (!locationMap[locSite]) locationMap[locSite] = {};
    if (!locationMap[locSite][room]) locationMap[locSite][room] = {t:0, d:0};
    locationMap[locSite][room].t += qty;
    locationMap[locSite][room].d += mig;
  }

  // ── fabrics / sites ──
  const COLORS = ['#f97316','#0ea5e9','#10b981','#a855f7','#f43f5e','#eab308','#06b6d4'];
  const fabrics = Object.entries(siteMap)
    .filter(([,v]) => v.total > 0)
    .sort((a,b) => (b[1].done/b[1].total||0) - (a[1].done/a[1].total||0))
    .map(([n,v],i) => ({
      n, t:v.total, d:v.done, p:Math.round(v.done/v.total*1000)/10,
      h:0, r:v.total-v.done, c:COLORS[i%COLORS.length],
      s: v.start ? fmtLbl2(v.start) : '–',
      e: v.end   ? fmtLbl2(v.end)   : '–',
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
    locations: Object.fromEntries(
      Object.entries(locationMap).map(([site, rooms]) => [
        site,
        Object.entries(rooms).map(([room, v]) => ({
          l: room,
          t: v.t,
          d: v.d,
          p: v.t > 0 ? Math.round(v.d/v.t*100) : 0,
        }))
      ])
    ),
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
    upcoming: (()=>{
      // แผนการติดตั้ง 2 สัปดาห์หน้า (14 วัน นับจากวันนี้)
      const up = {};
      const todayStr = today.toISOString().slice(0,10);
      const end14 = new Date(today.getTime() + 14*86400000).toISOString().slice(0,10);

      // loop ทุก row หา plan date ใน range วันนี้ - 14 วัน
      let upSite = null;
      for (let i = 2; i < hktRows.length; i++) {
        const r = hktRows[i]; if (!r) continue;
        if (r[0]) upSite = String(r[0]).trim();
        if (!upSite) continue;
        const qty  = typeof r[6]==='number' ? r[6] : 0;
        if (qty <= 0) continue;
        const cat  = r[18] ? String(r[18]).trim() : '';
        if (cat === 'AP') continue;
        const mig  = typeof r[15]==='number' ? r[15] : 0;
        const hDt  = toDate(r[19]);
        if (!hDt) continue;
        const hStr = isoDate(hDt);
        if (hStr < todayStr || hStr > end14) continue;

        const dev = r[3] ? String(r[3]).slice(0,50) : 'อุปกรณ์';
        if (!up[hStr]) up[hStr] = {};
        if (!up[hStr][upSite]) up[hStr][upSite] = {qty:0, rem:0, cats:[], types:[], locs:new Set()};
        up[hStr][upSite].qty += qty;
        up[hStr][upSite].rem += Math.max(0, qty - mig);
        if (!up[hStr][upSite].cats.includes(cat)) up[hStr][upSite].cats.push(cat);
        if (!up[hStr][upSite].types.includes(dev)) up[hStr][upSite].types.push(dev);
        if (r[1]) up[hStr][upSite].locs.add(String(r[1]).trim());
      }

      // convert Set to Array
      Object.values(up).forEach(day =>
        Object.values(day).forEach(v => { v.locs = [...v.locs]; })
      );
      return up;
    })(),
  };
}

async function getData() {
  const now = Date.now();
  if (cachedData && (now-cacheTime) < CACHE_TTL) return cachedData;
  cachedData = parseData();
  cacheTime = now;
  return cachedData;
}

// ── Webhook รับ Excel จาก Make.com ──
// รับ multipart/form-data ด้วย
const multer = require('multer');
const upload = multer({storage: multer.memoryStorage(), limits:{fileSize:50*1024*1024}});

app.post('/api/webhook/excel', upload.single('file'), (req,res)=>{
  try {
    // รับจาก multer (multipart) หรือ JSON
    if (req.file) {
      const buf = req.file.buffer;
      if (buf.length < 1000) return res.status(400).json({error:'File too small: '+buf.length});
      fs.writeFileSync(EXCEL_PATH, buf);
      cachedData = null; cacheTime = 0; excelUpdatedAt = Date.now();
      console.log('Webhook multipart: Excel updated, size='+buf.length);
      return res.json({success:true, size:buf.length, updated: new Date().toISOString()});
    }
    const body = req.body;
    // debug: ดูว่าได้รับอะไร
    const keys = Object.keys(body||{});
    const sample = keys.map(k=>({k, type:typeof body[k], len:typeof body[k]==='string'?body[k].length:'-'}));
    console.log('Webhook received:', JSON.stringify(sample));

    let buf;
    if (body && body.content) {
      buf = Buffer.from(body.content, 'base64');
    } else if (body && body.data) {
      buf = Buffer.isBuffer(body.data) ? body.data : Buffer.from(body.data, 'base64');
    } else {
      return res.status(400).json({error:'Unknown format', keys, sample});
    }
    if (buf.length < 1000) return res.status(400).json({error:'File too small: '+buf.length, keys, sample});
    fs.writeFileSync(EXCEL_PATH, buf);
    cachedData = null; cacheTime = 0; excelUpdatedAt = Date.now();
    console.log('Webhook: Excel updated, size='+buf.length);
    res.json({success:true, size:buf.length, updated: new Date().toISOString()});
  } catch(e) {
    console.error('Webhook error:', e);
    res.status(500).json({error:String(e)});
  }
});

app.get('/api/dashboard', async (req,res) => {
  try { res.json(await getData()); }
  catch(e) { res.status(500).json({error:e.message}); }
});
app.get('/api/health', (req,res) => res.json({status:'ok'}));
app.get('/api/ready', (req,res)=>{
  const now = Date.now();
  const uptime = now - serverStartedAt;
  res.json({ready: excelUpdatedAt!==null || uptime>30000, updated: excelUpdatedAt, uptime: Math.round(uptime/1000)});
});
app.post('/api/cache/refresh', (req,res) => {
  cacheTime=0; cachedData=null;
  try { res.json({success:true, data:parseData()}); }
  catch(e) { res.json({success:false, error:e.message}); }
});

app.use(express.static(path.join(__dirname,'../frontend')));
app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'../frontend/index.html')));

const PORT = process.env.PORT||3001;
app.listen(PORT,()=>{
  console.log(`HKT Dashboard on port ${PORT}`);
  // trigger Make.com ให้ส่ง Excel ใหม่ตอน server start
  const MAKE_WEBHOOK = 'https://hook.eu1.make.com/6mangbq9f8j4evhdv8252x1pjvnhlwsj';
  require('https').request(MAKE_WEBHOOK,{method:'POST'},r=>{
    console.log('Make.com startup trigger:', r.statusCode);
  }).on('error',e=>console.log('Make.com trigger err:',e.message)).end();
});
