// ==UserScript==
// @name         DPD Dispatcher – KPI Monitor (Depot flexibel)
// @namespace    bodo.dpd.custom
// @version      2.1.4
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool_dispatcher_KPI.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool_dispatcher_KPI.user.js
// @description  Dispatcher KPI Monitor
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-start
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @connect      dispatcher2-de.geopost.com
// @connect      *
// ==/UserScript==

(function(){
'use strict';
if (window.__FVKPI_RUNNING) return; window.__FVKPI_RUNNING = true;

const NS='fvkpi-';
const PANEL_ID = '#fvkpi-panel';
const RENDER_DEBOUNCE = 300;

const DEPOT_LS_KEY = 'fvkpi-depot';
const DEFAULT_DEPOT = '195';

const padLeft = (s, len, ch='0') => { s=String(s||''); return s.length>=len ? s : (ch.repeat(len-s.length)+s); };

function depotToHost(input){
  const v = String(input||'').trim();
  if(v.includes('.') && !/^\d+$/.test(v)) return v;

  const digits = v.replace(/\D/g,'');
  if(!digits) return `scanserver-d0010195.ssw.dpdit.de`;

  let d7 = '';
  if(digits.length === 7){
    d7 = digits;
  } else if(digits.length === 5){
    d7 = padLeft(digits, 7);
  } else {
    const depot3 = padLeft(digits, 3);
    const d5 = `10${depot3}`;
    d7 = padLeft(d5, 7);
  }
  return `scanserver-d${d7}.ssw.dpdit.de`;
}

function getDepotValue(){
  return localStorage.getItem(DEPOT_LS_KEY) || DEFAULT_DEPOT;
}
function getScanHost(){
  return depotToHost(getDepotValue());
}

/** ===== Dispatcher URLs ===== */
const PICKUP_DELIVERY_BASE = '/dispatcher/api/pickup-delivery';

const DEBUG = localStorage.getItem('fvkpi-debug')==='1';
const LOG=(...a)=>{ if(DEBUG) console.log('[fvkpi]',...a); };

const norm = s => String(s||'').replace(/\s+/g,' ').trim();
const pad2 = n => String(n).padStart(2,'0');
const todayDE = () => new Date().toLocaleDateString('de-DE');
const timeHM = () => { const d=new Date(); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; };

const fmtInt=v=>v==null?'—':String(Math.round(v||0)).replace(/\B(?=(\d{3})+(?!\d))/g,'.');
const fmtDec1=v=>Number.isFinite(v)?String(v.toFixed(1)).replace('.',','):'—';
const fmtPct=v=>Number.isFinite(v)?`${String(v.toFixed(1)).replace('.',',')} %`:'—';
const fmtKg=v=>Number.isFinite(v)?`${fmtDec1(v)} kg`:'—';
const fmtDurMin=m=>{
  if(!Number.isFinite(m)) return '—';
  const mm=Math.max(0, Math.round(m));
  const h=Math.floor(mm/60), r=mm%60;
  return `${h} Std ${pad2(r)} Min`;
};

const qsaMain = sel => Array.from(document.querySelectorAll(sel)).filter(el=>!el.closest(PANEL_ID));
const sleep = ms => new Promise(r=>setTimeout(r,ms));

function toast(msg, ok=true){
  const el=document.createElement('div');
  el.style.cssText='position:fixed;right:16px;bottom:16px;padding:10px 14px;border-radius:10px;font:600 13px system-ui;color:#fff;z-index:2147483647;'+(ok?'background:#16a34a':'background:#b91c1c');
  el.textContent=msg;
  document.body.appendChild(el);
  setTimeout(()=>{ el.style.opacity='0'; el.style.transition='opacity .3s'; setTimeout(()=>el.remove(),300); }, 1400);
}

/* ====== BLOCK 02/11 – Styles + UI Grundgerüst ====== */

let PANEL=null, CONTENT=null;
let renderTimer=null;
let LAST_USER_SORT_TS=0;
const SUPPRESS_RENDER_MS=4000;

let LOADING_OV=null;
let LOADING_COUNT=0;

function ensureStyles(){
  if(document.getElementById(NS+'style')) return;
  const s=document.createElement('style');
  s.id=NS+'style';
  s.textContent=`
.${NS}panel{
  position:fixed; top:48px; left:50%; transform:translateX(-50%);
  width:min(1200px,96vw); max-height:78vh; overflow:auto;
  background:#fff; border:1px solid rgba(0,0,0,.12);
  box-shadow:0 12px 28px rgba(0,0,0,.18); border-radius:12px;
  z-index:2147483646;
}
.${NS}hdr{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
.${NS}pill{display:inline-flex;gap:6px;align-items:center;flex-wrap:wrap}
.${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer;margin-left:6px}
.${NS}btn-sm[disabled]{opacity:.55;cursor:not-allowed}
.${NS}tbl{width:100%;border-collapse:collapse}
.${NS}tbl thead th{
  position:sticky;top:0;z-index:1;padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.12);
  font:700 12px system-ui;text-align:right;white-space:nowrap;background:#ffe2e2;color:#8b0000;cursor:pointer;user-select:none
}
.${NS}tbl thead th:first-child,.${NS}tbl tbody td:first-child,.${NS}tbl tfoot td:first-child{text-align:left}
.${NS}tbl tbody tr{cursor:pointer}
.${NS}tbl tbody tr:hover{background:#f8fafc}
.${NS}tbl tbody td{padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.06);font:500 12px system-ui;text-align:right;white-space:nowrap;vertical-align:top}
.${NS}tbl tfoot td{padding:8px 10px;border-top:1px solid rgba(0,0,0,.12);font:700 12px system-ui;background:#e0f2ff;color:#003366;text-align:right;white-space:nowrap}
.${NS}act{text-align:center}
.${NS}iconbtn{padding:4px 8px;border:1px solid rgba(0,0,0,.15);border-radius:999px;background:#fff;cursor:pointer}
.${NS}mini{font-size:12px;opacity:.8}
.${NS}note{opacity:.75;font-size:12px;padding:8px 12px;border-top:1px solid rgba(0,0,0,.06);background:#fafafa}
.${NS}empty{padding:12px;text-align:center;opacity:.7;display:flex;gap:10px;align-items:center;justify-content:center}
.${NS}modal{position:fixed;inset:0;background:rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;z-index:2147483647}
.${NS}modal-box{
  background:#fff;
  width:min(1650px,99vw);
  max-width:99vw;
  border-radius:12px;
  box-shadow:0 20px 40px rgba(0,0,0,.25);
  padding:14px
}
.${NS}modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}
.${NS}sort-ind{margin-left:6px;opacity:.7}
.${NS}wrap{max-height:70vh;overflow:auto;border:1px solid #e5e7eb;border-radius:8px;background:#fff}

/* Spinner / Loading */
.${NS}spinner{width:18px;height:18px;border-radius:999px;border:3px solid rgba(0,0,0,.15);border-top-color:rgba(0,0,0,.65);animation:${NS}spin .9s linear infinite;display:inline-block}
@keyframes ${NS}spin{to{transform:rotate(360deg)}}
.${NS}loading-ov{position:fixed;inset:0;background:rgba(255,255,255,.65);z-index:2147483647;display:none;align-items:center;justify-content:center}
.${NS}loading-box{background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;padding:14px 16px;display:flex;gap:10px;align-items:center;font:700 13px system-ui;color:#111827}
`;
  document.head.appendChild(s);
}

function ensureLoadingOverlay(){
  ensureStyles();
  if(LOADING_OV) return;
  LOADING_OV=document.createElement('div');
  LOADING_OV.className=NS+'loading-ov';
  LOADING_OV.innerHTML=`
    <div class="${NS}loading-box">
      <div class="${NS}spinner" aria-hidden="true"></div>
      <div class="${NS}loading-msg">Daten werden geladen…</div>
    </div>`;
  document.body.appendChild(LOADING_OV);
}

function loadingOn(msg='Daten werden geladen…'){
  ensureLoadingOverlay();
  LOADING_COUNT++;
  const m=LOADING_OV.querySelector('.'+NS+'loading-msg');
  if(m) m.textContent=msg;
  LOADING_OV.style.display='flex';
}
function loadingOff(){
  if(LOADING_COUNT>0) LOADING_COUNT--;
  if(LOADING_COUNT===0 && LOADING_OV) LOADING_OV.style.display='none';
}

function esc(s){ return String(s||'').replace(/[&<>"]/g, c=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;' }[c])); }

function mountUI(){
  if(document.querySelector(PANEL_ID)) return;

  PANEL=document.createElement('div');
  PANEL.className=NS+'panel';
  PANEL.id=PANEL_ID.slice(1);
  PANEL.style.display='none';

  PANEL.innerHTML=`
    <div class="${NS}hdr">
      <div>
        KPI Monitor – Systempartner
        <span class="${NS}mini">[Stand: ${todayDE()} ${timeHM()}]</span>
        <span class="${NS}mini" style="margin-left:10px;opacity:.85">Scan: <b id="${NS}scanHost">${esc(getScanHost())}</b></span>
      </div>

      <div class="${NS}pill" style="justify-content:flex-end">
        <label class="${NS}mini" style="display:inline-flex;gap:6px;align-items:center">
          Depot:
          <input id="${NS}depotInput" type="text"
                 style="width:90px;padding:6px 8px;border:1px solid rgba(0,0,0,.2);border-radius:8px;font:600 12px system-ui"
                 placeholder="${esc(DEFAULT_DEPOT)}"
                 value="${esc(getDepotValue())}">
        </label>
        <button class="${NS}btn-sm" data-act="save-depot" title="Depot speichern (lokal)">Speichern</button>

        <button class="${NS}btn-sm" data-act="refresh">Aktualisieren</button>
        <button class="${NS}btn-sm" data-act="copy-all">Alles kopieren</button>
        <button class="${NS}btn-sm" data-act="close">Schließen</button>
      </div>
    </div>
    <div id="${NS}content"></div>
    <div class="${NS}note" id="${NS}note">
      Daten: Fahrzeugübersicht + pickup-delivery (Mengen/PLZ) + scanserver (Gewicht/Tour).
    </div>
  `;
  document.body.appendChild(PANEL);
  CONTENT=PANEL.querySelector('#'+NS+'content');

  PANEL.addEventListener('click', async e=>{
    const b=e.target.closest('button[data-act]'); if(!b) return;

    if(b.dataset.act==='save-depot'){
      const inp = PANEL.querySelector('#'+NS+'depotInput');
      const val = String(inp?.value||'').trim();
      if(!val){
        toast('Depot leer – bitte z.B. 195 eingeben.', false);
        return;
      }
      localStorage.setItem(DEPOT_LS_KEY, val);

      WEIGHT_CACHE = null;
      AGG_CACHE = null;
      PD_CACHE = null;

      const hostEl = PANEL.querySelector('#'+NS+'scanHost');
      if(hostEl) hostEl.textContent = getScanHost();

      toast('Depot gespeichert: '+val, true);
      render(true);
      return;
    }

    if(b.dataset.act==='refresh') render(true);
    if(b.dataset.act==='copy-all') copyMainTable();
    if(b.dataset.act==='close') closePanel();
  },{passive:false});

  PANEL.addEventListener('click', async e=>{
    const cpy=e.target.closest('button[data-copy-partner]');
    if(cpy){
      e.preventDefault();
      copyPartnerRowOnly(cpy.dataset.copyPartner);
      return;
    }
    const row=e.target.closest('tr[data-partner]');
    if(row && !e.target.closest('button')){
      e.preventDefault();
      openPartnerModal(row.getAttribute('data-partner'));
      return;
    }
  },{passive:false});
}

/* ====== BLOCK 03/11 – Sortier-Helper + Modal ====== */

function makeTableSortable(root){
  const table = (root?.tagName === 'TABLE') ? root : root.querySelector('table');
  if(!table) return;
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  if(!thead || !tbody) return;

  Array.from(tbody.querySelectorAll('tr')).forEach(tr=>tr.setAttribute('data-row','1'));
  const ths = Array.from(thead.querySelectorAll('th'));

  const parseVal=(txt)=>{
    const t=String(txt||'').trim();
    const clean=t.replace(/\./g,'').replace(',', '.').replace(/[^\d.\-]/g,'').trim();
    if(clean!=='' && !isNaN(clean)) return parseFloat(clean);
    return t.toLowerCase();
  };
  const updateIndicators=(col,dir)=>{
    ths.forEach((th,i)=>{
      th.querySelector(`span.${NS}sort-ind`)?.remove();
      th.dataset.sortDir='0';
      if(i===col){
        const sp=document.createElement('span');
        sp.className=`${NS}sort-ind`;
        sp.textContent = dir>0 ? '▲' : '▼';
        th.appendChild(sp);
        th.dataset.sortDir=String(dir);
      }
    });
  };

  ths.forEach((th,colIndex)=>{
    th.addEventListener('click', ()=>{
      LAST_USER_SORT_TS=Date.now();
      const asc = (Number(th.dataset.sortDir||'0') !== 1);
      const dir = asc ? 1 : -1;

      const rows = Array.from(tbody.querySelectorAll('tr[data-row="1"]'));
      rows.sort((a,b)=>{
        const aText=a.children[colIndex]?.textContent ?? '';
        const bText=b.children[colIndex]?.textContent ?? '';
        const av=parseVal(aText), bv=parseVal(bText);
        if(typeof av==='number' && typeof bv==='number') return (av-bv)*dir;
        return aText.localeCompare(bText,'de',{numeric:true})*dir;
      });
      rows.forEach(r=>tbody.appendChild(r));
      updateIndicators(colIndex,dir);
      LAST_USER_SORT_TS=Date.now();
    },{passive:true});
  });
}

function modalCreate(html){
  const ov=document.createElement('div');
  ov.className=NS+'modal';
  ov.innerHTML=`<div class="${NS}modal-box">${html}</div>`;
  document.body.appendChild(ov);

  const box=ov.querySelector('.'+NS+'modal-box');
  ov.addEventListener('mousedown', (e)=>{
    if(!box.contains(e.target)) ov.remove();
  }, {passive:true});

  return ov;
}
function modalSet(ov, html){
  const box=ov.querySelector('.'+NS+'modal-box');
  if(box) box.innerHTML = html;
}

/* ====== BLOCK 04/11 – Fahrzeugübersicht: Grid/Tabelle lesen (Spalten nach Name) ====== */

let CACHED_COLS_OV=null;

function findColumnsOverviewDatagrid(){
  const ths=qsaMain('thead th,[role="columnheader"]');
  if(!ths.length) return null;
  if (CACHED_COLS_OV && CACHED_COLS_OV._hdrCount===ths.length) return CACHED_COLS_OV;

  const hdr = Array.from(ths).map((el,i)=>{
    const t=(el.textContent||'').replace(/\s+/g,' ').trim().toLowerCase();
    const tt=(el.querySelector('input[title], [title]')?.getAttribute('title')||'').trim().toLowerCase();
    return {i,text:t,title:tt};
  });

  const byName = (needles)=>{
    needles = Array.isArray(needles) ? needles : [needles];
    for(const h of hdr){
      const s = `${h.text} ${h.title}`;
      for(const n of needles){
        if(new RegExp(n,'i').test(s)) return h.i;
      }
    }
    return -1;
  };

  const cols = {
    systempartner: byName(['^systempartner$','systempartner']),
    tour:          byName(['^tour$','\\btour\\b']),
    tourstart:     byName(['^tourstart$','tourstart']),
    tourende:      byName(['^tourende$','tourende']),
    zusteller:     byName(['^zustellername$','zustellername','fahrername','fahrer']),
    stopps:        byName(['zustellstopps\\s*gesamt','stopps\\s*gesamt']),
    offen:         byName(['offene\\s*zustellstopps','offene\\s*stopps']),
    lieferquote:   byName(['lieferquote']),
    _hdrCount: ths.length
  };

  if(cols.systempartner<0 || cols.tour<0) return null;
  CACHED_COLS_OV=cols;
  return cols;
}

function parseIntSafe(s){
  if(s==null) return null;
  const t=String(s).replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'').trim();
  if(!t) return null;
  const v=Math.round(parseFloat(t));
  return Number.isFinite(v)?v:null;
}
function parsePctSafe(s){
  if(s==null) return null;
  const t=String(s).replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'').trim();
  if(!t) return null;
  const v=parseFloat(t);
  return Number.isFinite(v)?v:null;
}
function timeToMin(hm){
  const m=String(hm||'').match(/(\d{1,2})\s*:\s*(\d{2})/);
  if(!m) return null;
  return (+m[1])*60 + (+m[2]);
}
function diffMinWrap(startHM,endHM){
  const a=timeToMin(startHM), b=timeToMin(endHM);
  if(a==null || b==null) return null;
  let d=b-a; if(d<0) d+=24*60;
  return d;
}

function readRowsOverviewDatagrid(){
  const C=findColumnsOverviewDatagrid();
  if(!C) return {ok:false, rows:[]};

  const trs=qsaMain('tbody tr,[role="row"]').filter(tr=>{
    const cells=tr.querySelectorAll('td,[role="gridcell"]');
    const maxNeed = Math.max(C.systempartner,C.tour,C.tourstart,C.tourende,C.zusteller,C.stopps,C.offen,C.lieferquote);
    return cells && cells.length>maxNeed;
  });

  const out=[];
  for(const tr of trs){
    const tds=Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
    const partner=norm(tds[C.systempartner]?.textContent||'');
    const tour=norm(tds[C.tour]?.textContent||'');
    if(!partner || !tour) continue;

    out.push({
      partner,
      tour,
      tourstart: norm(tds[C.tourstart]?.textContent||''),
      tourende:  norm(tds[C.tourende]?.textContent||''),
      zusteller: norm((tds[C.zusteller]?.querySelector('div[title]')?.getAttribute('title'))||tds[C.zusteller]?.textContent||''),
      stopps:    parseIntSafe(tds[C.stopps]?.textContent),
      offen:     parseIntSafe(tds[C.offen]?.textContent),
      lieferquote: parsePctSafe(tds[C.lieferquote]?.textContent)
    });
  }
  return {ok:true, rows:out};
}

function readOverviewAll(){
  const tryA=readRowsOverviewDatagrid();
  if(tryA.ok && tryA.rows.length) return tryA;
  return {ok:false, rows:[]};
}

/* ====== BLOCK 05/11 – pickup-delivery: Token Capture (fetch+XHR) + Daten laden ====== */

let AUTH_BEARER = '';
let PD_CACHE = null; // {ts, items[]}
const PD_CACHE_TTL = 45_000;

function storeAuthFromHeaders(headers){
  try{
    let auth='';
    if(!headers) return;
    if(typeof headers.get==='function') auth = headers.get('Authorization') || headers.get('authorization') || '';
    else auth = headers.Authorization || headers.authorization || '';
    auth = String(auth||'').trim();
    if(auth.startsWith('Bearer ')){
      AUTH_BEARER = auth;
    }
  }catch{}
}

(function hookFetch(){
  if(window.__fvkpi_fetchHooked) return;
  window.__fvkpi_fetchHooked=true;
  const orig = window.fetch;
  window.fetch = function(input, init){
    try{ storeAuthFromHeaders(init?.headers); }catch{}
    return orig.apply(this, arguments);
  };
})();

(function hookXHR(){
  if(window.__fvkpi_xhrHooked) return;
  window.__fvkpi_xhrHooked=true;

  const origSet = XMLHttpRequest.prototype.setRequestHeader;
  XMLHttpRequest.prototype.setRequestHeader = function(k,v){
    try{
      if(String(k).toLowerCase()==='authorization'){
        storeAuthFromHeaders({ Authorization: v });
      }
    }catch{}
    return origSet.apply(this, arguments);
  };
})();

function buildPDUrl(){
  const d=new Date();
  const ds=`${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
  const u=new URL(location.origin + PICKUP_DELIVERY_BASE);
  u.searchParams.set('page','1');
  u.searchParams.set('pageSize','500');
  u.searchParams.set('dateFrom', ds);
  u.searchParams.set('dateTo', ds);
  u.searchParams.set('active','true');
  return u.toString();
}

async function loadPickupDeliveryAllPages(){
  const now=Date.now();
  if(PD_CACHE && (now-PD_CACHE.ts)<PD_CACHE_TTL) return PD_CACHE.items;

  const items=[];
  let page=1, totalCount=null;

  for(let guard=0; guard<80; guard++){
    const u=new URL(buildPDUrl());
    u.searchParams.set('page', String(page));

    const headers={};
    if(AUTH_BEARER) headers['Authorization']=AUTH_BEARER;

    const r=await fetch(u.toString(), { method:'GET', headers, credentials:'include' });
    if(!r.ok){
      toast('pickup-delivery nicht lesbar (Token fehlt). Einmal irgendeinen Dispatcher-Tab öffnen und dann KPI erneut.', false);
      break;
    }
    const j=await r.json().catch(()=>null);
    if(!j) break;

    const arr = Array.isArray(j.results) ? j.results : [];
    for(const it of arr) items.push(it);

    if(totalCount==null && typeof j.totalCount==='number') totalCount=j.totalCount;
    if(!arr.length) break;
    if(totalCount!=null && items.length>=totalCount) break;

    page++;
    await sleep(40);
  }

  PD_CACHE={ts:Date.now(), items};
  return items;
}

function plz5(s){
  const m=String(s||'').match(/(\d{5})/);
  return m ? m[1] : '';
}

function classifyItem(it){
  const pr = String(it?.priority ?? it?.priorityCode ?? '').toLowerCase();

  const scArr = Array.isArray(it?.serviceCodes) ? it.serviceCodes : [];
  const acArr = Array.isArray(it?.additionalCodes) ? it.additionalCodes : [];
  const seArr = Array.isArray(it?.serviceElements) ? it.serviceElements : [];
  const scan = String(it?.scanCode ?? it?.scan_code ?? '').toLowerCase();

  const sc = (scArr.join(',') + ',' + acArr.join(',') + ',' + seArr.join(',') + ',' +
              String(it?.serviceCode ?? it?.service ?? it?.serviceType ?? '') + ',' +
              String(it?.serviceTypes ?? '') + ',' +
              String(it?.serviceSystemId ?? '') + ',' +
              String(it?.asCodes ?? '') + ',' +
              String(it?.asCode ?? '')).toLowerCase();

  if (it?.timeCritical === true) return 'prio';
  if (pr.includes('prio') || pr.includes('time') || pr.includes('tc')) return 'prio';

  if (pr.includes('express') || pr === 'exp') return 'express';
  if (sc.includes('e12') || sc.includes('express') || sc.includes('exp12') || sc.includes('exp')) return 'express';
  if (scan.includes('e12') || scan.includes('express')) return 'express';

  return 'other';
}

function parcelsCount(it){
  const pick = [it?.realParcels,it?.estimatedParcels,it?.completeParcels,it?.parcels,it?.parcelCount]
    .find(x=>typeof x==='number' && Number.isFinite(x));
  return (typeof pick==='number' && Number.isFinite(pick)) ? pick : 0;
}

/* ====== BLOCK 06/11 – Gewicht: scanserver report_weight.cgi (Depot flexibel + https/http Fallback) ====== */

let WEIGHT_CACHE = null; // {ts, map: Map(tour->kg)}
const WEIGHT_TTL = 6*60*60*1000; // 6h

function buildWeightUrls(){
  const host = getScanHost();
  const path = '/cgi-bin/report_weight.cgi';
  return [
    `https://${host}${path}`,
    `http://${host}${path}`
  ];
}

function parseWeightsFromHtml(html){
  const doc = new DOMParser().parseFromString(html || '', 'text/html');
  const rows = Array.from(doc.querySelectorAll('table tr'));
  const map = new Map();

  for(const tr of rows){
    const tds = Array.from(tr.querySelectorAll('td')).map(td=>norm(td.textContent));
    if(tds.length < 2) continue;
    const tour = tds[0];
    const last = tds[tds.length-1];
    const kg = parseIntSafe(last);
    if(tour && kg != null) map.set(String(tour), kg);
  }
  return map;
}

function loadWeights(){
  const now=Date.now();
  if(WEIGHT_CACHE && (now-WEIGHT_CACHE.ts)<WEIGHT_TTL) return Promise.resolve(WEIGHT_CACHE.map);

  const urls = buildWeightUrls();
  const host = getScanHost();

  return new Promise((resolve)=>{
    let i = 0;

    const tryNext = () => {
      if(i >= urls.length){
        toast(`Gewicht: keine Verbindung zu ${host} (https/http). Prüfe @connect + Erreichbarkeit.`, false);
        resolve(new Map());
        return;
      }

      const url = urls[i++];
      LOG('WEIGHT_URL try', url);

      GM_xmlhttpRequest({
        method: 'GET',
        url,
        timeout: 20000,
        onload: (res) => {
          const status = res?.status || 0;
          if(status < 200 || status >= 300){
            LOG('WEIGHT_URL status', status, url);
            tryNext();
            return;
          }

          try{
            const map = parseWeightsFromHtml(res.responseText || '');
            if(map.size === 0){
              toast(`Gewicht: ${host} erreichbar, aber keine Werte gefunden.`, false);
            }
            WEIGHT_CACHE = { ts: Date.now(), map };
            resolve(map);
          }catch(e){
            console.error('[fvkpi] weight parse', e);
            tryNext();
          }
        },
        onerror: (e) => { LOG('WEIGHT_URL error', url, e); tryNext(); },
        ontimeout: () => { LOG('WEIGHT_URL timeout', url); tryNext(); }
      });
    };

    tryNext();
  });
}

/* ====== BLOCK 07/11 – Aggregation (Partner + Tour) ====== */

function groupByPartnerTour(overviewRows){
  const P=new Map();
  for(const r of overviewRows){
    const partner=r.partner;
    const tour=r.tour;
    if(!P.has(partner)) P.set(partner, new Map());
    const T=P.get(partner);
    if(!T.has(tour)){
      T.set(tour,{
        partner, tour,
        zusteller: r.zusteller||'',
        tourstart: r.tourstart||'',
        tourende:  r.tourende||'',
        tourzeitMin: diffMinWrap(r.tourstart, r.tourende),
        stopps: r.stopps||0,
        offen: r.offen||0,
        lieferquote: r.lieferquote,
        prio:0, express:0, other:0, gesamtpakete:0,
        abholstopps:0, geplAbholpakete:0,
        plzSet: new Set(),
        gewichtKg: null
      });
    }
  }
  return P;
}

function applyPickupDeliveryToMap(P, pdItems){
  const idx=new Map();
  for(const [,T] of P.entries()){
    for(const [tour,obj] of T.entries()){
      const k=String(tour);
      if(!idx.has(k)) idx.set(k, []);
      idx.get(k).push(obj);
    }
  }

  for(const it of (pdItems||[])){
    const tour = norm(it?.tour||'');
    if(!tour) continue;

    const list=idx.get(String(tour));
    if(!list || !list.length) continue;

    const postal = plz5(it?.postalCode || it?.countryPostalCode || it?.countryPostal || it?.dpdPlz || it?.pcode || '');
    const typ = classifyItem(it);
    const cnt = parcelsCount(it);
    const orderType = String(it?.orderType || it?.type || '').toUpperCase();

    for(const obj of list){
      if(orderType.includes('DELIV')){
        if(typ==='prio') obj.prio += cnt;
        else if(typ==='express') obj.express += cnt;
        else obj.other += cnt;

        obj.gesamtpakete += cnt;
        if(postal) obj.plzSet.add(postal);

      }else if(orderType.includes('PICK')){
        const key = (it?.stop!=null ? `S${it.stop}` : '') || (it?.id!=null ? `I${it.id}` : '');
        if(key){
          if(!obj.__pickupStopSet) obj.__pickupStopSet=new Set();
          if(!obj.__pickupStopSet.has(key)){
            obj.__pickupStopSet.add(key);
            obj.abholstopps += 1;
          }
        }
        obj.geplAbholpakete += cnt;
        if(postal) obj.plzSet.add(postal);
      }
    }
  }

  for(const [,T] of P.entries()){
    for(const [,obj] of T.entries()){
      delete obj.__pickupStopSet;
    }
  }
}

function applyWeightsToMap(P, weightMap){
  for(const [,T] of P.entries()){
    for(const [tour,obj] of T.entries()){
      const kg = weightMap.get(String(tour));
      if(kg!=null) obj.gewichtKg = kg;
    }
  }
}

function summarizePartner(P){
  const per=[];
  let totals={
    tours:0, avgStops:0, avgDur:0,
    prio:0, express:0, other:0, gesamtpakete:0,
    abholstopps:0, geplAbholpakete:0,
    plzSet:new Set(),
    lieferquoteAvg:null,
    gewichtSum:0,
    gewichtAvg:null
  };

  for(const [partner, T] of P.entries()){
    const toursArr=Array.from(T.values());
    const tourCount=toursArr.length;

    const sumStops=toursArr.reduce((a,x)=>a+(x.stopps||0),0);
    const avgStops=tourCount? (sumStops/tourCount) : 0;

    const durArr=toursArr.map(x=>x.tourzeitMin).filter(Number.isFinite);
    const avgDur=durArr.length ? (durArr.reduce((a,b)=>a+b,0)/durArr.length) : null;

    const prio=toursArr.reduce((a,x)=>a+(x.prio||0),0);
    const express=toursArr.reduce((a,x)=>a+(x.express||0),0);
    const other=toursArr.reduce((a,x)=>a+(x.other||0),0);
    const gesamtpakete=toursArr.reduce((a,x)=>a+(x.gesamtpakete||0),0);

    const abholstopps=toursArr.reduce((a,x)=>a+(x.abholstopps||0),0);
    const geplAbholpakete=toursArr.reduce((a,x)=>a+(x.geplAbholpakete||0),0);

    const plzSet=new Set();
    for(const x of toursArr) for(const pz of (x.plzSet||[])) plzSet.add(pz);

    const lqArr=toursArr.map(x=>x.lieferquote).filter(Number.isFinite);
    const lqAvg=lqArr.length ? (lqArr.reduce((a,b)=>a+b,0)/lqArr.length) : null;

    const gewichtSum=toursArr.reduce((a,x)=>a+(Number.isFinite(x.gewichtKg)?x.gewichtKg:0),0);
    const gewichtAvg=tourCount ? (gewichtSum/tourCount) : null;

    per.push({
      partner,
      tourCount,
      avgStops,
      avgDur,
      prio, express, other, gesamtpakete,
      abholstopps, geplAbholpakete,
      plzSet,
      plzCount: plzSet.size,
      lieferquoteAvg:lqAvg,
      gewichtSum,
      gewichtAvg
    });

    totals.tours += tourCount;
    totals.prio += prio;
    totals.express += express;
    totals.other += other;
    totals.gesamtpakete += gesamtpakete;
    totals.abholstopps += abholstopps;
    totals.geplAbholpakete += geplAbholpakete;
    totals.gewichtSum += gewichtSum;
    for(const pz of plzSet) totals.plzSet.add(pz);
  }

  per.sort((a,b)=>a.partner.localeCompare(b.partner,'de'));
  totals.avgStops = per.length ? (per.reduce((a,x)=>a+(x.avgStops||0),0) / per.length) : 0;

  const durAll = per.map(x=>x.avgDur).filter(Number.isFinite);
  totals.avgDur = durAll.length ? (durAll.reduce((a,b)=>a+b,0)/durAll.length) : null;

  const lqAll = per.map(x=>x.lieferquoteAvg).filter(Number.isFinite);
  totals.lieferquoteAvg = lqAll.length ? (lqAll.reduce((a,b)=>a+b,0)/lqAll.length) : null;

  totals.plzCount = totals.plzSet.size;
  totals.gewichtAvg = totals.tours ? (totals.gewichtSum / totals.tours) : null;

  return {per, totals};
}

/* ====== BLOCK 08/11 – HTML Builder (UI) ====== */

function buildPartnerTableHtmlUI(per, totals){
  const head=`
  <thead><tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
    <th style="text-align:left;padding:8px 10px;border-bottom:1px solid #e5e7eb;">Systempartner</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Touren</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Ø Stopps/Tour</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Ø Tourzeit</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Prio</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Express</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Other</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Gesamtpakete</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Abholstopps</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">gepl.Abholp.</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">PLZ (unique)</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Gewicht</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Lieferquote Ø</th>
    <th style="padding:8px 10px;border-bottom:1px solid #e5e7eb;">Aktion</th>
  </tr></thead>`;

  const body = per.map(p=>`
    <tr data-partner="${esc(p.partner)}" data-row="1">
      <td style="text-align:left;">${esc(p.partner)}</td>
      <td>${fmtInt(p.tourCount)}</td>
      <td>Ø ${fmtDec1(p.avgStops)}</td>
      <td>Ø ${fmtDurMin(p.avgDur)}</td>
      <td>${fmtInt(p.prio)}</td>
      <td>${fmtInt(p.express)}</td>
      <td>${fmtInt(p.other)}</td>
      <td>${fmtInt(p.gesamtpakete)}</td>
      <td>${fmtInt(p.abholstopps)}</td>
      <td>${fmtInt(p.geplAbholpakete)}</td>
      <td>${fmtInt(p.plzCount)}</td>
      <td>Ø ${fmtKg(p.gewichtAvg)}</td>
      <td>Ø ${fmtPct(p.lieferquoteAvg)}</td>
      <td class="${NS}act">
        <button class="${NS}iconbtn" title="Zeile kopieren" data-copy-partner="${esc(p.partner)}">⧉</button>
      </td>
    </tr>`).join('');

  const foot=`
  <tfoot><tr data-partner="__TOTAL__" data-row="1">
    <td>Gesamt</td>
    <td>${fmtInt(totals.tours)}</td>
    <td>Ø ${fmtDec1(totals.avgStops)}</td>
    <td>Ø ${fmtDurMin(totals.avgDur)}</td>
    <td>${fmtInt(totals.prio)}</td>
    <td>${fmtInt(totals.express)}</td>
    <td>${fmtInt(totals.other)}</td>
    <td>${fmtInt(totals.gesamtpakete)}</td>
    <td>${fmtInt(totals.abholstopps)}</td>
    <td>${fmtInt(totals.geplAbholpakete)}</td>
    <td>${fmtInt(totals.plzCount)}</td>
    <td>Ø ${fmtKg(totals.gewichtAvg)}</td>
    <td>Ø ${fmtPct(totals.lieferquoteAvg)}</td>
    <td class="${NS}act">—</td>
  </tr></tfoot>`;

  return `
  <div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
    <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayDE()} ${timeHM()}</div>
    <table class="${NS}tbl" data-kind="partner" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:13px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
      ${head}<tbody>${body}</tbody>${foot}
    </table>
  </div>`;
}

/* ====== BLOCK 08.1/11 – COPY Builder (INLINE-STYLES, ohne Aktion/Buttons) ====== */

function buildPartnerTableHtmlCOPY(per, totals){
  const C=baseCopyCssTable();
  const head = `
    <thead><tr>
      <th style="${C.thL}">Systempartner</th>
      <th style="${C.th}">Touren</th>
      <th style="${C.th}">Ø Stopps/Tour</th>
      <th style="${C.th}">Ø Tourzeit</th>
      <th style="${C.th}">Prio</th>
      <th style="${C.th}">Express</th>
      <th style="${C.th}">Other</th>
      <th style="${C.th}">Gesamtpakete</th>
      <th style="${C.th}">Abholstopps</th>
      <th style="${C.th}">gepl.Abholp.</th>
      <th style="${C.th}">PLZ (unique)</th>
      <th style="${C.th}">Gewicht</th>
      <th style="${C.th}">Lieferquote Ø</th>
    </tr></thead>`;

  const body = per.map(p=>`
    <tr>
      <td style="${C.tdL}">${esc(p.partner)}</td>
      <td style="${C.td}">${fmtInt(p.tourCount)}</td>
      <td style="${C.td}">Ø ${fmtDec1(p.avgStops)}</td>
      <td style="${C.td}">Ø ${fmtDurMin(p.avgDur)}</td>
      <td style="${C.td}">${fmtInt(p.prio)}</td>
      <td style="${C.td}">${fmtInt(p.express)}</td>
      <td style="${C.td}">${fmtInt(p.other)}</td>
      <td style="${C.td}">${fmtInt(p.gesamtpakete)}</td>
      <td style="${C.td}">${fmtInt(p.abholstopps)}</td>
      <td style="${C.td}">${fmtInt(p.geplAbholpakete)}</td>
      <td style="${C.td}">${fmtInt(p.plzCount)}</td>
      <td style="${C.td}">Ø ${fmtKg(p.gewichtAvg)}</td>
      <td style="${C.td}">Ø ${fmtPct(p.lieferquoteAvg)}</td>
    </tr>`).join('');

  const foot = `
    <tfoot><tr>
      <td style="${C.tfL}">Gesamt</td>
      <td style="${C.tf}">${fmtInt(totals.tours)}</td>
      <td style="${C.tf}">Ø ${fmtDec1(totals.avgStops)}</td>
      <td style="${C.tf}">Ø ${fmtDurMin(totals.avgDur)}</td>
      <td style="${C.tf}">${fmtInt(totals.prio)}</td>
      <td style="${C.tf}">${fmtInt(totals.express)}</td>
      <td style="${C.tf}">${fmtInt(totals.other)}</td>
      <td style="${C.tf}">${fmtInt(totals.gesamtpakete)}</td>
      <td style="${C.tf}">${fmtInt(totals.abholstopps)}</td>
      <td style="${C.tf}">${fmtInt(totals.geplAbholpakete)}</td>
      <td style="${C.tf}">${fmtInt(totals.plzCount)}</td>
      <td style="${C.tf}">Ø ${fmtKg(totals.gewichtAvg)}</td>
      <td style="${C.tf}">Ø ${fmtPct(totals.lieferquoteAvg)}</td>
    </tr></tfoot>`;

  return `
    <div style="${C.wrap}">
      <div style="${C.stand}">Stand: ${todayDE()} ${timeHM()}</div>
      <table cellpadding="0" cellspacing="0" style="${C.table}">
        ${head}<tbody>${body}</tbody>${foot}
      </table>
    </div>`;
}

/* ====== Modal Daten + COPY Builder ===== */

function summarizeToursForModal(toursArr){
  const tourCount = toursArr.length;

  const zustSet = new Set();
  for(const t of toursArr){
    const z = norm(t.zusteller||'');
    if(z) zustSet.add(z);
  }
  const zustCount = zustSet.size;

  const sumOffen = toursArr.reduce((a,x)=>a+(x.offen||0),0);
  const sumPrio = toursArr.reduce((a,x)=>a+(x.prio||0),0);
  const sumExpress = toursArr.reduce((a,x)=>a+(x.express||0),0);
  const sumOther = toursArr.reduce((a,x)=>a+(x.other||0),0);

  const sumAbholstopps = toursArr.reduce((a,x)=>a+(x.abholstopps||0),0);
  const sumGeplAbholp = toursArr.reduce((a,x)=>a+(x.geplAbholpakete||0),0);

  const sumStopps = toursArr.reduce((a,x)=>a+(x.stopps||0),0);
  const avgStopps = tourCount ? (sumStopps / tourCount) : 0;

  const sumGesamtpakete = toursArr.reduce((a,x)=>a+(x.gesamtpakete||0),0);
  const avgGesamtpakete = tourCount ? (sumGesamtpakete / tourCount) : 0;

  const durArr = toursArr.map(x=>x.tourzeitMin).filter(Number.isFinite);
  const avgDur = durArr.length ? (durArr.reduce((a,b)=>a+b,0)/durArr.length) : null;

  const plzSet = new Set();
  for(const t of toursArr){
    for(const pz of (t.plzSet||[])) plzSet.add(pz);
  }
  const plzCount = plzSet.size;

  const kgArr = toursArr.map(x=>x.gewichtKg).filter(Number.isFinite);
  const avgKg = kgArr.length ? (kgArr.reduce((a,b)=>a+b,0)/kgArr.length) : null;

  const lqArr = toursArr.map(x=>x.lieferquote).filter(Number.isFinite);
  const lqAvg = lqArr.length ? (lqArr.reduce((a,b)=>a+b,0)/lqArr.length) : null;

  return {
    tourCount,
    zustCount,
    avgStopps,
    sumOffen,
    avgDur,
    sumPrio,
    sumExpress,
    sumOther,
    avgGesamtpakete,
    sumAbholstopps,
    sumGeplAbholp,
    plzCount,
    avgKg,
    lqAvg
  };
}

function buildToursTableHtmlCOPY(toursArr, footerSum, onlyRowObj=null){
  const C=baseCopyCssTable();

  const head = `
    <thead><tr>
      <th style="${C.thL}">Tour</th>
      <th style="${C.thL}">Zusteller</th>
      <th style="${C.th}">Stopps</th>
      <th style="${C.th}">Offen</th>
      <th style="${C.th}">Tourzeit</th>
      <th style="${C.th}">Prio</th>
      <th style="${C.th}">Express</th>
      <th style="${C.th}">Other</th>
      <th style="${C.th}">Gesamtpakete</th>
      <th style="${C.th}">Abholstopps</th>
      <th style="${C.th}">gepl.Abholp.</th>
      <th style="${C.thL}">PLZ</th>
      <th style="${C.th}">Gewicht</th>
      <th style="${C.th}">Lieferquote</th>
    </tr></thead>`;

  const rows = (onlyRowObj ? [onlyRowObj] : toursArr);

  const body = rows.map(t=>{
    const plz = Array.from(t.plzSet||[]).sort().join(', ');
    return `
      <tr>
        <td style="${C.tdL}">${esc(t.tour)}</td>
        <td style="${C.tdL}">${esc(t.zusteller||'')}</td>
        <td style="${C.td}">${fmtInt(t.stopps)}</td>
        <td style="${C.td}">${fmtInt(t.offen)}</td>
        <td style="${C.td}">${fmtDurMin(t.tourzeitMin)}</td>
        <td style="${C.td}">${fmtInt(t.prio)}</td>
        <td style="${C.td}">${fmtInt(t.express)}</td>
        <td style="${C.td}">${fmtInt(t.other)}</td>
        <td style="${C.td}">${fmtInt(t.gesamtpakete)}</td>
        <td style="${C.td}">${fmtInt(t.abholstopps)}</td>
        <td style="${C.td}">${fmtInt(t.geplAbholpakete)}</td>
        <td style="${C.tdL}">${esc(plz||'')}</td>
        <td style="${C.td}">${fmtKg(t.gewichtKg)}</td>
        <td style="${C.td}">${fmtPct(t.lieferquote)}</td>
      </tr>`;
  }).join('');

  const foot = onlyRowObj ? '' : `
    <tfoot><tr>
      <td style="${C.tfL}">Gesamt (${fmtInt(footerSum.tourCount)})</td>
      <td style="${C.tfL}">${fmtInt(footerSum.zustCount)}</td>
      <td style="${C.tf}">Ø ${fmtDec1(footerSum.avgStopps)}</td>
      <td style="${C.tf}">${fmtInt(footerSum.sumOffen)}</td>
      <td style="${C.tf}">Ø ${fmtDurMin(footerSum.avgDur)}</td>
      <td style="${C.tf}">${fmtInt(footerSum.sumPrio)}</td>
      <td style="${C.tf}">${fmtInt(footerSum.sumExpress)}</td>
      <td style="${C.tf}">${fmtInt(footerSum.sumOther)}</td>
      <td style="${C.tf}">Ø ${fmtDec1(footerSum.avgGesamtpakete)}</td>
      <td style="${C.tf}">${fmtInt(footerSum.sumAbholstopps)}</td>
      <td style="${C.tf}">${fmtInt(footerSum.sumGeplAbholp)}</td>
      <td style="${C.tfL}">${fmtInt(footerSum.plzCount)}</td>
      <td style="${C.tf}">Ø ${fmtKg(footerSum.avgKg)}</td>
      <td style="${C.tf}">Ø ${fmtPct(footerSum.lqAvg)}</td>
    </tr></tfoot>`;

  return `
    <div style="${C.wrap}">
      <div style="${C.stand}">Stand: ${todayDE()} ${timeHM()}</div>
      <table cellpadding="0" cellspacing="0" style="${C.table}">
        ${head}<tbody>${body}</tbody>${foot}
      </table>
    </div>`;
}

/* ====== BLOCK 09/11 – Aggregates Master (HÄNGER-SICHER + TIMEOUTS) ====== */

let AGG_CACHE=null; // {ts, data}
const AGG_TTL=20_000;

let AGG_INFLIGHT = null;
let MODAL_BUSY = false;

function withTimeout(promise, ms, label){
  let t=null;
  const timeout = new Promise((_,rej)=>{
    t=setTimeout(()=>rej(new Error(`Timeout: ${label} (${ms}ms)`)), ms);
  });
  return Promise.race([promise, timeout]).finally(()=>{ if(t) clearTimeout(t); });
}

async function getAggregates(){
  const now=Date.now();
  if(AGG_CACHE && (now-AGG_CACHE.ts)<AGG_TTL) return AGG_CACHE.data;
  if(AGG_INFLIGHT) return AGG_INFLIGHT;

  AGG_INFLIGHT = (async () => {
    let didLoadingOn=false;
    try{
      const ov=readOverviewAll();
      if(!ov.ok || !ov.rows.length) return null;

      loadingOn('Daten werden geladen…');
      didLoadingOn=true;

      let pdItems=[];
      try{
        pdItems = await withTimeout(loadPickupDeliveryAllPages(), 25000, 'pickup-delivery');
      }catch(e){
        console.error('[fvkpi] PD timeout/error', e);
        toast('pickup-delivery: Timeout/Fehler (Console).', false);
        pdItems=[];
      }

      let wMap=new Map();
      try{
        wMap = await withTimeout(loadWeights(), 25000, 'Gewicht (scanserver)');
      }catch(e){
        console.error('[fvkpi] WEIGHT timeout/error', e);
        toast('Gewicht: Timeout/Fehler (Console).', false);
        wMap=new Map();
      }

      const P = groupByPartnerTour(ov.rows);
      if(pdItems && pdItems.length) applyPickupDeliveryToMap(P, pdItems);
      if(wMap) applyWeightsToMap(P, wMap);

      const sum = summarizePartner(P);
      const data = { P, per: sum.per, totals: sum.totals };

      AGG_CACHE={ts:Date.now(), data};
      return data;

    }catch(e){
      console.error('[fvkpi] getAggregates fatal', e);
      toast('KPI: Fehler beim Laden (Console).', false);
      return null;
    } finally {
      AGG_INFLIGHT = null;
      if(didLoadingOn) loadingOff();
    }
  })();

  try{
    return await withTimeout(AGG_INFLIGHT, 35000, 'Aggregates gesamt');
  }catch(e){
    console.error('[fvkpi] Aggregates overall timeout', e);
    toast('KPI: Laden hängt (Timeout).', false);
    try{ loadingOff(); }catch{}
    AGG_INFLIGHT = null;
    return null;
  }
}

/* ====== BLOCK 10/11 – Render + Copy + Drilldown ====== */

async function render(force=false){
  if(Date.now() - LAST_USER_SORT_TS < SUPPRESS_RENDER_MS && !force) return;
  if(renderTimer){ clearTimeout(renderTimer); renderTimer=null; }

  const run=async ()=>{
    if(!CONTENT) return;
    CONTENT.innerHTML=`<div class="${NS}empty"><span class="${NS}spinner" aria-hidden="true"></span> <span>Daten werden geladen…</span></div>`;

    try{
      const agg=await getAggregates();
      if(!agg){
        CONTENT.innerHTML=`<div class="${NS}empty">Keine Daten / Timeout / Fehler. (F12 → Console)</div>`;
        return;
      }
      CONTENT.innerHTML = buildPartnerTableHtmlUI(agg.per, agg.totals);
      makeTableSortable(CONTENT);
    }catch(e){
      console.error('[fvkpi] render error', e);
      CONTENT.innerHTML=`<div class="${NS}empty">Fehler beim Rendern. (F12 → Console)</div>`;
      toast('Render-Fehler (Console).', false);
    }
  };

  if(force) await run();
  else renderTimer=setTimeout(run, RENDER_DEBOUNCE);
}

async function copyMainTable(){
  const agg = await getAggregates();
  if(!agg){ toast('Keine Daten', false); return; }

  const html = buildPartnerTableHtmlCOPY(agg.per, agg.totals);
  const ok = await copyHtmlToClipboard(html);
  toast(ok?'Tabelle kopiert':'Kopieren fehlgeschlagen', ok);
}

async function copyPartnerRowOnly(partner){
  const agg=await getAggregates();
  if(!agg){ toast('Keine Daten', false); return; }

  if(partner==='__TOTAL__'){
    const html = buildPartnerTableHtmlCOPY(agg.per, agg.totals);
    const ok = await copyHtmlToClipboard(html);
    toast(ok?'Tabelle kopiert':'Kopieren fehlgeschlagen', ok);
    return;
  }

  const p=agg.per.find(x=>x.partner===partner);
  if(!p){ toast('Partner nicht gefunden', false); return; }

  const one = [{
    partner: p.partner,
    tourCount: p.tourCount,
    avgStops: p.avgStops,
    avgDur: p.avgDur,
    prio: p.prio,
    express: p.express,
    other: p.other,
    gesamtpakete: p.gesamtpakete,
    abholstopps: p.abholstopps,
    geplAbholpakete: p.geplAbholpakete,
    plzCount: p.plzCount,
    gewichtSum: p.gewichtSum,
    gewichtAvg: (p.tourCount ? (p.gewichtSum / p.tourCount) : null),
    lieferquoteAvg: p.lieferquoteAvg
  }];

  const totals = {
    tours: p.tourCount,
    avgStops: p.avgStops,
    avgDur: p.avgDur,
    prio: p.prio,
    express: p.express,
    other: p.other,
    gesamtpakete: p.gesamtpakete,
    abholstopps: p.abholstopps,
    geplAbholpakete: p.geplAbholpakete,
    plzCount: p.plzCount,
    gewichtSum: p.gewichtSum,
    gewichtAvg: (p.tourCount ? (p.gewichtSum / p.tourCount) : null),
    lieferquoteAvg: p.lieferquoteAvg
  };

  const html = buildPartnerTableHtmlCOPY(one, totals);
  const ok = await copyHtmlToClipboard(html);
  toast(ok?'Zeile kopiert':'Kopieren fehlgeschlagen', ok);
}

function buildModalLoadingHtml(title='Daten werden geladen…'){
  return `
    <h3 style="margin:0 0 10px 0;font:700 16px system-ui">${esc(title)}</h3>
    <div style="display:flex;gap:10px;align-items:center;padding:12px;border:1px solid #e5e7eb;border-radius:10px;background:#fff">
      <span class="${NS}spinner" aria-hidden="true"></span>
      <div style="font:700 13px system-ui;color:#111827">Daten werden geladen… bitte warten.</div>
    </div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>`;
}

async function openPartnerModal(partner){
  if(MODAL_BUSY){
    toast('Bitte warten – lädt noch…', false);
    return;
  }
  MODAL_BUSY = true;

  const title = (partner==='__TOTAL__') ? 'Touren – Gesamt' : `Touren – ${partner}`;
  const ov = modalCreate(buildModalLoadingHtml(title));

  loadingOn('Daten werden geladen…');

  try{
    const agg=await getAggregates();
    if(!agg){
      modalSet(ov, `<div class="${NS}empty">Keine Daten / Timeout.</div><div class="${NS}modal-actions"><button class="${NS}btn-sm" data-act="close">Schließen</button></div>`);
      return;
    }

    let toursArr = [];
    let displayPartner = partner;

    if (partner === '__TOTAL__') {
      for (const [,T] of agg.P.entries()){
        for (const obj of T.values()) toursArr.push(obj);
      }
      displayPartner = 'Gesamt';
    } else {
      const T = agg.P.get(partner);
      if(!T){
        modalSet(ov, `<div class="${NS}empty">Partner nicht gefunden.</div><div class="${NS}modal-actions"><button class="${NS}btn-sm" data-act="close">Schließen</button></div>`);
        return;
      }
      toursArr = Array.from(T.values());
    }

    toursArr.sort((a,b)=>String(a.tour).localeCompare(String(b.tour),'de',{numeric:true}));
    const sum = summarizeToursForModal(toursArr);

    const uiTableHead = `
      <thead><tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
        <th style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Tour</th>
        <th style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Zusteller</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Stopps</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Offen</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Tourzeit</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Prio</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Express</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Other</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Gesamtpakete</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Abholstopps</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">gepl.Abholp.</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">PLZ</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Gewicht</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Lieferquote</th>
        <th style="padding:8px;border:1px solid #e5e7eb;">Aktion</th>
      </tr></thead>`;

    const uiBody = toursArr.map(t=>{
      const plz = Array.from(t.plzSet||[]).sort().join(', ');
      const rowKey = `${t.tour}||${t.zusteller||''}`;
      return `
        <tr data-row="1">
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">${esc(t.tour)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">${esc(t.zusteller||'')}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.stopps)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.offen)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtDurMin(t.tourzeitMin)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.prio)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.express)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.other)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.gesamtpakete)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.abholstopps)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(t.geplAbholpakete)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;white-space:normal;max-width:560px;">${esc(plz||'')}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtKg(t.gewichtKg)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(t.lieferquote)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:center;">
            <button class="${NS}iconbtn" title="Zeile kopieren" data-copy-tourrow="${esc(rowKey)}">⧉</button>
          </td>
        </tr>`;
    }).join('');

    const uiFoot = `
      <tfoot>
        <tr data-row="0" style="background:#e0f2ff;color:#003366;font-weight:700;">
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt (${fmtInt(sum.tourCount)})</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">${fmtInt(sum.zustCount)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">Ø ${fmtDec1(sum.avgStopps)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(sum.sumOffen)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">Ø ${fmtDurMin(sum.avgDur)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(sum.sumPrio)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(sum.sumExpress)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(sum.sumOther)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">Ø ${fmtDec1(sum.avgGesamtpakete)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(sum.sumAbholstopps)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(sum.sumGeplAbholp)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">${fmtInt(sum.plzCount)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">Ø ${fmtKg(sum.avgKg)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">Ø ${fmtPct(sum.lqAvg)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:center;">—</td>
        </tr>
      </tfoot>`;

    modalSet(ov, `
      <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Touren – ${esc(displayPartner)} (Touren: ${fmtInt(sum.tourCount)})</h3>
      <div class="${NS}wrap">
        <table class="${NS}tbl" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;font:12px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
          ${uiTableHead}<tbody>${uiBody}</tbody>${uiFoot}
        </table>
      </div>
      <div class="${NS}modal-actions">
        <button class="${NS}btn-sm" data-act="copy-table">Tabelle kopieren</button>
        <button class="${NS}btn-sm" data-act="close">Schließen</button>
      </div>
    `);

    const table = ov.querySelector('table');
    if(table) makeTableSortable(table);

    ov.addEventListener('click', async (e)=>{
      const b=e.target.closest('button[data-act]');
      if(b){
        if(b.dataset.act==='close'){ ov.remove(); return; }
        if(b.dataset.act==='copy-table'){
          const html = buildToursTableHtmlCOPY(toursArr, sum, null);
          const ok=await copyHtmlToClipboard(html);
          toast(ok?'Tabelle kopiert':'Kopieren fehlgeschlagen', ok);
          return;
        }
      }

      const rowBtn = e.target.closest('button[data-copy-tourrow]');
      if(rowBtn){
        e.preventDefault();
        const key = rowBtn.getAttribute('data-copy-tourrow') || '';
        const [tourKey, zustKey] = key.split('||');

        const obj = toursArr.find(t => norm(t.tour)===norm(tourKey) && norm(t.zusteller||'')===norm(zustKey||''));
        if(!obj){ toast('Zeile nicht gefunden', false); return; }

        const html = buildToursTableHtmlCOPY(toursArr, sum, obj);
        const ok = await copyHtmlToClipboard(html);
        toast(ok?'Zeile kopiert':'Kopieren fehlgeschlagen', ok);
        return;
      }
    }, {passive:false});

  } finally {
    loadingOff();
    MODAL_BUSY = false;
  }
}

/* ====== BLOCK 11/11 – Boot + Loader-Button oben ====== */

function openPanel(){
  ensureStyles();
  mountUI();
  const p=document.querySelector(PANEL_ID);
  if(p) p.style.display='';
  render(true);
}
function closePanel(){
  const p=document.querySelector(PANEL_ID);
  if(p) p.style.display='none';
}

function registerWithLoader(){
  const def = {
    id: 'kpi-monitor',
    label: 'KPI Monitor',
    panels: [PANEL_ID],
    run: () => {
      const el=document.querySelector(PANEL_ID);
      if(el && getComputedStyle(el).display!=='none') closePanel();
      else openPanel();
    }
  };

  const G = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;

  if(G.TM && typeof G.TM.register==='function'){
    G.TM.register(def);
    return true;
  }

  const KEY='__tmQueue';
  G[KEY]=Array.isArray(G[KEY])?G[KEY]:[];
  G[KEY].push(def);

  let tries=0;
  const t=setInterval(()=>{
    tries++;
    if(G.TM && typeof G.TM.register==='function'){
      try{ G.TM.register(def); }catch{}
      clearInterval(t);
    }
    if(tries>80) clearInterval(t);
  }, 500);

  return false;
}

try{ registerWithLoader(); }catch(e){ console.error('[fvkpi] boot', e); }

})();
