/* ====== TEIL 1/12 – Metadaten + Start IIFE (LOADER-fähig) ====== */

// ==UserScript==
// @name         DPD Dispatcher – Partner-Report Mailer
// @namespace    bodo.dpd.custom
// @version      5.4.1
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher.partner.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher.partner.user.js
// @description  ✉ je Partner mit Bestätigung + „Änderungen speichern“; Zeilenklick = Vorschau; Gesamt an „gesamt“. Lokale Empfänger (IndexedDB), Export/Import. Robust (Datagrid ODER normale Tabelle). Fix: Abholstops robust + Status-Spalte in Partnerseiten. Loader-Integration (TM).
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      10.14.7.169
// ==/UserScript==

(function(){
'use strict';
if (window.__FVPR_RUNNING) return; window.__FVPR_RUNNING = true;
// ganz oben bei den Konstanten
const ENABLE_STANDALONE = false; // <- kein eigener Button

const USE_LOADER = !!(window.TM && typeof window.TM.register === 'function'); // <— NEU
const NS='fvpr-', PANEL_ID='#fvpr-panel', OFFSET_PX=-240;
const REFRESH_MS=60_000, RENDER_DEBOUNCE=300, SCAN_MAX_STEPS=40, SCAN_STAG_LIMIT=2;
const GATEWAY_DEFAULT='http://10.14.7.169/mail.php', GATEWAY_API_KEY='fvpr-SECRET-123';

const DEBUG = localStorage.getItem('fvpr-debug')==='1';
const LOG=(...a)=>{ if(DEBUG) console.log('[fvpr]',...a); };

const norm=s=>String(s||'').replace(/\s+/g,' ').trim();
const parsePct=s=>{ if(s==null)return null; const t=String(s).replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'').trim(); if(!t)return null; const v=parseFloat(t); return Number.isFinite(v)?v:null; };
const parseIntDe=s=>{ if(s==null)return null; const t=String(s).replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'').trim(); if(!t)return null; const v=Math.round(parseFloat(t)); return Number.isFinite(v)?v:null; };
const fmtPct=v=>Number.isFinite(v)?String(v.toFixed(1)).replace('.',','):'—';
const fmtInt=v=>v==null?'—':String(Math.round(v||0)).replace(/\B(?=(\d{3})+(?!\d))/g,'.');

const todayStr=()=>new Date().toLocaleDateString('de-DE');
const pad2=n=>String(n).padStart(2,'0');
const timeHM=()=>{ const d=new Date(); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; };
const dateStamp=()=>{ const d=new Date(); return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; };

const sum = (vs,proj)=>vs.reduce((a,v)=>a+(proj(v)||0),0);
const avg = (vs,proj)=>{ const arr=vs.map(proj).filter(x=>x!=null); return arr.length?arr.reduce((a,b)=>a+b,0)/arr.length:0; };

const groupByPartner=rows=>{ const m=new Map(); for(const r of rows){ if(!m.has(r.partner)) m.set(r.partner,[]); m.get(r.partner).push(r); } return m; };

const qsaMain=sel=>Array.from(document.querySelectorAll(sel)).filter(el=>!el.closest(PANEL_ID));
function toast(msg, ok=true){ const el=document.createElement('div'); el.style.cssText='position:fixed;right:16px;bottom:16px;padding:10px 14px;border-radius:10px;font:600 13px system-ui;color:#fff;z-index:2147483647;'+(ok?'background:#16a34a':'background:#b91c1c'); el.textContent=msg; document.body.appendChild(el); setTimeout(()=>{ el.style.opacity='0'; el.style.transition='opacity .3s'; setTimeout(()=>el.remove(),300); }, 1400); }

/* ====== TEIL 2/12 – IndexedDB ====== */

const IDB_NAME='fvpr_db', IDB_VER=1;

function idbOpen(){ return new Promise((res,rej)=>{ const req=indexedDB.open(IDB_NAME,IDB_VER); req.onupgradeneeded=()=>{ const db=req.result; if(!db.objectStoreNames.contains('partners')) db.createObjectStore('partners',{keyPath:'name'}); if(!db.objectStoreNames.contains('settings')){ const s=db.createObjectStore('settings',{keyPath:'id'}); s.put({id:'global', subjectPrefix:'Aktueller Tour.Report', distTo:'', distCc:'', signature:'', httpGateway:'', apiKey:''}); } }; req.onsuccess=()=>res(req.result); req.onerror=()=>rej(req.error); }); }
async function idbGet(store,key){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readonly').objectStore(store).get(key); r.onsuccess=()=>res(r.result||null); r.onerror=()=>rej(r.error); }); }
async function idbPut(store,val){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readwrite').objectStore(store).put(val); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
async function idbDel(store,key){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readwrite').objectStore(store).delete(key); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
async function idbAll(store){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readonly').objectStore(store).getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error); }); }

async function ensureSettingsRecord(){
  const def={id:'global', subjectPrefix:'Aktueller Tour.Report', distTo:'', distCc:'', signature:'', httpGateway:GATEWAY_DEFAULT, apiKey:GATEWAY_API_KEY};
  const cur=await idbGet('settings','global');
  if(!cur){ await idbPut('settings',def); return def; }
  for(const k of Object.keys(def)) if(!(k in cur)) cur[k]=def[k];
  await idbPut('settings',cur);
  return cur;
}

async function exportDb(){
  try{
    const settings = await idbGet('settings','global').catch(()=>null)||{};
    const partners = await idbAll('partners').catch(()=>[])||[];
    const data = {
      version:'3.3.5',
      exportedAt:new Date().toISOString(),
      settings:{ subjectPrefix: typeof settings.subjectPrefix==='string'?settings.subjectPrefix:'Aktueller Tour.Report' },
      partners: partners.map(p=>({ name:String(p.name||''), to:String(p.to||''), cc:String(p.cc||''), alias:String(p.alias||'') }))
    };
    const json=JSON.stringify(data,null,2);
    const blob=new Blob([json],{type:'application/json'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=`fvpr_export_${dateStamp()}.json`; document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); },200);
    toast('Export erstellt');
  }catch(e){ console.error('[fvpr] Export',e); alert('Export fehlgeschlagen (Konsole ansehen).'); }
}

/* ====== TEIL 3/12 – Grid/Tabelle Erkennung ====== */

let GRID_VP=null, CACHED_COLS=null;

function getGridViewport(){
  if (GRID_VP && document.body.contains(GRID_VP)) return GRID_VP;
  const grid = qsaMain('[role="grid"], [data-testid*="grid"], .Datagrid__Root-sc-')[0] || document.querySelector('[role="grid"]');
  if(!grid) return null;
  GRID_VP = grid.closest('[class*="Datagrid"]') || grid;
  return GRID_VP;
}
function getAnyTable(){
  const tables = qsaMain('table');
  for (const t of tables){
    const ths = Array.from(t.querySelectorAll('thead th')).map(th=>norm(th.textContent).toLowerCase());
    if (ths.length && ths.some(x=>x.includes('systempartner'))) return t;
  }
  return null;
}
const includesAll=(s,arr)=>arr.every(w=>new RegExp(w,'i').test(s||''));

/* ====== TEIL 4/12 – Datagrid: Spalten finden ====== */

function findColumnsDatagrid(){
  const ths=qsaMain('thead th,[role="columnheader"]');
  if(!ths.length) return null;
  if (CACHED_COLS && CACHED_COLS._hdrCount===ths.length) return CACHED_COLS;

  const normTxt = el => (el?.textContent||'').replace(/\s+/g,' ').trim().toLowerCase();
  const titleOf = el => (el?.querySelector('input[title], [title]')?.getAttribute('title')||'').trim().toLowerCase();
  const H = Array.from({length: ths.length}, (_,i)=>({ i, text: normTxt(ths[i]), title: titleOf(ths[i]) }));

  const byEither = fn => { for(const h of H){ if(fn(h.text)||fn(h.title)) return h.i; } return -1; };

  function pickTourIndex(){
    const cand = H.map(h=>h.i).filter(i=>/\btour\b/.test(H[i].text) || /\btour\b/.test(H[i].title));
    if (!cand.length) return -1;
    const withNr = cand.filter(i => /\b(tour|tour\-?)\s*(nr|nummer)\b/.test(H[i].text) || /\bnr\b/.test(H[i].text) || /\bnummer\b/.test(H[i].text));
    if (withNr.length) return withNr[0];

    const rows = qsaMain('tbody tr,[role="row"]');
    const score = i => {
      let nNum=0, nTime=0, nSeen=0;
      for (const tr of rows.slice(0, 40)){
        const tds=tr.querySelectorAll('td,[role="gridcell"]');
        if (!tds || !tds[i]) continue;
        const v = (tds[i].textContent||'').trim();
        if (!v) continue;
        nSeen++;
        if (/^\d{1,2}:\d{2}$/.test(v)) nTime++;
        else if (/^\d{2,6}$/.test(v.replace(/\s+/g,''))) nNum++;
      }
      return nNum*2 - nTime*3 + nSeen*0.1;
    };
    let best=cand[0], bestScore=-1e9;
    for (const i of cand){ const s=score(i); if (s>bestScore){ best=i; bestScore=s; } }
    return best;
  }

  const sys = byEither(s=>/\bsystempartner\b/.test(s));
  const tour = pickTourIndex();
  const driver = byEither(s=>/(zustellername|fahrername|fahrer|driver)/.test(s));
  const eta = byEither(s=>/(^eta\b|eta\s*%)/.test(s));
  const status = byEither(s=>/\bstatus\b/.test(s));
  const stopsTotal = byEither(s=>includesAll(s, [/zustell?stopps?/, /gesamt/]));
  const stopsOpen  = byEither(s=>includesAll(s, [/offen/, /zustell?stopps?/]));
  const pkgsTotal  = byEither(s=>includesAll(s, [/pakete|geplante/, /(gesamt|zustell)/]));
  const obstacles  = byEither(s=>/\bzustellhindernisse\b/.test(s)||/\bhinderniss?e?\b/.test(s));
  const pickupOpen = (()=>{ const i1=byEither(s=>includesAll(s, [/offen|open/, /abhol(stopp|stopps|ung|ungen)|pickup(s)?/])); if (i1>=0) return i1; return byEither(s=>/abhol(stopp|stopps|ung|ungen)/.test(s)); })();

  const cols={sys,tour,driver,eta,status,stopsTotal,stopsOpen,pkgsTotal,obstacles,pickupOpen,_hdrCount:ths.length};
  if (cols.sys<0) return null;
  CACHED_COLS=cols;
  return cols;
}

/* ====== TEIL 5/12 – Datagrid lesen (inkl. Status-Farben) ====== */

async function readAllRowsDatagrid(maxSteps=SCAN_MAX_STEPS){
  const vp=getGridViewport();
  if(!vp) return {ok:false,rows:[]};
  if (vp.scrollHeight<=vp.clientHeight) return readRowsDatagrid();

  const seen=new Set(), acc=[];
  const sleep=ms=>new Promise(r=>setTimeout(r,ms));
  vp.scrollTop=0; let steps=0,last=-1,stag=0;

  while(steps++<maxSteps){
    const {ok,rows}=readRowsDatagrid();
    if(ok){
      for(const r of rows){
        const key=`${r.partner}||${r.tour}||${r.driver}`;
        if(!seen.has(key)){ seen.add(key); acc.push(r); }
      }
    }
    const atEnd = Math.ceil(vp.scrollTop + vp.clientHeight) >= vp.scrollHeight;
    if(acc.length===last) stag++; else { stag=0; last=acc.length; }
    if(atEnd||stag>=SCAN_STAG_LIMIT) break;
    vp.scrollTop = Math.min(vp.scrollTop + Math.max(200, vp.clientHeight*0.9), vp.scrollHeight);
    await sleep(40);
  }
  return {ok:acc.length>0, rows:acc};
}

function readRowsDatagrid(){
  const C=findColumnsDatagrid();
  if(!C) return {ok:false,rows:[]};

  const trs=qsaMain('tbody tr,[role="row"]').filter(tr=>{
    const cells=tr.querySelectorAll('td,[role="gridcell"]');
    const needed=Math.max(...Object.values(C).filter(v=>typeof v==='number'&&v>=0));
    return cells && cells.length>needed;
  });

  const out=[];
  for(const tr of trs){
    const tds=Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
    const partner=norm(tds[C.sys]?.textContent||'');
    if(!partner) continue;

    let statusText='', statusBg='', statusFg='';
    if (C.status>=0) {
      const cell=tds[C.status];
      const badge = cell?.querySelector('[title]') || cell?.querySelector('*') || cell;
      statusText = norm( cell?.getAttribute('title') || badge?.getAttribute('title') || cell?.textContent );
      try{
        const cs = window.getComputedStyle(badge);
        statusBg = cs.backgroundColor || '';
        statusFg = cs.color || '';
      }catch{}
    }

    out.push({
      partner,
      tour:    C.tour>=0?norm(tds[C.tour]?.textContent):'',
      driver:  C.driver>=0?norm((tds[C.driver]?.querySelector('div[title]')?.getAttribute('title'))||tds[C.driver]?.textContent):'',
      status:  statusText || '',
      statusBg, statusFg,
      eta:     C.eta>=0?parsePct(tds[C.eta]?.textContent):null,
      stops:   C.stopsTotal>=0?parseIntDe(tds[C.stopsTotal]?.textContent):null,
      open:    C.stopsOpen>=0?parseIntDe(tds[C.stopsOpen]?.textContent):null,
      pkgs:    C.pkgsTotal>=0?parseIntDe(tds[C.pkgsTotal]?.textContent):null,
      obstacles:C.obstacles>=0?parseIntDe(tds[C.obstacles]?.textContent):null,
      pOpen:   C.pickupOpen>=0?parseIntDe(tds[C.pickupOpen]?.textContent):null,
    });
  }
  return {ok:true,rows:out};
}

/* ====== TEIL 6/12 – Plain-Table Reader (inkl. Status-Farben) ====== */

function readRowsPlainTable(){
  const table = getAnyTable();
  if(!table) return {ok:false,rows:[]};

  const head = Array.from(table.querySelectorAll('thead th')).map(th=>norm(th.textContent).toLowerCase());
  const idx = (labelOpts)=>{ for(const l of labelOpts){ const i=head.findIndex(t=>t.includes(l)); if(i>=0) return i; } return -1; };

  function pickTourIdx(){
    let i = idx(['tour nr','tour-nr','tournr','tournummer']);
    if (i>=0) return i;
    const cand = head.map((t,ix)=>({t,ix})).filter(x=>x.t.includes('tour')).map(x=>x.ix);
    if (!cand.length) return -1;
    const rows = Array.from(table.querySelectorAll('tbody tr')).slice(0,40);
    const score = ci=>{
      let nNum=0,nTime=0,nSeen=0;
      for(const tr of rows){
        const tds = tr.querySelectorAll('td');
        if(!tds[ci]) continue;
        const v = (tds[ci].textContent||'').trim();
        if(!v) continue;
        nSeen++;
        if (/^\d{1,2}:\d{2}$/.test(v)) nTime++;
        else if (/^\d{2,6}$/.test(v.replace(/\s+/g,''))) nNum++;
      }
      return nNum*2 - nTime*3 + nSeen*0.1;
    };
    let best=cand[0], bestScore=-1e9;
    for (const ci of cand){ const s=score(ci); if (s>bestScore){ best=ci; bestScore=s; } }
    return best;
  }

  const C = {
    sys: idx(['systempartner']),
    tour: pickTourIdx(),
    driver: idx(['zustellername','fahrername','fahrer','driver']),
    status: idx(['status']),
    eta: idx(['eta']),
    stopsTotal: idx(['stopps gesamt','zustellstopps gesamt','stopp gesamt']),
    stopsOpen: idx(['offene stopps','offene zustellstopps']),
    pkgsTotal: idx(['pakete gesamt','geplante zustellpakete']),
    obstacles: idx(['zustellhindernisse','hindernis']),
    pickupOpen: (()=>{ const i1=idx(['offene abholstops','offene abholstopps','abholungen offen','open pickups','pickup open']); if(i1>=0) return i1; return idx(['abholstops','abholstopps','abholung']); })()
  };
  if (C.sys<0) return {ok:false,rows:[]};

  const trs = Array.from(table.querySelectorAll('tbody tr'));
  const out=[];
  for(const tr of trs){
    const tds = Array.from(tr.querySelectorAll('td'));
    const partner = norm(tds[C.sys]?.textContent||'');
    if(!partner) continue;

    let statusBg='', statusFg='';
    if (C.status>=0){
      const cell=tds[C.status];
      const badge = cell?.querySelector('[title]') || cell?.querySelector('*') || cell;
      try{
        const cs=getComputedStyle(badge);
        statusBg=cs.backgroundColor||'';
        statusFg=cs.color||'';
      }catch{}
    }

    out.push({
      partner,
      tour:    C.tour>=0?norm(tds[C.tour]?.textContent):'',
      driver:  C.driver>=0?norm(tds[C.driver]?.textContent):'',
      status:  C.status>=0?norm(tds[C.status]?.textContent):'',
      statusBg, statusFg,
      eta:     C.eta>=0?parsePct(tds[C.eta]?.textContent):null,
      stops:   C.stopsTotal>=0?parseIntDe(tds[C.stopsTotal]?.textContent):null,
      open:    C.stopsOpen>=0?parseIntDe(tds[C.stopsOpen]?.textContent):null,
      pkgs:    C.pkgsTotal>=0?parseIntDe(tds[C.pkgsTotal]?.textContent):null,
      obstacles:C.obstacles>=0?parseIntDe(tds[C.obstacles]?.textContent):null,
      pOpen:   C.pickupOpen>=0?parseIntDe(tds[C.pickupOpen]?.textContent):null,
    });
  }
  return {ok:out.length>0, rows:out};
}

async function readAllRows(){
  const tryA = await readAllRowsDatagrid();
  if (tryA.ok && tryA.rows.length) return tryA;
  const tryB = readRowsPlainTable();
  if (tryB.ok && tryB.rows.length) return tryB;
  return {ok:false, rows:[]};
}

/* ====== TEIL 7/12 – Styles ====== */

function ensureStyles(){
  if(document.getElementById(NS+'style')) return;
  const s=document.createElement('style');
  s.id=NS+'style';
  s.textContent= `
.${NS}wrap{position:fixed;top:8px;left:calc(50% + ${OFFSET_PX}px);display:flex;gap:8px;z-index:2147483647}
.${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:8px 14px;border-radius:999px;font:600 13px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
.${NS}panel{position:fixed;top:48px;left:calc(50% + ${OFFSET_PX}px);width:min(1150px,96vw);max-height:76vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:2147483646}
.${NS}hdr{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
.${NS}pill{display:inline-flex;gap:6px;align-items:center;flex-wrap:wrap}
.${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer;margin-left:6px}
.${NS}tbl{width:100%;border-collapse:collapse}
.${NS}tbl thead th{position:sticky;top:0;z-index:1;padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.12);font:700 12px system-ui;text-align:right;white-space:nowrap;background:#ffe2e2;color:#8b0000;cursor:pointer;user-select:none}
.${NS}tbl thead th:first-child,.${NS}tbl tbody td:first-child{text-align:left}
.${NS}tbl tbody tr{cursor:pointer}
.${NS}tbl tbody tr:hover{background:#f8fafc}
.${NS}tbl tbody td{padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.06);font:500 12px system-ui;text-align:right;white-space:nowrap}
.${NS}tbl tbody td.${NS}act{text-align:center}
.${NS}tbl tfoot td{padding:8px 10px;border-top:1px solid rgba(0,0,0,.12);font:700 12px system-ui;background:#e0f2ff;color:#003366;text-align:right;white-space:nowrap}
.${NS}empty{padding:12px;text-align:center;opacity:.7}
.${NS}cfg{padding:10px;border-top:1px solid rgba(0,0,0,.06);background:#fafafa}
.${NS}cfg input{width:100%;padding:6px;border:1px solid #e5e7eb;border-radius:8px}
.${NS}row{display:grid;grid-template-columns:1fr 2fr;gap:10px;margin:6px 0}
.${NS}modal{position:fixed;inset:0;background:rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;z-index:2147483647}
.${NS}modal-box{background:#fff;min-width:min(560px,96vw);max-width:96vw;border-radius:12px;box-shadow:0 20px 40px rgba(0,0,0,.25);padding:14px}
.${NS}modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}
.${NS}iconbtn{padding:4px 8px;border:1px solid rgba(0,0,0,.15);border-radius:999px;background:#fff;cursor:pointer}
.${NS}mini{font-size:12px;opacity:.8}
.${NS}note{opacity:.75;font-size:12px}
.${NS}sort-ind{margin-left:6px;opacity:.7}
`;
  document.head.appendChild(s);
}

/* ====== TEIL 8/12 – UI (Panel/Buttons/Import-Export) + Sortier-Helper ====== */
let PANEL, CONTENT, CFGBOX;

// Unterdrückung von Auto-Render nach Nutzersortierung (ms)
let LAST_USER_SORT_TS = 0;
const SUPPRESS_RENDER_MS = 4000;

function makeTableSortable(root){
  const table = (root instanceof HTMLElement) ? root.querySelector('table') : root;
  if(!table) return;
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  if(!thead || !tbody) return;

  Array.from(tbody.querySelectorAll('tr')).forEach(tr=>{ tr.setAttribute('data-row','1'); });

  const ths = Array.from(thead.querySelectorAll('th'));

  const parseVal=(txt)=>{ const t = String(txt||'').trim(); const clean = t.replace(/\./g,'').replace(',', '.').replace(/[^\d.\-]/g,'').trim(); if(clean!=='' && !isNaN(clean)) return parseFloat(clean); return t.toLowerCase(); };

  const updateIndicators=(col,dir)=>{ ths.forEach((th,i)=>{ th.querySelector(`span.${NS}sort-ind`)?.remove(); th.dataset.sortDir = '0'; if(i===col){ const sp=document.createElement('span'); sp.className=`${NS}sort-ind`; sp.textContent = dir>0 ? '▲' : '▼'; th.appendChild(sp); th.dataset.sortDir = String(dir); } }); };

  ths.forEach((th, colIndex)=>{ th.style.cursor='pointer'; th.addEventListener('click', ()=>{ LAST_USER_SORT_TS = Date.now(); const asc = (Number(th.dataset.sortDir||'0') !== 1); const dir = asc ? 1 : -1; const rows = Array.from(tbody.querySelectorAll('tr[data-row="1"]')); rows.sort((a,b)=>{ const aText = a.children[colIndex]?.textContent ?? ''; const bText = b.children[colIndex]?.textContent ?? ''; const aVal = parseVal(aText); const bVal = parseVal(bText); if(typeof aVal==='number' && typeof bVal==='number'){ return (aVal - bVal)*dir; } return aText.localeCompare(bText,'de',{numeric:true})*dir; }); rows.forEach(r=>tbody.appendChild(r)); updateIndicators(colIndex, dir); LAST_USER_SORT_TS = Date.now(); },{passive:true}); });
}

function mountUI(forLoader=false){
  if (document.getElementById(NS+'wrap') && document.getElementById(NS+'panel')) return;

  // — PANEL immer erstellen (für Loader geschlossen starten)
  PANEL=document.createElement('div');
  PANEL.className=NS+'panel';
  PANEL.style.display = forLoader ? 'none' : '';
  PANEL.id=PANEL_ID.slice(1);
  PANEL.innerHTML= `
    <div class="${NS}hdr">
      <div>Auswertung – Systempartner (Fahrzeugübersicht) <span class="${NS}mini">[Stand: ${todayStr()} ${timeHM()}]</span></div>
      <div class="${NS}pill">
        <button class="${NS}btn-sm" data-act="refresh">Aktualisieren</button>
        <span class="${NS}btn-sm" style="opacity:.0;cursor:default"></span>
        <button class="${NS}btn-sm" data-act="send-partner-and-total">✉ Pro Partner und gesamt an uns</button>
        <button class="${NS}btn-sm" data-act="send-total-only">✉ Gesamt</button>
        <button class="${NS}btn-sm" data-act="settings">Einstellungen</button>
      </div>
    </div>
    <div id="${NS}content"></div>
    <div class="${NS}cfg" style="display:none" id="${NS}cfgbox">
      <h4 style="margin:0 0 6px 0;font:700 14px system-ui">Globale Einstellungen</h4>
      <div class="${NS}row"><label>Betreff-Prefix</label><input id="${NS}cfg-subj" type="text"></div>
      <button class="${NS}btn-sm" data-act="cfg-save">Speichern</button>
      <button class="${NS}btn-sm" data-act="cfg-hide">Schließen</button>
      <button class="${NS}btn-sm" data-act="export">Export</button>
      <button class="${NS}btn-sm" data-act="import">Import</button>
    </div>
    <div class="${NS}note">Pro Partner wird nur versendet, wenn im Partner-Eintrag eine gültige Adresse hinterlegt ist. Die Gesamt-Mail geht an den Eintrag mit Name = "gesamt".</div>
    <input type="file" id="${NS}impfile" accept="application/json" style="display:none">
  `;
  document.body.appendChild(PANEL);
  CONTENT=PANEL.querySelector('#'+NS+'content');
  CFGBOX=PANEL.querySelector('#'+NS+'cfgbox');

  // — Standalone-Fallback: eigenen Button nur erstellen, wenn KEIN Loader aktiv
  if (!forLoader && ENABLE_STANDALONE){
    if (!document.getElementById(NS+'wrap')){
      const wrap=document.createElement('div'); wrap.id=NS+'wrap'; wrap.className=NS+'wrap';
      const btn=document.createElement('button'); btn.className=NS+'btn'; btn.textContent='Partner-Report';
      wrap.append(btn); document.body.appendChild(wrap);
      btn.addEventListener('click',()=>{ const will=PANEL.style.display==='none'; PANEL.style.display=will?'':'none'; if(will) fillCfg(); },{passive:true});
    }
  }

  // Actions
  PANEL.addEventListener('click', async e=>{
    const b=e.target.closest('button[data-act]'); if(!b) return;
    if(b.dataset.act==='refresh') render(true);
    if(b.dataset.act==='send-partner-and-total') await sendPartnerAndTotalConfirm();
    if(b.dataset.act==='send-total-only') await sendTotalOnlyConfirm();
    if(b.dataset.act==='settings'){ CFGBOX.style.display=CFGBOX.style.display==='none'?'':''; if(CFGBOX.style.display!=='none') await fillCfg(); }
    if(b.dataset.act==='cfg-hide') CFGBOX.style.display='none';
    if(b.dataset.act==='cfg-save') await saveCfgFromUI();
    if(b.dataset.act==='export') await exportDb();
    if(b.dataset.act==='import') document.getElementById(NS+'impfile').click();
  },{passive:false});

  PANEL.querySelector('#'+NS+'impfile').addEventListener('change', async (ev)=>{
    const f=ev.target.files?.[0]; if(!f) return;
    try{
      const data=JSON.parse(await f.text());
      if(data.settings) await idbPut('settings',{id:'global',...data.settings});
      if(Array.isArray(data.partners)){
        const db=await idbOpen();
        const tx=db.transaction('partners','readwrite');
        const st=tx.objectStore('partners');
        const keys=await new Promise(r=>{ const k=st.getAllKeys(); k.onsuccess=()=>r(k.result||[]); });
        await Promise.all(keys.map(k=>new Promise(r=>{ const d=st.delete(k); d.onsuccess=r; })));
        for(const p of data.partners) st.put(p);
        await new Promise(r=>{ tx.oncomplete=r; });
      }
      alert('Import erfolgreich.');
      await fillCfg();
      render(true);
    }catch(e){ console.error(e); alert('Import fehlgeschlagen (ungültiges JSON).'); }
  },{passive:true});

  // Delegation: ✉ / 👁 / ⧉ in Zeilen + Gesamtleiste
  PANEL.addEventListener('click', e=>{
    const sendBtn=e.target.closest('button[data-sp]'); if(sendBtn){ e.preventDefault(); sendSinglePartnerConfirm(sendBtn.dataset.sp); return; }
    const eyeBtn=e.target.closest('button[data-eye]'); if(eyeBtn){ e.preventDefault(); openPreview(eyeBtn.dataset.eye); return; }
    const copyBtn=e.target.closest('button[data-copy]'); if(copyBtn){ e.preventDefault(); copyPartnerHtml(copyBtn.dataset.copy); return; }
    const totalEye=e.target.closest('button[data-total-eye]'); if(totalEye){ e.preventDefault(); openTotalPreview(); return; }
    const totalCopy=e.target.closest('button[data-total-copy]'); if(totalCopy){ e.preventDefault(); copyTotalHtml(); return; }
    const totalBtn=e.target.closest('button[data-total-send]'); if(totalBtn){ e.preventDefault(); sendTotalOnlyConfirm(); return; }
    const row=e.target.closest('tr[data-partner]'); if(row && !e.target.closest('button')){ e.preventDefault(); openPreview(row.getAttribute('data-partner')); }
  },{passive:false});
}

/* ====== TEIL 9/12 – Modale, Vorschau, Partner-Dialog ====== */
function removeStandaloneButton(){
  const wrap = document.getElementById(NS+'wrap');
  if (wrap) wrap.remove();
}

function modal(html){ const ov=document.createElement('div'); ov.className=NS+'modal'; ov.innerHTML=`<div class="${NS}modal-box">${html}</div>`; document.body.appendChild(ov); return ov; }
function softenColor(rgb, alpha=0.18){ if(!rgb) return ''; const m = rgb.match(/rgba?\s*\(\s*(\d+)[, ]\s*(\d+)[, ]\s*(\d+)/i); if(!m) return ''; const [_,r,g,b]=m; return `rgba(${r},${g},${b},${alpha})`; }
function etaBg(v){ if(v==null) return ''; if(v>=100) return 'rgba(22,163,74,0.18)'; if(v>=94) return 'rgba(202,138,4,0.18)'; return 'rgba(185,28,28,0.18)'; }

/* (… ALLES aus deinem Original ab hier unverändert …)
   — openPreview, openConfirm, openPartnerDialog, Mail/Clipboard/Gateway,
   — HTML-Builder (partnerHtml, summaryHtml, mailPartnerHtml, mailSummaryHtml),
   — getAggregates, sendSinglePartnerConfirm, sendTotalOnlyConfirm, sendPartnerAndTotalConfirm
   — render()
   (Ich lasse die Funktionskörper unverändert, nur der Platz ist hier knapp.)
   ↓ Die Funktionen stehen vollständig in deinem ursprünglichen Codeblock – ich kürze nur diese Erklärung.
*/

/* ====== TEIL 9/12 – Modale, Vorschau, Partner-Dialog (komplett) ====== */
function removeStandaloneButton(){
  const wrap = document.getElementById(NS+'wrap');
  if (wrap) wrap.remove();
}
function modal(html){
  const ov=document.createElement('div');
  ov.className=NS+'modal';
  ov.innerHTML=`<div class="${NS}modal-box">${html}</div>`;
  document.body.appendChild(ov);
  return ov;
}
function softenColor(rgb, alpha=0.18){
  if(!rgb) return '';
  const m = rgb.match(/rgba?\s*\(\s*(\d+)[, ]\s*(\d+)[, ]\s*(\d+)/i);
  if(!m) return '';
  const [_,r,g,b]=m;
  return `rgba(${r},${g},${b},${alpha})`;
}
function etaBg(v){
  if(v==null) return '';
  if(v>=100) return 'rgba(22,163,74,0.18)';
  if(v>=94)  return 'rgba(202,138,4,0.18)';
  return 'rgba(185,28,28,0.18)';
}

/* ====== Mail/Clipboard/Helpers ====== */
function splitEmails(raw){ return (raw||'').split(/[,;\s]+/).map(s=>s.trim()).filter(Boolean); }
function isEmail(s){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }
function normalizeEmailList(raw){
  const arr=splitEmails(raw); const valid=[],invalid=[]; const seen=new Set();
  for(const a of arr){ const low=a.toLowerCase(); if(seen.has(low)) continue; seen.add(low); (isEmail(a)?valid:invalid).push(a); }
  return {valid,invalid};
}
async function copyHtmlToClipboard(html){
  try{
    if(navigator.clipboard&&window.ClipboardItem){
      const item=new ClipboardItem({'text/html':new Blob([html],{type:'text/html'})});
      await navigator.clipboard.write([item]);
    } else {
      const d=document.createElement('div'); d.style.position='fixed'; d.style.left='-99999px'; d.innerHTML=html; document.body.appendChild(d);
      const r=document.createRange(); r.selectNodeContents(d);
      const sel=window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('copy'); sel.removeAllRanges(); d.remove();
    }
    return true;
  }catch(e){ console.error(e); return false; }
}
function openMailto(subject,to='',cc=''){
  const href=`mailto:${encodeURIComponent(to||'')}?subject=${encodeURIComponent(subject)}${cc?`&cc=${encodeURIComponent(cc)}`:''}`;
  window.open(href,'_blank');
}
async function deliverMail({subject, html, to, cc}){
  const toL=normalizeEmailList(to), ccL=normalizeEmailList(cc);
  if(toL.valid.length===0){ toast('Keine gültige Empfängeradresse (An). Bitte prüfen.', false); return; }
  if(toL.invalid.length||ccL.invalid.length){ toast(`Ungültige Adressen ignoriert: ${[...toL.invalid,...ccL.invalid].join(', ')}`, false); }

  const g=await ensureSettingsRecord();
  const url=(g.httpGateway||GATEWAY_DEFAULT).trim();
  const key=(g.apiKey||GATEWAY_API_KEY||'').trim();

  if(typeof GM_xmlhttpRequest==='function' && /^https?:\/\//i.test(url)){
    try{
      const res=await new Promise((resolve,reject)=>{
        GM_xmlhttpRequest({
          method:'POST', url, headers:{'Content-Type':'application/json','X-Api-Key':key},
          data:JSON.stringify({subject,html,to:toL.valid.join(','),cc:ccL.valid.join(',')}),
          onload:r=>resolve(r), onerror:e=>reject(e), ontimeout:()=>reject(new Error('timeout')), timeout:10000
        });
      });
      let j=null; try{ j=JSON.parse(res.responseText||''); }catch{}
      if(!(res.status>=200&&res.status<300) || !j || j.ok!==true) throw new Error(`Gateway-Fehler ${res.status}: ${res.responseText}`);
      toast('Mail über Gateway gesendet'); return;
    }catch(e){ console.error('[fvpr] GM gateway error',e); toast('Gateway nicht erreichbar – Fallback Outlook',false); }
  }
  if(/^https:\/\//i.test(url)){
    try{
      const r=await fetch(url,{method:'POST', headers:{'Content-Type':'application/json','X-Api-Key':key},
        body:JSON.stringify({subject,html,to:toL.valid.join(','),cc:ccL.valid.join(',')}), mode:'cors', keepalive:true});
      const t=await r.text(); let j=null; try{ j=JSON.parse(t);}catch{}
      if(!r.ok||!j||j.ok!==true) throw new Error(`HTTP ${r.status}: ${t}`);
      toast('Mail über Gateway gesendet'); return;
    }catch(e){ console.error('[fvpr] fetch gateway error',e); toast('Gateway nicht erreichbar – Fallback Outlook',false); }
  }
  await copyHtmlToClipboard(html);
  openMailto(subject,toL.valid.join(','),ccL.valid.join(','));
  alert('Entwurf geöffnet. HTML ist in der Zwischenablage – Strg+V drücken.');
}

/* ====== HTML-Builder ====== */
function partnerHtml(partner,list,signature){
  const rows=list.map(r=>{
    const etaStyle = `background:${etaBg(r.eta)};`;
    return `
      <tr data-row="1">
        <td style="padding:6px 8px;border:1px solid #e5e7eb;">${r.tour||'—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${r.driver||'—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.stops)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pkgs)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.open)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.obstacles)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;${etaStyle}">${fmtPct(r.eta)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pOpen)}</td>
      </tr>`;
  }).join('');
  const totals={
    tours:list.length,
    etaAvg:avg(list,r=>r.eta),
    stops:sum(list,r=>r.stops),
    pkgs:sum(list,r=>r.pkgs),
    open:sum(list,r=>r.open),
    obstacles:sum(list,r=>r.obstacles),
    pOpen:sum(list,r=>r.pOpen)
  };
  const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';
  return `
  <div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
    <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
    <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
      <thead><tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Tour</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Fahrername</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">offen</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">offene Abholstops</th>
      </tr></thead>
      <tbody>${rows}</tbody>
      <tfoot>
        <tr style="background:#e0f2ff;color:#003366;font-weight:700;">
          <td style="padding:8px;border:1px solid #e5e7eb;"></td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Touren: ${fmtInt(totals.tours)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">ETA Ø: ${fmtPct(totals.etaAvg)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
        </tr>
      </tfoot>
    </table>
    ${signatureHtml}
  </div>`;
}
function summaryHtml(per,totals,signature){
  const head=`
    <thead><tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
      <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Systempartner</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Touren</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA % (Ø)</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps gesamt</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Offene Stopps</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete gesamt</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Offene Abholstops</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Aktion</th>
    </tr></thead>`;
  const body=per.map(p=>{
    const etaStyle=`background:${etaBg(p.etaAvg)};`;
    return `
      <tr data-partner="${p.partner.replace(/"/g,'&quot;')}" data-row="1">
        <td style="text-align:left">${p.partner}</td>
        <td>${fmtInt(p.tours)}</td>
        <td style="${etaStyle}">${fmtPct(p.etaAvg)}</td>
        <td>${fmtInt(p.stops)}</td>
        <td>${fmtInt(p.open)}</td>
        <td>${fmtInt(p.pkgs)}</td>
        <td>${fmtInt(p.obstacles)}</td>
        <td>${fmtInt(p.pOpen)}</td>
        <td class="${NS}act">
          <button class="${NS}iconbtn" title="Mail an Partner (Bestätigung)" data-sp="${p.partner}">✉︎</button>
          <button class="${NS}iconbtn" title="Vorschau öffnen" data-eye="${p.partner}">👁</button>
          <button class="${NS}iconbtn" title="Partner-HTML in Zwischenablage" data-copy="${p.partner}">⧉</button>
        </td>
      </tr>`;
  }).join('');
  const foot=`
    <tfoot><tr>
      <td>Gesamt (alle)</td>
      <td>${fmtInt(totals.tours)}</td>
      <td>${fmtPct(totals.etaAvg)}</td>
      <td>${fmtInt(totals.stops)}</td>
      <td>${fmtInt(totals.open)}</td>
      <td>${fmtInt(totals.pkgs)}</td>
      <td>${fmtInt(totals.obstacles)}</td>
      <td>${fmtInt(totals.pOpen)}</td>
      <td class="${NS}act">
        <button class="${NS}iconbtn" title="Gesamtübersicht an „gesamt“ (Bestätigung)" data-total-send="1">✉︎</button>
        <button class="${NS}iconbtn" title="Gesamt-Vorschau öffnen" data-total-eye="1">👁</button>
        <button class="${NS}iconbtn" title="Gesamt-HTML in Zwischenablage" data-total-copy="1">⧉</button>
      </td>
    </tr></tfoot>`;
  const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';
  return `
  <div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
    <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
    <table class="${NS}tbl" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
      ${head}<tbody>${body}</tbody>${foot}
    </table>
    ${signatureHtml}
  </div>`;
}
function mailPartnerHtml(partner,list,signature){
  const rows=list.map(r=>{
    const etaStyle = `background:${etaBg(r.eta)};`;
    return `
      <tr>
        <td style="padding:8px;border:1px solid #e5e7eb;">${r.tour||'—'}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">${r.driver||'—'}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.stops)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pkgs)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.open)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.obstacles)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;${etaStyle}">${fmtPct(r.eta)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pOpen)}</td>
      </tr>`;
  }).join('');
  const totals={
    tours:list.length,
    etaAvg:avg(list,r=>r.eta),
    stops:sum(list,r=>r.stops),
    pkgs:sum(list,r=>r.pkgs),
    open:sum(list,r=>r.open),
    obstacles:sum(list,r=>r.obstacles),
    pOpen:sum(list,r=>r.pOpen)
  };
  const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';
  return `
  <div style="font:13px/1.5 -apple-system,Segoe UI,Arial,sans-serif; color:#111;">
    <div style="margin:0 0 8px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
    <div style="max-width:100%; overflow-x:auto;">
      <table cellpadding="0" cellspacing="0" style="width:100%; min-width:560px; border-collapse:collapse; table-layout:fixed; font:13px/1.45 -apple-system,Segoe UI,Arial,sans-serif;">
        <thead>
          <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
            <th align="center"style="padding:8px;border:1px solid #e5e7eb;">Tour</th>
            <th align="left"  style="padding:8px;border:1px solid #e5e7eb;">Fahrername</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Stopps</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Pakete</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">offen</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">ETA</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">offene Abholstops</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
        <tfoot>
          <tr style="background:#e0f2ff;color:#003366;font-weight:700;">
            <td style="padding:8px;border:1px solid #e5e7eb;"></td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Touren: ${fmtInt(totals.tours)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">ETA Ø: ${fmtPct(totals.etaAvg)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
    ${signatureHtml}
  </div>`;
}
function mailSummaryHtml(per,totals,signature){
  const body=per.map(p=>{
    const etaStyle=`background:${etaBg(p.etaAvg)};`;
    return `
      <tr>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">${p.partner}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.tours)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;${etaStyle}">${fmtPct(p.etaAvg)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.stops)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.open)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.pkgs)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.obstacles)}</td>
        <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.pOpen)}</td>
      </tr>`;
  }).join('');
  const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';
  return `
  <div style="font:13px/1.5 -apple-system,Segoe UI,Arial,sans-serif; color:#111;">
    <div style="margin:0 0 8px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
    <div style="max-width:100%; overflow-x:auto;">
      <table cellpadding="0" cellspacing="0" style="width:100%; min-width:640px; border-collapse:collapse; table-layout:fixed; font:13px/1.45 -apple-system,Segoe UI,Arial,sans-serif;">
        <thead>
          <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
            <th align="left"  style="padding:8px;border:1px solid #e5e7eb;">Systempartner</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Touren</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">ETA % (Ø)</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Stopps gesamt</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Offene Stopps</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Pakete gesamt</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Offene Abholstops</th>
          </tr>
        </thead>
        <tbody>${body}</tbody>
        <tfoot>
          <tr style="background:#e0f2ff;color:#00366;font-weight:700;">
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt (alle)</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.tours)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(totals.etaAvg)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
          </tr>
        </tfoot>
      </table>
    </div>
    ${signatureHtml}
  </div>`;
}

/* ---------- Aggregation + Flows ---------- */
async function getAggregates(){
  const {ok,rows}=await readAllRows();
  if(!ok||rows.length===0) return null;
  const groups=groupByPartner(rows);
  const per=[];
  for(const [partner,list] of groups){
    per.push({
      partner, list,
      tours:list.length,
      etaAvg:avg(list,r=>r.eta),
      stops:sum(list,r=>r.stops),
      open:sum(list,r=>r.open),
      pkgs:sum(list,r=>r.pkgs),
      obstacles:sum(list,r=>r.obstacles),
      pOpen:sum(list,r=>r.pOpen)
    });
  }
  per.sort((a,b)=>a.partner.localeCompare(b.partner,'de'));
  const totals={
    tours:per.reduce((a,p)=>a+(p.tours||0),0),
    etaAvg:avg(rows,r=>r.eta),
    stops:per.reduce((a,p)=>a+(p.stops||0),0),
    open:per.reduce((a,p)=>a+(p.open||0),0),
    pkgs:per.reduce((a,p)=>a+(p.pkgs||0),0),
    obstacles:per.reduce((a,p)=>a+(p.obstacles||0),0),
    pOpen:per.reduce((a,p)=>a+(p.pOpen||0),0)
  };
  return {per, totals};
}
async function copyPartnerHtml(partner){
  const agg=await getAggregates(); if(!agg) return;
  const g=await ensureSettingsRecord();
  const p=agg.per.find(x=>x.partner===partner); if(!p) return;
  const html=partnerHtml(partner, p.list, g.signature||'');
  const ok=await copyHtmlToClipboard(html);
  toast(ok?'Vorschau in Zwischenablage':'Kopieren fehlgeschlagen', ok);
}
async function openTotalPreview(){
  const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
  const g=await ensureSettingsRecord();
  const html=summaryHtml(agg.per, agg.totals, g.signature||'');
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Vorschau – Gesamt</h3>
    <div style="max-height:60vh;overflow:auto;border:1px solid #e5e7eb;border-radius:8px;padding:8px;margin-bottom:10px;background:#fff">${html}</div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>
  `);
  const innerTable = ov.querySelector('table');
  if(innerTable) makeTableSortable(innerTable.parentElement);
  ov.addEventListener('click', e=>{ if(e.target.closest('button[data-act="close"]')) ov.remove(); },{passive:false});
}
async function copyTotalHtml(){
  const agg=await getAggregates(); if(!agg) return;
  const g=await ensureSettingsRecord();
  const html=summaryHtml(agg.per, agg.totals, g.signature||'');
  const ok=await copyHtmlToClipboard(html);
  toast(ok?'Gesamt in Zwischenablage':'Kopieren fehlgeschlagen', ok);
}
async function fillCfg(){
  const g=await ensureSettingsRecord();
  const el=PANEL.querySelector('#'+NS+'cfg-subj');
  if(el) el.value=g.subjectPrefix||'Aktueller Tour.Report';
}
async function saveCfgFromUI(){
  try{
    const cur=await ensureSettingsRecord();
    await idbPut('settings',{
      id:'global',
      subjectPrefix:(PANEL.querySelector('#'+NS+'cfg-subj')?.value||'').trim(),
      signature:cur.signature||'',
      httpGateway:cur.httpGateway||GATEWAY_DEFAULT,
      apiKey:cur.apiKey||GATEWAY_API_KEY
    });
    toast('Einstellungen gespeichert');
  }catch(e){
    console.error(e); toast('Fehler beim Speichern',false);
  }
}
async function openPreview(partner){
  const agg=await getAggregates();
  if(!agg){ alert('Keine Daten gefunden.'); return; }
  const g=await ensureSettingsRecord();
  const p=agg.per.find(x=>x.partner===partner);
  if(!p){ alert('Partner nicht gefunden.'); return; }
  const content=partnerHtml(partner, p.list, g.signature||'');
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Vorschau – ${partner}</h3>
    <div style="max-height:60vh;overflow:auto;border:1px solid #e5e7eb;border-radius:8px;padding:8px;margin-bottom:10px;background:#fff">${content}</div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="edit">Einstellungen</button>
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>
  `);
  const innerTable = ov.querySelector('table');
  if(innerTable) makeTableSortable(innerTable.parentElement);
  ov.addEventListener('click', e=>{
    const btn=e.target.closest('button[data-act]'); if(!btn) return;
    if(btn.dataset.act==='close') ov.remove();
    if(btn.dataset.act==='edit'){ ov.remove(); openPartnerDialog(partner); }
  },{passive:false});
}
function openConfirm({title, subject, to, cc, saveKey, onOk}){
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">${title||'Bestätigen'}</h3>
    <div style="display:grid;grid-template-columns:1fr 2fr;gap:8px;margin:8px 0">
      <label>Betreff</label><input id="${NS}cf-subj" type="text" value="${subject||''}">
      <label>An</label><input id="${NS}cf-to" type="text" value="${to||''}">
      <label>CC</label><input id="${NS}cf-cc" type="text" value="${cc||''}">
    </div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="save">Änderungen speichern</button>
      <button class="${NS}btn-sm" data-act="ok">OK</button>
      <button class="${NS}btn-sm" data-act="cancel">Abbrechen</button>
    </div>
  `);
  ov.addEventListener('click', async e=>{
    const b=e.target.closest('button[data-act]'); if(!b) return;
    const subj=ov.querySelector('#'+NS+'cf-subj').value.trim();
    const toV =ov.querySelector('#'+NS+'cf-to').value.trim();
    const ccV =ov.querySelector('#'+NS+'cf-cc').value.trim();
    if(b.dataset.act==='cancel'){ ov.remove(); return; }
    if(b.dataset.act==='save'){
      const key = saveKey || 'gesamt';
      const cur = await idbGet('partners', key) || {name:key, alias:''};
      await idbPut('partners', { name:key, to:toV, cc:ccV, alias:cur.alias||'' });
      toast('Adressen gespeichert');
      return;
    }
    if(b.dataset.act==='ok'){ onOk({subject:subj, to:toV, cc:ccV}); ov.remove(); }
  },{passive:false});
}
async function openPartnerDialog(partner){
  const key=(partner==='gesamt')?'gesamt':partner;
  const cur=await idbGet('partners',key)||{name:key,to:'',cc:'',alias:''};
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Einstellungen – ${partner}</h3>
    <div class="${NS}row"><label>Alias (optional)</label><input id="${NS}pd-alias" type="text" value="${cur.alias||''}"></div>
    <div class="${NS}row"><label>An</label><input id="${NS}pd-to" type="text" value="${cur.to||''}" placeholder="a@b.de, c@d.de"></div>
    <div class="${NS}row"><label>CC</label><input id="${NS}pd-cc" type="text" value="${cur.cc||''}" placeholder="optional"></div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="save">Speichern</button>
      <button class="${NS}btn-sm" data-act="clear">Löschen</button>
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>
  `);
  ov.addEventListener('click', async e=>{
    const btn=e.target.closest('button[data-act]'); if(!btn) return;
    if(btn.dataset.act==='close'){ ov.remove(); return; }
    if(btn.dataset.act==='clear'){ await idbDel('partners',key); ov.remove(); return; }
    if(btn.dataset.act==='save'){
      const to=ov.querySelector('#'+NS+'pd-to').value.trim();
      const cc=ov.querySelector('#'+NS+'pd-cc').value.trim();
      const alias=ov.querySelector('#'+NS+'pd-alias').value.trim();
      await idbPut('partners',{name:key,to,cc,alias});
      alert('Gespeichert.'); ov.remove();
    }
  },{passive:false});
}

/* ====== Versand-Flows ====== */
async function sendSinglePartnerConfirm(partner){
  if(partner==='gesamt'){ await sendTotalOnlyConfirm(); return; }
  const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
  const g=await ensureSettingsRecord();
  const p=agg.per.find(x=>x.partner===partner); if(!p){ alert('Partner nicht gefunden.'); return; }
  const rec=await idbGet('partners', partner)||{};
  const alias=rec.alias||partner;
  const subject=`${g.subjectPrefix||'Aktueller Tour.Report'} – ${alias} – ${todayStr()}`;
  const html=mailPartnerHtml(partner, p.list, g.signature||'');
  openConfirm({
    title:`Senden an ${alias}?`,
    subject, to:rec.to||'', cc:rec.cc||'', saveKey:partner,
    onOk:({subject,to,cc})=>deliverMail({subject, html, to, cc})
  });
}
async function sendTotalOnlyConfirm(){
  const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
  const g=await ensureSettingsRecord();
  const rec=await idbGet('partners','gesamt')||{};
  const subject=`${g.subjectPrefix||'Aktueller Tour.Report'} – Gesamt – ${todayStr()}`;
  const html=mailSummaryHtml(agg.per, agg.totals, g.signature||'');
  openConfirm({
    title:'Gesamtübersicht senden?',
    subject, to:rec.to||'', cc:rec.cc||'', saveKey:'gesamt',
    onOk:({subject,to,cc})=>deliverMail({subject, html, to, cc})
  });
}
async function sendPartnerAndTotalConfirm(){
  const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
  const g=await ensureSettingsRecord();
  const ready=[];
  for(const p of agg.per){
    const ov=await idbGet('partners', p.partner);
    if(ov && normalizeEmailList(ov.to||'').valid.length>0) ready.push(p.partner);
  }
  const rec=await idbGet('partners','gesamt')||{};
  const subjectTotal=`${g.subjectPrefix||'Aktueller Tour.Report'} – Gesamt – ${todayStr()}`;
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Sammelversand bestätigen</h3>
    <div style="font:13px;margin-bottom:10px">Es werden <b>${ready.length}</b> Partner-Mails gesendet (nur mit gültiger Adresse) und <b>1</b> Gesamt-Mail an „gesamt“.</div>
    <div style="display:grid;grid-template-columns:1fr 2fr;gap:8px;margin:8px 0">
      <label>Gesamt – Betreff</label><input id="${NS}sammel-subj" type="text" value="${subjectTotal}">
      <label>Gesamt – An</label><input id="${NS}sammel-to" type="text" value="${rec.to||''}">
      <label>Gesamt – CC</label><input id="${NS}sammel-cc" type="text" value="${rec.cc||''}">
    </div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="save">Änderungen speichern</button>
      <button class="${NS}btn-sm" data-act="ok">Senden</button>
      <button class="${NS}btn-sm" data-act="cancel">Abbrechen</button>
    </div>
  `);
  ov.addEventListener('click', async e=>{
    const b=e.target.closest('button[data-act]'); if(!b) return;
    const subj=ov.querySelector('#'+NS+'sammel-subj').value.trim();
    const to  =ov.querySelector('#'+NS+'sammel-to').value.trim();
    const cc  =ov.querySelector('#'+NS+'sammel-cc').value.trim();
    if(b.dataset.act==='cancel'){ ov.remove(); return; }
    if(b.dataset.act==='save'){
      const cur=await idbGet('partners','gesamt')||{name:'gesamt',alias:''};
      await idbPut('partners',{name:'gesamt', to, cc, alias:cur.alias||''});
      toast('Adressen „gesamt“ gespeichert'); return;
    }
    if(b.dataset.act==='ok'){
      for(const pname of ready){
        const ovRec=await idbGet('partners', pname);
        const part=agg.per.find(x=>x.partner===pname);
        const html=mailPartnerHtml(pname, part.list, g.signature||'');
        const subjP=`${g.subjectPrefix||'Aktueller Tour.Report'} – ${(ovRec.alias||pname)} – ${todayStr()}`;
        await deliverMail({subject:subjP, html, to:ovRec.to||'', cc:ovRec.cc||''});
      }
      const htmlTot=mailSummaryHtml(agg.per, agg.totals, g.signature||'');
      await deliverMail({subject:subj, html:htmlTot, to, cc});
      ov.remove();
    }
  },{passive:false});
}

// ---------- Render ----------
let renderTimer=null;
async function render(force=false){
  if(Date.now() - LAST_USER_SORT_TS < SUPPRESS_RENDER_MS && !force) return;
  if(renderTimer){ clearTimeout(renderTimer); renderTimer=null; }

  const run=async ()=>{
    if(!CONTENT) return;
    const res=await getAggregates();
    if(!res){
      CONTENT.innerHTML=`<div class="${NS}empty">Keine Daten gefunden (Tab „Fahrzeugübersicht“ sichtbar?).</div>`;
      return;
    }
    const {per, totals}=res;
    const html = summaryHtml(per, totals, '');
    CONTENT.innerHTML = html;
    makeTableSortable(CONTENT);
  };
  if(force) await run(); else renderTimer=setTimeout(run, RENDER_DEBOUNCE);
}

/* ====== TEIL 12/12 – Boot + Loader-Integration ====== */

function openPanel(){
  try { ensureStyles(); } catch {}
  try { if (!document.querySelector(PANEL_ID)) mountUI(true); } catch {}
  const p = document.querySelector(PANEL_ID);
  if (p) p.style.display = '';
  try { fillCfg(); } catch {}
  try { render(true); } catch {}
}

function closePanel(){
  const p = document.querySelector(PANEL_ID);
  if (p) p.style.display = 'none';
}

async function bootStandalone(){
  try { ensureStyles(); mountUI(false); await ensureSettingsRecord(); await render(true); }
  catch(e){ console.error('[fvpr] init', e); }
  let mo = null, refreshInterval = null;
  if (!mo){
    mo = new MutationObserver(()=>{ if (Date.now() - LAST_USER_SORT_TS < SUPPRESS_RENDER_MS) return; render(); });
    mo.observe(document.body, { childList:true, subtree:true });
  }
  if (!refreshInterval){
    refreshInterval = setInterval(()=>{ if (Date.now() - LAST_USER_SORT_TS < SUPPRESS_RENDER_MS) return; render(); }, REFRESH_MS);
  }
}

// ====== Start ======
function registerWithLoader(){
  const def = {
    id: 'partner-report',
    label: 'Partner-Report',
    panels: [PANEL_ID],
    run: () => {
      const el = document.querySelector(PANEL_ID);
      if (el && getComputedStyle(el).display !== 'none') closePanel();
      else openPanel();
    }
  };
  const G = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
  if (G.TM && typeof G.TM.register === 'function'){
    G.TM.register(def);
    removeStandaloneButton();        // <- Button sicher entfernen
  } else {
    const KEY='__tmQueue';
    G[KEY] = Array.isArray(G[KEY]) ? G[KEY] : [];
    G[KEY].push(def);                // <- nur vormerken, KEIN bootStandalone()
  }
  G.fvpr_open = openPanel;
  G.fvpr_close = closePanel;
}

// Nur registrieren (und ggf. Button entfernen). KEIN bootStandalone() mehr!
registerWithLoader();




})();
