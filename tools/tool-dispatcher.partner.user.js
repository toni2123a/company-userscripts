/* ====== TEIL 1/12 – Metadaten + Start IIFE (LOADER-fähig) ====== */

// ==UserScript==
// @name         DPD Dispatcher – Partner-Report Mailer
// @namespace    bodo.dpd.custom
// @version      5.7.3
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher.partner.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher.partner.user.js
// @description  ✉ je Partner mit Bestätigung + „Änderungen speichern“; Zeilenklick = Vorschau; Gesamt an „gesamt“. Lokale Empfänger (IndexedDB), Export/Import. Robust (Datagrid ODER normale Tabelle). Fix: robuste Spalten-Erkennung je Header-Reihenfolge + ETA-Prozentspalte statt ETA-Zeit + Abholstops robust + Status-Spalte in Partnerseiten. Klick-Details Stopps/Pakete/offen bis Paket-Lebenslauf mit Prio/Express-Markierung. Loader-Integration (TM).
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      10.14.7.169
// ==/UserScript==

(function(){
'use strict';
if (window.__FVPR_RUNNING) return; window.__FVPR_RUNNING = true;
console.log('[fvpr] Script gestartet', { ua: navigator.userAgent, url: location.href, ts: new Date().toISOString() });


// ganz oben bei den Konstanten
const ENABLE_STANDALONE = false; // <- kein eigener Button

const NS='fvpr-', PANEL_ID='#fvpr-panel', OFFSET_PX=-240;
const REFRESH_MS=60_000, RENDER_DEBOUNCE=300, SCAN_MAX_STEPS=40, SCAN_STAG_LIMIT=2;
const GATEWAY_DEFAULT='http://10.14.7.169/mail.php', GATEWAY_API_KEY='fvpr-SECRET-123';

const DEBUG = localStorage.getItem('fvpr-debug')==='1';
const LOG=(...a)=>{ if(DEBUG) console.log('[fvpr]',...a); };

// ====== Visuelles Debug-Log (im Panel sichtbar) ======
const DIAG_ENABLED = localStorage.getItem('fvpr-diag')==='1';
const _diagEntries = [];
const DIAG = (tag, msg, data) => {
  const ts = new Date().toLocaleTimeString('de-DE');
  const entry = { ts, tag, msg, data: data !== undefined ? data : '' };
  _diagEntries.push(entry);
  if (_diagEntries.length > 200) _diagEntries.shift();
  if (DEBUG) console.log(`[fvpr-diag][${tag}]`, msg, data ?? '');
  // Live-Update falls Panel offen
  try { _diagRenderLive(); } catch {}
};
function _diagRenderLive(){
  const box = document.getElementById(NS+'diagbox');
  if (!box || box.style.display === 'none') return;
  const pre = box.querySelector('pre');
  if (pre) pre.textContent = _diagFormat();
}
function _diagFormat(){
  return _diagEntries.map(e => `[${e.ts}] [${e.tag}] ${e.msg}${e.data !== '' ? '  → ' + (typeof e.data === 'object' ? JSON.stringify(e.data) : String(e.data)) : ''}`).join('\n');
}

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
const tourKey = t => String(t || '').replace(/[^\dA-Za-z]/g,'').trim();

const groupByPartner=rows=>{ const m=new Map(); for(const r of rows){ if(!m.has(r.partner)) m.set(r.partner,[]); m.get(r.partner).push(r); } return m; };

const qsaMain=sel=>Array.from(document.querySelectorAll(sel)).filter(el=>!el.closest(PANEL_ID));
function toast(msg, ok=true){ const el=document.createElement('div'); el.style.cssText='position:fixed;right:16px;bottom:16px;padding:10px 14px;border-radius:10px;font:600 13px system-ui;color:#fff;z-index:2147483647;'+(ok?'background:#16a34a':'background:#b91c1c'); el.textContent=msg; document.body.appendChild(el); setTimeout(()=>{ el.style.opacity='0'; el.style.transition='opacity .3s'; setTimeout(()=>el.remove(),300); }, 1400); }

/* ====== TEIL 2/12 – IndexedDB ====== */

const IDB_NAME='fvpr_db', IDB_VER=3;

function idbOpen(){
  return new Promise((res,rej)=>{
    const req = indexedDB.open(IDB_NAME, IDB_VER);

    req.onupgradeneeded = ()=>{
      const db=req.result;

      if(!db.objectStoreNames.contains('partners'))
        db.createObjectStore('partners',{keyPath:'name'});

      if(!db.objectStoreNames.contains('tourMap'))
        db.createObjectStore('tourMap',{keyPath:'tour'}); // {tour, partner, updatedAt}

      if(!db.objectStoreNames.contains('settings')){
        const s=db.createObjectStore('settings',{keyPath:'id'});
        s.put({
          id:'global',
          subjectPrefix:'Aktueller Tour.Report',
          distTo:'',
          distCc:'',
          signature:'',
          httpGateway:'',
          apiKey:''
        });
      }
    };

    req.onsuccess=()=>res(req.result);
    req.onerror=()=>rej(req.error);
  });
}

async function idbGet(store,key){
  const db=await idbOpen();
  return new Promise((res,rej)=>{
    const r=db.transaction(store,'readonly').objectStore(store).get(key);
    r.onsuccess=()=>res(r.result||null);
    r.onerror=()=>rej(r.error);
  });
}
async function idbPut(store,val){
  const db=await idbOpen();
  return new Promise((res,rej)=>{
    const r=db.transaction(store,'readwrite').objectStore(store).put(val);
    r.onsuccess=()=>res(true);
    r.onerror=()=>rej(r.error);
  });
}
async function idbDel(store,key){
  const db=await idbOpen();
  return new Promise((res,rej)=>{
    const r=db.transaction(store,'readwrite').objectStore(store).delete(key);
    r.onsuccess=()=>res(true);
    r.onerror=()=>rej(r.error);
  });
}
async function idbAll(store){
  const db=await idbOpen();
  return new Promise((res,rej)=>{
    const r=db.transaction(store,'readonly').objectStore(store).getAll();
    r.onsuccess=()=>res(r.result||[]);
    r.onerror=()=>rej(r.error);
  });
}

async function idbPutMany(store, arr){
  if(!arr || !arr.length) return;
  const db = await idbOpen();
  await new Promise((res,rej)=>{
    const tx = db.transaction(store,'readwrite');
    const st = tx.objectStore(store);
    for(const v of arr) st.put(v);
    tx.oncomplete = ()=>res(true);
    tx.onerror = ()=>rej(tx.error);
  });
}

async function cacheTourPartner(rows){
  try{
    const now = Date.now();
    const m = new Map(); // tourKey -> partner
    for(const r of (rows||[])){
      const t = tourKey(r.tour);
      const p = norm(r.partner);
      if(t && p) m.set(t,p);
    }
    if(!m.size) return;

    const recs = [];
    for(const [tour, partner] of m.entries()){
      recs.push({ tour, partner, updatedAt: now });
    }
    await idbPutMany('tourMap', recs);
  }catch(e){
    console.warn('[fvpr] cacheTourPartner', e);
  }
}


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
  DIAG('grid', 'getGridViewport', { found: !!grid, tagName: grid?.tagName, class: grid?.className?.slice?.(0,80) });
  if(!grid) return null;
  GRID_VP = grid.closest('[class*="Datagrid"]') || grid;
  return GRID_VP;
}
function getAnyTable(){
  const tables = qsaMain('table');
  DIAG('table', 'getAnyTable – Tabellen im DOM', { count: tables.length });
  for (const t of tables){
    const ths = Array.from(t.querySelectorAll('thead th')).map(th=>norm(th.textContent).toLowerCase());
    DIAG('table', 'Spalten einer Tabelle', { cols: ths.length, headers: ths.slice(0,8).join(' | ') });
    if (ths.length && ths.some(x=>x.includes('systempartner'))) return t;
  }
  DIAG('table', 'Keine passende Tabelle mit "systempartner" gefunden');
  return null;
}
const includesAll=(s,arr)=>arr.every(w=>new RegExp(w,'i').test(s||''));

/* ====== TEIL 4/12 – Datagrid: Spalten finden ====== */

function findColumnsDatagrid(){
  const ths=qsaMain('thead th,[role="columnheader"]');
  if(!ths.length) return null;

  const normTxt = el => (el?.textContent||'').replace(/\s+/g,' ').trim().toLowerCase();
  const titleOf = el => (el?.querySelector('input[title], [title]')?.getAttribute('title')||'').trim().toLowerCase();
  const H = Array.from({length: ths.length}, (_,i)=>({ i, text: normTxt(ths[i]), title: titleOf(ths[i]) }));

  // Wichtig: Cache nicht nur nach Spaltenanzahl, sondern nach kompletter Header-Signatur.
  // Sonst können bei Kollegen mit anderer Dispatcher-Spaltenreihenfolge falsche Spalten wiederverwendet werden.
  const headerSig = H.map(h => `${h.text}|${h.title}`).join('||');
  if (CACHED_COLS && CACHED_COLS._headerSig === headerSig) return CACHED_COLS;

  const byEither = fn => { for(const h of H){ if(fn(h.text)||fn(h.title)) return h.i; } return -1; };
  const pickHeader = (must=[], any=[], not=[]) => byEither(s => {
    if(!s) return false;
    if(must.length && !must.every(re => re.test(s))) return false;
    if(any.length && !any.some(re => re.test(s))) return false;
    if(not.length && not.some(re => re.test(s))) return false;
    return true;
  });

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
  function pickEtaIndex(){
    // Es gibt im Dispatcher mehrere ETA-Spalten, z. B. ETA % und ETA-Zeit.
    // Deshalb nicht nur nach Header-Text gehen, sondern die sichtbaren Zellwerte prüfen.
    // Gewünscht ist die Prozent-ETA-Spalte: Zellen enthalten überwiegend "100 %", "94 %" usw.
    // Die Zeit-ETA-Spalte enthält Werte wie "-00:02" oder "00:38" und wird negativ bewertet.
    const cand = H.map(h=>h.i).filter(i=>{
      const s = `${H[i].text} ${H[i].title}`;
      return /\beta\b/.test(s) && !/tour|route|zeitfenster|ankunft/.test(s);
    });
    if(!cand.length) return -1;

    const rows = qsaMain('tbody tr,[role="row"]');
    const score = i => {
      let sc = 0, seen = 0;
      const hs = `${H[i].text} ${H[i].title}`;
      if (/%|prozent|percent/.test(hs)) sc += 20;
      if (/zeit|time|hh:mm|ankunft/.test(hs)) sc -= 20;

      for (const tr of rows.slice(0, 60)){
        const tds = tr.querySelectorAll('td,[role="gridcell"]');
        if(!tds || !tds[i]) continue;
        const v = (tds[i].textContent || '').trim();
        if(!v) continue;
        seen++;
        if (/%/.test(v)) sc += 8;
        if (/^-?\d{1,2}:\d{2}$/.test(v) || /^-?\d{2}:\d{2}$/.test(v)) sc -= 10;
        const n = parsePct(v);
        if (Number.isFinite(n) && n >= 0 && n <= 130 && !/:/.test(v)) sc += 3;
        if (Number.isFinite(n) && n < 0) sc -= 5;
      }
      return sc + seen * 0.05;
    };

    let best = cand[0], bestScore = -1e9;
    for(const i of cand){
      const s = score(i);
      if(s > bestScore){ best = i; bestScore = s; }
    }
    DIAG('cols', 'ETA-Kandidaten bewertet', cand.map(i=>({i, header:H[i].text, title:H[i].title, score:score(i)})));
    return best;
  }

  const eta = pickEtaIndex();
  const status = byEither(s=>/\bstatus\b/.test(s));

  // Streng getrennte Erkennung: Zustellstopps gesamt darf nicht mit offenen Stopps oder Abholstopps verwechselt werden.
  const stopsTotal = (() => {
    const i = pickHeader([/(zustell)?stopps?/, /(gesamt|total)/], [], [/offen|open|abhol|pickup|hindern/]);
    if(i >= 0) return i;
    return pickHeader([/stopps?/, /(gesamt|total)/], [], [/offen|open|abhol|pickup|hindern/]);
  })();
  const stopsOpen  = (() => {
    const i = pickHeader([/(offen|open)/, /(zustell)?stopps?/], [], [/abhol|pickup|gesamt|total|hindern/]);
    if(i >= 0) return i;
    return pickHeader([/(offen|open)/, /stopps?/], [], [/abhol|pickup|gesamt|total|hindern/]);
  })();
  const pkgsTotal  = (() => {
    const i = pickHeader([/pakete|geplante/, /(gesamt|zustell|total)/], [], [/offen|open|abhol|pickup|hindern/]);
    if(i >= 0) return i;
    return byEither(s=>includesAll(s, [/pakete|geplante/, /(gesamt|zustell)/]));
  })();
  const obstacles  = byEither(s=>/\bzustellhindernisse\b/.test(s)||/\bhinderniss?e?\b/.test(s));
  const pickupOpen = (()=>{
    const i1 = pickHeader([/offen|open/, /abhol(stopp|stopps|ung|ungen)|pickup(s)?/], [], [/zustell/]);
    if (i1>=0) return i1;
    return byEither(s=>/abhol(stopp|stopps|ung|ungen)/.test(s));
  })();

  const cols={sys,tour,driver,eta,status,stopsTotal,stopsOpen,pkgsTotal,obstacles,pickupOpen,_headerSig:headerSig};
  DIAG('cols', 'findColumnsDatagrid', cols);
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
  const idxStrict = (must=[], not=[]) => head.findIndex(t => must.every(x => t.includes(x)) && !not.some(x => t.includes(x)));

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
    eta: (()=>{
      // Auch in normalen Tabellen gibt es teilweise zwei ETA-Spalten.
      // Wir wählen gezielt die Prozent-Spalte anhand Header und Zellwerten.
      const cand = head.map((t,ix)=>({t,ix})).filter(x=>x.t.includes('eta') && !/tour|route|zeitfenster|ankunft/.test(x.t)).map(x=>x.ix);
      if(!cand.length) return -1;
      const rows = Array.from(table.querySelectorAll('tbody tr')).slice(0,60);
      const score = ci => {
        let sc=0, seen=0;
        const h=head[ci]||'';
        if (h.includes('%') || h.includes('prozent') || h.includes('percent')) sc += 20;
        if (/zeit|time|hh:mm|ankunft/.test(h)) sc -= 20;
        for(const tr of rows){
          const tds=tr.querySelectorAll('td');
          if(!tds[ci]) continue;
          const v=(tds[ci].textContent||'').trim();
          if(!v) continue;
          seen++;
          if (/%/.test(v)) sc += 8;
          if (/^-?\d{1,2}:\d{2}$/.test(v) || /^-?\d{2}:\d{2}$/.test(v)) sc -= 10;
          const n=parsePct(v);
          if (Number.isFinite(n) && n>=0 && n<=130 && !/:/.test(v)) sc += 3;
          if (Number.isFinite(n) && n<0) sc -= 5;
        }
        return sc + seen*0.05;
      };
      let best=cand[0], bestScore=-1e9;
      for(const ci of cand){ const s=score(ci); if(s>bestScore){ best=ci; bestScore=s; } }
      DIAG('cols','PlainTable ETA-Kandidaten bewertet', cand.map(i=>({i, header:head[i], score:score(i)})));
      return best;
    })(),
    stopsTotal: (()=>{ const i=idxStrict(['stopps','gesamt'], ['offen','abhol','pickup','hindern']); if(i>=0) return i; return idx(['zustellstopps gesamt','stopps gesamt','stopp gesamt']); })(),
    stopsOpen: (()=>{ const i=idxStrict(['offen','stopps'], ['abhol','pickup','gesamt','hindern']); if(i>=0) return i; return idx(['offene zustellstopps','offene stopps']); })(),
    pkgsTotal: (()=>{ const i=idxStrict(['pakete','gesamt'], ['offen','abhol','pickup','hindern']); if(i>=0) return i; return idx(['pakete gesamt','geplante zustellpakete']); })(),
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
  DIAG('read', 'readAllRows gestartet');
  const tryA = await readAllRowsDatagrid();
  DIAG('read', 'Datagrid-Ergebnis', { ok: tryA.ok, rows: tryA.rows.length });
  if (tryA.ok && tryA.rows.length) return tryA;
  const tryB = readRowsPlainTable();
  DIAG('read', 'PlainTable-Ergebnis', { ok: tryB.ok, rows: tryB.rows.length });
  if (tryB.ok && tryB.rows.length) return tryB;
  DIAG('read', 'KEINE Daten gefunden – weder Datagrid noch Tabelle');
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
.${NS}panel{
  position:fixed;
  top:48px;                /* vertikal bleibt wie gehabt */
  left:50%;                /* horizontale Mitte */
  transform:translateX(-50%);
  width:min(1150px,96vw);
  max-height:76vh;
  overflow:auto;
  background:#fff;
  border:1px solid rgba(0,0,0,.12);
  box-shadow:0 12px 28px rgba(0,0,0,.18);
  border-radius:12px;
  z-index:2147483646;
}


.${NS}hdr{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
.${NS}pill{display:inline-flex;gap:6px;align-items:center;flex-wrap:wrap}
.${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer;margin-left:6px}
.${NS}tbl{width:100%;border-collapse:collapse}
.${NS}tbl thead th{position:sticky;top:0;z-index:1;padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.12);font:700 12px system-ui;text-align:right;white-space:nowrap;background:#ffe2e2;color:#8b0000;cursor:pointer;user-select:none}
.${NS}tbl thead th:first-child,.${NS}tbl thead td:first-child,.${NS}tbl tbody td:first-child{text-align:left}
.${NS}tbl tbody tr{cursor:pointer}
.${NS}tbl tbody tr:hover{background:#f8fafc}
.${NS}tbl tbody td{padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.06);font:500 12px system-ui;text-align:right;white-space:nowrap}
.${NS}tbl tbody td.${NS}act{text-align:center}
.${NS}tbl tfoot td{padding:8px 10px;border-top:1px solid rgba(0,0,0,.12);font:700 12px system-ui;background:#e0f2ff;color:#003366;text-align:right;white-space:nowrap}
.${NS}empty{padding:12px;text-align:center;opacity:.7}
.${NS}cfg{padding:10px;border-top:1px solid rgba(0,0,0,.06);background:#fafafa}
.${NS}cfg input{width:100%;padding:6px;border:1px solid #e5e7eb;border-radius:8px}
.${NS}row{display:grid;grid-template-columns:1fr 2fr;gap:10px;margin:6px 0}
.${NS}modal{position:fixed;top:0;left:0;right:0;bottom:0;width:100%;height:100%;background:rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;z-index:2147483647}
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
  const table = (root instanceof HTMLTableElement) ? root : ((root instanceof HTMLElement) ? root.querySelector('table') : root);
  if(!table) return;

  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  if(!thead || !tbody) return;

  Array.from(tbody.querySelectorAll('tr')).forEach(tr=>{ tr.setAttribute('data-row','1'); });

  const ths = Array.from(thead.querySelectorAll('th'));
  if(!ths.length) return;

  const parseVal = (txt) => {
    const t = String(txt || '').trim();
    const clean = t.replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, '').trim();
    if (clean !== '' && !isNaN(clean)) return parseFloat(clean);
    return t.toLowerCase();
  };

  const updateIndicators = (col, dir) => {
    ths.forEach((th, i) => {
      th.querySelector(`span.${NS}sort-ind`)?.remove();
      if (i === col) {
        const sp = document.createElement('span');
        sp.className = `${NS}sort-ind`;
        sp.textContent = dir === 1 ? '▲' : '▼';
        th.appendChild(sp);
      }
    });
    table.dataset.sortCol = String(col);
    table.dataset.sortDir = String(dir);
  };

  const doSort = (colIndex) => {
    const currentCol = parseInt(table.dataset.sortCol || '-1', 10);
    const currentDir = parseInt(table.dataset.sortDir || '0', 10);

    let dir = 1;
    if (currentCol === colIndex) {
      dir = currentDir === 1 ? -1 : 1;
    }

    const rows = Array.from(tbody.querySelectorAll('tr[data-row="1"]'));

    rows.sort((a, b) => {
      const aText = a.children[colIndex]?.textContent ?? '';
      const bText = b.children[colIndex]?.textContent ?? '';

      const aVal = parseVal(aText);
      const bVal = parseVal(bText);

      if (typeof aVal === 'number' && typeof bVal === 'number') {
        return (aVal - bVal) * dir;
      }

      return String(aVal).localeCompare(String(bVal), 'de', { numeric: true }) * dir;
    });

    rows.forEach(r => tbody.appendChild(r));
    updateIndicators(colIndex, dir);
    LAST_USER_SORT_TS = Date.now();
  };

  ths.forEach((th, colIndex) => {
    if (th.dataset.fvprSortBound === '1') return;
    th.dataset.fvprSortBound = '1';
    th.style.cursor = 'pointer';

    th.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      doSort(colIndex);
    }, false);
  });
}

function mountUI(forLoader=false){
  DIAG("mount", "mountUI aufgerufen", { forLoader, alreadyExists: !!document.getElementById(NS+"panel") });
  if (document.getElementById(NS+'panel')) return;

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

  if (!forLoader && ENABLE_STANDALONE){
    if (!document.getElementById(NS+'wrap')){
      const wrap=document.createElement('div'); wrap.id=NS+'wrap'; wrap.className=NS+'wrap';
      const btn=document.createElement('button'); btn.className=NS+'btn'; btn.textContent='Partner-Report';
      wrap.append(btn); document.body.appendChild(wrap);
      btn.addEventListener('click',()=>{ const will=PANEL.style.display==='none'; PANEL.style.display=will?'':'none'; if(will) fillCfg(); },{passive:true});
    }
  }

  PANEL.addEventListener('click', async e=>{
    const b=e.target.closest('button[data-act]'); if(!b) return;
    if(b.dataset.act==='refresh') { const hide=(typeof fvprShowBlockingLoader==='function')?fvprShowBlockingLoader('Auswertung wird aktualisiert …'):(()=>{}); try{ await render(true); } finally{ hide(); } }
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

  PANEL.addEventListener('click', e=>{
    try {
      const sendBtn=e.target.closest('button[data-sp]');
      if(sendBtn){
        e.preventDefault();
        sendSinglePartnerConfirm(sendBtn.dataset.sp);
        return;
      }

      const eyeBtn=e.target.closest('button[data-eye]');
      if(eyeBtn){
        e.preventDefault();
        DIAG('click','👁-Button geklickt',{partner:eyeBtn.dataset.eye});
        openPreview(eyeBtn.dataset.eye).catch(err=>{
          DIAG('error','openPreview fehlgeschlagen',{err:String(err)});
          toast('Fehler: '+err.message,false);
        });
        return;
      }

      const copyBtn=e.target.closest('button[data-copy]');
      if(copyBtn){
        e.preventDefault();
        copyPartnerHtml(copyBtn.dataset.copy);
        return;
      }

      const totalEye=e.target.closest('button[data-total-eye]');
      if(totalEye){
        e.preventDefault();
        openTotalPreview().catch(err=>{
          DIAG('error','openTotalPreview fehlgeschlagen',{err:String(err)});
          toast('Fehler: '+err.message,false);
        });
        return;
      }

      const totalCopy=e.target.closest('button[data-total-copy]');
      if(totalCopy){
        e.preventDefault();
        copyTotalHtml();
        return;
      }

      const totalBtn=e.target.closest('button[data-total-send]');
      if(totalBtn){
        e.preventDefault();
        sendTotalOnlyConfirm();
        return;
      }

      const totalRow=e.target.closest('tr[data-total-row="1"]');
      if(totalRow && !e.target.closest('button')){
        e.preventDefault();
        DIAG('click','Gesamtzeile geklickt',{ tagName: e.target.tagName });
        toast('Lade Gesamtübersicht…', true);
        openTotalPreview().catch(err=>{
          DIAG('error','openTotalPreview Zeilen-Klick fehlgeschlagen',{
            err: String(err),
            stack: err?.stack?.slice?.(0,300)
          });
          toast('Gesamt-Fehler: '+err.message, false);
        });
        return;
      }

      const row=e.target.closest('tr[data-partner]');
      if(row && !e.target.closest('button')){
        e.preventDefault();
        const partner = row.getAttribute('data-partner');
        DIAG('click','Zeile geklickt',{ partner, tagName: e.target.tagName });
        toast('Lade Detail: '+partner+'…', true);
        openPreview(partner).catch(err=>{
          DIAG('error','openPreview Zeilen-Klick fehlgeschlagen',{
            err: String(err),
            stack: err?.stack?.slice?.(0,300)
          });
          toast('Detail-Fehler: '+err.message, false);
        });
      }
    } catch(err) {
      DIAG('error','Klick-Handler Exception',{
        err: String(err),
        stack: err?.stack?.slice?.(0,300)
      });
      toast('Klick-Fehler: '+err.message, false);
    }
  },{passive:false});
}

function removeStandaloneButton(){
  const wrap = document.getElementById(NS+'wrap');
  if (wrap) wrap.remove();
}

function findFahrzeuguebersichtTrigger(){
  const selectors = [
    '[role="tab"]',
    '.mat-mdc-tab',
    '.mat-tab-label',
    '.mat-mdc-tab-link',
    '.mat-tab-link',
    'button',
    'a',
    'div'
  ];

  for(const sel of selectors){
    const nodes = Array.from(document.querySelectorAll(sel)).filter(el=>!el.closest(PANEL_ID));
    for(const el of nodes){
      const txt = norm(el.textContent || '');
      if(!/^fahrzeugübersicht$/i.test(txt) && !/fahrzeugübersicht/i.test(txt)) continue;

      const clickable =
        el.closest('[role="tab"]') ||
        el.closest('.mat-mdc-tab') ||
        el.closest('.mat-tab-label') ||
        el.closest('.mat-mdc-tab-link') ||
        el.closest('.mat-tab-link') ||
        el.closest('button') ||
        el.closest('a') ||
        el;

      if(clickable) return clickable;
    }
  }
  return null;
}

function findFirstTopTabLikeElement(){
  const selectors = [
    '[role="tab"]',
    '.mat-mdc-tab',
    '.mat-tab-label',
    '.mat-mdc-tab-link',
    '.mat-tab-link'
  ];

  for(const sel of selectors){
    const nodes = Array.from(document.querySelectorAll(sel)).filter(el=>!el.closest(PANEL_ID));
    if(!nodes.length) continue;

    const visible = nodes.filter(el=>{
      const r = el.getBoundingClientRect();
      return r.width > 40 && r.height > 20 && r.top >= 0 && r.top < 250;
    });

    if(visible.length){
      visible.sort((a,b)=>a.getBoundingClientRect().left - b.getBoundingClientRect().left);
      return visible[0];
    }
  }
  return null;
}

function isLikelyActiveTab(el){
  if(!el) return false;

  const aria = String(el.getAttribute?.('aria-selected') || '').toLowerCase();
  const cls = String(el.className || '');
  if(aria === 'true') return true;
  if(/active|selected|mdc-tab--active|mat-mdc-tab-active|mat-tab-label-active/i.test(cls)) return true;

  try{
    const cs = window.getComputedStyle(el);
    const bg = String(cs.backgroundColor || '');
    const color = String(cs.color || '');
    if(
      bg === 'rgb(225, 6, 50)' ||
      bg === 'rgb(229, 0, 54)' ||
      bg === 'rgb(230, 0, 50)' ||
      color === 'rgb(255, 255, 255)'
    ) return true;
  }catch{}

  return false;
}

function isFahrzeuguebersichtActive(){
  const trigger = findFahrzeuguebersichtTrigger();
  if(trigger && isLikelyActiveTab(trigger)) return true;

  const firstTab = findFirstTopTabLikeElement();
  if(firstTab){
    const txt = norm(firstTab.textContent || '');
    if(/fahrzeugübersicht/i.test(txt) && isLikelyActiveTab(firstTab)) return true;
  }

  return false;
}

function isFahrzeuguebersichtReady(){
  const tryA = findColumnsDatagrid();
  if(tryA){
    const rows = qsaMain('tbody tr,[role="row"]').filter(tr=>{
      const cells = tr.querySelectorAll('td,[role="gridcell"]');
      const needed = Math.max(...Object.values(tryA).filter(v=>typeof v==='number'&&v>=0));
      return cells && cells.length > needed;
    });
    if(rows.length > 0) return true;
  }

  const tryB = getAnyTable();
  if(tryB){
    const trs = Array.from(tryB.querySelectorAll('tbody tr'));
    if(trs.length > 0) return true;
  }

  return false;
}

async function waitForOverviewReady(timeoutMs = 15000){
  const start = Date.now();
  while(Date.now() - start < timeoutMs){
    if(isFahrzeuguebersichtReady()) return true;
    await new Promise(r=>setTimeout(r,250));
  }
  return false;
}

function fireRealClick(el){
  if(!el) return false;

  const target =
    el.closest?.('[role="tab"]') ||
    el.closest?.('.mat-mdc-tab') ||
    el.closest?.('.mat-tab-label') ||
    el.closest?.('.mat-mdc-tab-link') ||
    el.closest?.('.mat-tab-link') ||
    el;

  try { target.scrollIntoView({ block:'center', inline:'center' }); } catch {}

  const evOpts = { bubbles:true, cancelable:true, view:window };
  try { target.dispatchEvent(new PointerEvent('pointerdown', evOpts)); } catch {}
  try { target.dispatchEvent(new MouseEvent('mousedown', evOpts)); } catch {}
  try { target.dispatchEvent(new PointerEvent('pointerup', evOpts)); } catch {}
  try { target.dispatchEvent(new MouseEvent('mouseup', evOpts)); } catch {}
  try { target.dispatchEvent(new MouseEvent('click', evOpts)); } catch {}
  try { if(typeof target.click === 'function') target.click(); } catch {}

  return true;
}

async function ensureFahrzeuguebersichtActive(){
  if(isFahrzeuguebersichtActive() && isFahrzeuguebersichtReady()) return true;

  let trigger = findFahrzeuguebersichtTrigger();

  if(!trigger){
    const firstTab = findFirstTopTabLikeElement();
    if(firstTab){
      const txt = norm(firstTab.textContent || '');
      if(/fahrzeugübersicht/i.test(txt)) trigger = firstTab;
    }
  }

  if(!trigger){
    console.warn('[fvpr] Fahrzeugübersicht-Tab nicht gefunden');
    return false;
  }

  fireRealClick(trigger);
  await new Promise(r=>setTimeout(r,1000));

  if(!isFahrzeuguebersichtActive()){
    const firstTab = findFirstTopTabLikeElement();
    if(firstTab){
      fireRealClick(firstTab);
      await new Promise(r=>setTimeout(r,1000));
    }
  }

  const ok = await waitForOverviewReady(15000);
  if(!ok) console.warn('[fvpr] Fahrzeugübersicht wurde nicht rechtzeitig geladen');
  return ok;
}

function modal(html){
  DIAG("modal", "Modal wird erstellt", { htmlSnippet: (html||"").slice(0,100) });
  const ov=document.createElement('div');
  ov.className=NS+'modal';
  ov.innerHTML=`<div class="${NS}modal-box">${html}</div>`;

  // Klick auf den dunklen Bereich außerhalb des aktuellen Fensters schließt nur dieses Fenster.
  ov.addEventListener('mousedown', e=>{
    if(e.target === ov) ov.remove();
  }, {passive:true});

  document.body.appendChild(ov);
  DIAG("modal", "Modal ins DOM eingefügt", { display: getComputedStyle(ov).display, zIndex: getComputedStyle(ov).zIndex, rect: JSON.stringify(ov.getBoundingClientRect()) });
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
      <tr data-row="1" data-partner="${String(partner).replace(/"/g,'&quot;')}" data-tour="${String(r.tour||'').replace(/"/g,'&quot;')}">
        <td style="padding:6px 8px;border:1px solid #e5e7eb;">${r.tour||'—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${r.driver||'—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;"><span class="${NS}numlink">${fmtInt(r.stops)}</span></td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;"><span class="${NS}numlink">${fmtInt(r.pkgs)}</span></td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;"><span class="${NS}numlink">${fmtInt(r.open)}</span></td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;"><span class="${NS}numlink">${fmtInt(r.obstacles)}</span></td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;${etaStyle}">${fmtPct(r.eta)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;"><span class="${NS}numlink">${fmtInt(r.pOpen)}</span></td>
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
    <table class="${NS}tbl" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
      <thead>
        <tr data-total-row="1" style="background:#e0f2ff;color:#003366;font-weight:700;font-size:12px;cursor:pointer;">
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Touren: ${fmtInt(totals.tours)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">ETA Ø: ${fmtPct(totals.etaAvg)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
        </tr>
        <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Tour</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Fahrername</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">offen</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">offene Abholstops</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
    ${signatureHtml}
  </div>`;
}

function summaryHtml(per,totals,signature){
  const head=`
    <thead>
      <tr data-total-row="1" style="background:#e0f2ff;color:#003366;font-weight:700;font-size:12px;cursor:pointer;">
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:left;">Gesamt (alle)</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.tours)}</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(totals.etaAvg)}</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
        <td style="padding:8px 10px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
        <td class="${NS}act" style="padding:8px 10px;border:1px solid #e5e7eb;text-align:center;">
          <button class="${NS}iconbtn" title="Gesamtübersicht an „gesamt“ (Bestätigung)" data-total-send="1">✉︎</button>
          <button class="${NS}iconbtn" title="Gesamt-Vorschau öffnen" data-total-eye="1">👁</button>
          <button class="${NS}iconbtn" title="Gesamt-HTML in Zwischenablage" data-total-copy="1">⧉</button>
        </td>
      </tr>
      <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
        <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Systempartner</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Touren</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA % (Ø)</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps gesamt</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Offene Stopps</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete gesamt</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Offene Abholstops</th>
        <th style="padding:6px 8px;border:1px solid #e5e7eb;">Aktion</th>
      </tr>
    </thead>`;

  const body=per.map(p=>{
    const etaStyle=`background:${etaBg(p.etaAvg)};`;
    return `
      <tr data-partner="${p.partner.replace(/"/g,'&quot;')}" data-row="1">
        <td style="text-align:left">${p.partner}</td>
        <td>${fmtInt(p.tours)}</td>
        <td style="${etaStyle}">${fmtPct(p.etaAvg)}</td>
        <td><span class="${NS}numlink">${fmtInt(p.stops)}</span></td>
        <td><span class="${NS}numlink">${fmtInt(p.open)}</span></td>
        <td><span class="${NS}numlink">${fmtInt(p.pkgs)}</span></td>
        <td><span class="${NS}numlink">${fmtInt(p.obstacles)}</span></td>
        <td><span class="${NS}numlink">${fmtInt(p.pOpen)}</span></td>
        <td class="${NS}act">
          <button class="${NS}iconbtn" title="Mail an Partner (Bestätigung)" data-sp="${p.partner}">✉︎</button>
          <button class="${NS}iconbtn" title="Vorschau öffnen" data-eye="${p.partner}">👁</button>
          <button class="${NS}iconbtn" title="Partner-HTML in Zwischenablage" data-copy="${p.partner}">⧉</button>
        </td>
      </tr>`;
  }).join('');

  const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';

  return `
  <div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
    <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
    <table class="${NS}tbl" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
      ${head}<tbody>${body}</tbody>
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
          <tr style="background:#e0f2ff;color:#003366;font-weight:700;font-size:12px;">
            <td style="padding:8px;border:1px solid #e5e7eb;"></td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Touren: ${fmtInt(totals.tours)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">ETA Ø: ${fmtPct(totals.etaAvg)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
          </tr>
          <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
            <th align="center" style="padding:8px;border:1px solid #e5e7eb;">Tour</th>
            <th align="left" style="padding:8px;border:1px solid #e5e7eb;">Fahrername</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Stopps</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Pakete</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">offen</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">ETA</th>
            <th align="right" style="padding:8px;border:1px solid #e5e7eb;">offene Abholstops</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
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
          <tr style="background:#e0f2ff;color:#003366;font-weight:700;font-size:12px;">
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt (alle)</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.tours)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(totals.etaAvg)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
            <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
          </tr>
          <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
            <th align="left" style="padding:8px;border:1px solid #e5e7eb;">Systempartner</th>
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
      </table>
    </div>
    ${signatureHtml}
  </div>`;
}

function totalToursHtml(rows, signature){
  const sorted = [...(rows || [])].sort((a,b)=>{
    const pa = String(a.partner || '').localeCompare(String(b.partner || ''), 'de');
    if (pa !== 0) return pa;
    return String(a.tour || '').localeCompare(String(b.tour || ''), 'de', { numeric:true });
  });

  const body = sorted.map(r=>{
    const etaStyle = `background:${etaBg(r.eta)};`;
    return `
      <tr data-row="1">
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${r.partner || '—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;">${r.tour || '—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${r.driver || '—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.stops)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pkgs)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.open)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.obstacles)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;${etaStyle}">${fmtPct(r.eta)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pOpen)}</td>
      </tr>`;
  }).join('');

  const totals = {
    tours: sorted.length,
    etaAvg: avg(sorted, r=>r.eta),
    stops: sum(sorted, r=>r.stops),
    pkgs: sum(sorted, r=>r.pkgs),
    open: sum(sorted, r=>r.open),
    obstacles: sum(sorted, r=>r.obstacles),
    pOpen: sum(sorted, r=>r.pOpen)
  };

  const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';

  return `
  <div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
    <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
    <table class="${NS}tbl" cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
      <thead>
        <tr data-total-row="1" style="background:#e0f2ff;color:#003366;font-weight:700;font-size:12px;">
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt (alle)</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.tours)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">—</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">ETA Ø: ${fmtPct(totals.etaAvg)}</td>
          <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
        </tr>
        <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
          <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Systempartner</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Tour</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Fahrername</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">offen</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA</th>
          <th style="padding:6px 8px;border:1px solid #e5e7eb;">offene Abholstops</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
    ${signatureHtml}
  </div>`;
}

/* ---------- Aggregation + Flows ---------- */
async function getAggregates(){
  DIAG('agg', 'getAggregates gestartet');

  const tabReady = await ensureFahrzeuguebersichtActive();
  DIAG('agg', 'ensureFahrzeuguebersichtActive', { tabReady });

  const {ok,rows}=await readAllRows();
  DIAG('agg', 'readAllRows Ergebnis', { ok, rowCount: rows.length });
  if(!ok||rows.length===0) return null;

  cacheTourPartner(rows);

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
  DIAG("detail", "openTotalPreview aufgerufen");
  const hideLoader = (typeof fvprShowBlockingLoader === 'function') ? fvprShowBlockingLoader('Gesamt-Vorschau wird geladen …') : (()=>{});
  try {
    const agg = await getAggregates();
    if(!agg){
      alert('Keine Daten gefunden.');
      return;
    }

    const g = await ensureSettingsRecord();
  const allRows = agg.per.flatMap(p => p.list || []);
  const html = totalToursHtml(allRows, g.signature || '');

  const ov = modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Vorschau – Gesamt (alle Touren)</h3>
    <div style="max-height:60vh;overflow:auto;border:1px solid #e5e7eb;border-radius:8px;padding:8px;margin-bottom:10px;background:#fff">${html}</div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="copy">Kopieren</button>
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>
  `);

  const innerTable = ov.querySelector('table');
  if(innerTable) makeTableSortable(innerTable.parentElement);

  ov.addEventListener('click', async e=>{
    const btn = e.target.closest('button[data-act]');
    if(!btn) return;

    if(btn.dataset.act === 'copy'){
      const ok = await copyHtmlToClipboard(html);
      toast(ok ? 'Gesamt kopiert' : 'Kopieren fehlgeschlagen', ok);
    }

    if(btn.dataset.act === 'close') ov.remove();
  }, {passive:false});
  } finally {
    hideLoader();
  }
}

async function copyTotalHtml(){
  const agg = await getAggregates(); if(!agg) return;
  const g = await ensureSettingsRecord();
  const allRows = agg.per.flatMap(p => p.list || []);
  const html = totalToursHtml(allRows, g.signature || '');
  const ok = await copyHtmlToClipboard(html);
  toast(ok ? 'Gesamt in Zwischenablage' : 'Kopieren fehlgeschlagen', ok);
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
  DIAG("detail", "openPreview aufgerufen", { partner });
  const hideLoader = (typeof fvprShowBlockingLoader === 'function') ? fvprShowBlockingLoader(`Vorschau wird geladen – ${partner} …`) : (()=>{});
  try {
    const agg=await getAggregates();
    if(!agg){ DIAG("detail","getAggregates lieferte null"); alert('Keine Daten gefunden.'); return; }
    const g=await ensureSettingsRecord();
    const p=agg.per.find(x=>x.partner===partner);
    if(!p){ DIAG("detail", "Partner nicht in Aggregates gefunden", { partner, verfuegbar: agg.per.map(x=>x.partner) }); alert('Partner nicht gefunden.'); return; }
    DIAG("detail", "Partner gefunden", { partner, touren: p.tours, rows: p.list.length });
    const content=partnerHtml(partner, p.list, g.signature||'');
    DIAG("detail", "partnerHtml erzeugt", { htmlLen: content.length });
    const ov=modal(`
      <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Vorschau – ${partner}</h3>
      <div style="max-height:60vh;overflow:auto;border:1px solid #e5e7eb;border-radius:8px;padding:8px;margin-bottom:10px;background:#fff">${content}</div>
      <div class="${NS}modal-actions">
        <button class="${NS}btn-sm" data-act="copy">Kopieren</button>
        <button class="${NS}btn-sm" data-act="edit">Einstellungen</button>
        <button class="${NS}btn-sm" data-act="close">Schließen</button>
      </div>
    `);
    const ovRect = ov.getBoundingClientRect();
    const ovStyle = getComputedStyle(ov);
    DIAG("detail", "Modal nach Erstellung", {
      inDOM: document.body.contains(ov),
      display: ovStyle.display,
      visibility: ovStyle.visibility,
      opacity: ovStyle.opacity,
      zIndex: ovStyle.zIndex,
      rect: `${Math.round(ovRect.width)}x${Math.round(ovRect.height)} @ ${Math.round(ovRect.left)},${Math.round(ovRect.top)}`,
      childNodes: ov.childNodes.length,
      boxHTML: ov.querySelector('.'+NS+'modal-box')?.innerHTML?.length || 0,
    });
    if (ovRect.width === 0 || ovRect.height === 0) {
      DIAG("detail", "⚠ Modal hat Größe 0! Erzwinge Inline-Styles");
      ov.style.cssText = 'position:fixed!important;top:0!important;left:0!important;right:0!important;bottom:0!important;width:100vw!important;height:100vh!important;background:rgba(0,0,0,.35)!important;display:flex!important;align-items:center!important;justify-content:center!important;z-index:2147483647!important;';
    }
    const innerTable = ov.querySelector('table');
    if(innerTable) makeTableSortable(innerTable.parentElement);
    ov.addEventListener('click', async e=>{
      const btn=e.target.closest('button[data-act]'); if(!btn) return;
      if(btn.dataset.act==='copy'){
        const ok=await copyHtmlToClipboard(content);
        toast(ok?'Vorschau kopiert':'Kopieren fehlgeschlagen', ok);
      }
      if(btn.dataset.act==='close') ov.remove();
      if(btn.dataset.act==='edit'){ ov.remove(); openPartnerDialog(partner); }
    },{passive:false});
  } catch(err) {
    DIAG("error", "openPreview EXCEPTION", { err: String(err), stack: err?.stack?.slice?.(0,400) });
    console.error('[fvpr] openPreview', err);
    toast('Detail-Fehler: '+err.message, false);
  } finally {
    hideLoader();
  }
}

function openConfirm({title, subject, to, cc, saveKey, onOk, htmlToCopy=''}){
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">${title||'Bestätigen'}</h3>
    <div style="display:grid;grid-template-columns:1fr 2fr;gap:8px;margin:8px 0">
      <label>Betreff</label><input id="${NS}cf-subj" type="text" value="${subject||''}">
      <label>An</label><input id="${NS}cf-to" type="text" value="${to||''}">
      <label>CC</label><input id="${NS}cf-cc" type="text" value="${cc||''}">
    </div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="copy">Kopieren</button>
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
    if(b.dataset.act==='copy'){
      if(!htmlToCopy){ toast('Nichts zum Kopieren', false); return; }
      const ok=await copyHtmlToClipboard(htmlToCopy);
      toast(ok?'HTML kopiert':'Kopieren fehlgeschlagen', ok);
      return;
    }
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
      <button class="${NS}btn-sm" data-act="copy-mails">Adressen kopieren</button>
      <button class="${NS}btn-sm" data-act="save">Speichern</button>
      <button class="${NS}btn-sm" data-act="clear">Löschen</button>
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>
  `);
  ov.addEventListener('click', async e=>{
    const btn=e.target.closest('button[data-act]'); if(!btn) return;
    if(btn.dataset.act==='close'){ ov.remove(); return; }
    if(btn.dataset.act==='clear'){ await idbDel('partners',key); ov.remove(); return; }
    if(btn.dataset.act==='copy-mails'){
      const to=ov.querySelector('#'+NS+'pd-to').value.trim();
      const cc=ov.querySelector('#'+NS+'pd-cc').value.trim();
      try{
        await navigator.clipboard.writeText(`An: ${to}\nCC: ${cc}`);
        toast('Adressen kopiert');
      }catch{
        toast('Kopieren fehlgeschlagen', false);
      }
      return;
    }
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
    subject, to:rec.to||'', cc:rec.cc||'', saveKey:partner, htmlToCopy:html,
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
    subject, to:rec.to||'', cc:rec.cc||'', saveKey:'gesamt', htmlToCopy:html,
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
  const htmlTot=mailSummaryHtml(agg.per, agg.totals, g.signature||'');

  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Sammelversand bestätigen</h3>
    <div style="font:13px;margin-bottom:10px">Es werden <b>${ready.length}</b> Partner-Mails gesendet (nur mit gültiger Adresse) und <b>1</b> Gesamt-Mail an „gesamt“.</div>
    <div style="display:grid;grid-template-columns:1fr 2fr;gap:8px;margin:8px 0">
      <label>Gesamt – Betreff</label><input id="${NS}sammel-subj" type="text" value="${subjectTotal}">
      <label>Gesamt – An</label><input id="${NS}sammel-to" type="text" value="${rec.to||''}">
      <label>Gesamt – CC</label><input id="${NS}sammel-cc" type="text" value="${rec.cc||''}">
    </div>
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="copy-total">Gesamt kopieren</button>
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
    if(b.dataset.act==='copy-total'){
      const ok=await copyHtmlToClipboard(htmlTot);
      toast(ok?'Gesamt kopiert':'Kopieren fehlgeschlagen', ok);
      return;
    }
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
    DIAG('render', 'run() gestartet', { contentExists: !!CONTENT });
    if(!CONTENT) return;
    const res=await getAggregates();
    if(!res){
      CONTENT.innerHTML=`<div class="${NS}empty">Keine Daten gefunden (Tab „Fahrzeugübersicht“ sichtbar?).</div>`;
      return;
    }
    const {per, totals}=res;
    const html = summaryHtml(per, totals, '');
    CONTENT.innerHTML = html;
    makeTableSortable(CONTENT.querySelector('table'));
    DIAG('render', 'Panel-Inhalt gerendert', { htmlLen: html.length });
  };
  if(force) await run(); else renderTimer=setTimeout(run, RENDER_DEBOUNCE);
}

/* ====== TEIL 12/12 – Boot + Loader-Integration ====== */

async function openPanel(){
  DIAG("boot", "openPanel() aufgerufen");
  try { ensureStyles(); } catch {}
  try { if (!document.querySelector(PANEL_ID)) mountUI(true); } catch {}
  const p = document.querySelector(PANEL_ID);
  DIAG("boot", "Panel-Element", { found: !!p, display: p?.style?.display, id: p?.id, parentTag: p?.parentElement?.tagName });
  if (p) p.style.removeProperty('display');

  try { await ensureFahrzeuguebersichtActive(); } catch(e) { console.error('[fvpr] ensureFahrzeuguebersichtActive', e); }
  try { fillCfg(); } catch {}
  try { await render(true); } catch {}
  DIAG("boot", "openPanel() abgeschlossen", { panelVisible: p ? getComputedStyle(p).display !== 'none' : false });
}

function closePanel(){
  const p = document.querySelector(PANEL_ID);
  if (p) p.style.setProperty('display', 'none', 'important');
}

function installOutsideClickClose(){
  if (window.__FVPR_OUTSIDE_CLOSE_INSTALLED) return;
  window.__FVPR_OUTSIDE_CLOSE_INSTALLED = true;

  document.addEventListener('mousedown', e=>{
    const panel = document.querySelector(PANEL_ID);
    if(!panel) return;
    if(getComputedStyle(panel).display === 'none') return;

    // Wenn ein Vorschau-/Einstellungsfenster offen ist, kümmert sich dieses selbst ums Schließen.
    // Dadurch wird beim Klick auf den dunklen Bereich nicht zusätzlich das Hauptfenster geschlossen.
    if(e.target.closest && e.target.closest('.' + NS + 'modal')) return;

    // Klick im Hauptfenster: offen lassen. Klick irgendwo daneben: Hauptfenster schließen.
    if(e.target.closest && e.target.closest(PANEL_ID)) return;

    closePanel();
  }, true);
}

async function bootStandalone(){
  DIAG("boot", "bootStandalone()");
  try { ensureStyles(); mountUI(false); await ensureSettingsRecord(); await ensureFahrzeuguebersichtActive(); await render(true); }
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
installOutsideClickClose();

function registerWithLoader(){
  const def = {
    id: 'partner-report',
    label: 'Partner-Report',
    panels: [PANEL_ID],
    run: async () => {
      const el = document.querySelector(PANEL_ID);
      if (el && getComputedStyle(el).display !== 'none') closePanel();
      else await openPanel();
    },
    close: () => closePanel(),
    isOpen: () => {
      const el = document.querySelector(PANEL_ID);
      return !!el && el.offsetParent !== null && getComputedStyle(el).display !== 'none';
    }
  };
  const G = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
  if (G.TM && typeof G.TM.register === 'function'){
    G.TM.register(def);
    removeStandaloneButton();
  } else {
    const KEY='__tmQueue';
    G[KEY] = Array.isArray(G[KEY]) ? G[KEY] : [];
    G[KEY].push(def);
  }
  G.fvpr_open = openPanel;
  G.fvpr_close = closePanel;
}

DIAG('boot', 'Script initialisiert', {
  ua: navigator.userAgent,
  isEdge: /Edg\//.test(navigator.userAgent),
  isChrome: /Chrome\//.test(navigator.userAgent) && !/Edg\//.test(navigator.userAgent),
  url: location.href,
  screen: `${screen.width}x${screen.height}`,
  innerSize: `${innerWidth}x${innerHeight}`,
  dpr: devicePixelRatio,
  tmLoaderAvail: !!(window.TM && typeof window.TM.register === 'function'),
});

const G_BOOT = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;
if (G_BOOT.TM && typeof G_BOOT.TM.register === 'function') {
  registerWithLoader();
} else if (Array.isArray(G_BOOT.__tmQueue)) {
  registerWithLoader();
} else {
  G_BOOT['__tmQueue'] = G_BOOT['__tmQueue'] || [];
  registerWithLoader();

  setTimeout(() => {
    if (!(G_BOOT.TM && typeof G_BOOT.TM.register === 'function')) {
      LOG('Loader nicht erschienen – starte standalone');
      bootStandalone();
    }
  }, 3000);
}



/* ====== TEIL 13/13 – Klick-Details Stopps/Pakete/offen bis Paket-Lebenslauf ======
   Übernommen/angepasst aus der Tourenauswertung: liest /dispatcher/api/pickup-delivery,
   baut Tour-/Stopp-/Paketlisten und öffnet Paket-Lebenslauf per Paketnummer. */
const FVPR_DETAIL_PAGE_SIZE = 500;
const FVPR_DETAIL_HARD_MAX_PAGES = 300;
const FVPR_DETAIL_CACHE_MS = 60_000;
let FVPR_LAST_PD_REQUEST = null;
let FVPR_DETAIL_CACHE = { key:'', ts:0, rows:[] };

function escHtml(s){ return String(s ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch])); }
function fvprIsoToday(){ const d=new Date(); return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; }
function fvprStatusNorm(s){ return norm(s).toUpperCase().replace(/[_/\\|-]+/g,' ').replace(/\s+/g,' ').trim(); }
function fvprTypeOf(r){
  // Muss exakt zwischen Zustellung und Abholung trennen, sonst zeigt ein Klick auf offene Abholstopps falsche Daten.
  const pickupVals = [r?.pickup_status,r?.pickupStatus,r?.pickup_type,r?.pickupType,r?.pickupTime,r?.pickedUpTime,r?.pickupDateTime,r?.pickupTimestamp,r?.pickup?.time,r?.pickup?.dateTime];
  if(pickupVals.some(v=>norm(v))) return 'PICKUP';

  const deliveryVals = [r?.delivery_status,r?.deliveryStatus,r?.deliveredTime,r?.delivered_time,r?.deliveryTime,r?.deliveryDateTime,r?.deliveryTimestamp,r?.delivery?.time,r?.delivery?.dateTime];
  if(deliveryVals.some(v=>norm(v))) return 'DELIVERY';

  const txt = [r?.order_type,r?.orderType,r?.type,r?.stopType,r?.jobType,r?.missionType,r?.deliveryOrPickup,r?.kind,r?.category,r?.serviceType,r?.service_type].map(x=>String(x||'').toUpperCase()).join(' | ');
  if(/PICKUP|ABHOL/.test(txt)) return 'PICKUP';
  return 'DELIVERY';
}
function fvprStatusOf(r){
  const vals=[
    r?.delivery_status,r?.pickup_status,r?.deliveryStatus,r?.pickupStatus,
    r?.statusDisplay,r?.statusLabel,r?.statusDescription,r?.statusName,r?.statusText,r?.parcelStatus,
    typeof r?.status==='string'?r.status:'',typeof r?.stopStatus==='string'?r.stopStatus:'',typeof r?.orderStatus==='string'?r.orderStatus:'',
    r?.processingStatus,r?.scanStatus,r?.tourStatus,r?.reasonText,r?.statusReason,r?.problemReason,r?.problem_reason,
    r?.delivery?.status,r?.pickup?.status,r?.order?.status,r?.stop?.status
  ];
  for(const v of vals){ const x=norm(v); if(x) return x; }
  return '';
}
function fvprStatusDe(s){
  const n=fvprStatusNorm(s);
  const map={
    'DELIVERED':'ZUGESTELLT','DELIVERED TO PUDO':'ZUGESTELLT PUDO','DELIVERED_TO_PUDO':'ZUGESTELLT PUDO','ZUGESTELLT':'ZUGESTELLT',
    'PICKED UP':'ABGEHOLT','PICKED_UP':'ABGEHOLT','ABGEHOLT':'ABGEHOLT','COMPLETED':'ABGEHOLT',
    'DELIVERY PROBLEM':'ZUSTELLUNG PROBLEM','DELIVERY_PROBLEM':'ZUSTELLUNG PROBLEM','NOT DELIVERED':'NICHT ZUGESTELLT','NICHT ZUGESTELLT':'NICHT ZUGESTELLT',
    'DELIVERY CANCELLED AUTOMATICALLY':'ZUSTELLUNG STORNIERT (AUTOMATISCH)','PICKUP CANCELLED AUTOMATICALLY':'ABHOLUNG STORNIERT (AUTOMATISCH)','CANCELLED':'STORNIERT'
  };
  return map[n] || norm(s) || '—';
}
function fvprHasAnyRawValue(r, keys){
  const raw=r?.__raw||r||{};
  for(const k of keys){
    try{
      const v=k.split('.').reduce((a,p)=>a?.[p], raw);
      if(norm(v)) return true;
    }catch{}
  }
  return false;
}
function fvprIsDelivered(r){
  const n=fvprStatusNorm(r.__statusRaw||r.__status);
  return n==='DELIVERED'||n==='ZUGESTELLT'||n==='DELIVERED TO PUDO'||n==='DELIVERED_TO_PUDO'||n==='DELIVERED PUDO'||
    fvprHasAnyRawValue(r,['deliveredTime','delivered_time','deliveryTime','deliveryDateTime','deliveryTimestamp','delivery.time','delivery.dateTime']);
}
function fvprIsPicked(r){
  const n=fvprStatusNorm(r.__statusRaw||r.__status);
  return n==='PICKED UP'||n==='PICKED_UP'||n==='ABGEHOLT'||n==='COMPLETED'||
    fvprHasAnyRawValue(r,['pickupTime','pickedUpTime','pickupDateTime','pickupTimestamp','pickup.time','pickup.dateTime']);
}
function fvprIsCanceled(r){ const n=fvprStatusNorm(r.__statusRaw||r.__status); return /CANCEL|STORNI|AUTOMAT/.test(n); }
function fvprIsProblemDelivery(r){
  const n=fvprStatusNorm(r.__statusRaw||r.__status);
  const hasCode=!!norm(r.__additionalCode||r.__raw?.additional_code||r.__raw?.additionalCode||r.__raw?.problemReason||r.__raw?.problem_reason);
  return r.__type==='DELIVERY' && !fvprIsCanceled(r) && (/PROBLEM|HINDERN|NICHT|FAILED|REFUSED|NO ACCESS|ADDRESS|ADRESS|RETOUR|CLOSED|GESCHLOSSEN/.test(n) || hasCode);
}
function fvprIsOpenDelivery(r){
  // Offen bedeutet hier wirklich nur noch offene Zustellstopps: Zustellung, nicht zugestellt, nicht storniert, kein Zustellhindernis.
  return r.__type==='DELIVERY' && !fvprIsDelivered(r) && !fvprIsCanceled(r) && !fvprIsProblemDelivery(r);
}
function fvprIsOpenPickup(r){
  // Offen bedeutet hier wirklich nur noch offene Abholstopps: Abholung, nicht abgeholt, nicht storniert.
  return r.__type==='PICKUP' && !fvprIsPicked(r) && !fvprIsCanceled(r);
}
function fvprExtractTour(r){
  const vals=[r?.tour,r?.round,r?.route,r?.tourNo,r?.tourNumber,r?.routeNumber,r?.routeNo,r?.roundNo,r?.tripNo,r?.tripNumber,r?.tour?.number,r?.tour?.tourNumber,r?.tour?.routeNumber,r?.tour?.name,r?.route?.number,r?.route?.name,r?.stop?.tour,r?.stop?.tourNumber];
  for(const v of vals){ const x=norm(v); if(x) return x; }
  return '';
}
function fvprRawPartner(r){
  const vals=[r?.systemPartner,r?.systempartner,r?.systemPartnerName,r?.partner,r?.partnerName,r?.servicePartner,r?.servicePartnerName,r?.carrierName,r?.transportPartnerName,r?.contractorName,r?.subcontractor,r?.subcontractorName,r?.subcontractor_name,r?.tour?.partnerName,r?.tour?.systemPartner,r?.route?.partnerName];
  for(const v of vals){ const x=norm(v); if(x) return x; }
  return '';
}
function fvprDriverOf(r){
  const vals=[r?.driverName,r?.driver,r?.courierName,r?.courier_name,r?.courier,r?.employeeName,r?.driverFullName,r?.courierFullName,r?.vehicleDriverName,r?.vehicle_driver_name,r?.chauffeurName,r?.deliveryDriverName,r?.pickupDriverName,r?.driver?.name,r?.courier?.name,r?.employee?.name,r?.vehicle?.driverName,r?.vehicle?.driver?.name,r?.tour?.driverName,r?.tour?.driver,r?.route?.driverName];
  for(const v of vals){ const x=norm(v); if(x) return x; }
  return '';
}
function fvprServiceCodeOf(r){
  const out=new Set();
  const add=v=>{ if(v==null) return; String(v).split(/[^\dA-Za-z]+/).map(x=>x.trim()).filter(Boolean).forEach(x=>out.add(x)); };
  const arr=a=>Array.isArray(a)&&a.forEach(add);
  add(r?.serviceCode); add(r?.servicecode); add(r?.service_code); arr(r?.serviceCodes);
  if(r?.service&&typeof r.service==='object'){ add(r.service.code); add(r.service.serviceCode); add(r.service.id); arr(r.service.serviceCodes); }
  if(r?.product&&typeof r.product==='object'){ add(r.product.serviceCode); add(r.product.code); add(r.product.id); arr(r.product.serviceCodes); }
  return Array.from(out).sort((a,b)=>String(a).localeCompare(String(b),'de',{numeric:true})).join(' ');
}
function fvprPriorityKind(r){
  const txt=fvprStatusNorm([r.__serviceCode,r.__raw?.priority,r.__raw?.service_category,r.__raw?.serviceCategory,r.__raw?.service_type,r.__raw?.serviceType,r.__raw?.elements].join(' '));
  if(txt.includes('EXPRESS')) return 'EXPRESS';
  if(txt.includes('PRIO')) return 'PRIO';
  return '';
}
function fvprParcelListOf(r){
  const vals=[r?.parcel_number,r?.parcelNumber,r?.parcelNumbers,r?.parcels,r?.parcelNumberList,r?.parcelsList,r?.shipmentNumbers,r?.labels,r?.barcodes,r?.packages,r?.consignments,r?.shipments,r?.completed_parcel,r?.removed_parcel_numbers];
  const out=[];
  const push=v=>{
    if(v==null||v==='') return;
    if(typeof v==='string'||typeof v==='number'){
      const t=String(v).trim();
      if(/[;,|]/.test(t)){ t.split(/[;,|]/).forEach(push); return; }
      let x=t.replace(/\D+/g,'');
      if(!x) return; if(x.length===13) x='0'+x; if(x.length>=8) out.push(x); return;
    }
    if(Array.isArray(v)){ v.forEach(push); return; }
    if(typeof v==='object') [v.parcelNumber,v.number,v.psn,v.shipmentNumber,v.barcode,v.labelNumber,v.consignmentNumber].forEach(push);
  };
  vals.forEach(push); return [...new Set(out)];
}
function fvprParcelCountOf(r){ const l=fvprParcelListOf(r); if(l.length) return l.length; for(const v of [r?.estimated_parcels,r?.completed_parcel,r?.parcelCount,r?.parcelsCount,r?.numberOfParcels,r?.numberOfPackages,r?.packageCount,r?.packagesCount,r?.shipmentCount,r?.quantity,r?.qty,r?.pieces,r?.pieceCount,r?.itemCount,r?.count,r?.totalParcels,r?.totalPackages]){ const n=Number(v); if(Number.isFinite(n)&&n>0) return Math.trunc(n); } return 1; }
function fvprAddrOf(r){
  const street=norm(r?.street||r?.addressLine1||r?.address||r?.address?.street||r?.recipient?.street||'');
  const house=norm(r?.houseno||r?.houseNo||r?.houseNumber||r?.address?.houseNumber||r?.recipient?.houseNumber||'');
  const postal=norm(r?.postalCode||r?.zipCode||r?.zip||r?.postal_code||r?.address?.postalCode||r?.recipient?.postalCode||'');
  const city=norm(r?.city||r?.town||r?.address?.city||r?.recipient?.city||'');
  return [[street,house].filter(Boolean).join(' '),[postal,city].filter(Boolean).join(' ')].filter(Boolean).join(' · ') || '—';
}
function fvprNameOf(r){ return norm(r?.name||r?.name1||r?.customer_name||r?.customerName||r?.recipient?.name||r?.consigneeName||r?.consignee_name||''); }
function fvprStopOf(r,idx){ return norm(r?.stop||r?.stopNo||r?.stopNumber||r?.sequence||r?.sequenceNo||r?.stop?.number||'') || String(idx+1); }
function fvprPickArray(payload){
  if(Array.isArray(payload)) return payload;
  if(payload&&Array.isArray(payload.items)) return payload.items;
  if(payload&&Array.isArray(payload.content)) return payload.content;
  if(payload&&payload.data){ if(Array.isArray(payload.data)) return payload.data; if(Array.isArray(payload.data.items)) return payload.data.items; if(Array.isArray(payload.data.content)) return payload.data.content; }
  if(payload&&Array.isArray(payload.results)) return payload.results;
  if(payload&&payload._embedded){ const v=Object.values(payload._embedded).find(Array.isArray); if(Array.isArray(v)) return v; }
  return [];
}
function fvprBuildHeaders(h){
  const H=new Headers();
  try{ Object.entries(h||{}).forEach(([k,v])=>{ const key=String(k).toLowerCase(); if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)) H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()),v); }); }catch{}
  if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
  if(!H.has('Authorization')){ const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/); if(m) H.set('Authorization','Bearer '+decodeURIComponent(m[1])); }
  return H;
}
function fvprBuildPickupDeliveryUrl(page){
  const base = FVPR_LAST_PD_REQUEST?.url ? new URL(FVPR_LAST_PD_REQUEST.url.href) : new URL('/dispatcher/api/pickup-delivery', location.origin);
  base.pathname = '/dispatcher/api/pickup-delivery';
  base.search = '';
  base.searchParams.set('page', String(page));
  base.searchParams.set('pageSize', String(FVPR_DETAIL_PAGE_SIZE));
  base.searchParams.set('dateFrom', fvprIsoToday());
  base.searchParams.set('dateTo', fvprIsoToday());
  base.searchParams.set('_ts', String(Date.now()+page));
  return base;
}
async function fvprFetchDetailRows(force=false){
  const key=fvprIsoToday();
  if(!force && FVPR_DETAIL_CACHE.key===key && Date.now()-FVPR_DETAIL_CACHE.ts < FVPR_DETAIL_CACHE_MS) return FVPR_DETAIL_CACHE.rows;
  const headers=fvprBuildHeaders(FVPR_LAST_PD_REQUEST?.headers||{});
  const raw=[]; const seen=new Set();
  for(let page=1; page<=FVPR_DETAIL_HARD_MAX_PAGES; page++){
    const url=fvprBuildPickupDeliveryUrl(page);
    const res=await fetch(url.toString(), {credentials:'include', headers, cache:'no-store'});
    if(!res.ok) throw new Error(`Detaildaten nicht geladen: HTTP ${res.status}`);
    const arr=fvprPickArray(await res.json());
    if(!arr.length) break;
    for(const item of arr){ const k=String(item?.id ?? `${item?.tour||''}|${item?.parcel_number||item?.parcelNumber||''}|${page}|${raw.length}`); if(seen.has(k)) continue; seen.add(k); raw.push(item); }
    if(arr.length < FVPR_DETAIL_PAGE_SIZE) break;
  }
  const tm=await idbAll('tourMap').catch(()=>[]);
  const pMap=new Map(), dMap=new Map();
  for(const r of tm){ const k=tourKey(r.tour); if(k&&r.partner&&!pMap.has(k)) pMap.set(k,norm(r.partner)); if(k&&(r.driver||r.driverName)&&!dMap.has(k)) dMap.set(k,norm(r.driver||r.driverName)); }
  const rows=raw.map((r,idx)=>{
    const tour=fvprExtractTour(r)||'—'; const tk=tourKey(tour);
    const statusRaw=fvprStatusOf(r);
    const o={ __raw:r, __idx:idx, __tour:tour, __partner:fvprRawPartner(r)||pMap.get(tk)||'Ohne Zuordnung', __driver:fvprDriverOf(r)||dMap.get(tk)||'', __type:fvprTypeOf(r), __statusRaw:statusRaw, __status:fvprStatusDe(statusRaw), __serviceCode:fvprServiceCodeOf(r), __additionalCode:norm(r?.additional_code||r?.additionalCode||r?.problemReason||r?.problem_reason||''), __parcelList:fvprParcelListOf(r), __pkgCount:fvprParcelCountOf(r), __addr:fvprAddrOf(r), __name:fvprNameOf(r), __stop:fvprStopOf(r,idx) };
    return o;
  });
  FVPR_DETAIL_CACHE={key,ts:Date.now(),rows};
  return rows;
}
function fvprOpenTracking(psn){ const id=String(psn||'').replace(/\D+/g,''); if(id) window.open(`https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${id}`,'_blank','noopener'); }
function fvprPsnButtons(list){
  const arr=(list&&list.length?list:['—']);
  return `<div class="${NS}psn-wrap">${arr.map(psn=> psn==='—' ? '—' : `<button class="${NS}psn-btn" data-fvpr-psn="${escHtml(psn)}" title="Lebenslauf öffnen">${escHtml(psn)}</button>`).join('')}</div>`;
}
function fvprKindBadge(row){
  const k=fvprPriorityKind(row);
  if(k==='EXPRESS') return `<span class="${NS}svc-badge ${NS}svc-express">Express</span>`;
  if(k==='PRIO') return `<span class="${NS}svc-badge ${NS}svc-prio">Prio</span>`;
  return '';
}
function fvprStatusBadge(s){
  const n=fvprStatusNorm(s); let cls='gray';
  if(/ZUGESTELLT|DELIVERED|ABGEHOLT|PICKED/.test(n)) cls='green';
  else if(/PROBLEM|HINDERN|NICHT|CANCEL|STORNI|FAILED|REFUSED/.test(n)) cls='red';
  else if(/OPEN|OFFEN|PLANNED|GEPLANT/.test(n)) cls='yellow';
  return `<span class="${NS}status ${cls}">${escHtml(s||'—')}</span>`;
}
function fvprMetricLabel(metric){ return ({stops:'Stopps',pkgs:'Pakete',open:'Offene Stopps',pOpen:'Offene Abholstopps',obstacles:'Zustellhindernisse'})[metric] || metric; }
function fvprRowsForMetric(rows, metric){
  // Jede Klickliste gibt nur den Inhalt der angeklickten Zahl zurück.
  // Stopps = alle Zustellstopps, Pakete = alle Zustellpakete paketgenau,
  // offen = nur offene Zustellstopps, pOpen = nur offene Abholstopps.
  if(metric==='stops') return rows.filter(r=>r.__type==='DELIVERY');
  if(metric==='open') return rows.filter(r=>fvprIsOpenDelivery(r));
  if(metric==='pOpen') return rows.filter(r=>fvprIsOpenPickup(r));
  if(metric==='obstacles') return rows.filter(r=>fvprIsProblemDelivery(r));
  if(metric==='pkgs') return rows
    .filter(r=>r.__type==='DELIVERY')
    .flatMap(r=>{
      const psns=r.__parcelList.length?r.__parcelList:['—'];
      return psns.map(psn=>Object.assign({},r,{__singlePsn:psn,__parcelList: psn==='—'?[]:[psn],__pkgCount:1}));
    });
  return [];
}
async function fvprOpenDetailList({partner='', tour='', metric='stops'}){
  let rows=await fvprFetchDetailRows(false);
  if(partner) rows=rows.filter(r=>r.__partner===partner);
  if(tour) rows=rows.filter(r=>String(r.__tour||'—')===String(tour||'—'));
  rows=fvprRowsForMetric(rows, metric);
  rows.sort((a,b)=>String(a.__tour).localeCompare(String(b.__tour),'de',{numeric:true}) || String(a.__stop).localeCompare(String(b.__stop),'de',{numeric:true}) || String(a.__addr).localeCompare(String(b.__addr),'de',{numeric:true}));
  const title=[fvprMetricLabel(metric), partner, tour?`Tour ${tour}`:''].filter(Boolean).join(' – ');
  const html=`
    <div style="font:13px system-ui">
      <div style="margin:0 0 8px 0;color:#334155;font-weight:700">${escHtml(title)} · ${fmtInt(rows.length)} Einträge · Stand ${todayStr()} ${timeHM()}</div>
      <div style="max-height:62vh;overflow:auto;border:1px solid #e5e7eb;border-radius:8px;background:#fff">
        <table class="${NS}tbl ${NS}detailtbl" style="width:100%;border-collapse:collapse">
          <thead><tr>
            <th>SP</th><th>Tour</th><th>Fahrer</th><th>Stopp</th><th>Empfänger / Adresse</th><th>Service</th><th>Status</th><th>Pakete</th><th>Paketnummer / Lebenslauf</th>
          </tr></thead>
          <tbody>${rows.map(r=>`
            <tr data-row="1">
              <td style="text-align:left">${escHtml(r.__partner)}</td>
              <td>${escHtml(r.__tour)}</td>
              <td style="text-align:left">${escHtml(r.__driver||'—')}</td>
              <td>${escHtml(r.__stop||'—')}</td>
              <td style="text-align:left"><b>${escHtml(r.__name||'—')}</b><br><span style="opacity:.8">${escHtml(r.__addr||'—')}</span></td>
              <td style="text-align:left">${fvprKindBadge(r)} <span>${escHtml(r.__serviceCode||'—')}</span></td>
              <td>${fvprStatusBadge(r.__status)}</td>
              <td>${fmtInt(r.__pkgCount)}</td>
              <td style="text-align:left">${fvprPsnButtons(r.__parcelList)}</td>
            </tr>`).join('') || `<tr><td colspan="9" class="${NS}empty">Keine passenden Detaildaten gefunden.</td></tr>`}</tbody>
        </table>
      </div>
    </div>`;
  const ov=modal(`
    <h3 style="margin:0 0 8px 0;font:700 16px system-ui">${escHtml(title)}</h3>
    ${html}
    <div class="${NS}modal-actions">
      <button class="${NS}btn-sm" data-act="copy-detail">Kopieren</button>
      <button class="${NS}btn-sm" data-act="close">Schließen</button>
    </div>`);
  const table=ov.querySelector('table'); if(table) makeTableSortable(table);
  ov.addEventListener('click', async e=>{
    const psn=e.target.closest('[data-fvpr-psn]'); if(psn){ e.preventDefault(); fvprOpenTracking(psn.dataset.fvprPsn||''); return; }
    const b=e.target.closest('button[data-act]'); if(!b) return;
    if(b.dataset.act==='copy-detail'){ const ok=await copyHtmlToClipboard(html); toast(ok?'Detail kopiert':'Kopieren fehlgeschlagen',ok); }
    if(b.dataset.act==='close') ov.remove();
  }, {passive:false});
}
function fvprCellMetricByIndex(row, cell){
  const cells=Array.from(row.children); const idx=cells.indexOf(cell); const total=row.dataset.totalRow==='1';
  // Wichtig: Partner-Vorschau-Zeilen haben data-partner UND data-tour.
  // Diese Zeilen sind anders aufgebaut als die Gesamtübersicht.
  // Alt wurde hier die Gesamtübersicht-Zuordnung benutzt. Dadurch wurde bei Tour 524/offen
  // nicht auf diese Tour gefiltert und es erschien die große Gesamtliste.
  const isTourDetailRow = row.hasAttribute('data-tour') || (cells.length>=8 && !total && norm(cells[0]?.textContent||''));
  if(total) return ({3:'stops',4:'open',5:'pkgs',6:'obstacles',7:'pOpen'})[idx] || '';
  if(isTourDetailRow) return ({2:'stops',3:'pkgs',4:'open',5:'obstacles',7:'pOpen'})[idx] || '';
  if(row.dataset.partner) return ({3:'stops',4:'open',5:'pkgs',6:'obstacles',7:'pOpen'})[idx] || '';
  return '';
}
function fvprFindPartnerFromModalRow(row){
  const h=row.closest(`.${NS}modal-box`)?.querySelector('h3')?.textContent||'';
  const m=h.match(/Vorschau\s*[–-]\s*(.+)$/); return m?norm(m[1]):'';
}

function fvprShowBlockingLoader(text='Detaildaten werden geladen …'){
  const ov=document.createElement('div');
  ov.className=NS+'modal '+NS+'loading-modal';
  ov.innerHTML=`<div class="${NS}modal-box" style="min-width:min(420px,92vw);text-align:center">
    <div style="font:800 16px system-ui;margin-bottom:8px">Bitte warten</div>
    <div style="font:600 13px system-ui;color:#334155;margin-bottom:10px">${escHtml(text)}</div>
    <div style="height:10px;border-radius:999px;background:#e5e7eb;overflow:hidden">
      <div class="${NS}loadingbar"></div>
    </div>
    <div style="font:500 12px system-ui;color:#64748b;margin-top:10px">Nicht erneut klicken. Die Dispatcher-Daten werden noch abgefragt.</div>
  </div>`;
  document.body.appendChild(ov);
  return ()=>{ try{ ov.remove(); }catch{} };
}
function fvprInstallDetailClickHandler(){
  if(window.__fvpr_detail_click_installed) return; window.__fvpr_detail_click_installed=true;
  document.addEventListener('click', async e=>{
    const psn=e.target.closest('[data-fvpr-psn]'); if(psn){ e.preventDefault(); e.stopPropagation(); fvprOpenTracking(psn.dataset.fvprPsn||''); return; }
    const cell=e.target.closest(`.${NS}tbl td`); if(!cell || e.target.closest('button')) return;
    const row=cell.closest('tr'); if(!row) return;
    const metric=fvprCellMetricByIndex(row, cell); if(!metric) return;
    const txt=norm(cell.textContent); if(!txt || txt==='—' || txt==='0') return;
    e.preventDefault(); e.stopPropagation();
    const partner=row.dataset.partner || fvprFindPartnerFromModalRow(row) || '';
    const tour = row.dataset.tour || (row.children.length>=8 && row.dataset.totalRow!=='1' ? norm(row.children[0]?.textContent||'') : '');
    const hideLoader=fvprShowBlockingLoader(`Detaildaten werden geladen${partner?' – '+partner:''}${tour?' – Tour '+tour:''} …`);
    try{ await fvprOpenDetailList({partner, tour, metric}); }
    catch(err){ console.error('[fvpr] Detail-Liste', err); alert(String(err?.message||err) + '\n\nHinweis: Bitte ggf. einmal die normale Pickup-/Delivery-Liste im Dispatcher öffnen, damit der API-Zugriff sicher erkannt wird.'); }
    finally{ hideLoader(); }
  }, true);
}
function fvprInstallPickupDeliveryHook(){
  if(!window.__fvpr_pd_fetch_hooked && window.fetch){
    const orig=window.fetch;
    window.fetch=async function(input, init={}){
      const res=await orig(input, init);
      try{
        const urlStr=typeof input==='string'?input:(input&&input.url)||'';
        if(urlStr.includes('/dispatcher/api/pickup-delivery') && res.ok){
          const u=new URL(urlStr, location.origin); const q=u.searchParams;
          if(!q.get('parcelNumber') && !q.get('parcel_number')){
            const headers={}; const src=(init&&init.headers)||(input&&input.headers);
            if(src){ if(src.forEach) src.forEach((v,k)=>headers[String(k).toLowerCase()]=String(v)); else if(Array.isArray(src)) src.forEach(([k,v])=>headers[String(k).toLowerCase()]=String(v)); else Object.entries(src).forEach(([k,v])=>headers[String(k).toLowerCase()]=String(v)); }
            FVPR_LAST_PD_REQUEST={url:u,headers};
          }
        }
      }catch{}
      return res;
    };
    window.__fvpr_pd_fetch_hooked=true;
  }
  if(!window.__fvpr_pd_xhr_hooked && window.XMLHttpRequest){
    const X=window.XMLHttpRequest, open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
    X.prototype.open=function(method,url){ this.__fvpr_pd_url=typeof url==='string'?new URL(url,location.origin):null; this.__fvpr_pd_headers={}; return open.apply(this, arguments); };
    X.prototype.setRequestHeader=function(k,v){ try{ this.__fvpr_pd_headers[String(k).toLowerCase()]=String(v); }catch{} return setH.apply(this, arguments); };
    X.prototype.send=function(){ const onload=()=>{ try{ if(this.__fvpr_pd_url && this.__fvpr_pd_url.href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300){ const q=this.__fvpr_pd_url.searchParams; if(!q.get('parcelNumber')&&!q.get('parcel_number')) FVPR_LAST_PD_REQUEST={url:this.__fvpr_pd_url, headers:this.__fvpr_pd_headers}; } }catch{} this.removeEventListener('load', onload); }; this.addEventListener('load', onload); return send.apply(this, arguments); };
    window.__fvpr_pd_xhr_hooked=true;
  }
}
(function fvprDetailAddonStart(){
  try{
    fvprInstallPickupDeliveryHook();
    fvprInstallDetailClickHandler();
    const css=document.createElement('style'); css.id=NS+'detail-style'; css.textContent=`
      .${NS}numlink{color:#0f3f75;text-decoration:none;cursor:pointer;font-weight:700}
      .${NS}psn-wrap{display:flex;flex-wrap:wrap;gap:4px;justify-content:flex-start}
      .${NS}psn-btn{border:1px solid rgba(0,0,0,.14);background:#f7f7f7;border-radius:6px;padding:2px 6px;cursor:pointer;font:600 11px system-ui;color:#0f3f75}
      .${NS}psn-btn:hover{background:#efefef}
      .${NS}svc-badge{display:inline-block;border-radius:999px;padding:2px 7px;margin-right:4px;font:800 11px system-ui;border:1px solid transparent}
      .${NS}svc-prio{background:#fff7ed;color:#9a3412;border-color:#fdba74}
      .${NS}svc-express{background:#fee2e2;color:#991b1b;border-color:#fca5a5}
      .${NS}status{display:inline-block;max-width:100%;padding:2px 8px;border-radius:999px;font:700 11px system-ui;border:1px solid transparent;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .${NS}status.gray{background:#f3f4f6;color:#374151;border-color:#d1d5db}.${NS}status.red{background:#fee2e2;color:#991b1b;border-color:#fca5a5}.${NS}status.green{background:#dcfce7;color:#166534;border-color:#86efac}.${NS}status.yellow{background:#fef3c7;color:#92400e;border-color:#fcd34d}
      .${NS}detailtbl th{position:sticky;top:0;z-index:2}
      .${NS}loadingbar{height:10px;width:35%;border-radius:999px;background:#64748b;animation:${NS}loadmove 1.05s infinite ease-in-out}
      @keyframes ${NS}loadmove{0%{transform:translateX(-110%)}50%{transform:translateX(120%)}100%{transform:translateX(310%)}}
    `; if(!document.getElementById(css.id)) document.head.appendChild(css);
  }catch(e){ console.warn('[fvpr] Detail-Addon konnte nicht gestartet werden', e); }
})();


})();
