// ==UserScript==
// @name         Dispatcher – Neu-/Rekla-Kunden Kontrolle
// @namespace    bodo.dpd.custom
// @version      3.9.0
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tools-Dispatcher-Abholer.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tools-Dispatcher-Abholer.user.js
// @description  Tagesliste per API (ohne Kundenfilter), lokal filtern; Hinweise: Predict außerhalb, schließt ≤30 Min, bereits geschlossen; COMPLETED grün; Telefon-Spalte; Fahrer-Telefon via vehicle-overview; Tour-Filter; Button dockt an #pm-wrap.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function(){
'use strict';

/* ======================= Basics & Storage ======================= */
const NS = 'kn-';
const LS_KEY = 'kn.saved.customers';
const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
const digits = s => String(s||'').replace(/\D+/g,'');
const normShort = s => digits(s).replace(/^0+/,'');

function loadList(){ try{ const a=JSON.parse(localStorage.getItem(LS_KEY)||'[]'); return Array.isArray(a)?a:[]; }catch{ return []; } }
function saveList(arr){ const clean=[...new Set((arr||[]).map(normShort).filter(Boolean))]; localStorage.setItem(LS_KEY, JSON.stringify(clean)); renderList(); }

function tourKeys(t){
  const s = String(t ?? '').trim();
  const no0 = s.replace(/^0+/, '');
  const pad2 = no0.padStart(2,'0');
  const pad3 = no0.padStart(3,'0');
  const pad4 = no0.padStart(4,'0');
  // eindeutige Liste zurückgeben
  return [...new Set([s, no0, pad2, pad3, pad4].filter(Boolean))];
}




/* ======================= Request Capture ======================= */
let lastOkRequest = null;

(function hook(){
  if (!window.__kn_fetch_hooked && window.fetch){
    const orig=window.fetch;
    window.fetch = async function(i, init={}){
      const res = await orig(i, init);
      try{
        const urlStr = typeof i==='string' ? i : (i && i.url) || '';
        if (urlStr.includes('/dispatcher/api/pickup-delivery') && res.ok){
          const u = new URL(urlStr, location.origin);
          const q = u.searchParams;
          if (!q.get('parcelNumber')) {
            const h = {};
            const src = (init && init.headers) || (i && i.headers);
            if (src){
              if (src.forEach) src.forEach((v,k)=>h[String(k).toLowerCase()]=String(v));
              else if (Array.isArray(src)) src.forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
              else Object.entries(src).forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
            }
            if (!h['authorization']){
              const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
              if (m) h['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
            }
            lastOkRequest = { url: u, headers: h };
            const n = document.getElementById(NS+'note'); if (n) n.style.display='none';
          }
        }
      }catch{}
      return res;
    };
    window.__kn_fetch_hooked = true;
  }

  if (!window.__kn_xhr_hooked && window.XMLHttpRequest){
    const X=window.XMLHttpRequest;
    const open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
    X.prototype.open=function(m,u){ this.__kn_url=(typeof u==='string')?new URL(u,location.origin):null; this.__kn_headers={}; return open.apply(this,arguments); };
    X.prototype.setRequestHeader=function(k,v){ try{ this.__kn_headers[String(k).toLowerCase()]=String(v); }catch{} return setH.apply(this,arguments); };
    X.prototype.send=function(){
      const onload=()=>{
        try{
          if (this.__kn_url && this.__kn_url.href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300){
            const q=this.__kn_url.searchParams;
            if (!q.get('parcelNumber')){
              if (!this.__kn_headers['authorization']){
                const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                if (m) this.__kn_headers['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
              }
              lastOkRequest = { url:this.__kn_url, headers:this.__kn_headers };
              const n = document.getElementById(NS+'note'); if (n) n.style.display='none';
            }
          }
        }catch{}
        this.removeEventListener('load', onload);
      };
      this.addEventListener('load', onload);
      return send.apply(this,arguments);
    };
    window.__kn_xhr_hooked = true;
  }
})();

/* ======================= Styles + UI ======================= */
function ensureStyles(){
  if (document.getElementById(NS+'style')) return;
  const s=document.createElement('style'); s.id=NS+'style';
  s.textContent=`
  .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:6px 12px;border-radius:999px;font:600 12px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
  .${NS}panel{
  position:fixed;
  top:72px;
  left:50%;
  transform:translateX(-50%);
  /* vorher: width:min(1200px,95vw); */
  width:min(1800px,98vw);          /* deutlich breiter */
  max-height:85vh;                 /* etwas höher */
  overflow-y:auto;                 /* NUR vertikales Scrollen */
  overflow-x:hidden;               /* horizontal aus */
  background:#fff;
  border:1px solid rgba(0,0,0,.12);
  box-shadow:0 12px 28px rgba(0,0,0,.18);
  border-radius:12px;
  z-index:100001;
  display:none
}

  .${NS}head{display:flex;gap:10px;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
  .${NS}row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .${NS}inp{padding:6px 8px;border-radius:8px;border:1px solid rgba(0,0,0,.15);font:13px system-ui}
  .${NS}chip{display:inline-flex;gap:6px;align-items:center;border:1px solid rgba(0,0,0,.12);background:#fff;padding:4px 8px;border-radius:999px;font:600 12px system-ui;cursor:pointer}
  .${NS}tbl{width:100%;border-collapse:collapse;font:12px system-ui}
  .${NS}tbl th,.${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top}
  .${NS}tbl th{position:sticky;top:0;background:#fafafa;text-align:left}
  .${NS}list{padding:8px 12px}
  .${NS}muted{opacity:.7}
  .${NS}warn-row { background: #fff3cd; }
  .${NS}status-completed { color: #0a7a2a; font-weight: 700; }
  .${NS}th-sort{ cursor:pointer; user-select:none; }
  .${NS}th-sort .arrow{ margin-left:6px; font-size:10px; opacity:.6; }
  .${NS}cust-collapser{display:flex;align-items:center;gap:8px;margin:8px 0 6px;font:600 12px system-ui;cursor:pointer;user-select:none}
  .${NS}chev{display:inline-block;transition:transform .18s ease}
  .${NS}chev.rot{transform:rotate(90deg)}
  .${NS}cust-wrap{overflow:hidden;transition:max-height .18s ease}
  .${NS}cust-wrap.collapsed{max-height:0}
  .${NS}dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px;background:#ccc;vertical-align:middle}
  .${NS}dot.on{background:#2ecc71}

  `;
  document.head.appendChild(s);
}

function initSavedCollapser(){
  const wrap = document.getElementById(NS+'saved-wrap');
  const chev = document.getElementById(NS+'saved-chev');
  const toggle = document.getElementById(NS+'saved-toggle');
  if (!wrap || !chev || !toggle) return;

  const LSKEY = 'pm.savedCust.collapsed'; // key kann so bleiben
  let collapsed = localStorage.getItem(LSKEY);
  collapsed = collapsed == null ? '1' : collapsed; // default: zugeklappt

  // Anfangszustand
  wrap.classList.toggle('collapsed', collapsed === '1');
  chev.classList.toggle('rot', collapsed !== '1');

  toggle.onclick = () => {
    const isColl = wrap.classList.toggle('collapsed');
    chev.classList.toggle('rot', !isColl);
    localStorage.setItem(LSKEY, isColl ? '1' : '0');
  };
}


let latestRows = [];

/* ======================= Auto-Refresh ======================= */
const AUTO_KEY = NS + 'auto.enabled';
const AUTO_MS  = 60 * 1000; // 60s
let autoTimer  = null;

function autoIsOn(){ return localStorage.getItem(AUTO_KEY) !== '0'; } // Default: EIN
function autoStart(){
  autoStop();
  autoTimer = setInterval(() => { try{ loadDetails(); }catch{} }, AUTO_MS);
}
function autoStop(){
  if (autoTimer){ clearInterval(autoTimer); autoTimer = null; }
}
function autoSet(on){
  localStorage.setItem(AUTO_KEY, on ? '1' : '0');
  on ? autoStart() : autoStop();
  updateAutoBtn();
}
function updateAutoBtn(){
  const b = document.getElementById(NS+'auto');
  if (!b) return;
  const on = autoIsOn();
  b.innerHTML = `<span class="${NS}dot ${on ? 'on' : ''}"></span> Auto 60s`;
  b.title = on ? 'Auto-Refresh ist EIN (alle 60s). Klicken zum Ausschalten.'
               : 'Auto-Refresh ist AUS. Klicken zum Einschalten.';
  b.setAttribute('aria-pressed', on ? 'true' : 'false');
}
window.addEventListener('beforeunload', autoStop);


let sortState = { key: null, dir: 1 }; // dir: 1=asc, -1=desc
const collator = new Intl.Collator('de-DE', { numeric: true, sensitivity: 'base' });

// Welche Spalte liest welchen Wert aus einem Zeilenobjekt?
const sortGetters = {
  number:       r => r.number,
  tour:         r => r.tour,
  name:         r => r.name,
  street:       r => r.street,
  plz:          r => Number(r.plz) || 0,
  ort:          r => r.ort,
  phone:        r => r.phone,
  driverName:   r => r.driverName,
  driverPhone:  r => r.driverPhone,
  predict:      r => r.predict,
  pickup:       r => r.pickup,
  status:       r => r.status,
  hintText:     r => r.hintText,
  lastScanRaw:  r => r.lastScanRaw ?? new Date(0) // Date-Objekt (s. unten)
};

function sortRows(rows){
  if (!sortState.key) return rows;
  const get = sortGetters[sortState.key];
  if (!get) return rows;
  const dir = sortState.dir;
  return [...rows].sort((a,b)=>{
    const va = get(a), vb = get(b);
    // Datum?
    if (va instanceof Date || vb instanceof Date){
      const ta = va instanceof Date ? va : new Date(va);
      const tb = vb instanceof Date ? vb : new Date(vb);
      return dir * (ta - tb);
    }
    // Zahl?
    if (typeof va === 'number' || typeof vb === 'number'){
      return dir * ((Number(va)||0) - (Number(vb)||0));
    }
    // natürlicher String-Vergleich (de-DE, numerisch)
    return dir * collator.compare(String(va ?? ''), String(vb ?? ''));
  });
}


function mountUI(){
  ensureStyles();
  if (!document.body || document.getElementById(NS+'btn')) return;

  const pmWrap = document.getElementById('pm-wrap');
  let host = pmWrap;
  if (!host){
    host = document.createElement('div');
    host.id = NS+'fallback-wrap';
    Object.assign(host.style, {position:'fixed', top:'8px', left:'50%', transform:'translateX(-50%)', display:'flex', gap:'8px', zIndex: 100000});
    document.body.appendChild(host);
  }

  const btn=document.createElement('button');
  btn.id=NS+'btn'; btn.type='button'; btn.className=NS+'btn';
  btn.textContent='Neu-/Rekla-Kunden Kontrolle';
  btn.addEventListener('click', ()=>togglePanel());
  host.appendChild(btn);

  const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel';
  panel.innerHTML=`
    <div class="${NS}head">
      <div class="${NS}row">
        <span>Kundennummern:</span>
        <input id="${NS}add" class="${NS}inp" placeholder="Nummer eingeben, Enter zum Hinzufügen">
        <button class="${NS}btn" id="${NS}addBtn" type="button">Hinzufügen</button>
        <button class="${NS}btn" id="${NS}import" type="button">Import</button>
        <button class="${NS}btn" id="${NS}export" type="button">Export</button>
        <button class="${NS}btn" id="${NS}clear"  type="button">Liste leeren</button>
      </div>
      <div class="${NS}row">
        <input id="${NS}tour" class="${NS}inp" placeholder="Tour-Filter (optional)">
        <button class="${NS}btn" id="${NS}load"  type="button">Liste laden (API)</button>
        <button class="${NS}btn" id="${NS}auto" type="button" aria-pressed="true"></button>
        <button class="${NS}btn" id="${NS}close" type="button">Schließen</button>
      </div>
    </div>
    <div class="${NS}list">
      <div id="${NS}note" class="${NS}muted" style="margin-bottom:8px">Hinweis: Einmal die normale Liste laden. Der letzte pickup-delivery-Request wird geklont.</div>
      <div id="${NS}saved"></div>
      <div id="${NS}out"></div>
    </div>`;
  document.body.appendChild(panel);

  const fileInput=document.createElement('input'); fileInput.type='file'; fileInput.accept='.txt,.csv,.json'; fileInput.style.display='none'; fileInput.id=NS+'file';
  document.body.appendChild(fileInput);

  document.getElementById(NS+'close').onclick = ()=>togglePanel(false);
  document.getElementById(NS+'addBtn').onclick = onAdd;
  document.getElementById(NS+'add').addEventListener('keydown', e=>{ if(e.key==='Enter') onAdd(); });
  document.getElementById(NS+'clear').onclick = ()=>{ if(confirm('Liste wirklich leeren?')) saveList([]); };
  document.getElementById(NS+'import').onclick = ()=>fileInput.click();
  document.getElementById(NS+'export').onclick = onExport;
  fileInput.addEventListener('change', onImport);
  document.getElementById(NS+'load').onclick = loadDetails;

  // >>> Auto-Button ERST JETZT verdrahten (existiert nun im DOM)
  const autoBtn = document.getElementById(NS+'auto');
  if (autoBtn){
    autoBtn.onclick = () => autoSet(!autoIsOn());
    updateAutoBtn();
  }

  // Live-Filter nach Tour
  document.getElementById(NS+'tour').addEventListener('input', ()=>renderTable(applyTourFilter(latestRows)));

  renderList();

  if (autoIsOn()) autoStart();
}





    function mountCustomerSavedCollapser(){
  try{
    // 1) Container mit "Gespeichert (" + Chips finden
    const labelNode = Array.from(document.querySelectorAll('div,section,header'))
      .find(el => /\bGespeichert\s*\(\d+\)/i.test(el.textContent||''));
    if (!labelNode) return;

    // Prüfe, ob schon eingebaut
    if (labelNode.__pm_collapser) return;

    // Chips-Container bestimmen: gleiche Ebene oder nächstes Sibling mit vielen Chips
    let chipsWrap = labelNode.nextElementSibling;
    if (!chipsWrap || chipsWrap.children.length < 1) {
      // Fallback: suche in Parent nach der Chipzeile
      chipsWrap = Array.from(labelNode.parentElement?.children||[])
        .find(el => el !== labelNode && (el.querySelector('.MuiChip-root,[data-testid="CancelIcon"],[class*="Chip"]') || '').length !== undefined) || null;
    }
    if (!chipsWrap) return;

    // Wrapper bauen
    const collapser = document.createElement('div');
    collapser.className = NS+'cust-collapser';
    const chev = document.createElement('span'); chev.className = NS+'chev'; chev.textContent = '▸';
    const title = document.createElement('span'); title.textContent = (labelNode.textContent||'Gespeichert').trim();
    collapser.append(chev, title);

    // Chipbereich in Hüll-Wrapper legen
    const wrap = document.createElement('div');
    wrap.className = NS+'cust-wrap';
    chipsWrap.parentNode.insertBefore(wrap, chipsWrap);
    wrap.appendChild(chipsWrap);

    // State laden: default collapsed = true
    const LSKEY = 'pm.savedCust.collapsed';
    let collapsed = localStorage.getItem(LSKEY);
    collapsed = collapsed == null ? '1' : collapsed; // default 1
    if (collapsed === '1') { wrap.classList.add('collapsed'); chev.classList.remove('rot'); }
    else { chev.classList.add('rot'); }

    // Collapser einfügen (über dem Wrapper)
    wrap.parentNode.insertBefore(collapser, wrap);

    // Toggle
    const apply = ()=> {
      const isColl = wrap.classList.toggle('collapsed');
      chev.classList.toggle('rot', !isColl);
      localStorage.setItem(LSKEY, isColl ? '1':'0');
    };
    collapser.addEventListener('click', apply);

    // Merker, damit wir nicht doppelt mounten
    labelNode.__pm_collapser = true;
  } catch(e){ /* silent */ }
}



function togglePanel(force){
  const panel=document.getElementById(NS+'panel'); if(!panel) return;
  const isHidden=getComputedStyle(panel).display==='none';
  const show = force!=null ? !!force : isHidden;
  panel.style.setProperty('display', show?'block':'none', 'important');
}
document.addEventListener('DOMContentLoaded', mountUI);
new MutationObserver(()=>mountUI()).observe(document.documentElement,{childList:true,subtree:true});

/* ======================= Import/Export ======================= */
function onAdd(){
  const inp=document.getElementById(NS+'add');
  const v=normShort(inp.value); if(!v) return;
  const list=loadList(); list.push(v); saveList(list); inp.value='';
}
async function onImport(e){
  const file=e.target.files && e.target.files[0]; e.target.value=''; if(!file) return;
  try{
    const txt=await file.text(); let nums=[];
    try{ const j=JSON.parse(txt); if(Array.isArray(j)) nums=j.map(normShort); }catch{}
    if(!nums.length) nums=[...String(txt).matchAll(/\d{5,}/g)].map(m=>normShort(m[0]));
    if(!nums.length){ alert('Im Import keine Kundennummern gefunden.'); return; }
    const list=loadList(); list.push(...nums); saveList(list);
    alert(`Import OK: ${nums.length} Nummern.`);
  }catch(err){ alert('Import fehlgeschlagen: '+(err&&err.message||err)); }
}
function onExport(){
  const content=(loadList()).join('\n');
  const blob=new Blob([content],{type:'text/plain;charset=utf-8'});
  const url=URL.createObjectURL(blob); const a=document.createElement('a');
  a.href=url; a.download='kundennummern.txt'; document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); },0);
}
function renderList(){
  const saved = loadList();
  const box = document.getElementById(NS+'saved');
  if (!box) return;

  if (!saved.length){
    box.innerHTML = `<div class="${NS}muted">Noch keine Nummern gespeichert.</div>`;
    return;
  }

  box.innerHTML = `
    <div class="${NS}cust-collapser" id="${NS}saved-toggle" style="margin-bottom:6px;">
      <span class="${NS}chev" id="${NS}saved-chev">▸</span>
      <span>Gespeichert (${saved.length}):</span>
    </div>

    <div class="${NS}cust-wrap" id="${NS}saved-wrap">
      <div style="display:flex;flex-wrap:wrap;gap:6px">
        ${saved.map(n=>`<span class="${NS}chip" data-n="${esc(n)}">${esc(n)} <button title="entfernen" data-n="${esc(n)}">×</button></span>`).join('')}
      </div>
    </div>
  `;

  // Entfernen-Buttons wie gehabt
  box.querySelectorAll('button[data-n]').forEach(b=>{
    b.onclick = ()=>{
      const n = String(b.dataset.n||'');
      saveList(loadList().filter(x=>x!==n));
    };
  });

  // Toggle initialisieren (liest/setzt localStorage + Pfeil)
  initSavedCollapser();
}

/* ======================= API-Header ======================= */
function buildHeaders(h){
  const H=new Headers();
  try{
    if(h){ Object.entries(h).forEach(([k,v])=>{
      const key=k.toLowerCase();
      if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)){
        H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()), v);
      }
    }); }
    if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
  }catch{}
  return H;
}

/* ======================= pickup-delivery: ohne Kundennummern, danach client-seitig filtern ======================= */
function buildUrlAll(base, page){
  const u=new URL(base.href);
  const q=u.searchParams;
  ['customerNo','customerNumber','customerNumbers','name','city','pcode','street','houseno'].forEach(k=>q.delete(k));
  ['sort','scanCode','receiptId','insertUserName','modifyUserName','elements'].forEach(k=>q.delete(k));
  q.set('page', String(page));
  q.set('pageSize','500');
  q.set('orderTypes','PICKUP');
  q.set('active','true');
  const today = new Date().toISOString().slice(0,10);
  q.set('dateFrom', u.searchParams.get('dateFrom') || today);
  q.set('dateTo',   u.searchParams.get('dateTo')   || today);
  u.search=q.toString();
  return u;
}
async function fetchPagedAll(){
  if(!lastOkRequest) throw new Error('Kein pickup-delivery Request erkannt. Bitte einmal die normale Liste laden.');
  const headers=buildHeaders(lastOkRequest.headers);
  const size=500, maxPages=60;
  let page=1, rows=[];
  while(page<=maxPages){
    const u=buildUrlAll(lastOkRequest.url, page);
    const r=await fetch(u.toString(), {credentials:'include', headers});
    if(!r.ok) break;
    const j=await r.json();
    const chunk=(j.items||j.content||j.data||j.results||j)||[];
    rows.push(...chunk);
    if(chunk.length<size) break;
    page++; await sleep(30);
  }
  return rows;
}

/* ======================= vehicle-overview: Fahrer-Telefon nach Tour ======================= */
function buildVehicleUrl(page){
  const u=new URL('https://dispatcher2-de.geopost.com/dispatcher/api/vehicle-overview');
  const q=u.searchParams;
  q.set('page', String(page));
  q.set('pageSize','500');
  // Standard-Filter beibehalten; wir ziehen alle Fahrzeuge des Tages
  q.set('withOrders','true');
  u.search=q.toString();
  return u;
}

async function fetchDriverPhoneMap(){
  if(!lastOkRequest) throw new Error('Kein Auth-Kontext. Bitte einmal die normale Liste laden.');
  const headers = buildHeaders(lastOkRequest.headers);
  const size=500, maxPages=20;
  let page=1;
  const map = new Map(); // key: tour-Variante -> {name, phone}

  while(page<=maxPages){
    const u = new URL('https://dispatcher2-de.geopost.com/dispatcher/api/vehicle-overview');
    u.searchParams.set('page', String(page));
    u.searchParams.set('pageSize', String(size));
    u.searchParams.set('withOrders','true');

    const r = await fetch(u.toString(), {credentials:'include', headers});
    if(!r.ok) break;
    const j = await r.json();
    const items = (j.items||j.content||j.data||j.results||j)||[];

    for(const it of items){
      const rawTour = String(it.tour ?? it.tourNumber ?? it.round ?? '').trim();
      if(!rawTour) continue;

      const info = {
        name:  String(it.courierName ?? it.driverName ?? '').trim(),
        phone: String(it.courierPhone ?? it.driverPhone ?? '').trim()
      };

      // alle Schlüsselvarianten mappen
      for (const k of tourKeys(rawTour)) {
        if (!map.has(k)) map.set(k, info);
        else {
          // ggf. fehlende Felder ergänzen
          const cur = map.get(k);
          if (!cur.name  && info.name)  cur.name  = info.name;
          if (!cur.phone && info.phone) cur.phone = info.phone;
        }
      }
    }

    if(items.length<size) break;
    page++; await sleep(30);
  }
  return map;
}


/* ======================= Zeit-/Interval-Tools + Feld-Getter ======================= */
function hhmm(ts){ if (!ts) return ''; if (typeof ts==='string' && /^\d{1,2}:\d{2}(:\d{2})?$/.test(ts)) return ts.slice(0,5); const d=new Date(ts); if (isNaN(d)) return ''; return d.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}); }
function buildWindow(from,to){ const a=hhmm(from), b=hhmm(to); return (a||b) ? `${a}–${b}` : ''; }

function toMin(s){ if(!s) return null; const m=String(s).match(/^(\d{1,2}):(\d{2})/); if(!m) return null; const h=Math.min(23,parseInt(m[1],10)); const mi=Math.min(59,parseInt(m[2],10)); return h*60+mi; }
function makeInterval(f,t){ let a=toMin(f), b=toMin(t); if(a==null && b==null) return null; if(a==null) a=b; if(b==null) b=a; if(b<=a) b+=1440; return [a,b]; }
function intervalsOverlap(a,b){ if(!a||!b) return false; return a[0] < b[1] && b[0] < a[1]; }

function predictOutsidePickup(row){
  const pred = makeInterval(row.from2 ?? row.timeFrom2, row.to2 ?? row.timeTo2);
  if(!pred) return false;
  const p1 = makeInterval(row.timeFrom1, row.timeTo1);
  const p2 = makeInterval(row.timeFrom2, row.timeTo2);
  const pickups = [p1,p2].filter(Boolean);
  if(!pickups.length) return false;
  return !pickups.some(p => intervalsOverlap(p, pred));
}

function closingHints(row){
  const now = new Date();
  const dateStr = row.date || new Date().toISOString().slice(0,10);
  function endOf(f,t){
    const iv=makeInterval(f,t); if(!iv) return null;
    const endMin = iv[1] % 1440;
    const h = Math.floor(endMin/60), m = endMin%60;
    const dt = new Date(`${dateStr}T${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}:00`);
    return isNaN(dt) ? null : dt;
  }
  const e1 = endOf(row.timeFrom1, row.timeTo1);
  const e2 = endOf(row.timeFrom2, row.timeTo2);
  const ends = [e1,e2].filter(Boolean);
  if (!ends.length) return {soon:false,closed:false};

  const soon = ends.some(e => e > now && (e - now) <= 30*60*1000);
  const closed = ends.every(e => e <= now);
  return {soon, closed};
}

// Feld-Getter
function getCustomerFromRow(r){ const raw = Array.isArray(r.customerNo) ? (r.customerNo[0] || '') : (r.customerNumber || r.customer || r.customerId || ''); return String(raw); }
function getName(r){ return r.name ?? (Array.isArray(r.customerName) ? r.customerName[0] : r.customerName) ?? ''; }
function getStreet(r){ const s=r.street ?? ''; const hn=r.houseno ?? r.houseNumber ?? ''; return [s,hn].filter(Boolean).join(' '); }
function getPredict(r){ const a=r.from2 ?? r.timeFrom2 ?? ''; const b=r.to2 ?? r.timeTo2 ?? ''; return buildWindow(a,b); }
function getPickup(r){ const w1=buildWindow(r.timeFrom1, r.timeTo1); const w2=buildWindow(r.timeFrom2, r.timeTo2); return [w1,w2].filter(Boolean).join(' | '); }
function getStatus(r){ return r.pickupStatus ?? r.deliveryStatus ?? r.status ?? ''; }
function getPhone(r){
  if (Array.isArray(r.addressPhone) && r.addressPhone.length) return String(r.addressPhone[0]);
  return r.phone || r.contactPhone || '';
}
/* ======================= Letzter Scan + Tracking (tour-basiert) ======================= */
// kleine Helfer

 // ---- letzter Scan aus einem Tracking-Stopp bestimmen ----
function _toDateSafe(v){ if(!v) return null; const d=new Date(v); return isNaN(d)?null:d; }
function lastScanFromStop(s){
  // bevorzugt: scanDate + scanTime (so kommt es bei dir)
  if (s.scanDate && s.scanTime) {
    const d = new Date(`${s.scanDate}T${s.scanTime}`);
    if (!isNaN(d)) return d;
  }
  // Fallbacks
  const cands = [
    _toDateSafe(s.deliveredTime),
    _toDateSafe(s.modifyDate),
    (s.firstStopHandlingScanTime && s.scanDate) ? _toDateSafe(`${s.scanDate}T${s.firstStopHandlingScanTime}`) : null,
  ].filter(Boolean);
  if(!cands.length) return null;
  return cands.reduce((a,b)=> a>b ? a : b);
}
function toDate(v, dateStr=null){
  if (!v) return null;
  if (/^\d{2}:\d{2}:\d{2}$/.test(v) && dateStr) return new Date(`${dateStr}T${v}`);
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function lastScanInfo(r){
  const cands = [
    toDate(r.deliveredTime),
    toDate(r.etaScanTime, r.scanDate),
    toDate(r.firstStopHandlingScanTime, r.scanDate),
    toDate(r.scanTime, r.scanDate),
    toDate(r.modifyDate)
  ].filter(Boolean);
  if (!cands.length) return { dateTime:null, source:null };
  const max = cands.reduce((a,b)=> a>b ? a : b);
  let source = 'modifyDate';
  if (max.getTime() === toDate(r.deliveredTime)?.getTime()) source = 'deliveredTime';
  else if (max.getTime() === toDate(r.etaScanTime, r.scanDate)?.getTime()) source = 'etaScanTime';
  else if (max.getTime() === toDate(r.firstStopHandlingScanTime, r.scanDate)?.getTime()) source = 'firstStopHandlingScanTime';
  else if (max.getTime() === toDate(r.scanTime, r.scanDate)?.getTime()) source = 'scanTime';
  return { dateTime:max, source };
}
function gmapsDirections(lat, lon, address, country='DE'){
  const origin = (lat!=null && lon!=null) ? `${lat},${lon}` : '';
  const dest = encodeURIComponent([address, country].filter(Boolean).join(', '));
  return origin
    ? `https://www.google.com/maps/dir/?api=1&origin=${origin}&destination=${dest}&travelmode=driving`
    : `https://www.google.com/maps/search/?api=1&query=${dest}`;
}
function getFirstParcel(r){
  return r.parcelNumber || (Array.isArray(r.parcelNumbers) ? r.parcelNumbers[0] : '');
}

// Holt das Tracking für EINE Tour und gibt nur den aktuellsten Scan (GPS) dieser Tour zurück.
async function fetchTrackingByTour(depot, tour, dateStr){
  if(!lastOkRequest) throw new Error('Kein Auth-Kontext. Bitte einmal die normale Liste laden.');
  const headers = buildHeaders(lastOkRequest.headers);

  // Depot
  let dep = String(depot || '').trim();
  if (!dep) dep = lastOkRequest?.url?.searchParams?.get('depot') || '';
  if (!dep) { console.warn('[tracking] Kein Depot -> übersprungen'); return null; }

  // Tour 3-stellig
  const tourPadded = String(tour ?? '').trim().replace(/^0+/, '').padStart(3,'0');

  // URL
  const u = new URL('https://dispatcher2-de.geopost.com/dispatcher/api/vehicle-overview/tracking');
  u.searchParams.set('depot', dep);
  u.searchParams.set('tour', tourPadded);
  const d = (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) ? dateStr : new Date().toISOString().slice(0,10);
  u.searchParams.set('date', d);

  const r = await fetch(u.toString(), { credentials:'include', headers });
  if (!r.ok) { console.warn('[tracking] Request fehlgeschlagen', r.status); return null; }

  const json = await r.json();
  const stops = Array.isArray(json?.stops) ? json.stops
              : Array.isArray(json) ? json
              : (json?.items || json?.content || json?.data || json?.results || []);

  if (!stops || !stops.length) return null;

  // Jüngsten Stopp der Tour finden
  let latest = null, latestTs = null;
  for (const s of stops){
    const ts = lastScanFromStop(s);
    if(!ts) continue;
    if(!latestTs || ts > latestTs){
      latestTs = ts;
      latest = s;
    }
  }
  if (!latest) return null;

  // GPS des letzten Scans
  const lat = latest.gpsLat ?? latest.gpslat ?? latest.latitude ?? latest.plannedCoordinateLat ?? null;
  const lon = latest.gpsLong ?? latest.gpsLon ?? latest.gpslon ?? latest.longitude ?? latest.plannedCoordinateLong ?? null;

  return { lastScan: latestTs, lat, lon, tour: tourPadded };
}



/* ======================= Filtern & Rendern ======================= */
function matchesSaved(cellRaw, wantSet){
  const raw=digits(cellRaw); const pure=raw.replace(/^0+/, '');
  for(const w of wantSet){ if(!w) continue; if (pure===w) return true; if (pure.endsWith(w)) return true; if (raw.endsWith(w)) return true; }
  return false;
}

function applyTourFilter(rows){
  const inp = document.getElementById(NS+'tour');
  const v = (inp && inp.value || '').trim().toLowerCase();
  if(!v) return rows;
  return (rows||[]).filter(r => String(r.tour||'').toLowerCase().includes(v));
}

async function loadDetails(){
  const out=document.getElementById(NS+'out');
  out.innerHTML = '<div class="'+NS+'muted">Lade …</div>';

  const savedShorts = loadList().map(normShort).filter(Boolean);
  if(!savedShorts.length){ out.innerHTML='<div class="'+NS+'muted">Keine Nummern gespeichert.</div>'; return; }
  const wantSet = new Set(savedShorts);

  let rowsApi=[], driverMap=null;
  try{
    const [rows, map] = await Promise.all([fetchPagedAll(), fetchDriverPhoneMap()]);
    rowsApi = rows;
    driverMap = map;
  }catch{
    out.innerHTML = `<div class="${NS}muted">Kein API-Request erkannt. Bitte einmal die normale Liste laden und erneut versuchen.</div>`;
    return;
  }

  // Zuerst auf gespeicherte Kunden filtern (damit wir nur nötige Touren fürs Tracking ziehen)
  const filteredSrc = rowsApi.filter(r => matchesSaved(getCustomerFromRow(r), wantSet));

// --- Tracking pro Tour holen: wir wollen NUR den letzten Scan je Tour ---
const depotFromCtx = lastOkRequest?.url?.searchParams?.get('depot') || '';
const tourSet = new Set(filteredSrc.map(r => String(r.tour||'').trim()).filter(Boolean));
const latestByTour = new Map();

for (const t of tourSet) {
  const one = filteredSrc.find(r => String(r.tour||'').trim() === t);
  const depotGuess = String(one?.depot || depotFromCtx || '').trim();
  const dateGuess  = (one?.date && /^\d{4}-\d{2}-\d{2}$/.test(one.date)) ? one.date : null;

  try{
    const info = await fetchTrackingByTour(depotGuess, t, dateGuess);
    if (info) latestByTour.set(String(t).trim(), info);
  }catch(e){
    console.warn('[tracking] Fehler für Tour', t, e);
  }
}



  const rows = filteredSrc.map(r => {
    const status = getStatus(r) || '—';
    const warnPredict = predictOutsidePickup(r);
    const {soon, closed} = closingHints(r);
    const hints=[];
    if (warnPredict) hints.push('Predict außerhalb Abholfenster');
    if (closed) hints.push('bereits geschlossen');
    else if (soon) hints.push('schließt in ≤30 Min');

    // Fahrer-Infos per Tour
    const tour = String(r.tour || '').trim();
    let drv = null;
    for (const k of tourKeys(tour)) { if (driverMap && driverMap.has(k)) { drv = driverMap.get(k); break; } }

// tracking für diese Tour
const tkey = String(r.tour || '').trim();
const tinfo = latestByTour.get(tkey) || null;

// Zieladresse (Kunde)
const address = [
  getStreet(r),
  [digits(r.postalCode || r.zip || ''), r.city || ''].filter(Boolean).join(' ')
].filter(Boolean).join(', ');

// Google-Maps: von letzter Scan-Position (Fahrer) zum Kunden
const gmaps = (tinfo && tinfo.lat != null && tinfo.lon != null)
  ? `https://www.google.com/maps/dir/?api=1&origin=${tinfo.lat},${tinfo.lon}&destination=${encodeURIComponent(address)}&travelmode=driving`
  : `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(address)}`;

return {
  number: normShort(getCustomerFromRow(r)),
  tour,
  name:   getName(r)   || '—',
  street: getStreet(r) || '—',
  plz:    digits(r.postalCode || r.zip || ''),
  ort:    r.city || '',
  phone:  getPhone(r) || '—',
  driverName:  (drv && drv.name)  ? drv.name  : '—',
  driverPhone: (drv && drv.phone) ? drv.phone : '—',
  predict: getPredict(r) || '—',
  pickup:  getPickup(r)  || '—',
  status,
  isCompleted: String(status).toUpperCase()==='COMPLETED',
  warnRow: warnPredict,
  hintText: hints.join(' • ') || '—',

  // NEU: pro Tour einheitlich
  lastScan: tinfo?.lastScan
  ? tinfo.lastScan.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' })
  : '—',
  lastScanRaw: tinfo?.lastScan ?? null,   // <— für Sortierung (Date)
  gpsLat: tinfo?.lat ?? '',
  gpsLon: tinfo?.lon ?? '',
  gmaps
};

  });

  latestRows = rows;
  renderTable(applyTourFilter(latestRows));
}

function renderTable(rows){
  const out = document.getElementById(NS+'out');

  // Spaltendefinitionen: Label + Schlüssel für Sortierung (key=null => nicht sortierbar)
  const cols = [
    {label:'Kundennr.',            key:'number'},
    {label:'Tour',                 key:'tour'},
    {label:'Kundenname',           key:'name'},
    {label:'Straße',               key:'street'},
    {label:'PLZ',                  key:'plz'},
    {label:'Ort',                  key:'ort'},
    {label:'Telefon',              key:'phone'},
    {label:'Fahrer',               key:'driverName'},
    {label:'Fahrer Telefon',       key:'driverPhone'},
    {label:'Predict Zeitfenster',  key:'predict'},
    {label:'Zeitfenster Abholung', key:'pickup'},
    {label:'Status',               key:'status'},    // HTML in Zelle, Sort über Text
    {label:'Hinweise',             key:'hintText'},
    {label:'Letzter Scan',         key:'lastScanRaw'},
    {label:'Karte',                key:null}         // nicht sortierbar (Link)
  ];

  // Header HTML inkl. Pfeil
  const headHtml = cols.map(c=>{
    if (!c.key) return `<th>${esc(c.label)}</th>`;
    const isSorted = sortState.key === c.key;
    const arrow = isSorted ? (sortState.dir === 1 ? '▲' : '▼') : '';
    return `<th class="${NS}th-sort" data-key="${esc(c.key)}">${esc(c.label)}<span class="arrow">${arrow}</span></th>`;
  }).join('');

  // Vor dem Rendern sortieren
  const rowsSorted = sortRows(rows || []);

  // Body
  const bodyHtml = (rowsSorted||[]).map(r=>{
    const statusHtml = r.isCompleted
      ? `<span class="${NS}status-completed">${esc(r.status||'—')}</span>`
      : esc(r.status||'—');
    const linkHtml = r.gmaps
      ? `<a href="${esc(r.gmaps)}" target="_blank" rel="noopener" title="${esc(r.gpsLat||'')}, ${esc(r.gpsLon||'')}">Karte</a>`
      : '—';

    // Reihenfolge MUSS zu 'cols' passen
    const cells = [
      esc(r.number||'—'),
      esc(r.tour||'—'),
      esc(r.name||'—'),
      esc(r.street||'—'),
      esc(r.plz||'—'),
      esc(r.ort||'—'),
      esc(r.phone||'—'),
      esc(r.driverName||'—'),
      esc(r.driverPhone||'—'),
      esc(r.predict||'—'),
      esc(r.pickup||'—'),
      statusHtml,                 // HTML
      esc(r.hintText||'—'),
      esc(r.lastScan||'—'),
      linkHtml                    // HTML
    ];

    const trClass = r.warnRow ? ` class="${NS}warn-row"` : '';
    // Status (index 11) & Karte (index 14) sind HTML – NICHT escapen
    const html = cells.map((v,i)=> (i===11 || i===14) ? `<td>${v}</td>` : `<td>${v}</td>`).join('');
    return `<tr${trClass}>${html}</tr>`;
  }).join('');

  out.innerHTML = `
    <table class="${NS}tbl">
      <thead><tr>${headHtml}</tr></thead>
      <tbody>${bodyHtml || `<tr><td colspan="${cols.length}">Keine Treffer.</td></tr>`}</tbody>
    </table>`;

  // Click-Handler für Sortierung
  out.querySelectorAll('th[data-key]').forEach(th=>{
    th.onclick = () => {
      const k = th.getAttribute('data-key');
      if (sortState.key === k) {
        sortState.dir = -sortState.dir;   // Toggle
      } else {
        sortState.key = k;
        sortState.dir = 1;
      }
      // Neu rendern (mit evtl. aktivem Tour-Filter)
      renderTable(applyTourFilter(latestRows));
    };
  });
}

/* ======================= Boot ======================= */
function onReady(){ ensureStyles(); mountUI(); }
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', onReady);
else onReady();

})();
