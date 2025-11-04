// ==UserScript==
// @name         Dispatcher – Neu-/Rekla-Kunden Kontrolle
// @namespace    bodo.dpd.custom
// @version      3.5.0
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
  .${NS}panel{position:fixed;top:72px;left:50%;transform:translateX(-50%);width:min(1200px,95vw);max-height:78vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100001;display:none}
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
  `;
  document.head.appendChild(s);
}

let latestRows = [];

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

  // Live-Filter nach Tour
  document.getElementById(NS+'tour').addEventListener('input', ()=>renderTable(applyTourFilter(latestRows)));

  renderList();
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
  const saved=loadList();
  const box=document.getElementById(NS+'saved'); if(!box) return;
  if(!saved.length){ box.innerHTML=`<div class="${NS}muted">Noch keine Nummern gespeichert.</div>`; return; }
  box.innerHTML=`
    <div style="margin-bottom:6px;font:600 12px system-ui">Gespeichert (${saved.length}):</div>
    <div style="display:flex;flex-wrap:wrap;gap:6px">
      ${saved.map(n=>`<span class="${NS}chip" data-n="${esc(n)}">${esc(n)} <button title="entfernen" data-n="${esc(n)}">×</button></span>`).join('')}
    </div>`;
  box.querySelectorAll('button[data-n]').forEach(b=>{
    b.onclick = ()=>{ const n=String(b.dataset.n||''); saveList(loadList().filter(x=>x!==n)); };
  });
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
    // parallel laden
    const [rows, map] = await Promise.all([fetchPagedAll(), fetchDriverPhoneMap()]);
    rowsApi = rows;
    driverMap = map;
  }catch{
    out.innerHTML = `<div class="${NS}muted">Kein API-Request erkannt. Bitte einmal die normale Liste laden und erneut versuchen.</div>`;
    return;
  }

  const rows = rowsApi
    .filter(r => matchesSaved(getCustomerFromRow(r), wantSet))
    .map(r => {
      const status = getStatus(r) || '—';
      const warnPredict = predictOutsidePickup(r);
      const {soon, closed} = closingHints(r);
      const hints=[];
      if (warnPredict) hints.push('Predict außerhalb Abholfenster');
      if (closed) hints.push('bereits geschlossen');
      else if (soon) hints.push('schließt in ≤30 Min');

      const tour = String(r.tour || '').trim();
let drv = null;
for (const k of tourKeys(tour)) { if (driverMap && driverMap.has(k)) { drv = driverMap.get(k); break; } }


      return {
  number: normShort(getCustomerFromRow(r)),
  tour,
  name:   getName(r)   || '—',
  street: getStreet(r) || '—',
  plz:    digits(r.postalCode || r.zip || ''),
  ort:    r.city || '',
  phone:  getPhone(r) || '—',

  // >>> NEU / KORRIGIERT <<<
  driverName:  (drv && drv.name)  ? drv.name  : '—',
  driverPhone: (drv && drv.phone) ? drv.phone : '—',

  predict: getPredict(r) || '—',
  pickup:  getPickup(r)  || '—',
  status,
  isCompleted: String(status).toUpperCase()==='COMPLETED',
  warnRow: warnPredict,
  hintText: hints.join(' • ') || '—'
};

    });

  latestRows = rows;
  renderTable(applyTourFilter(latestRows));
}

function renderTable(rows){
  const out=document.getElementById(NS+'out');
  const head=['Kundennr.','Tour','Kundenname','Straße','PLZ','Ort','Telefon','Fahrer','Fahrer Telefon','Predict Zeitfenster','Zeitfenster Abholung','Status','Hinweise'];
  const body=(rows||[]).map(r=>{
    const statusHtml = r.isCompleted ? `<span class="${NS}status-completed">${esc(r.status||'—')}</span>` : esc(r.status||'—');
    const tds = [
  r.number||'—',  r.tour||'—', r.name||'—',
  r.street||'—', r.plz||'—', r.ort||'—', r.phone||'—',
  r.driverName||'—', r.driverPhone||'—',
  r.predict||'—', r.pickup||'—', statusHtml, esc(r.hintText||'—')
].map((v,i)=> i===11 ? `<td>${v}</td>` : `<td>${esc(v)}</td>`).join(''); // i===11: Status ist HTML
    const trClass = r.warnRow ? ` class="${NS}warn-row"` : '';
    return `<tr${trClass}>${tds}</tr>`;
  }).join('');

  out.innerHTML = `
    <table class="${NS}tbl">
      <thead><tr>${head.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${body || `<tr><td colspan="13">Keine Treffer.</td></tr>`}</tbody>
    </table>`;
}

/* ======================= Boot ======================= */
function onReady(){ ensureStyles(); mountUI(); }
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', onReady);
else onReady();

})();
