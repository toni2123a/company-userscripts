// ==UserScript==
// @name         DPD Dispatcher – Neu-/Rekla-Kunden Kontrolle
// @namespace    bodo.dpd.custom
// @version      2.0.0
// @description  Kundennummern importieren/exportieren, per captured pickup-delivery-Request laden, DOM-Werte (Name/Status/Zeitfenster) mitverwenden.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function(){
'use strict';

/* ======================= Basics & Storage (localStorage) ======================= */
const NS = 'kn-';
const LS_KEY = 'kn.saved.customers';
const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
const digits = s => String(s||'').replace(/\D+/g,'');
const normShort = s => digits(s).replace(/^0+/,'');
function loadList(){ try{ const a=JSON.parse(localStorage.getItem(LS_KEY)||'[]'); return Array.isArray(a)?a:[]; }catch{ return []; } }
function saveList(arr){ const clean=[...new Set((arr||[]).map(normShort).filter(Boolean))]; localStorage.setItem(LS_KEY, JSON.stringify(clean)); renderList(); }

/* ======================= Request-Capture wie PRIO/EXP12 ======================= */
let lastOkRequest = null;

// fetch hook
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
  .${NS}wrap{position:fixed;top:8px;left:calc(50% - 240px);display:flex;gap:8px;z-index:2147483646}
  .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:8px 14px;border-radius:999px;font:600 13px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
  .${NS}panel{position:fixed;top:56px;left:50%;transform:translateX(-50%);width:min(1200px,95vw);max-height:78vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:2147483645;display:none}
  .${NS}head{display:flex;gap:10px;align-items:center;justify-content:space-between;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
  .${NS}row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  .${NS}inp{padding:6px 8px;border-radius:8px;border:1px solid rgba(0,0,0,.15);font:13px system-ui}
  .${NS}chip{display:inline-flex;gap:6px;align-items:center;border:1px solid rgba(0,0,0,.12);background:#fff;padding:4px 8px;border-radius:999px;font:600 12px system-ui;cursor:pointer}
  .${NS}tbl{width:100%;border-collapse:collapse;font:12px system-ui}
  .${NS}tbl th,.${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top}
  .${NS}tbl th{position:sticky;top:0;background:#fafafa;text-align:left}
  .${NS}list{padding:8px 12px}
  .${NS}muted{opacity:.7}
  #fvpr-wrap{ left: calc(50% - 420px) !important; }
  `;
  document.head.appendChild(s);
}

function mountUI(){
  if (!document.body || document.getElementById(NS+'btn')) return;

  const fvprWrap = document.getElementById('fvpr-wrap');
  const btn=document.createElement('button');
  btn.id=NS+'btn'; btn.type='button'; btn.className=NS+'btn';
  btn.textContent='Neu-/Rekla-Kunden Kontrolle';
  btn.addEventListener('click', ()=>togglePanel());

  if (fvprWrap) fvprWrap.insertBefore(btn, fvprWrap.firstElementChild||null);
  else { const wrap=document.createElement('div'); wrap.className=NS+'wrap'; wrap.appendChild(btn); document.body.appendChild(wrap); }

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

  // hidden file input
  const fileInput=document.createElement('input'); fileInput.type='file'; fileInput.accept='.txt,.csv,.json'; fileInput.style.display='none'; fileInput.id=NS+'file';
  document.body.appendChild(fileInput);

  // Events
  document.getElementById(NS+'close').onclick = ()=>togglePanel(false);
  document.getElementById(NS+'addBtn').onclick = onAdd;
  document.getElementById(NS+'add').addEventListener('keydown', e=>{ if(e.key==='Enter') onAdd(); });
  document.getElementById(NS+'clear').onclick = ()=>{ if(confirm('Liste wirklich leeren?')) saveList([]); };
  document.getElementById(NS+'import').onclick = ()=>fileInput.click();
  document.getElementById(NS+'export').onclick = onExport;
  fileInput.addEventListener('change', onImport);
  document.getElementById(NS+'load').onclick = loadDetails;

  ensureStyles();
  renderList();
}

function togglePanel(force){
  const panel=document.getElementById(NS+'panel'); if(!panel) return;
  const isHidden=getComputedStyle(panel).display==='none';
  const show = force!=null ? !!force : isHidden;
  panel.style.setProperty('display', show?'block':'none', 'important');
  if (show) renderList();
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

/* ======================= API helpers (wie im PRIO/EXP12) ======================= */
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
function buildUrlSameView(base, page){
  const u=new URL(base.href);
  const q=u.searchParams;
  q.set('page', String(page));
  q.set('pageSize','500');
  q.delete('parcelNumber'); // nur Listen
  u.search=q.toString();
  return u;
}
async function fetchPaged(builder){
  if(!lastOkRequest) throw new Error('Kein pickup-delivery Request erkannt. Bitte einmal die Liste laden.');
  const headers=buildHeaders(lastOkRequest.headers);
  const size=500, maxPages=60;
  let page=1, rows=[];
  while(page<=maxPages){
    const u=builder(lastOkRequest.url, page);
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

/* ======================= DOM Map: Kundenname/Status/Zeitfenster ======================= */
const lower = s => String(s||'').toLowerCase();
const normTxt = s => String(s||'').replace(/\s+/g,' ').trim();

function readDomMap(){
  const map=new Map();
  const thead=document.querySelector('thead[class*="Datagrid"]');
  const tbody=thead?.parentElement?.parentElement?.querySelector('tbody');
  if(!thead||!tbody) return map;

  const ths=Array.from(thead.querySelectorAll('th')).filter(th=>th.className.includes('DraggableHeader__CustomTh'));
  const labels=ths.map(th=>lower(normTxt(th.querySelector('label')?.textContent || th.textContent)));

  const firstRow = tbody.querySelector('tr'); if(!firstRow) return map;
  const tdsFirst = Array.from(firstRow.children);
  const selectionOffset = (tdsFirst[0] && tdsFirst[0].querySelector('input[type="checkbox"]')) ? 1 : 0;

  const colIndex={}; labels.forEach((label, idx)=> colIndex[label]=selectionOffset+idx);
  const idx = name => (colIndex[lower(name)] ?? -1);

  const C = {
    kundennummer: idx('kundennummer'),
    kundenname:   idx('kundenname'),
    predict:      idx('predict-zeitfenster'),
    standard:     idx('standard-zeitfenster'),
    zeit1:        idx('zeitfenster 1 z / a'),
    zeit2:        idx('zeitfenster 2 z / a'),
    status:       idx('status')
  };

  const get = (tr,i)=> i>=0 ? normTxt(tr.children[i]?.textContent || '') : '';

  function getStatus(tr){
    if (C.status < 0) return '';
    const cell = tr.children[C.status];
    const title = cell?.querySelector('div[title]')?.getAttribute('title');
    return normTxt(title || cell?.textContent || '');
  }

  for(const tr of Array.from(tbody.querySelectorAll('tr'))){
    const raw = get(tr, C.kundennummer);
    const key = normShort(raw);
    if(!key) continue;
    const name = get(tr, C.kundenname);
    const pred = get(tr, C.predict) || get(tr, C.standard);
    const pickup = [get(tr, C.zeit1), get(tr, C.zeit2)].filter(Boolean).join(' | ');
    const status = getStatus(tr);
    map.set(key, { name, pred, pickup, status });
  }
  return map;
}

/* ======================= Laden + Rendern ======================= */
function getCustomerFromRow(r){
  return r.customerNumber || r.customerNo || r.customer || r.customerId || r.kundennummer || '';
}
function hhmm(ts){
  if (!ts) return '';
  if (typeof ts==='string' && /\d{1,2}:\d{2}/.test(ts)) return ts;
  const d=new Date(ts); if (isNaN(d)) return '';
  return d.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'});
}
function buildWindow(from,to){
  const a=hhmm(from), b=hhmm(to);
  return (a||b) ? `${a}–${b}` : '';
}
function matchesSaved(cellRaw, wantSet){
  const raw=digits(cellRaw); const pure=raw.replace(/^0+/,'');
  for(const w of wantSet){
    if(!w) continue;
    if (pure===w) return true;
    if (pure.endsWith(w)) return true;
    if (raw.endsWith(w)) return true;
  }
  return false;
}

async function loadDetails(){
  const out=document.getElementById(NS+'out');
  out.innerHTML = '<div class="'+NS+'muted">Lade …</div>';

  const wantSet=new Set(loadList().map(normShort));
  if(!wantSet.size){ out.innerHTML='<div class="'+NS+'muted">Keine Nummern gespeichert.</div>'; return; }

  let rowsApi=[];
  try{
    rowsApi = await fetchPaged(buildUrlSameView);
  }catch(e){
    out.innerHTML = `<div class="${NS}muted">Kein API-Request erkannt. Bitte einmal die normale Liste laden und erneut versuchen.</div>`;
    return;
  }

  const domMap = readDomMap();

  const rows = rowsApi
    .filter(r => matchesSaved(getCustomerFromRow(r), wantSet))
    .map(r=>{
      const numRaw = getCustomerFromRow(r);
      const key = normShort(numRaw);
      const dom = domMap.get(key) || {};
      return {
        number: key,
        systempartner: r.depot || r.systemPartner || '',
        tour: r.tour || '',
        name: dom.name || r.customerName || r.name || '',
        street: [r.street || r.addressLine1 || '', r.houseNumber || r.houseno || ''].filter(Boolean).join(' '),
        plz: digits(r.postalCode || r.zip || ''),
        ort: r.city || '',
        predict: dom.pred || buildWindow(r.from2||r.predictFrom||r.predictStart, r.to2||r.predictTo||r.predictEnd),
        pickup: dom.pickup || buildWindow(r.pickupFrom||r.pickupStart, r.pickupTo||r.pickupEnd),
        status: dom.status || (r.statusName||r.statusText||r.stateText||r.status||r.deliveryStatus||r.parcelStatus||'—')
      };
    });

  renderTable(rows);
}

function renderTable(rows){
  const out=document.getElementById(NS+'out');
  const head=['Kundennr.','Systempartner','Tour','Kundenname','Straße','PLZ','Ort','Predict Zeitfenster','Zeitfenster Abholung','Status'];
  const body=(rows||[]).map(r=>[
    r.number||'—', r.systempartner||'—', r.tour||'—', r.name||'—',
    r.street||'—', r.plz||'—', r.ort||'—', r.predict||'—', r.pickup||'—', r.status||'—'
  ].map(v=>`<td>${esc(v)}</td>`).join(''));

  out.innerHTML = `
    <table class="${NS}tbl">
      <thead><tr>${head.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${body.length? body.map(tr=>`<tr>${tr}</tr>`).join('') : `<tr><td colspan="10">Keine Treffer (prüf gespeicherte Nummern oder Ansicht).</td></tr>`}</tbody>
    </table>`;
}

/* ======================= Boot ======================= */
function onReady(){ ensureStyles(); mountUI(); }
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', onReady);
else onReady();

})();
