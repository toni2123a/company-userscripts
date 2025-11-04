// ==UserScript==
// @name         Dispatcher – Neu-/Rekla-Kunden Kontrolle (robust)
// @namespace    bodo.dpd.custom
// @version      1.2.0
// @description  Kundennummern robust in GM-Speicher sichern und Detail-Liste abrufen.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_listValues
// @grant        GM_deleteValue
// ==/UserScript==

(function(){
'use strict';

/* ======================= Grundkonstanten/Helper ======================= */
const NS = 'kn-';                            // Namensraum (kollisionsfrei zu fvpr-)
const GM_KEY_LIST = 'kn.saved.customers';    // Nummernliste
const GM_KEY_CFG  = 'kn.config';             // Konfig
const LS_KEY      = 'kn-saved-customers';    // Alt (Migration)
const LS_CFG      = 'kn-config';             // Alt (Migration)

const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
const fmtHHMM = d => d ? new Date(d).toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}) : '';

/* ======================= GM-Storage (robust) ========================= */
const gmGet = (k, def=null) => { try { const v=GM_getValue(k); return (v==null||v==='')?def:JSON.parse(v); } catch { return def; } };
const gmSet = (k, v) => { try { GM_setValue(k, JSON.stringify(v)); } catch {} };
const gmDel = (k) => { try { GM_deleteValue(k); } catch {} };

// Einmalige Migration von localStorage -> GM_*
(function migrateOnce(){
  const FLAG='kn.migrated.v1';
  if (gmGet(FLAG,false)) return;
  try{
    let oldList=[]; let oldCfg=null;
    try{ oldList=JSON.parse(localStorage.getItem(LS_KEY)||'[]')||[]; }catch{}
    try{ oldCfg =JSON.parse(localStorage.getItem(LS_CFG)||'null'); }catch{}
    if(Array.isArray(oldList)&&oldList.length) gmSet(GM_KEY_LIST, Array.from(new Set(oldList)));
    if(oldCfg && typeof oldCfg==='object') gmSet(GM_KEY_CFG, oldCfg);
    localStorage.removeItem(LS_KEY);
    localStorage.removeItem(LS_CFG);
  }catch{}
  gmSet(FLAG,true);
})();

function loadCfg(){
  const base={ shrinkZeroBlocks:false };
  const cur=gmGet(GM_KEY_CFG,null);
  if(!cur || typeof cur!=='object'){ gmSet(GM_KEY_CFG, base); return base; }
  for(const k of Object.keys(base)) if(!(k in cur)) cur[k]=base[k];
  gmSet(GM_KEY_CFG, cur);
  return cur;
}
function saveCfg(){ gmSet(GM_KEY_CFG, cfg); }

function loadList(){
  const arr=gmGet(GM_KEY_LIST, []);
  return Array.isArray(arr)?arr:[];
}
function saveList(arr){
  const clean=Array.from(new Set((arr||[]).filter(Boolean)));
  gmSet(GM_KEY_LIST, clean);
  renderList();
}

/* ======================= Nummern-Normalisierung ====================== */
const stripLeadingZeros = s => String(s||'').replace(/\D+/g,'').replace(/^0+/,'');
function shrinkZeroBlocks(s){ return String(s||'').replace(/0{5,}/g,'0'); } // „vor den füllenden Nullen“
function normalizeCustomer(raw){
  let x = stripLeadingZeros(raw);
  if (cfg.shrinkZeroBlocks) x = shrinkZeroBlocks(x);
  return x;
}

/* ======================= Styles ===================================== */
function ensureStyles(){
  if(document.getElementById(NS+'style')) return;
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
  .${NS}muted{opacity:.7}`;
  document.head.appendChild(s);
}

/* ======================= UI Mounting ================================= */
const cfg = loadCfg();
tryMountUI();
setInterval(tryMountUI, 1500);
new MutationObserver(()=>tryMountUI()).observe(document.documentElement,{childList:true,subtree:true});

function tryMountUI(){
  if (!document.body) return;
  if (document.getElementById(NS+'btn')) return;

  // Button direkt in fvpr-wrap setzen, falls vorhanden
  const fvprWrap = document.getElementById('fvpr-wrap');
  const btn = document.createElement('button');
  btn.id = NS+'btn';
  btn.type = 'button';
  btn.className = NS+'btn';
  btn.textContent = 'Neu-/Rekla-Kunden Kontrolle';
  btn.addEventListener('click', ()=>togglePanel());

  if (fvprWrap) {
    fvprWrap.appendChild(btn);
  } else {
    // Fallback: eigener Wrap, gleiche Position/Optik wie fvpr
    const wrap = document.createElement('div');
    wrap.className = NS+'wrap';
    wrap.appendChild(btn);
    document.body.appendChild(wrap);
  }

  // Panel
  const panel = document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel';
  panel.innerHTML = `
    <div class="${NS}head">
      <div class="${NS}row">
        <span>Kundennummern:</span>
        <input id="${NS}add" class="${NS}inp" placeholder="Nummer eingeben, Enter zum Hinzufügen">
        <button class="${NS}btn" id="${NS}addBtn" type="button">Hinzufügen</button>
        <button class="${NS}btn" id="${NS}fromDom" type="button">+ Aus Auswahl</button>
        <button class="${NS}btn" id="${NS}paste" type="button">+ Aus Zwischenablage</button>
        <button class="${NS}btn" id="${NS}clear" type="button">Liste leeren</button>
      </div>
      <div class="${NS}row">
        <label class="${NS}chip"><input type="checkbox" id="${NS}shrink"> Nullen-Block kürzen</label>
        <button class="${NS}btn" id="${NS}load" type="button">Liste laden (API)</button>
        <button class="${NS}btn" id="${NS}close" type="button">Schließen</button>
      </div>
    </div>
    <div class="${NS}list">
      <div class="${NS}row ${NS}muted" style="margin-bottom:8px">Hinweis: Einmal eine normale Suche/Liste in Dispatcher laden, damit ein API-Request geklont werden kann.</div>
      <div id="${NS}saved"></div>
      <div id="${NS}out"></div>
    </div>`;
  document.body.appendChild(panel);

  // Events
  document.getElementById(NS+'close').onclick = ()=>togglePanel(false);
  document.getElementById(NS+'addBtn').onclick = onAdd;
  document.getElementById(NS+'add').addEventListener('keydown', e=>{ if(e.key==='Enter') onAdd(); });
  document.getElementById(NS+'clear').onclick = ()=>{ if(confirm('Liste wirklich leeren?')) saveList([]); };
  document.getElementById(NS+'paste').onclick = onPaste;
  document.getElementById(NS+'fromDom').onclick = addFromDom;
  document.getElementById(NS+'load').onclick = loadDetails;
  document.getElementById(NS+'shrink').checked = !!cfg.shrinkZeroBlocks;
  document.getElementById(NS+'shrink').onchange = (e)=>{ cfg.shrinkZeroBlocks=!!e.target.checked; saveCfg(); };

  ensureStyles();
  renderList();
}

function togglePanel(force){
  const panel=document.getElementById(NS+'panel');
  const show = force!=null ? !!force : panel.style.display==='none';
  panel.style.display = show ? '' : 'none';
  if (show) renderList();
}

/* ======================= Eingabe / Liste ============================= */
function onAdd(){
  const inp=document.getElementById(NS+'add');
  const v=normalizeCustomer(inp.value);
  if(!v) return;
  const list=loadList(); list.push(v); saveList(list);
  inp.value='';
}
async function onPaste(){
  try{
    const t = await navigator.clipboard.readText();
    if (!t) return;
    const found = Array.from(t.matchAll(/\d{6,}/g)).map(m=>normalizeCustomer(m[0]));
    if (!found.length) return;
    const list = loadList(); list.push(...found); saveList(list);
  }catch(e){ alert('Zwischenablage nicht lesbar: ' + e); }
}
// Liest Kundennummern aus sichtbarer Tabelle/Zeilen
function addFromDom(){
  const rows = Array.from(document.querySelectorAll('tbody tr, [role="row"]'));
  const nums = [];
  for (const tr of rows){
    const txt = (tr.textContent||'').trim();
    const matches = txt.match(/\b\d{6,}\b/g);
    if (matches) nums.push(...matches.map(normalizeCustomer));
  }
  if (!nums.length) { alert('Keine Kundennummern in der Tabelle erkannt.'); return; }
  const list = loadList(); list.push(...nums); saveList(list);
}
function renderList(){
  const saved=loadList();
  const box=document.getElementById(NS+'saved'); if(!box) return;
  if(!saved.length){ box.innerHTML='<div class="'+NS+'muted">Noch keine Nummern gespeichert.</div>'; return; }
  box.innerHTML=`
    <div style="margin-bottom:6px;font:600 12px system-ui">Gespeichert (${saved.length}):</div>
    <div style="display:flex;flex-wrap:wrap;gap:6px">
      ${saved.map(n=>`<span class="${NS}chip" data-n="${esc(n)}">${esc(n)} <button title="entfernen" data-n="${esc(n)}">×</button></span>`).join('')}
    </div>`;
  box.querySelectorAll('button[data-n]').forEach(b=>{
    b.onclick = ()=>{ const n=String(b.dataset.n||''); const arr=loadList().filter(x=>x!==n); saveList(arr); };
  });
}

/* ======================= API-Hooks (clone) =========================== */
// Klont letzte erfolgreiche Dispatcher-API-Anfrage (URL + Header inkl. JWT)
let lastOkRequest=null;
(function hook(){
  if (!window.__kn_fetch_hooked && window.fetch) {
    const orig=window.fetch;
    window.fetch = async function(i, init={}){
      const res = await orig(i, init);
      try{
        const uStr = typeof i==='string' ? i : (i&&i.url)||'';
        if (uStr.includes('/dispatcher/api/') && res.ok) {
          const u = new URL(uStr, location.origin);
          const h = {};
          const src = (init && init.headers) || (i && i.headers);
          if (src) {
            if (src.forEach) src.forEach((v,k)=>h[String(k).toLowerCase()]=String(v));
            else if (Array.isArray(src)) src.forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
            else Object.entries(src).forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
          }
          if (!h['authorization']) {
            const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
            if (m) h['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
          }
          lastOkRequest = { url:u, headers:h };
        }
      }catch{}
      return res;
    };
    window.__kn_fetch_hooked=true;
  }
  if (!window.__kn_xhr_hooked && window.XMLHttpRequest){
    const X=window.XMLHttpRequest;
    const open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
    X.prototype.open=function(m,u){ this.__kn_url=(typeof u==='string')?new URL(u,location.origin):null; this.__kn_headers={}; return open.apply(this,arguments); };
    X.prototype.setRequestHeader=function(k,v){ try{ this.__kn_headers[String(k).toLowerCase()]=String(v); }catch{} return setH.apply(this,arguments); };
    X.prototype.send=function(){
      const onload=()=>{
        try{
          if (this.__kn_url && this.__kn_url.href.includes('/dispatcher/api/') && this.status>=200 && this.status<300){
            if (!this.__kn_headers['authorization']) {
              const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
              if (m) this.__kn_headers['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
            }
            lastOkRequest = { url:this.__kn_url, headers:this.__kn_headers };
          }
        }catch{}
        this.removeEventListener('load', onload);
      };
      this.addEventListener('load', onload);
      return send.apply(this,arguments);
    };
    window.__kn_xhr_hooked=true;
  }
})();

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
// Parameternamen-Kandidaten für „Kundennummer“
function buildUrlByCustomer(base, customerNumber, page=1){
  const candidates=['customerNumber','customerNo','customer','customerId'];
  const u=new URL(base.href);
  const q=u.searchParams;
  q.set('page', String(page));
  q.set('pageSize','100');
  ['parcelNumber','priority','elements'].forEach(k=>q.delete(k));
  for(const k of candidates) q.set(k, customerNumber);
  u.search=q.toString();
  return u;
}

/* ======================= Daten holen + rendern ======================= */
const pick = (o,...keys)=> keys.map(k=>o?.[k]).find(v=>v!=null && v!=='') || '';
const parsePredict = r => {
  const from=pick(r,'from2','predictFrom','predictStart');
  const to  =pick(r,'to2','predictTo','predictEnd');
  return (from||to)?`${fmtHHMM(from)}–${fmtHHMM(to)}`:'';
};
const parsePickup = r => {
  const s=pick(r,'pickupFrom','pickupStart'); const e=pick(r,'pickupTo','pickupEnd');
  return (s||e)?`${fmtHHMM(s)}–${fmtHHMM(e)}`:'';
};
const parseStatus = r => pick(r,'statusName','statusText','stateText','status','deliveryStatus','parcelStatus') || '—';

async function fetchByCustomer(number){
  if(!lastOkRequest) throw new Error('Kein API-Request zum Klonen gefunden.');
  const headers=buildHeaders(lastOkRequest.headers);
  let page=1, rows=[];
  while(page<=5){
    const u=buildUrlByCustomer(lastOkRequest.url, number, page);
    const r=await fetch(u.toString(), {credentials:'include', headers});
    if(!r.ok) break;
    const j=await r.json();
    const chunk=(j.items||j.content||j.data||j.results||j)||[];
    rows.push(...chunk);
    if(chunk.length<100) break;
    page++; await sleep(50);
  }
  return rows;
}

async function loadDetails(){
  const out=document.getElementById(NS+'out');
  out.innerHTML='<div class="'+NS+'muted">Lade …</div>';
  const nums=loadList();
  if(!nums.length){ out.innerHTML='<div class="'+NS+'muted">Keine Nummern gespeichert.</div>'; return; }
  if(!lastOkRequest){ out.innerHTML='<div class="'+NS+'muted">Bitte zuerst eine normale Dispatcher-Liste laden/suchen, damit ein API-Request geklont werden kann.</div>'; return; }

  const allRows=[];
  for(const n of nums){
    try{
      const rows=await fetchByCustomer(n);
      const r=rows[0]||{};
      allRows.push({
        number:n,
        systempartner: pick(r,'systemPartner','partnerName','partner'),
        tour: pick(r,'tour','route','round'),
        name: pick(r,'receiverName','customerName','name'),
        street: [pick(r,'street','street1'), pick(r,'houseno','houseNumber')].filter(Boolean).join(' '),
        plz: pick(r,'postalCode','zip','zipCode'),
        ort: pick(r,'city','town'),
        predict: parsePredict(r),
        pickup: parsePickup(r),
        status: parseStatus(r)
      });
    }catch(e){
      allRows.push({number:n, systempartner:'', tour:'', name:'', street:'', plz:'', ort:'', predict:'', pickup:'', status:'(Fehler)'});
    }
    await sleep(60);
  }

  const head=['Kundennr.','Systempartner','Tour','Kundenname','Straße','PLZ','Ort','Predict Zeitfenster','Zeitfenster Abholung','Status'];
  const body=allRows.map(r=>[
    esc(r.number), esc(r.systempartner||'—'), esc(r.tour||'—'), esc(r.name||'—'),
    esc(r.street||'—'), esc(r.plz||'—'), esc(r.ort||'—'),
    esc(r.predict||'—'), esc(r.pickup||'—'), esc(r.status||'—')
  ].map(v=>`<td>${v}</td>`).join(''));

  document.getElementById(NS+'out').innerHTML=`
    <table class="${NS}tbl">
      <thead><tr>${head.map(h=>`<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${body.map(tr=>`<tr>${tr}</tr>`).join('') || `<tr><td colspan="10">Keine Daten gefunden.</td></tr>`}</tbody>
    </table>`;
}

/* ======================= Start/Init ================================== */
ensureStyles();

})();

