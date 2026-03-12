// ==UserScript==
// @name         Dispatcher – AbholKunden Kontrolle
// @namespace    bodo.dpd.custom
// @version      1.3.0
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher-abholer.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher-abholer.user.js
// @description  Tagesliste per API (ohne Kundenfilter), Kundennummern automatisch aus Abholung (API/DOM-Fallback), Alert-Monitoring (Scheduled/Complaint/Not Accepted/Critical/Distance), lokal filtern; Hinweise: Predict außerhalb, schließt ≤30 Min, bereits geschlossen; COMPLETED grün; Telefon-Spalte; Fahrer-Telefon via vehicle-overview; Tour-Filter; Button dockt an #pm-wrap.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==


(function () {
  'use strict';



  // ---------- Modul-Metadaten für das Hauptpanel ----------
  const def = {
    id: 'dispatcher-abholer-kontrolle',
    label: 'Abholkunden Kontrolle',
    run: () => startModuleOnce()
  };

  // Sofort registrieren (oder queue’n, falls Loader noch nicht da)
  if (window.TM && typeof window.TM.register === 'function') {
    window.TM.register(def);
  } else {
    window.__tmQueue = window.__tmQueue || [];
    window.__tmQueue.push(def);
  }

  // ============== AB HIER: dein ursprüngliches Script – auf Lazy-Start umgebaut ==============
  const NS = 'abholer-';
  const STORAGE_PREFIX = 'dispatcher.abholer.';
  const LEGACY_STORAGE_PREFIX = 'kn.';
  const LAST_REQ_SESSION_KEY = STORAGE_PREFIX + 'lastReq';
  const LAST_REQ_GLOBAL_KEY = '__dispatcherAbholerLastOkRequest';
  const FETCH_HOOK_FLAG = '__dispatcherAbholerFetchHooked';
  const XHR_HOOK_FLAG = '__dispatcherAbholerXhrHooked';
  const XHR_URL_KEY = '__dispatcherAbholerUrl';
  const XHR_HEADERS_KEY = '__dispatcherAbholerHeaders';

  try {
   const s = sessionStorage.getItem(LAST_REQ_SESSION_KEY);
   if (s) {
     const o = JSON.parse(s);
      window[LAST_REQ_GLOBAL_KEY] = {
       url: new URL(o.url, location.origin),
       headers: o.headers || {}
     };
   }
 } catch {}

 // Hooks sofort installieren, nicht erst im Lazy-Start
 installHooksOnce?.();

  // Guard, damit UI/Hook nur 1x gebaut werden
  let started = false;
  function startModuleOnce() {
    if (started) { togglePanel(true); return; }
    started = true;
    boot();             // alles initialisieren
    togglePanel(true);  // Panel anzeigen
  }

  /* ======================= Basics & Storage ======================= */
  const LS_KEY = STORAGE_PREFIX + 'saved.customers';
  const LEGACY_LS_KEY = LEGACY_STORAGE_PREFIX + 'saved.customers';
  const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
  const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const digits = s => String(s||'').replace(/\D+/g,'');
  const normShort = s => digits(s).replace(/^0+/,'');

  function loadList(){
    try{
      const raw = localStorage.getItem(LS_KEY) ?? localStorage.getItem(LEGACY_LS_KEY) ?? '[]';
      const a = JSON.parse(raw);
      return Array.isArray(a) ? a : [];
    }catch{ return []; }
  }
  function saveList(arr){ const clean=[...new Set((arr||[]).map(normShort).filter(Boolean))]; localStorage.setItem(LS_KEY, JSON.stringify(clean)); renderList(); }

  function tourKeys(t){
    const s = String(t ?? '').trim();
    const no0 = s.replace(/^0+/, '');
    const pad2 = no0.padStart(2,'0');
    const pad3 = no0.padStart(3,'0');
    const pad4 = no0.padStart(4,'0');
    return [...new Set([s, no0, pad2, pad3, pad4].filter(Boolean))];
  }

  /* ======================= Request Capture (lazy installiert) ======================= */
 let lastOkRequest = window[LAST_REQ_GLOBAL_KEY] || null;
  function installHooksOnce(){
    if (!window[FETCH_HOOK_FLAG] && window.fetch){
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
              window[LAST_REQ_GLOBAL_KEY] = lastOkRequest;
                 try {
   sessionStorage.setItem(LAST_REQ_SESSION_KEY, JSON.stringify({ url: u.href, headers: h }));
 } catch {}
              const n = document.getElementById(NS+'note'); if (n) n.style.display='none';
            }
          }
        }catch{}
        return res;
      };
      window[FETCH_HOOK_FLAG] = true;
    }

    if (!window[XHR_HOOK_FLAG] && window.XMLHttpRequest){
      const X=window.XMLHttpRequest;
      const open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
      X.prototype.open=function(m,u){ this[XHR_URL_KEY]=(typeof u==='string')?new URL(u,location.origin):null; this[XHR_HEADERS_KEY]={}; return open.apply(this,arguments); };
      X.prototype.setRequestHeader=function(k,v){ try{ this[XHR_HEADERS_KEY][String(k).toLowerCase()]=String(v); }catch{} return setH.apply(this,arguments); };
      X.prototype.send=function(){
        const onload=()=>{
          try{
            if (this[XHR_URL_KEY] && this[XHR_URL_KEY].href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300){
              const q=this[XHR_URL_KEY].searchParams;
              if (!q.get('parcelNumber')){
                if (!this[XHR_HEADERS_KEY]['authorization']){
                  const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                  if (m) this[XHR_HEADERS_KEY]['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
                }
                lastOkRequest = { url:this[XHR_URL_KEY], headers:this[XHR_HEADERS_KEY] };
                window[LAST_REQ_GLOBAL_KEY] = lastOkRequest;
                const n = document.getElementById(NS+'note'); if (n) n.style.display='none';
              }
            }
          }catch{}
          this.removeEventListener('load', onload);
        };
        this.addEventListener('load', onload);
        return send.apply(this,arguments);
      };
      window[XHR_HOOK_FLAG] = true;
    }
  }

  /* ======================= Styles + UI ======================= */
  function ensureStyles(){
    if (document.getElementById(NS+'style')) return;
    const s=document.createElement('style'); s.id=NS+'style';
    s.textContent=`
      .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:6px 12px;border-radius:999px;font:600 12px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
      .${NS}panel{position:fixed;top:72px;left:50%;transform:translateX(-50%);width:min(1800px,98vw);max-height:85vh;overflow-y:auto;overflow-x:hidden;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100001;display:none}
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
      .${NS}status-problem { color:#c62828; font-weight:700; }
      .${NS}tabs{display:flex;gap:6px;align-items:center}
      .${NS}btn.active{background:#111827;color:#fff;border-color:#111827}
      .${NS}kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:8px;margin:8px 0 10px}
      .${NS}kpi{border:1px solid rgba(0,0,0,.12);border-radius:10px;padding:8px;background:#fff}
      .${NS}kpi h4{margin:0 0 6px;font:700 12px system-ui}
      .${NS}kpi .n{font:700 18px/1 system-ui}
      .${NS}kpi .m{margin-top:4px;font:12px system-ui;opacity:.75}
      .${NS}badge{display:inline-block;padding:2px 6px;border-radius:999px;border:1px solid rgba(0,0,0,.12);font:600 11px system-ui;background:#fff}
      .${NS}badge.open{background:#fff8e1}
      .${NS}badge.ack{background:#e8f5e9}
      .${NS}tour-group{background:#f3f4f6;font:700 12px system-ui;cursor:pointer}
      .${NS}tour-group:hover{background:#e8eaef}
      .${NS}tour-title{display:inline-flex;align-items:center;gap:8px}
      .${NS}tour-arrow{display:inline-block;transition:transform .18s ease}
      .${NS}tour-arrow.open{transform:rotate(90deg)}
      .${NS}tour-meta{display:flex;gap:14px;flex-wrap:wrap;align-items:center}
      .${NS}count-problem{color:#c62828;font-weight:700}
      .${NS}count-hint{color:#9a6700;font-weight:700}
      .${NS}bridge-state{font:700 11px system-ui;color:#4b5563;padding:2px 0}
    `;
    document.head.appendChild(s);
  }

  function initSavedCollapser(){
    const wrap = document.getElementById(NS+'saved-wrap');
    const chev = document.getElementById(NS+'saved-chev');
    const toggle = document.getElementById(NS+'saved-toggle');
    if (!wrap || !chev || !toggle) return;
    const LSKEY = STORAGE_PREFIX + 'saved.collapsed';
    let collapsed = localStorage.getItem(LSKEY);
    if (collapsed == null) collapsed = localStorage.getItem('pm.savedCust.collapsed');
    collapsed = collapsed == null ? '1' : collapsed;
    wrap.classList.toggle('collapsed', collapsed === '1');
    chev.classList.toggle('rot', collapsed !== '1');
    toggle.onclick = () => {
      const isColl = wrap.classList.toggle('collapsed');
      chev.classList.toggle('rot', !isColl);
      localStorage.setItem(LSKEY, isColl ? '1' : '0');
    };
  }

  let latestRows = [];
  let latestAlertRows = [];
  let latestAlertSummary = null;
  let currentView = 'customers';
  const TOUR_COLLAPSE_KEY = NS + 'tour.collapsed';
  let tourCollapseState = null;

  /* ======================= Auto-Refresh ======================= */
  const AUTO_KEY = NS + 'auto.enabled';
  const AUTO_MS  = 60 * 1000;
  let autoTimer  = null;
  function autoIsOn(){ return localStorage.getItem(AUTO_KEY) !== '0'; }
  function refreshActiveView(){ if (currentView === 'alerts') loadAlerts(); else loadDetails(); }
  function autoStart(){ autoStop(); refreshActiveView(); autoTimer = setInterval(() => { try{ refreshActiveView(); }catch{} }, AUTO_MS);}
  function autoStop(){ if (autoTimer){ clearInterval(autoTimer); autoTimer = null; } }
  function autoSet(on){ localStorage.setItem(AUTO_KEY, on ? '1' : '0'); on ? autoStart() : autoStop(); updateAutoBtn(); }
  function updateAutoBtn(){
    const b = document.getElementById(NS+'auto'); if (!b) return;
    const on = autoIsOn();
    b.innerHTML = `<span class="${NS}dot ${on ? 'on' : ''}"></span> Auto 60s`;
    b.title = on ? 'Auto-Refresh ist EIN (alle 60s). Klicken zum Ausschalten.' : 'Auto-Refresh ist AUS. Klicken zum Einschalten.';
    b.setAttribute('aria-pressed', on ? 'true' : 'false');
  }

  /* ======================= Alerts Bridge (Tamper -> Dashboard) ======================= */
  const BRIDGE_ENABLED = true;
  const BRIDGE_PRIMARY_ENDPOINT = 'http://10.14.7.169/DPD_Dashboard/dispatcher_alerts_push.php';
  const BRIDGE_FALLBACK_ENDPOINTS = [];
  const BRIDGE_ENDPOINT_STORAGE_KEY = 'da_alerts_bridge_endpoint';
  const BRIDGE_DEPOT_STORAGE_KEY = 'da_alerts_bridge_depot';
  const BRIDGE_CLIENT_ID = 'roadnet-default';
  const BRIDGE_DEPOT = '';
  const BRIDGE_TOKEN = '12c87869d52fc4651843512b595fc79ba8251d450988cbf4199e5918b33ecbd0';
  const BRIDGE_PUSH_INTERVAL_MS = 60 * 1000;
  const BRIDGE_MAX_ROWS = 500;
  const BRIDGE_REQUIRED_STATUSES = 'OPEN';
  let bridgePollTimer = null;
  const bridgeState = {
    inFlight: false,
    lastSig: '',
    lastSentAt: 0
  };

  let sortState = { key: null, dir: 1 };
  const collator = new Intl.Collator('de-DE', { numeric: true, sensitivity: 'base' });
  const sortGetters = {
    number:r=>r.number,tour:r=>r.tour,name:r=>r.name,street:r=>r.street,plz:r=>Number(r.plz)||0,ort:r=>r.ort,phone:r=>r.phone,
    driverName:r=>r.driverName,driverPhone:r=>r.driverPhone,predict:r=>r.predict,pickup:r=>r.pickup,status:r=>r.status,
    hintText:r=>r.hintText,lastScanRaw:r=>r.lastScanRaw??new Date(0)
  };
  const ALERT_TYPES = ['COLLECTION_SCHEDULED','COLLECTION_COMPLAINT','COLLECTION_NOT_ACCEPTED','COLLECTION_CRITICAL','COLLECTION_DISTANCE_DEVIATION'];
  const ALERT_TYPE_LABELS = {
    COLLECTION_SCHEDULED: 'Unbestätigt (Scheduled)',
    COLLECTION_COMPLAINT: 'Beschwerden',
    COLLECTION_NOT_ACCEPTED: 'Zu disponieren / Nicht akzeptiert',
    COLLECTION_CRITICAL: 'Zeitkritisch',
    COLLECTION_DISTANCE_DEVIATION: 'Distanzabweichung'
  };
  function alertTypeLabel(v){ return ALERT_TYPE_LABELS[v] || String(v || '—'); }
  function sortRows(rows){
    if (!sortState.key) return rows;
    const get = sortGetters[sortState.key]; if (!get) return rows;
    const dir = sortState.dir;
    return [...rows].sort((a,b)=>{
      const va = get(a), vb = get(b);
      if (va instanceof Date || vb instanceof Date){ return dir * ((va instanceof Date?va:new Date(va)) - (vb instanceof Date?vb:new Date(vb))); }
      if (typeof va === 'number' || typeof vb === 'number'){ return dir * ((Number(va)||0) - (Number(vb)||0)); }
      return dir * collator.compare(String(va ?? ''), String(vb ?? ''));
    });
  }

  function mountUI(){
    ensureStyles();
    if (document.getElementById(NS+'panel')) return;

    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel';
    panel.innerHTML=`
      <div class="${NS}head">
        <div class="${NS}row" style="width:100%;justify-content:space-between">
          <span class="${NS}muted">Kundennummern automatisch aus Abholung + Alert-Monitoring gemäß Dispatcher.</span>
          <div class="${NS}tabs">
            <button class="${NS}btn" id="${NS}tab-customers" type="button">Kunden</button>
            <button class="${NS}btn" id="${NS}tab-alerts" type="button">Alerts</button>
          </div>
        </div>

        <div class="${NS}row" id="${NS}ctrl-customers">
          <input id="${NS}tour" class="${NS}inp" placeholder="Tour-Filter (optional)">
          <button class="${NS}btn" id="${NS}load"  type="button">Liste laden (API)</button>
        </div>

        <div class="${NS}row" id="${NS}ctrl-alerts" style="display:none">
          <select id="${NS}alertType" class="${NS}inp" title="Alerttyp">
            <option value="">Alle Alerttypen</option>
            ${ALERT_TYPES.map(t => `<option value="${esc(t)}">${esc(alertTypeLabel(t))}</option>`).join('')}
          </select>
          <select id="${NS}alertStatus" class="${NS}inp" title="Alertstatus">
            <option value="OPEN">OPEN</option>
          </select>
          <input id="${NS}alertTour" class="${NS}inp" placeholder="Tour-Filter Alerts">
          <input id="${NS}alertSearch" class="${NS}inp" placeholder="Suche Kunde/Adresse/Depot">
          <button class="${NS}btn" id="${NS}loadAlerts" type="button">Alerts laden</button>
          <button class="${NS}btn" id="${NS}bridgeTest" type="button">Bridge Test</button>
          <button class="${NS}btn" id="${NS}bridgeDepot" type="button">Depot setzen</button>
          <span id="${NS}bridgeState" class="${NS}bridge-state">Bridge: Idle</span>
        </div>

        <div class="${NS}row">
          <button class="${NS}btn" id="${NS}auto" type="button" aria-pressed="true"></button>
          <button class="${NS}btn" id="${NS}close" type="button">Schließen</button>
        </div>
      </div>
      <div class="${NS}list">
        <div id="${NS}note" class="${NS}muted" style="margin-bottom:8px">Hinweis: Kundenliste bleibt bei leeren Zwischenständen erhalten. Alerts laden nur OPEN für PICKUP.</div>
        <div id="${NS}saved"></div>
        <div id="${NS}alerts-kpi" style="display:none"></div>
        <div id="${NS}out-customers"></div>
        <div id="${NS}out-alerts" style="display:none"></div>
      </div>`;
    document.body.appendChild(panel);

    document.getElementById(NS+'close').onclick = ()=>togglePanel(false);
    document.getElementById(NS+'tab-customers').onclick = ()=>setView('customers');
    document.getElementById(NS+'tab-alerts').onclick = ()=>setView('alerts');
    document.getElementById(NS+'load').onclick = loadDetails;
    document.getElementById(NS+'loadAlerts').onclick = loadAlerts;
    document.getElementById(NS+'bridgeTest').onclick = async () => {
      await loadAlerts({ silent: currentView !== 'alerts', forceStatuses: BRIDGE_REQUIRED_STATUSES, skipBridgeSync: true });
      syncAlertBridgeSnapshot(true, 'manual');
    };
    document.getElementById(NS+'bridgeDepot').onclick = () => {
      ensureBridgeDepotConfigured(true);
      syncAlertBridgeStatusLabel();
    };

    const autoBtn = document.getElementById(NS+'auto');
    if (autoBtn){ autoBtn.onclick = () => autoSet(!autoIsOn()); updateAutoBtn(); }

    document.getElementById(NS+'tour').addEventListener('input', ()=>{ if (currentView === 'customers') renderTable(applyTourFilter(latestRows)); });
    document.getElementById(NS+'alertType').addEventListener('change', ()=>renderAlertTable(applyAlertFilter(latestAlertRows)));
    document.getElementById(NS+'alertTour').addEventListener('input', ()=>renderAlertTable(applyAlertFilter(latestAlertRows)));
    document.getElementById(NS+'alertSearch').addEventListener('input', ()=>renderAlertTable(applyAlertFilter(latestAlertRows)));
    document.getElementById(NS+'alertStatus').addEventListener('change', ()=>loadAlerts());

    renderList();
    updateAutoBtn();
    setView('customers');
    if (autoIsOn()) autoStart();
    if (BRIDGE_ENABLED) bridgeStart();

    document.addEventListener('visibilitychange', () => {
      if (document.hidden) autoStop();
      else if (autoIsOn()) autoStart();
    });
  }

  function togglePanel(force){
    const panel=document.getElementById(NS+'panel'); if(!panel) return;
    const isHidden=getComputedStyle(panel).display==='none';
    const show = force!=null ? !!force : isHidden;
    panel.style.setProperty('display', show?'block':'none', 'important');
  }

  function setView(view){
    currentView = view === 'alerts' ? 'alerts' : 'customers';
    const isAlerts = currentView === 'alerts';

    const ctrlCustomers = document.getElementById(NS+'ctrl-customers');
    const ctrlAlerts = document.getElementById(NS+'ctrl-alerts');
    const saved = document.getElementById(NS+'saved');
    const outCustomers = document.getElementById(NS+'out-customers');
    const outAlerts = document.getElementById(NS+'out-alerts');
    const alertsKpi = document.getElementById(NS+'alerts-kpi');
    const tabCustomers = document.getElementById(NS+'tab-customers');
    const tabAlerts = document.getElementById(NS+'tab-alerts');

    if (ctrlCustomers) ctrlCustomers.style.display = isAlerts ? 'none' : 'flex';
    if (ctrlAlerts) ctrlAlerts.style.display = isAlerts ? 'flex' : 'none';
    if (saved) saved.style.display = isAlerts ? 'none' : 'block';
    if (outCustomers) outCustomers.style.display = isAlerts ? 'none' : 'block';
    if (alertsKpi) alertsKpi.style.display = isAlerts ? 'block' : 'none';
    if (outAlerts) outAlerts.style.display = isAlerts ? 'block' : 'none';
    if (tabCustomers) tabCustomers.classList.toggle('active', !isAlerts);
    if (tabAlerts) tabAlerts.classList.toggle('active', isAlerts);

    if (isAlerts) {
      if (!latestAlertRows.length) loadAlerts();
      else {
        renderAlertKpis(latestAlertSummary, latestAlertRows);
        renderAlertTable(applyAlertFilter(latestAlertRows));
      }
    } else {
      if (!latestRows.length) loadDetails();
      else renderTable(applyTourFilter(latestRows));
    }
  }
  /* ======================= Sichtbare Kundennummern ======================= */
  function renderList(){
    const saved = loadList();
    const box = document.getElementById(NS+'saved');
    if (!box) return;
    if (!saved.length){ box.innerHTML = `<div class="${NS}muted">Noch keine Kundennummern erkannt.</div>`; return; }
    box.innerHTML = `
      <div class="${NS}cust-collapser" id="${NS}saved-toggle" style="margin-bottom:6px;">
        <span class="${NS}chev" id="${NS}saved-chev">▸</span>
        <span>Automatisch erkannt (${saved.length}):</span>
      </div>
      <div class="${NS}cust-wrap" id="${NS}saved-wrap">
        <div style="display:flex;flex-wrap:wrap;gap:6px">
          ${saved.map(n=>`<span class="${NS}chip">${esc(n)}</span>`).join('')}
        </div>
      </div>`;
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
    // Autorisierung immer aus aktuellem Cookie aktualisieren (Token-Rotation)
    try {
     const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
      if (m) H.set('Authorization', 'Bearer ' + decodeURIComponent(m[1]));    } catch {}
    }catch{}
    return H;
  }

  // Versucht „Zusatzcode“/Problemcode aus dem Datensatz zu lesen
    function getZusatzcode(r){
    // robust – berücksichtigt Array-Felder wie additionalCodes
    try {
      if (Array.isArray(r.additionalCodes) && r.additionalCodes.length > 0) {
        return String(r.additionalCodes[0]).trim();
      }
      // Fallbacks für mögliche andere Varianten:
      const cand = r.zusatzcode || r.zusatzCode || r.problemCode || r.additionalCode || r.addCode || '';
      return String(cand || '').trim();
    } catch {
      return '';
    }
  }

  /* ======================= pickup-delivery ======================= */
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

  /* ======================= vehicle-overview ======================= */
  async function fetchDriverPhoneMap(){
    if(!lastOkRequest) throw new Error('Kein Auth-Kontext. Bitte einmal die normale Liste laden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size=500, maxPages=20;
    let page=1;
    const map = new Map();

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
          name:String(it.courierName ?? it.driverName ?? '').trim(),
          phone:String(it.courierPhone ?? it.driverPhone ?? '').trim(),
          openDeliveryStops: Number(it.openDeliveryStops ?? it.deliveryOpenStops ?? 0),
          openPickupStops: Number(it.openPickupStops ?? 0)
        };
        for (const k of tourKeys(rawTour)) {
          if (!map.has(k)) map.set(k, info);
          else {
            const cur = map.get(k);
            if (!cur.name  && info.name)  cur.name  = info.name;
            if (!cur.phone && info.phone) cur.phone = info.phone;
            if (Number.isFinite(info.openDeliveryStops) && (!Number.isFinite(cur.openDeliveryStops) || info.openDeliveryStops > cur.openDeliveryStops)) cur.openDeliveryStops = info.openDeliveryStops;
            if (Number.isFinite(info.openPickupStops) && (!Number.isFinite(cur.openPickupStops) || info.openPickupStops > cur.openPickupStops)) cur.openPickupStops = info.openPickupStops;
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

  function getCustomerFromRow(r){ const raw = Array.isArray(r.customerNo) ? (r.customerNo[0] || '') : (r.customerNumber || r.customer || r.customerId || ''); return String(raw); }
  function getName(r){ return r.name ?? (Array.isArray(r.customerName) ? r.customerName[0] : r.customerName) ?? ''; }
  function getStreet(r){ const s=r.street ?? ''; const hn=r.houseno ?? r.houseNumber ?? ''; return [s,hn].filter(Boolean).join(' '); }
  function getPredict(r){ const a=r.from2 ?? r.timeFrom2 ?? ''; const b=r.to2 ?? r.timeTo2 ?? ''; return buildWindow(a,b); }
  function getPickup(r){ const w1=buildWindow(r.timeFrom1, r.timeTo1); const w2=buildWindow(r.timeFrom2, r.timeTo2); return [w1,w2].filter(Boolean).join(' | '); }
  function getStatus(r){ return r.pickupStatus ?? r.deliveryStatus ?? r.status ?? ''; }
  function getPhone(r){ if (Array.isArray(r.addressPhone) && r.addressPhone.length) return String(r.addressPhone[0]); return r.phone || r.contactPhone || ''; }

  function _toDateSafe(v){ if(!v) return null; const d=new Date(v); return isNaN(d)?null:d; }
  function lastScanFromStop(s){
    if (s.scanDate && s.scanTime) {
      const d = new Date(`${s.scanDate}T${s.scanTime}`);
      if (!isNaN(d)) return d;
    }
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
  function getFirstParcel(r){ return r.parcelNumber || (Array.isArray(r.parcelNumbers) ? r.parcelNumbers[0] : ''); }

  async function fetchTrackingByTour(depot, tour, dateStr){
    if(!lastOkRequest) throw new Error('Kein Auth-Kontext. Bitte einmal die normale Liste laden.');
    const headers = buildHeaders(lastOkRequest.headers);
    let dep = String(depot || '').trim();
    if (!dep) dep = lastOkRequest?.url?.searchParams?.get('depot') || '';
    if (!dep) { console.warn('[tracking] Kein Depot -> übersprungen'); return null; }
    const tourPadded = String(tour ?? '').trim().replace(/^0+/, '').padStart(3,'0');
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
    let latest = null, latestTs = null;
    for (const s of stops){
      const ts = lastScanFromStop(s);
      if(!ts) continue;
      if(!latestTs || ts > latestTs){ latestTs = ts; latest = s; }
    }
    if (!latest) return null;
    const lat = latest.gpsLat ?? latest.gpslat ?? latest.latitude ?? latest.plannedCoordinateLat ?? null;
    const lon = latest.gpsLong ?? latest.gpsLon ?? latest.gpslon ?? latest.longitude ?? latest.plannedCoordinateLong ?? null;
    return { lastScan: latestTs, lat, lon, tour: tourPadded };
  }

  function matchesSaved(cellRaw, wantSet){
    const raw=digits(cellRaw); const pure=raw.replace(/^0+/, '');
    for(const w of wantSet){ if(!w) continue; if (pure===w) return true; if (pure.endsWith(w)) return true; if (raw.endsWith(w)) return true; }
    return false;
  }
  function extractNumbersFromText(text){
    if (!text) return [];
    return [...String(text).matchAll(/\d{5,}/g)].map(m => normShort(m[0])).filter(Boolean);
  }
  function extractCustomerNumbersFromRows(rows){
    const found = new Set();
    for (const r of (rows || [])) {
      const n = normShort(getCustomerFromRow(r));
      if (n) found.add(n);
    }
    return [...found];
  }
  function extractCustomerNumbersFromDom(){
    const found = new Set();
    const addText = (value) => {
      for (const n of extractNumbersFromText(value)) found.add(n);
    };
    const addNode = (node) => {
      if (!node) return;
      addText(node.textContent || '');
      if (typeof node.getAttribute === 'function') addText(node.getAttribute('title') || '');
    };

    const columnSelectors = [
      '[data-column="customerNo"]',
      '[data-column="customerNumber"]',
      '[data-field="customerNo"]',
      '[data-field="customerNumber"]',
      '[data-col-id*="customer"]',
      '[data-col-id*="kunde"]',
      '[col-id*="customer"]',
      '[col-id*="kunde"]',
      '.ag-cell[col-id*="customer"]',
      '.ag-cell[col-id*="kunde"]'
    ];
    for (const sel of columnSelectors) {
      document.querySelectorAll(sel).forEach(addNode);
    }

    document.querySelectorAll('table').forEach(table => {
      const headers = [...table.querySelectorAll('thead th')].map(th => String(th.textContent || '').trim().toLowerCase());
      const idx = headers.findIndex(h => /kund|customer/.test(h));
      if (idx < 0) return;
      table.querySelectorAll('tbody tr').forEach(tr => {
        const cells = tr.querySelectorAll('td,th');
        if (cells[idx]) addNode(cells[idx]);
      });
    });

    return [...found];
  }
  function collectCustomerNumbers(rowsApi){
    const fromApi = extractCustomerNumbersFromRows(rowsApi);
    if (fromApi.length) return { numbers: fromApi, source: 'api' };
    const fromDom = extractCustomerNumbersFromDom();
    if (fromDom.length) return { numbers: fromDom, source: 'dom' };
    return { numbers: [], source: 'none' };
  }
  function mergeDetectedCustomers(existingList, detectedList){
    const set = new Set((existingList || []).map(normShort).filter(Boolean));
    let added = 0;
    for (const n of (detectedList || []).map(normShort).filter(Boolean)) {
      if (!set.has(n)) {
        set.add(n);
        added++;
      }
    }
    return { list:[...set], added, changed: added > 0 };
  }

  function loadTourCollapseState(){
    if (tourCollapseState && typeof tourCollapseState === 'object') return tourCollapseState;
    try {
      const parsed = JSON.parse(localStorage.getItem(TOUR_COLLAPSE_KEY) || '{}');
      tourCollapseState = (parsed && typeof parsed === 'object') ? parsed : {};
    } catch {
      tourCollapseState = {};
    }
    return tourCollapseState;
  }
  function isTourCollapsed(tour){
    const map = loadTourCollapseState();
    if (Object.prototype.hasOwnProperty.call(map, tour)) return !!map[tour];
    return true; // default: zugeklappt
  }
  function setTourCollapsed(tour, collapsed){
    const map = loadTourCollapseState();
    map[tour] = !!collapsed;
    try { localStorage.setItem(TOUR_COLLAPSE_KEY, JSON.stringify(map)); } catch {}
  }

  function applyTourFilter(rows){
    const inp = document.getElementById(NS+'tour');
    const v = (inp && inp.value || '').trim().toLowerCase();
    if(!v) return rows;
    return (rows||[]).filter(r => String(r.tour||'').toLowerCase().includes(v));
  }

  function toArray(v){
    if (Array.isArray(v)) return v;
    if (v == null || v === '') return [];
    return [v];
  }
  function getAlertStatusesSelection(){
    const el = document.getElementById(NS+'alertStatus');
    return (el && el.value) ? String(el.value) : 'OPEN';
  }
  function formatScanDateTime(scanDate, scanTime){
    if (scanDate && scanTime) {
      const d = new Date(`${scanDate}T${scanTime}`);
      if (!isNaN(d)) return d.toLocaleString('de-DE', { hour:'2-digit', minute:'2-digit', day:'2-digit', month:'2-digit' });
    }
    return '—';
  }
  function buildAlertWindow(r){
    const w1 = buildWindow(r.from1, r.to1);
    const w2 = buildWindow(r.from2, r.to2);
    const val = [w1,w2].filter(Boolean).join(' | ');
    return val || '—';
  }
  function buildAlertAddress(r){
    return [
      [r.street || '', r.houseno || r.houseNo || ''].filter(Boolean).join(' '),
      [digits(r.postalCode || r.zip || ''), r.city || ''].filter(Boolean).join(' ')
    ].filter(Boolean).join(', ');
  }
  function buildAlertMapLink(r){
    const lat = r.gpsLat ?? r.gpslat ?? r.latitude ?? r.plannedGpsLat ?? null;
    const lon = r.gpsLong ?? r.gpsLon ?? r.gpslon ?? r.longitude ?? r.plannedGpsLong ?? null;
    const addr = buildAlertAddress(r);
    if (lat != null && lon != null && addr) return `https://www.google.com/maps/dir/?api=1&origin=${lat},${lon}&destination=${encodeURIComponent(addr)}&travelmode=driving`;
    if (lat != null && lon != null) return `https://www.google.com/maps/search/?api=1&query=${lat},${lon}`;
    if (addr) return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(addr)}`;
    return '';
  }
  function normalizeAlertRow(r){
    const customerNoRaw = toArray(r.customerNo)[0] || r.customerNumber || '';
    const customerNo = normShort(customerNoRaw) || String(customerNoRaw || '—');
    const customerName = toArray(r.customerName)[0] || r.name || '—';
    const additionalCodes = toArray(r.additionalCodes).filter(Boolean).map(String);
    const distNum = Number(r.plannedRealDistanceDeviationMeters);
    const distText = Number.isFinite(distNum) ? `${Math.round(distNum)} m` : '—';

    const hints = [];
    if (r.alertType === 'COLLECTION_DISTANCE_DEVIATION' && Number.isFinite(distNum) && distNum > 250) hints.push('Distanz > 250m');
    if (r.alertType === 'COLLECTION_CRITICAL' || r.timeCritical) hints.push('zeitkritisch');
    if (String(r.pickupStatus || '').toUpperCase() === 'PROBLEM' && additionalCodes.length) hints.push(`Problemcode: ${additionalCodes.join(',')}`);

    return {
      id: r.id || `${r.alertType || ''}-${customerNo}-${r.tour || ''}-${r.scanDate || ''}-${r.scanTime || ''}`,
      alertType: String(r.alertType || ''),
      alertTypeLabel: alertTypeLabel(r.alertType),
      alertStatus: String(r.alertStatus || '—'),
      pickupStatus: String(r.pickupStatus || '—'),
      customerNo,
      customerName: String(customerName || '—'),
      tour: String(r.tour || '—'),
      depot: String(r.depot || '—'),
      parcels: `${r.completedParcels ?? 0}/${r.realParcels ?? 0}/${r.estimatedParcels ?? 0}`,
      additionalCodes: additionalCodes.join(', ') || '—',
      distanceNum: Number.isFinite(distNum) ? distNum : null,
      distanceText: distText,
      scanText: formatScanDateTime(r.scanDate, r.scanTime),
      windowText: buildAlertWindow(r),
      hintText: hints.join(' • ') || '—',
      mapLink: buildAlertMapLink(r),
      searchBlob: [customerNo, customerName, r.street, r.city, r.depot, r.tour].filter(Boolean).join(' ').toLowerCase()
    };
  }
  function applyAlertFilter(rows){
    const typeEl = document.getElementById(NS+'alertType');
    const tourEl = document.getElementById(NS+'alertTour');
    const qEl = document.getElementById(NS+'alertSearch');
    const type = (typeEl && typeEl.value || '').trim();
    const tour = (tourEl && tourEl.value || '').trim().toLowerCase();
    const q = (qEl && qEl.value || '').trim().toLowerCase();

    return (rows || []).filter(r => {
      if (type && r.alertType !== type) return false;
      if (tour && !String(r.tour || '').toLowerCase().includes(tour)) return false;
      if (q && !String(r.searchBlob || '').includes(q)) return false;
      return true;
    });
  }

  function setAlertBridgeStatus(label, color){
    const el = document.getElementById(NS+'bridgeState');
    if (!el) return;
    el.textContent = `Bridge: ${String(label || 'Idle')}`;
    el.style.color = color || '#4b5563';
  }
  function syncAlertBridgeStatusLabel(){
    const depot = getConfiguredBridgeDepot();
    if (!depot) {
      setAlertBridgeStatus('Depot fehlt', '#991b1b');
      return;
    }
    setAlertBridgeStatus(`Depot ${depot}`, '#4b5563');
  }
  function normalizeBridgeEndpoint(raw){
    const value = String(raw || '').trim();
    if (!/^https?:\/\//i.test(value)) return '';
    return value.replace(/\/+$/, '');
  }
  function buildBridgeEndpoints(){
    const out = [];
    const seen = new Set();
    const push = (value) => {
      const endpoint = normalizeBridgeEndpoint(value);
      if (!endpoint || seen.has(endpoint)) return;
      seen.add(endpoint);
      out.push(endpoint);
    };
    push(BRIDGE_PRIMARY_ENDPOINT);
    (Array.isArray(BRIDGE_FALLBACK_ENDPOINTS) ? BRIDGE_FALLBACK_ENDPOINTS : []).forEach(push);
    try { push(localStorage.getItem(BRIDGE_ENDPOINT_STORAGE_KEY) || ''); } catch {}
    return out;
  }
  function postBridgePayload(payload, endpoints, idx = 0){
    if (!Array.isArray(endpoints) || idx >= endpoints.length) {
      return Promise.reject(new Error('no_endpoint_available'));
    }
    const endpoint = endpoints[idx];
    return fetch(endpoint, {
      method: 'POST',
      mode: 'cors',
      credentials: 'omit',
      headers: {
        'Content-Type': 'application/json',
        'X-Dispatcher-Token': BRIDGE_TOKEN
      },
      body: JSON.stringify(payload)
    })
      .then(r => r.text().then(body => ({ r, body })))
      .then(({ r, body }) => {
        if (!r.ok) {
          return postBridgePayload(payload, endpoints, idx + 1);
        }
        try { localStorage.setItem(BRIDGE_ENDPOINT_STORAGE_KEY, endpoint); } catch {}
        return { endpoint, body };
      })
      .catch(() => postBridgePayload(payload, endpoints, idx + 1));
  }
  function normalizeBridgeDepot(raw){
    const t = String(raw || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (/^D\d{4}$/.test(t)) return t;
    if (/^D\d{3}$/.test(t)) return `D0${t.slice(1)}`;
    if (/^\d{4}$/.test(t)) return `D${t}`;
    if (/^\d{3}$/.test(t)) return `D0${t}`;
    return '';
  }
  function readStoredBridgeDepot(){
    try { return normalizeBridgeDepot(localStorage.getItem(BRIDGE_DEPOT_STORAGE_KEY) || ''); }
    catch { return ''; }
  }
  function writeStoredBridgeDepot(raw){
    const depot = normalizeBridgeDepot(raw);
    try {
      if (depot) localStorage.setItem(BRIDGE_DEPOT_STORAGE_KEY, depot);
      else localStorage.removeItem(BRIDGE_DEPOT_STORAGE_KEY);
    } catch {}
    return depot;
  }
  function getConfiguredBridgeDepot(){
    const fromConfig = normalizeBridgeDepot(BRIDGE_DEPOT);
    if (fromConfig) return fromConfig;
    const fromStorage = readStoredBridgeDepot();
    if (fromStorage) return fromStorage;
    return '';
  }
  function ensureBridgeDepotConfigured(forcePrompt){
    const configured = getConfiguredBridgeDepot();
    if (configured) return configured;
    if (!forcePrompt) return '';
    while (true) {
      const input = window.prompt('Bitte Standort/Depot eingeben (z.B. D0157, 0157 oder 157).', '');
      if (input == null) return '';
      const depot = normalizeBridgeDepot(input);
      if (!depot) {
        alert('Ungueltiges Depot. Erlaubt sind z.B. D0157, 0157 oder 157.');
        continue;
      }
      return writeStoredBridgeDepot(depot);
    }
  }
  function buildAlertBridgeRows(rows, depot){
    const out = [];
    for (const row of (rows || [])) {
      const rowDepot = normalizeBridgeDepot(row && row.depot ? row.depot : '');
      if (rowDepot && rowDepot !== depot) continue;
      const status = String(row && row.alertStatus || '').toUpperCase().trim();
      if (status !== 'OPEN') continue;
      out.push({
        id: String(row && row.id || ''),
        alertType: String(row && row.alertType || ''),
        alertTypeLabel: String(row && row.alertTypeLabel || ''),
        alertStatus: String(row && row.alertStatus || ''),
        pickupStatus: String(row && row.pickupStatus || ''),
        customerNo: String(row && row.customerNo || ''),
        customerName: String(row && row.customerName || ''),
        tour: String(row && row.tour || ''),
        depot: depot,
        parcels: String(row && row.parcels || ''),
        additionalCodes: String(row && row.additionalCodes || ''),
        distanceText: String(row && row.distanceText || ''),
        scanText: String(row && row.scanText || ''),
        windowText: String(row && row.windowText || ''),
        hintText: String(row && row.hintText || ''),
        mapLink: String(row && row.mapLink || '')
      });
      if (out.length >= BRIDGE_MAX_ROWS) break;
    }
    return out;
  }
  function buildAlertBridgeKpis(rows){
    const byType = {};
    const byStatus = { OPEN: 0, ACKNOWLEDGED: 0, OTHER: 0 };
    for (const t of ALERT_TYPES) byType[t] = 0;
    for (const row of (rows || [])) {
      const status = String(row && row.alertStatus || '').toUpperCase();
      if (status === 'OPEN') byStatus.OPEN += 1;
      else if (status === 'ACKNOWLEDGED') byStatus.ACKNOWLEDGED += 1;
      else byStatus.OTHER += 1;
      const type = String(row && row.alertType || '');
      if (!Object.prototype.hasOwnProperty.call(byType, type)) byType[type] = 0;
      byType[type] += 1;
    }
    return {
      total: (rows || []).length,
      open: byStatus.OPEN,
      acknowledged: byStatus.ACKNOWLEDGED,
      other: byStatus.OTHER,
      byStatus,
      byType
    };
  }
  function buildAlertBridgeSignature(payload){
    let hash = 2166136261 >>> 0;
    const push = (value) => {
      const text = String(value == null ? '' : value);
      for (let i = 0; i < text.length; i++) {
        hash ^= text.charCodeAt(i);
        hash = Math.imul(hash, 16777619) >>> 0;
      }
      hash ^= 124;
      hash = Math.imul(hash, 16777619) >>> 0;
    };
    push(payload.depot);
    push(payload.statusScope);
    push((payload.rows || []).length);
    for (const row of (payload.rows || [])) {
      push(row.id);
      push(row.alertType);
      push(row.alertStatus);
      push(row.tour);
      push(row.customerNo);
      push(row.scanText);
    }
    return `${payload.depot}|${(payload.rows || []).length}|${hash.toString(16)}`;
  }
  function syncAlertBridgeSnapshot(force = false, trigger = 'auto'){
    if (!BRIDGE_ENABLED) return;
    if (bridgeState.inFlight) {
      if (trigger === 'manual') setAlertBridgeStatus('Senden laeuft bereits', '#b45309');
      return;
    }

    const depot = ensureBridgeDepotConfigured(trigger === 'manual');
    if (!depot) {
      setAlertBridgeStatus('Depot fehlt', '#991b1b');
      return;
    }

    const rows = buildAlertBridgeRows(latestAlertRows, depot);
    const kpis = buildAlertBridgeKpis(rows);
    const payload = {
      source: 'dispatcher_abholer_alerts',
      clientId: String(BRIDGE_CLIENT_ID || '').trim(),
      depot,
      statusScope: BRIDGE_REQUIRED_STATUSES,
      title: 'Dispatcher-Abholer-Alerts',
      stamp: new Date().toLocaleString('de-DE'),
      kpis,
      rows,
      sourceUrl: location.href,
      exportedAt: new Date().toISOString()
    };
    const signature = buildAlertBridgeSignature(payload);
    const nowMs = Date.now();
    if (!force && signature === bridgeState.lastSig && (nowMs - bridgeState.lastSentAt) < BRIDGE_PUSH_INTERVAL_MS - 1000) {
      return;
    }

    const endpoints = buildBridgeEndpoints();
    if (!endpoints.length) {
      setAlertBridgeStatus('Kein Endpoint', '#991b1b');
      return;
    }

    bridgeState.inFlight = true;
    setAlertBridgeStatus(`Sende (${rows.length})...`, '#1d4ed8');
    postBridgePayload(payload, endpoints, 0)
      .then(() => {
        bridgeState.lastSig = signature;
        bridgeState.lastSentAt = nowMs;
        setAlertBridgeStatus(`OK ${new Date().toLocaleTimeString('de-DE', { hour:'2-digit', minute:'2-digit', second:'2-digit' })}`, '#065f46');
      })
      .catch(() => {
        setAlertBridgeStatus('Fehler', '#991b1b');
      })
      .finally(() => {
        bridgeState.inFlight = false;
      });
  }
  function bridgeTick(){
    loadAlerts({ silent: currentView !== 'alerts', forceStatuses: BRIDGE_REQUIRED_STATUSES });
  }
  function bridgeStart(){
    bridgeStop();
    if (!BRIDGE_ENABLED) return;
    syncAlertBridgeStatusLabel();
    bridgeTick();
    bridgePollTimer = setInterval(() => { try { bridgeTick(); } catch {} }, BRIDGE_PUSH_INTERVAL_MS);
  }
  function bridgeStop(){
    if (!bridgePollTimer) return;
    clearInterval(bridgePollTimer);
    bridgePollTimer = null;
  }

  function buildAlertingUrl(alertType, alertStatuses, page){
    const u = new URL('https://dispatcher2-de.geopost.com/dispatcher/api/alerting');
    const today = new Date().toISOString().slice(0,10);
    u.searchParams.set('page', String(page));
    u.searchParams.set('pageSize', '250');
    u.searchParams.set('dateFrom', today);
    u.searchParams.set('dateTo', today);
    u.searchParams.set('depots', '');
    u.searchParams.set('tours', '');
    u.searchParams.set('customerNumbers', '');
    u.searchParams.set('orderTypes', 'PICKUP');
    u.searchParams.set('pudoIds', '');
    u.searchParams.set('pickupTypes', '');
    u.searchParams.set('pickupStatuses', '');
    u.searchParams.set('deliveryStatuses', '');
    u.searchParams.set('additionalCodes', '');
    u.searchParams.set('alertTypes', alertType || '');
    u.searchParams.set('alertStatuses', alertStatuses || 'OPEN');
    u.searchParams.set('eventIds', '');
    u.searchParams.set('elements', '');
    u.searchParams.set('sort', '');
    return u;
  }
  async function fetchAlertSummary(){
    const headers = buildHeaders(lastOkRequest?.headers);
    const u = new URL('https://dispatcher2-de.geopost.com/dispatcher/api/alerting/alerts');
    const r = await fetch(u.toString(), { credentials:'include', headers });
    if (!r.ok) throw new Error(`alerting/alerts ${r.status}`);
    return await r.json();
  }
  async function fetchAlertRowsByType(alertType, alertStatuses){
    const headers = buildHeaders(lastOkRequest?.headers);
    const size = 250;
    const maxPages = 20;
    const rows = [];
    let page = 1;

    while (page <= maxPages){
      const u = buildAlertingUrl(alertType, alertStatuses, page);
      const r = await fetch(u.toString(), { credentials:'include', headers });
      if (!r.ok) throw new Error(`alerting ${alertType} ${r.status}`);
      const j = await r.json();
      const chunk = (j.results || j.items || j.content || j.data || j.rows || []);
      if (!Array.isArray(chunk)) break;
      rows.push(...chunk);
      if (chunk.length < size) break;
      page++;
      await sleep(30);
    }
    return rows;
  }
  function renderAlertKpis(summary, rows){
    const box = document.getElementById(NS+'alerts-kpi');
    if (!box) return;
    const countsFromRows = {};
    for (const t of ALERT_TYPES) countsFromRows[t] = 0;
    for (const r of (rows || [])) {
      const t = String(r.alertType || '');
      if (Object.prototype.hasOwnProperty.call(countsFromRows, t)) countsFromRows[t] += 1;
    }

    const cards = ALERT_TYPES.map(t => {
      const n = countsFromRows[t] || 0;
      return `<div class="${NS}kpi"><h4>${esc(alertTypeLabel(t))}</h4><div class="n">${esc(n)}</div><div class="m">${esc(t)}</div></div>`;
    }).join('');

    const updated = new Date().toLocaleTimeString('de-DE', { hour:'2-digit', minute:'2-digit' });
    box.innerHTML = `<div class="${NS}muted" style="margin-bottom:6px">Aktualisiert: ${esc(updated)} · Statusfilter: ${esc(getAlertStatusesSelection())}</div><div class="${NS}kpi-grid">${cards}</div>`;
  }
  function renderAlertTable(rows){
    const out = document.getElementById(NS+'out-alerts');
    if (!out) return;

    const bodyHtml = (rows || []).map(r => {
      const statusClass = String(r.alertStatus || '').toUpperCase() === 'OPEN' ? 'open' : (String(r.alertStatus || '').toUpperCase() === 'ACKNOWLEDGED' ? 'ack' : '');
      const badge = `<span class="${NS}badge ${statusClass}">${esc(r.alertStatus || '—')}</span>`;
      const mapLink = r.mapLink ? `<a href="${esc(r.mapLink)}" target="_blank" rel="noopener">Karte</a>` : '—';
      return `<tr>
        <td>${esc(r.alertTypeLabel)}</td>
        <td>${badge}</td>
        <td>${esc(r.pickupStatus)}</td>
        <td>${esc(r.customerNo)}</td>
        <td>${esc(r.customerName)}</td>
        <td>${esc(r.tour)}</td>
        <td>${esc(r.depot)}</td>
        <td>${esc(r.parcels)}</td>
        <td>${esc(r.additionalCodes)}</td>
        <td>${esc(r.distanceText)}</td>
        <td>${esc(r.scanText)}</td>
        <td>${esc(r.windowText)}</td>
        <td>${esc(r.hintText)}</td>
        <td>${mapLink}</td>
      </tr>`;
    }).join('');

    out.innerHTML = `<table class="${NS}tbl"><thead><tr>
      <th>Alerttyp</th><th>Alertstatus</th><th>Pickupstatus</th><th>Kundennr.</th><th>Name</th><th>Tour</th><th>Depot</th><th>Pakete C/R/E</th><th>Zusatzcodes</th><th>Distanz</th><th>Scan</th><th>Zeitfenster</th><th>Hinweise</th><th>Karte</th>
    </tr></thead><tbody>${bodyHtml || '<tr><td colspan="14">Keine Alerts für den aktuellen Filter.</td></tr>'}</tbody></table>`;
  }

  let isLoadingAlerts = false;
  async function loadAlerts(opts = {}){
    if (isLoadingAlerts) return;
    const silent = !!opts.silent;
    isLoadingAlerts = true;
    const out = document.getElementById(NS+'out-alerts');
    if (out && !silent) out.innerHTML = `<div class="${NS}muted">Alerts werden geladen …</div>`;

    try {
      const statuses = String(opts.forceStatuses || getAlertStatusesSelection() || BRIDGE_REQUIRED_STATUSES);
      const summaryPromise = fetchAlertSummary();
      const detailPromises = ALERT_TYPES.map(t => fetchAlertRowsByType(t, statuses));
      const settled = await Promise.allSettled([summaryPromise, ...detailPromises]);

      const summaryRes = settled[0];
      latestAlertSummary = summaryRes.status === 'fulfilled' ? summaryRes.value : null;

      const combined = [];
      for (let i=1; i<settled.length; i++) {
        const res = settled[i];
        if (res.status === 'fulfilled' && Array.isArray(res.value)) combined.push(...res.value);
      }

      const normalized = combined.map(normalizeAlertRow).filter((r) => String(r && r.alertStatus || '').toUpperCase() === 'OPEN');
      const dedup = [];
      const seen = new Set();
      for (const r of normalized) {
        const k = `${r.id}|${r.alertType}|${r.alertStatus}`;
        if (seen.has(k)) continue;
        seen.add(k);
        dedup.push(r);
      }
      latestAlertRows = dedup;

      if (!silent || currentView === 'alerts') {
        renderAlertKpis(latestAlertSummary, latestAlertRows);
        renderAlertTable(applyAlertFilter(latestAlertRows));
      }
      if (!opts.skipBridgeSync && BRIDGE_ENABLED && statuses === BRIDGE_REQUIRED_STATUSES) {
        syncAlertBridgeSnapshot(false, 'auto');
      }
    } catch (e) {
      if (!silent && out) out.innerHTML = `<div class="${NS}muted">Alerts konnten nicht geladen werden: ${esc(e && e.message || e)}</div>`;
      if (BRIDGE_ENABLED) setAlertBridgeStatus('Fehler beim Laden', '#991b1b');
    } finally {
      isLoadingAlerts = false;
    }
  }

  let isLoading = false;
  async function loadDetails(){
    if (isLoading) return;           // Reentrancy-Schutz
    isLoading = true;
    const out=document.getElementById(NS+'out-customers');
    out.innerHTML = '<div class="'+NS+'muted">Lade …</div>';

    try{
      let rowsApi = [];
      let driverMap = new Map();
      let rowsApiFailed = false;
      const [rowsResult, mapResult] = await Promise.allSettled([fetchPagedAll(), fetchDriverPhoneMap()]);
      if (rowsResult.status === 'fulfilled') rowsApi = Array.isArray(rowsResult.value) ? rowsResult.value : [];
      else rowsApiFailed = true;
      if (mapResult.status === 'fulfilled') driverMap = mapResult.value || new Map();

      const existingShorts = loadList().map(normShort).filter(Boolean);
      const collected = collectCustomerNumbers(rowsApi);
      const detectedShorts = collected.numbers.map(normShort).filter(Boolean);
      const merged = mergeDetectedCustomers(existingShorts, detectedShorts);
      const savedShorts = merged.list;
      if (merged.changed) saveList(savedShorts);

      if (!savedShorts.length) {
        latestRows = [];
        out.innerHTML = `<div class="${NS}muted">Keine Kundennummern in Abholung gefunden (API/DOM).</div>`;
        return;
      }

      if (!rowsApi.length) {
        latestRows = [];
        const reason = rowsApiFailed ? ' Kein API-Request erkannt. Bitte einmal die normale Liste laden.' : '';
        const sourceText = detectedShorts.length ? collected.source.toUpperCase() : 'BESTAND';
        const deltaText = merged.added > 0 ? ` (+${merged.added} neu)` : '';
        out.innerHTML = `<div class="${NS}muted">Keine pickup-delivery Daten verfügbar. Verwende bestehende Liste (${savedShorts.length}${deltaText}, Quelle: ${sourceText}).${reason}</div>`;
        return;
      }

      const wantSet = new Set(savedShorts);
      const filteredSrc = rowsApi.filter(r => matchesSaved(getCustomerFromRow(r), wantSet));
      if (!filteredSrc.length) {
        latestRows = [];
        out.innerHTML = `<div class="${NS}muted">Keine Abholaufträge für die erkannten Kundennummern gefunden.</div>`;
        return;
      }

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
        }catch(e){ console.warn('[tracking] Fehler für Tour', t, e); }
      }

      const rows = filteredSrc.map(r => {
        const status = getStatus(r) || '—';
        const warnPredict = predictOutsidePickup(r);
        const {soon, closed} = closingHints(r);
        const hints=[];
        if (warnPredict) hints.push('Predict außerhalb Abholfenster');
        if (closed) hints.push('bereits geschlossen');
        else if (soon) hints.push('schließt in ≤30 Min');
        // Problemcode in Hinweise verschieben
        const extraCode = getZusatzcode(r);
        if (String(status).toUpperCase() === 'PROBLEM' && extraCode) {
          hints.push(`Problemcode: ${extraCode}`);
        }

        const tour = String(r.tour || '').trim();
        let drv = null;
        for (const k of tourKeys(tour)) { if (driverMap && driverMap.has(k)) { drv = driverMap.get(k); break; } }

        const tkey = String(r.tour || '').trim();
        const tinfo = latestByTour.get(tkey) || null;

        const address = [
          getStreet(r),
          [digits(r.postalCode || r.zip || ''), r.city || ''].filter(Boolean).join(' ')
        ].filter(Boolean).join(', ');

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
        driverOpenDeliveryStops: (drv && Number.isFinite(Number(drv.openDeliveryStops))) ? Number(drv.openDeliveryStops) : null,
          predict: getPredict(r) || '—',
          pickup:  getPickup(r)  || '—',
          status,
          isCompleted: String(status).toUpperCase()==='COMPLETED',
          isProblem:   String(status).toUpperCase()==='PROBLEM',
       //   extraCode: getZusatzcode(r) || '',
          warnRow: warnPredict,
          hintText: hints.join(' • ') || '—',
          lastScan: tinfo?.lastScan ? tinfo.lastScan.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '—',
          lastScanRaw: tinfo?.lastScan ?? null,
          gpsLat: tinfo?.lat ?? '',
          gpsLon: tinfo?.lon ?? '',
          gmaps
        };
      });

      latestRows = rows;
      renderTable(applyTourFilter(latestRows));
    } finally {
      isLoading = false;
    }
  }

  function renderTable(rows){
    const out = document.getElementById(NS+'out-customers');
    const cols = [
      {label:'Kundennr.', key:'number'},
      {label:'Tour', key:'tour'},
      {label:'Kundenname', key:'name'},
      {label:'Straße', key:'street'},
      {label:'PLZ', key:'plz'},
      {label:'Ort', key:'ort'},
      {label:'Telefon', key:'phone'},
      {label:'Fahrer', key:'driverName'},
      {label:'Fahrer Telefon', key:'driverPhone'},
      {label:'Predict Zeitfenster', key:'predict'},
      {label:'Zeitfenster Abholung', key:'pickup'},
      {label:'Status', key:'status'},
      {label:'Hinweise', key:'hintText'},
      {label:'Letzter Scan', key:'lastScanRaw'},
      {label:'Karte', key:null}
    ];

    const headHtml = cols.map(c=>{
      if (!c.key) return `<th>${esc(c.label)}</th>`;
      const isSorted = sortState.key === c.key;
      const arrow = isSorted ? (sortState.dir === 1 ? '▲' : '▼') : '';
      return `<th class="${NS}th-sort" data-key="${esc(c.key)}">${esc(c.label)}<span class="arrow">${arrow}</span></th>`;
    }).join('');

    const openRows = (rows || []).filter(r => !r.isCompleted);
    const grouped = new Map();
    for (const r of openRows) {
      const tour = String(r.tour || '').trim() || 'ohne Tour';
      if (!grouped.has(tour)) grouped.set(tour, []);
      grouped.get(tour).push(r);
    }

    const groups = [...grouped.entries()].map(([tour, groupRows]) => {
      const hintCount = groupRows.filter(r => {
        const h = String(r.hintText || '').trim();
        return h && h !== '—';
      }).length;
      const problemCount = groupRows.filter(r => !!r.isProblem).length;
      const driverOpenDeliveryStopsVals = groupRows
        .map(r => r.driverOpenDeliveryStops)
        .filter(v => Number.isFinite(Number(v)))
        .map(v => Number(v));
      const driverOpenDeliveryStops = driverOpenDeliveryStopsVals.length ? Math.max(...driverOpenDeliveryStopsVals) : null;
      const severity = (problemCount > 0 ? 2 : 0) + (hintCount > 0 ? 1 : 0);
      return { tour, rows: groupRows, hintCount, problemCount, severity, driverOpenDeliveryStops };
    });

    groups.sort((a,b) => {
      if (b.severity !== a.severity) return b.severity - a.severity;
      if (b.problemCount !== a.problemCount) return b.problemCount - a.problemCount;
      if (b.hintCount !== a.hintCount) return b.hintCount - a.hintCount;
      return collator.compare(String(a.tour), String(b.tour));
    });

    const bodyHtml = groups.map(group => {
      const rowsSorted = (sortState.key && sortState.key !== 'tour')
        ? sortRows(group.rows)
        : [...group.rows].sort((a,b)=>collator.compare(String(a.number || ''), String(b.number || '')));

      const collapsed = isTourCollapsed(group.tour);
      const arrowClass = collapsed ? '' : 'open';
      const tourKey = encodeURIComponent(group.tour);
      const openStopsText = Number.isFinite(group.driverOpenDeliveryStops) ? String(group.driverOpenDeliveryStops) : '—';

      const groupHead = `<tr class="${NS}tour-group" data-tour="${esc(tourKey)}"><td colspan="${cols.length}"><div class="${NS}tour-meta"><span class="${NS}tour-title"><span class="${NS}tour-arrow ${arrowClass}">▸</span>Tour ${esc(group.tour)}</span><span>Aufträge: ${rowsSorted.length}</span><span>Offene Zustellstops Fahrer: ${esc(openStopsText)}</span><span class="${NS}count-hint">Hinweise: ${group.hintCount}</span><span class="${NS}count-problem">Probleme: ${group.problemCount}</span></div></td></tr>`;

      if (collapsed) return groupHead;

      const detailHtml = rowsSorted.map(r => {
        const statusHtml =
          r.isProblem
            ? `<span class="${NS}status-problem">${esc(r.status||'—')}</span>`
            : esc(r.status||'—');

        const linkHtml = r.gmaps ? `<a href="${esc(r.gmaps)}" target="_blank" rel="noopener" title="${esc(r.gpsLat||'')}, ${esc(r.gpsLon||'')}">Karte</a>` : '—';
        const cells = [
          esc(r.number||'—'), esc(r.tour||'—'), esc(r.name||'—'), esc(r.street||'—'), esc(r.plz||'—'), esc(r.ort||'—'),
          esc(r.phone||'—'), esc(r.driverName||'—'), esc(r.driverPhone||'—'), esc(r.predict||'—'), esc(r.pickup||'—'),
          statusHtml, esc(r.hintText||'—'), esc(r.lastScan||'—'), linkHtml
        ];
        const trClass = r.warnRow ? ` class="${NS}warn-row"` : '';
        return `<tr${trClass}>${cells.map(v=>`<td>${v}</td>`).join('')}</tr>`;
      }).join('');

      return groupHead + detailHtml;
    }).join('');

    const emptyHtml = `<tr><td colspan="${cols.length}">Keine offenen Treffer.</td></tr>`;
    out.innerHTML = `<table class="${NS}tbl"><thead><tr>${headHtml}</tr></thead><tbody>${bodyHtml || emptyHtml}</tbody></table>`;

    out.querySelectorAll('th[data-key]').forEach(th=>{
      th.onclick = () => {
        const k = th.getAttribute('data-key');
        if (sortState.key === k) { sortState.dir = -sortState.dir; }
        else { sortState.key = k; sortState.dir = 1; }
        renderTable(applyTourFilter(latestRows));
      };
    });

    out.querySelectorAll(`tr.${NS}tour-group[data-tour]`).forEach(tr => {
      tr.onclick = (ev) => {
        if (ev.target && ev.target.closest && ev.target.closest('a,button,input,select,label')) return;
        const key = String(tr.getAttribute('data-tour') || '');
        const tour = decodeURIComponent(key);
        setTourCollapsed(tour, !isTourCollapsed(tour));
        renderTable(applyTourFilter(latestRows));
      };
    });
  }

  // Modul-Boot: Hooks + UI vorbereiten (kein eigener Button)
  function boot(){
    installHooksOnce();
    mountUI();
    // Auto-Refresh sauber stoppen, wenn Seite verlassen wird
    window.addEventListener('beforeunload', () => { try{ autoStop(); bridgeStop(); }catch{} });
  }
})();
