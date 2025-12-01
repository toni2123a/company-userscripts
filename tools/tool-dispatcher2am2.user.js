// ==UserScript==
// @name         DPD Dispatcher – Prio/Express Monitoring
// @namespace    bodo.dpd.custom
// @version      6.6.0
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @description  PRIO/EXPRESS12: KPIs & Listen. Status/Servicecode direkt aus API, sortierbare Spalten, Predict-Zeitfenster, Zustellzeit, Button „EXPRESS12 >11:01“. Panel bleibt offen; PSN mit Auge-Button öffnet Scanserver.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  /* ====== TEIL 1/10 – Loader-Registrierung + Grundgerüst ====== */

  const moduleDef = {
    id: 'prio-express-monitor',
    label: 'Prioritäts-/Express-Überwachung',
    run: () => startModuleOnce()
  };

  if (window.TM && typeof window.TM.register === 'function') {
    window.TM.register(moduleDef);
  } else {
    window.__tmQueue = window.__tmQueue || [];
    window.__tmQueue.push(moduleDef);
  }

  let started = false;
  function startModuleOnce() {
    if (started) { togglePanel(true); return; }
    started = true;
    boot();
    togglePanel(true);
  }

  const NS  = 'pm-';
  const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const sleep = ms => new Promise(r=>setTimeout(r,ms));
  const norm = s => String(s||'').replace(/\s+/g,' ').trim();

  const state = {
    filterExpress: 'all',
    events: [],
    nextId: 1,
    _bootShown: false,
    _prioAllList: [],
    _prioOpenList: [],
    _expAllList: [],
    _expOpenList: [],
    _expLate11List: [],
    _modal: { rows: [], opts: {}, title: '' }
  };

  let lastOkRequest = null;
  let autoEnabled   = true;
  let autoTimer     = null;
  let isBusy        = false;
  let isLoading     = false;

  const collator = new Intl.Collator('de', { numeric:true, sensitivity:'base' });

  /* ====== TEIL 2/10 – Styles + Panel/Modal UI ====== */

  function ensureStyles(){
    if (document.getElementById(NS+'style')) return;
    const style=document.createElement('style'); style.id=NS+'style';
    style.textContent = `
    .${NS}panel{position:fixed;top:72px;left:50%;transform:translateX(-50%);width:min(1100px,95vw);max-height:78vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100000;display:none}
    .${NS}header{display:grid;grid-template-columns:1fr;gap:10px;align-items:start;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
    .${NS}toolbar{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:8px}
    .${NS}group{display:flex;flex-wrap:wrap;align-items:center;gap:8px;background:#f9fafb;border:1px solid rgba(0,0,0,.08);border-radius:12px;padding:6px 8px}
    .${NS}label{display:flex;align-items:center;gap:6px;font:600 12px system-ui}
    .${NS}select{padding:4px 8px;border-radius:8px;border:1px solid rgba(0,0,0,.15);background:#fff}
    .${NS}kpis{display:flex;flex-wrap:wrap;gap:8px;margin-top:4px}
    .${NS}kpi{flex:1 1 auto;background:#f5f5f5;border:1px solid rgba(0,0,0,.08);padding:6px 10px;border-radius:999px;font:600 12px system-ui;white-space:nowrap}
    .${NS}list{list-style:none;margin:0;padding:0}
    .${NS}empty{padding:14px 12px;opacity:.75;text-align:center;font:500 12px system-ui}
    .${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer}
    .${NS}chip{display:inline-flex;gap:6px;align-items:center;border:1px solid rgba(0,0,0,.12);background:#fff;padding:4px 8px;border-radius:999px;font:600 12px system-ui}
    .${NS}dot{width:8px;height:8px;border-radius:50%;background:#16a34a}
    .${NS}dot.off{background:#9ca3af}
    .${NS}loading{display:none;padding:8px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:600 12px system-ui;background:#fffbe6}
    .${NS}loading.on{display:block}
    .${NS}modal{position:fixed;inset:0;display:none;align-items:flex-start;justify-content:center;background:rgba(0,0,0,.35);z-index:100001}
    .${NS}modal-inner{background:#fff;width:min(1600px,96vw);height:min(88vh,1000px);overflow:auto;border-radius:12px;box-shadow:0 12px 28px rgba(0,0,0,.2);border:1px solid rgba(0,0,0,.12)}
    .${NS}modal-head{display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui;position:sticky;top:0;background:#fff;z-index:2}
    .${NS}modal-body{padding:8px 12px;max-height:calc(100% - 46px);overflow:auto}
    .${NS}tbl{width:100%;border-collapse:collapse;font:12px system-ui}
    .${NS}tbl th, .${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .${NS}tbl th{text-align:left;background:#fafafa;position:sticky;top:0;cursor:pointer;user-select:none;z-index:1}
    .${NS}sort-asc::after{content:" ▲";font-size:11px}
    .${NS}sort-desc::after{content:" ▼";font-size:11px}
    .${NS}eye{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border:1px solid rgba(0,0,0,.2);border-radius:6px;background:#fff;margin-right:6px;cursor:pointer;font-size:12px;line-height:1}
    .${NS}eye:hover{background:#f3f4f6}
    .${NS}badge{display:inline-block;padding:2px 6px;border-radius:999px;font-size:11px;border:1px solid rgba(0,0,0,.15);background:#f3f4f6}
    .${NS}badge-status-ok{background:#16a34a;color:#fff;border-color:#15803d}
    .${NS}badge-status-problem{background:#dc2626;color:#fff;border-color:#b91c1c}
    .${NS}badge-status-run{background:#eab308;color:#111827;border-color:#ca8a04}
    `;
    document.head.appendChild(style);
  }

  function mountUI(){
    ensureStyles();
    if (document.getElementById(NS+'panel')) return;

    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel';
    panel.innerHTML = `
      <div class="${NS}header">
        <div class="${NS}toolbar">
          <div class="${NS}group">
            <span class="${NS}label">Kommentare:</span>
            <select id="${NS}filter-comment" class="${NS}select">
              <option value="all">Alle</option>
              <option value="with">nur mit</option>
              <option value="without">nur ohne</option>
            </select>
            <span class="${NS}label">Express:</span>
            <select id="${NS}filter-express" class="${NS}select">
              <option value="all">alle</option>
              <option value="18">nur 18er</option>
              <option value="12">nur 12er</option>
            </select>
          </div>
          <div class="${NS}group">
            <button class="${NS}btn-sm" data-action="openSettings">Einstellungen</button>
            <span class="${NS}chip" id="${NS}auto-chip"><span class="${NS}dot" id="${NS}auto-dot"></span>Auto 60s</span>
            <button class="${NS}btn-sm" data-action="refreshApi">Aktualisieren (API)</button>
            <button class="${NS}btn-sm" data-action="showExpLate11">EXPRESS12 >11:01</button>
          </div>
        </div>
        <div class="${NS}kpis">
          <span class="${NS}kpi" id="${NS}chip-prio-all"  data-action="showPrioAll">PRIO in Ausrollung: <b id="${NS}kpi-prio-all">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-prio-open" data-action="showPrioOpen">PRIO noch nicht zugestellt: <b id="${NS}kpi-prio-open">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-exp-all"  data-action="showExpAll">EXPRESS in Ausrollung: <b id="${NS}kpi-exp-all">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-exp-open" data-action="showExpOpen">EXPRESS noch nicht zugestellt: <b id="${NS}kpi-exp-open">0</b></span>
        </div>
      </div>
      <div id="${NS}loading" class="${NS}loading">Lade …</div>
      <ul id="${NS}list" class="${NS}list"></ul>
      <div id="${NS}note-capture" style="padding:8px 12px;opacity:.7">Hinweis: einmal die normale Pickup-Liste laden – der letzte Request wird geklont.</div>
    `;
    document.body.appendChild(panel);

    const modal=document.createElement('div'); modal.id=NS+'modal'; modal.className=NS+'modal';
    modal.innerHTML = `
      <div class="${NS}modal-inner">
        <div class="${NS}modal-head">
          <div id="${NS}modal-title">Liste</div>
          <button class="${NS}btn-sm" data-action="closeModal">Schließen</button>
        </div>
        <div class="${NS}modal-body" id="${NS}modal-body"></div>
      </div>`;
    document.body.appendChild(modal);

    panel.addEventListener('click', async (e)=>{
      const k1=e.target.closest('#'+NS+'chip-prio-all');  if(k1){showPrioAll(); return;}
      const k2=e.target.closest('#'+NS+'chip-prio-open'); if(k2){showPrioOpen(); return;}
      const k3=e.target.closest('#'+NS+'chip-exp-all');   if(k3){showExpAll(); return;}
      const k4=e.target.closest('#'+NS+'chip-exp-open');  if(k4){showExpOpen(); return;}

      const b=e.target.closest('.'+NS+'btn-sm'); if(!b) return;
      const a=b.dataset.action;
      if(a==='openSettings'){ openSettingsModal(); return; }
      if(a==='refreshApi'){ await fullRefresh().catch(console.error); return; }
      if(a==='showExpLate11'){ showExpLate11(); return; }
    });

    const expSel = document.getElementById(NS+'filter-express');
    if (expSel){
      expSel.addEventListener('change', ()=>{
        state.filterExpress = expSel.value || 'all';
        if (document.getElementById(NS+'modal')?.style.display === 'flex') {
          if (/noch nicht zugestellt/i.test(state._modal.title)) showExpOpen();
          else if (/falsch einsortiert/i.test(state._modal.title)) showExpLate11();
          else showExpAll();
        }
        updateKpisForCurrentState();
      });
    }

    modal.addEventListener('click', e=>{
      if (e.target.dataset.action === 'closeModal' || e.target === modal) { hideModal(); return; }
      const eye=e.target.closest('button.'+NS+'eye[data-psn]'); if(eye){ openScanserver(String(eye.dataset.psn||'')); return; }
      const btn=e.target.closest('button[data-action]'); if(!btn) return;
      const a=btn.dataset.action;
      if(a==='guessDepot'){ guessDepotFromVehicles(); return; }
      if(a==='saveSettings'){ saveSettingsFromModal(); return; }
    });

    const autoDot  = document.getElementById(NS+'auto-dot');
    const autoChip = document.getElementById(NS+'auto-chip');
    function setAutoUI(){ autoDot.classList.toggle('off', !autoEnabled); }
    autoChip.addEventListener('click', ()=>{ autoEnabled=!autoEnabled; setAutoUI(); scheduleAuto(); });
    setAutoUI();

    if (!state._bootShown){
      addEvent({
        title:'Bereit',
        meta:'Status & Servicecode direkt aus API • Fahrer aus Fahrzeugübersicht • sortierbare Spalten • Predict-Zeitfenster • EXPRESS12 >11:01',
        sev:'info', read:true
      });
      state._bootShown=true;
    }
    render();
  }

  function togglePanel(force){
    const panel=document.getElementById(NS+'panel'); if(!panel){ mountUI(); return; }
    const isHidden=getComputedStyle(panel).display==='none';
    const show = force!=null ? !!force : isHidden;
    panel.style.setProperty('display', show?'block':'none', 'important');
  }

  /* ====== TEIL 3/10 – Einstellungen (Scanserver) ====== */

  const LSKEY = 'pmSettings';
  function loadSettings(){ try { return Object.assign({ scanserverPass:'', depotSuffix:'' }, JSON.parse(localStorage.getItem(LSKEY)||'{}')); } catch { return { scanserverPass:'', depotSuffix:'' }; } }
  function saveSettingsObj(s){ try { localStorage.setItem(LSKEY, JSON.stringify(s)); } catch{} }
  function setSetting(k,v){ const s=loadSettings(); s[k]=v; saveSettingsObj(s); }
  function getSetting(k){ return loadSettings()[k]; }

  function openSettingsModal(){
    const s = loadSettings();
    const html = `
      <div style="display:grid;gap:10px;max-width:520px">
        <label style="display:grid;gap:6px;font:600 12px system-ui">
          Scanserver-Passwort
          <input id="${NS}inp-pass" type="password" placeholder="••••••••" value="${esc(s.scanserverPass||'')}"
                 style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px"/>
        </label>
        <label style="display:grid;gap:6px;font:600 12px system-ui">
          Depotkennung (3-stellig, z. B. 157)
          <div style="display:flex;gap:8px;align-items:center">
            <input id="${NS}inp-depot" type="text" pattern="\\d{3}" maxlength="3" placeholder="157" value="${esc(String(s.depotSuffix||'').slice(-3))}"
                   style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px;width:100px;text-align:center;font-weight:700;letter-spacing:.5px"/>
            <button class="${NS}btn-sm" data-action="guessDepot">Auto erkennen</button>
          </div>
          <div style="opacity:.7;font:12px system-ui">Host: <code>scanserver-d0010<strong>${esc(String(s.depotSuffix||'').slice(-3)||'157')}</strong>.ssw.dpdit.de</code></div>
        </label>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button class="${NS}btn-sm" data-action="saveSettings">Speichern</button>
        </div>
      </div>`;
    openModal('Einstellungen', html);
  }

  function findVehicleGridContainer(){
    const cands = Array.from(document.querySelectorAll(
      '.MuiDataGrid-virtualScroller, .MuiDataGrid-main, [data-rttable="true"], [role="grid"], table'
    )).filter(el => el.offsetParent !== null);
    if (!cands.length) return null;
    const scrollable = cands.filter(el=>{
      const cs=getComputedStyle(el); const ov=cs.overflowY||cs.overflow;
      return /auto|scroll/i.test(ov||'') && el.scrollHeight>el.clientHeight;
    });
    const pool=scrollable.length?scrollable:cands;
    return pool.reduce((best,el)=>((el.scrollHeight||0) > ((best?.scrollHeight)||0) ? el : best), null);
  }

  function guessDepotSuffixFromVehicleTable(root){
    const grid = root || findVehicleGridContainer(); if (!grid) return '';
    const counts=new Map();
    const pick = (t) => {
      t = String(t||'');
      let m = t.match(/d0010(\d{3})/i) || t.match(/0010(\d{3})/) || t.match(/010(\d{3})/);
      if (m) return m[1];
      m = t.match(/\b(\d{3})\b/); if (m) return m[1];
      m = t.match(/0{0,2}(\d{3})\b/); if (m) return m[1];
      return '';
    };
    const ths = Array.from(grid.querySelectorAll('thead th,[role="columnheader"]'));
    const iDepot = ths.findIndex(th => /^Depot$/i.test((th.textContent||th.title||'').trim()));
    if (iDepot < 0) return '';
    const rows = Array.from(grid.querySelectorAll('tbody tr,[role="row"]'))
      .filter(r => r.querySelector('td,[role="gridcell"]'))
      .slice(0,800);
    for (const tr of rows){
      const tds = Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const raw = (tds[iDepot]?.getAttribute?.('aria-label') ||
                   tds[iDepot]?.getAttribute?.('data-title') ||
                   tds[iDepot]?.querySelector?.('[title]')?.getAttribute('title') ||
                   tds[iDepot]?.innerText || tds[iDepot]?.textContent || '').trim();
      if (!raw) continue;
      const suf = pick(raw);
      if (suf && suf!=='000') counts.set(suf,(counts.get(suf)||0)+1);
    }
    if (!counts.size) return '';
    let best='',bestN=-1; for(const [k,n] of counts){ if(n>bestN){best=k;bestN=n;} }
    return best;
  }

  function guessDepotFromVehicles(){
    const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
    if (g){
      const inp=document.getElementById(NS+'inp-depot'); if(inp) inp.value=g;
      addEvent({title:'Einstellungen', meta:`Depotkennung erkannt: ${g}`, sev:'info', read:true});
    } else {
      addEvent({title:'Einstellungen', meta:'Depotkennung konnte im aktiven Fahrzeug-Grid nicht ermittelt werden.', sev:'warn', read:true});
    }
  }

  function saveSettingsFromModal(){
    const pass = (document.getElementById(NS+'inp-pass')?.value || '');
    const dep  = (document.getElementById(NS+'inp-depot')?.value || '').replace(/\D+/g,'').slice(-3);
    saveSettingsObj({ ...loadSettings(), scanserverPass:pass, depotSuffix:dep });
    addEvent({title:'Einstellungen', meta:'Gespeichert.', sev:'info', read:true});
    hideModal();
  }

  function getScanserverBase(){
    let suf = String(getSetting('depotSuffix')||'').replace(/\D+/g,'').slice(-3);
    if (!suf){
      const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
      if (g){ suf=g; setSetting('depotSuffix',g); }
    }
    if (!suf) suf='157';
    return `https://scanserver-d0010${suf}.ssw.dpdit.de/cgi-bin/pa.cgi`;
  }

  function buildScanserverUrl(psnRaw){
    const pass = getSetting('scanserverPass')||'';
    if (!pass){
      addEvent({title:'Scanserver', meta:'Kein Passwort hinterlegt. Bitte in den Einstellungen setzen.', sev:'warn', read:true});
      openSettingsModal();
      return '';
    }
    let psn = String(psnRaw||'').replace(/\D+/g,'');
    if (psn.length===13) psn='0'+psn;
    const base=getScanserverBase();
    const params=new URLSearchParams();
    params.set('_url','file'); params.set('_passwd',pass);
    params.set('_disp','3'); params.set('_pivotxx','0'); params.set('_rastert','4');
    params.set('_rasteryt','0'); params.set('_rasterx','0'); params.set('_rastery','0');
    params.set('_pivot','0'); params.set('_pivotbp','0');
    params.set('_sortby','date|time'); params.set('_dca','0');
    params.set('_tabledef','psn|date|time|sa|tour|zc|sc|adr1|str|hno|plz1|city|dc|etafrom|etato');
    params.set('_arg59','dpd'); params.set('_arg0a',psn); params.set('_arg0b',psn); params.set('_arg0',psn+','+psn);
    params.set('_csv','0');
    return `${base}?${params.toString()}`;
  }
  function openScanserver(psn){ const url=buildScanserverUrl(psn); if(url) window.open(url,'_blank','noopener'); }

  /* ====== TEIL 4/10 – Network-Hook (API klonen) ====== */

  (function hookNetwork(){
    if (!window.__pm_fetch_hooked && window.fetch){
      const orig=window.fetch;
      window.fetch = async function(i, init={}){
        const res = await orig(i,init);
        try{
          const uStr = typeof i==='string' ? i : (i && i.url) || '';
          if (uStr.includes('/dispatcher/api/pickup-delivery') && res.ok){
            const u = new URL(uStr, location.origin);
            const q = u.searchParams;
            if (!q.get('parcelNumber')){
              const h={};
              const src=(init && init.headers) || (i && i.headers);
              if(src){
                if(src.forEach) src.forEach((v,k)=>h[String(k).toLowerCase()]=String(v));
                else if(Array.isArray(src)) src.forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
                else Object.entries(src).forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
              }
              if (!h['authorization']){
                const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                if(m) h['authorization']='Bearer '+decodeURIComponent(m[1]);
              }
              lastOkRequest={url:u, headers:h};
              const n=document.getElementById(NS+'note-capture'); if(n) n.style.display='none';
              scheduleAuto();
              if(autoEnabled && !isBusy && !document.hidden) fullRefresh().catch(()=>{});
            }
          }
        }catch{}
        return res;
      };
      window.__pm_fetch_hooked=true;
    }

    if (!window.__pm_xhr_hooked && window.XMLHttpRequest){
      const X=window.XMLHttpRequest;
      const open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
      X.prototype.open=function(m,u){ this.__pm_url=(typeof u==='string')?new URL(u,location.origin):null; this.__pm_headers={}; return open.apply(this,arguments); };
      X.prototype.setRequestHeader=function(k,v){ try{this.__pm_headers[String(k).toLowerCase()]=String(v);}catch{} return setH.apply(this,arguments); };
      X.prototype.send=function(){
        const onload = ()=>{
          try{
            if(this.__pm_url && this.__pm_url.href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300){
              const q=this.__pm_url.searchParams;
              if(!q.get('parcelNumber')){
                if(!this.__pm_headers['authorization']){
                  const m=document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                  if(m) this.__pm_headers['authorization']='Bearer '+decodeURIComponent(m[1]);
                }
                lastOkRequest={url:this.__pm_url, headers:this.__pm_headers};
                const n=document.getElementById(NS+'note-capture'); if(n) n.style.display='none';
                scheduleAuto();
                if(autoEnabled && !isBusy && !document.hidden) fullRefresh().catch(()=>{});
              }
            }
          }catch{}
          this.removeEventListener('load',onload);
        };
        this.addEventListener('load',onload);
        return send.apply(this,arguments);
      };
      window.__pm_xhr_hooked=true;
    }
  })();

  /* ====== TEIL 5/10 – Fahrer-Map aus Fahrzeugübersicht ====== */

  const gridIndex = {
    tour2driver: new Map()
  };

  function detectTourDriverCols(tbl){
    const ths = Array.from(tbl.querySelectorAll('thead th,[role="columnheader"]'))
      .map(el=>({el, txt:norm(el.textContent||el.title||'')}));
    let iTour=-1,iDrv=-1;
    ths.forEach((h,i)=>{
      if (iTour<0 && /\bTour(\s*nr|nummer)?\b/i.test(h.txt)) iTour=i;
      if (iDrv<0 && /(Zusteller(\s*name)?|Fahrer)/i.test(h.txt))   iDrv=i;
    });
    return { iTour, iDrv };
  }

  function collectTourDriverFromTable(tbl,map){
    const {iTour,iDrv}=detectTourDriverCols(tbl);
    if(iTour<0 || iDrv<0) return 0;
    const rows=Array.from(tbl.querySelectorAll('tbody tr,[role="row"]'))
      .filter(r=>r.querySelector('td,[role="gridcell"]'));
    let added=0;
    for(const tr of rows){
      const tds=Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const get = i => norm(
        tds[i]?.getAttribute?.('aria-label') ||
        tds[i]?.getAttribute?.('data-title') ||
        tds[i]?.querySelector?.('[title]')?.getAttribute('title') ||
        tds[i]?.innerText || tds[i]?.textContent || ''
      );
      const tour = get(iTour).replace(/[^\dA-Za-z]/g,'');
      const drv  = get(iDrv);
      if(tour && drv && !map.has(tour)){ map.set(tour,drv); added++; }
    }
    return added;
  }

  async function buildTourDriverMap(){
    try{
      const map=new Map();
      Array.from(document.querySelectorAll('table,[role="grid"]'))
        .filter(el=>el.offsetParent!==null)
        .forEach(tbl=>collectTourDriverFromTable(tbl,map));
      if(map.size){
        gridIndex.tour2driver = map;
        window.__pmTour2Driver = map;
      }
    }catch(e){ console.warn('[PM] Fahrer-Map Fehler',e); }
  }

  /* ====== TEIL 6/10 – API-Helfer, Paging ====== */

  function buildHeaders(h){
    const H=new Headers();
    try{
      if(h){
        Object.entries(h).forEach(([k,v])=>{
          const key=k.toLowerCase();
          if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)){
            H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()), v);
          }
        });
      }
      if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
    }catch{}
    return H;
  }

  function buildUrlPrio(base,page){
    const u=new URL(base.href); const q=u.searchParams;
    q.set('page',String(page)); q.set('pageSize','500'); q.set('priority','prio');
    q.delete('elements'); q.delete('parcelNumber'); u.search=q.toString(); return u;
  }
  function buildUrlElements(base,page,el){
    const u=new URL(base.href); const q=u.searchParams;
    q.set('page',String(page)); q.set('pageSize','500'); q.set('elements',String(el));
    q.delete('priority'); q.delete('parcelNumber'); u.search=q.toString(); return u;
  }

  function pickArray(payload){
    if(Array.isArray(payload)) return payload;
    if(payload && Array.isArray(payload.items)) return payload.items;
    if(payload && Array.isArray(payload.content)) return payload.content;
    if(payload && payload.data){
      if(Array.isArray(payload.data)) return payload.data;
      if(Array.isArray(payload.data.items)) return payload.data.items;
      if(Array.isArray(payload.data.content)) return payload.data.content;
    }
    if(payload && Array.isArray(payload.results)) return payload.results;
    if(payload && payload._embedded){
      const v=Object.values(payload._embedded).find(Array.isArray);
      if(Array.isArray(v)) return v;
    }
    return [];
  }
  function pickTotals(payload,sizeFallback){
    const p=payload||{};
    const pg=p.page||{};
    const totalElements = Number(p.totalElements ?? p.total ?? p.count ?? pg.totalElements ?? pg.total ?? 0);
    const totalPages    = Number(p.totalPages ?? pg.totalPages ?? (totalElements ? Math.ceil(totalElements/(sizeFallback||500)) : 0));
    return {
      totalElements: Number.isFinite(totalElements)?totalElements:0,
      totalPages:    Number.isFinite(totalPages)?totalPages:0
    };
  }

  async function fetchPagedFast(builder,{concurrency=6,size=500,hardMaxPages=200}={}) {
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers=buildHeaders(lastOkRequest.headers);

    const u1=builder(lastOkRequest.url,1);
    const r1=await fetch(u1.toString(),{credentials:'include',headers});
    if(!r1.ok) throw new Error(`HTTP ${r1.status}`);
    const j1=await r1.json();
    const chunk1=pickArray(j1);

    const {totalPages:tpRaw}=pickTotals(j1,size);
    let totalPages=tpRaw||0;

    if(totalPages>1){
      totalPages=Math.min(totalPages,hardMaxPages);
      const pages=[];
      for(let p=2;p<=totalPages;p++) pages.push(p);

      const out=chunk1.slice();
      let idx=0;
      async function worker(){
        while(idx<pages.length){
          const p=pages[idx++];
          const u=builder(lastOkRequest.url,p);
          const res=await fetch(u.toString(),{credentials:'include',headers});
          if(!res.ok) continue;
          const j=await res.json();
          const arr=pickArray(j);
          if(arr.length) out.push(...arr);
          if(arr.length<size){ idx=pages.length; break; }
        }
      }
      const workers=Array.from({length:Math.max(1,Math.min(concurrency,pages.length))},worker);
      await Promise.all(workers);
      return out;
    }

    if(chunk1.length<size) return chunk1;

    const out=chunk1.slice();
    let page=2;
    while(page<=hardMaxPages){
      const u=builder(lastOkRequest.url,page);
      const r=await fetch(u.toString(),{credentials:'include',headers});
      if(!r.ok) break;
      const j=await r.json();
      const arr=pickArray(j);
      if(!arr.length) break;
      out.push(...arr);
      if(arr.length<size) break;
      page++;
    }
    return out;
  }

  async function fetchPaged(builder){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers=buildHeaders(lastOkRequest.headers);
    const size=500,maxPages=50;
    let page=1,rows=[];
    while(page<=maxPages){
      const u=builder(lastOkRequest.url,page);
      const r=await fetch(u.toString(),{credentials:'include',headers});
      if(!r.ok) break;
      const j=await r.json();
      const chunk=pickArray(j);
      rows.push(...chunk);
      if(chunk.length<size) break;
      page++; await sleep(40);
    }
    return rows;
  }

  async function fetchPagedWithTotal(builder){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers=buildHeaders(lastOkRequest.headers);
    const size=500,maxPages=50;
    let page=1,rows=[],totalKnown=null;
    while(page<=maxPages){
      const u=builder(lastOkRequest.url,page);
      const r=await fetch(u.toString(),{credentials:'include',headers});
      if(!r.ok){ if(page===1) throw new Error(`HTTP ${r.status}`); break; }
      const j=await r.json();
      const chunk=pickArray(j);
      if(page===1){
        const {totalElements,totalPages}=pickTotals(j,size);
        if(Number.isFinite(totalElements)&&totalElements>=0) totalKnown=totalElements;
        if(Number.isFinite(totalPages)&&totalPages>0) totalKnown=totalKnown ?? totalPages*size;
      }
      rows.push(...chunk);
      if(chunk.length<size) break;
      page++; await sleep(40);
    }
    return {rows, total: totalKnown || rows.length};
  }


/* ====== TEIL 7/10 – Normalisierung, Status/Fahrer/Predict/Service ====== */

  const parcelId   = r => r.__pidOverride || r.parcelNumber || (Array.isArray(r.parcelNumbers)&&r.parcelNumbers[0]) || r.id || '';
  const addrOf     = r => [r.street,r.houseno].filter(Boolean).join(' ');
  const placeOf    = r => [r.postalCode,r.city].filter(Boolean).join(' ');
  const addCodes   = r => Array.isArray(r.additionalCodes)? r.additionalCodes.map(String): [];
  const isDelivery = r => String(r?.orderType||'').toUpperCase()==='DELIVERY';
  const isPRIO     = r => String(r?.priority||r?.prio||'').toUpperCase()==='PRIO';

  const hasExpress12 = r => {
    const el=r?.elements;
    if(Array.isArray(el)) return el.map(String).includes('023');
    return typeof el==='string'?/\b023\b/.test(el):false;
  };
  const hasExpress18 = r => {
    const el=r?.elements;
    if(Array.isArray(el)) return el.map(String).includes('010');
    return typeof el==='string'?/\b010\b/.test(el):false;
  };
  const hasExpressAny = r => hasExpress12(r)||hasExpress18(r);
  const expressTypeOf = r => hasExpress12(r)?'12':(hasExpress18(r)?'18':'');

  const composeDateTime = (dateStr,timeStr)=>{
    if(!dateStr || !timeStr) return null;
    const s=`${dateStr}T${String(timeStr).slice(0,8)}`;
    const d=new Date(s); return isNaN(d)?null:d;
  };
  const fromTime = r => r.from2 ? composeDateTime(r.date,r.from2) : null;
  const toTime   = r => r.to2   ? composeDateTime(r.date,r.to2)   : null;
  const deliveredTime = r => r.deliveredTime ? new Date(r.deliveredTime) : null;

  function apiStatus(r){
    const v =
      r?.statusDisplay || r?.statusLabel || r?.statusDescription ||
      r?.statusName || r?.statusText || r?.stateText ||
      r?.deliveryStatus || r?.parcelStatus || '';
    return String(v||'').trim();
  }
  const statusOf = r => apiStatus(r);

  const delivered = r => {
    if(r.deliveredTime) return true;
    const s=(statusOf(r)||'').toUpperCase();
    return /ZUGESTELLT|DELIVERED/.test(s);
  };

  const tourOf = r => r.tour ? String(r.tour) : '';

  function driverOf(r){
    const direct = r.driverName || r.driver || r.courierName || r.riderName || r.tourDriver || '';
    if(direct && direct.trim()) return direct.trim();
    const key = String(tourOf(r)||'').replace(/[^\dA-Za-z]/g,'');
    const viaGrid =
      (gridIndex.tour2driver && gridIndex.tour2driver.get(key)) ||
      (window.__pmTour2Driver instanceof Map ? window.__pmTour2Driver.get(key) : '');
    return (viaGrid || '—').trim() || '—';
  }

  function formatHHMM(d){
    return d ? d.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}) : '';
  }

  function buildPredictMeta(r){
    const pf=fromTime(r);
    const pt=toTime(r);
    const pfTs=pf?+pf:0;
    const ptTs=pt?+pt:0;
    let range='—';
    if(pf && pt) range=`${formatHHMM(pf)} - ${formatHHMM(pt)}`;
    else if(pf)  range=formatHHMM(pf);
    else if(pt)  range=formatHHMM(pt);
    return {pfTs,ptTs,range};
  }

  // Service-Code: zuerst aus der Fahrzeugübersicht (DOM), dann API-Felder, kein Zusatzcode-Fallback
 // Service-Code: zuerst aus der Fahrzeugübersicht (DOM), dann API-Felder inkl. serviceCodes
function serviceOf(r){
  if(!r) return '';

  const tryVal = v =>
    (typeof v==='number' || typeof v==='string') ? String(v).trim() : '';

  const tryArrFirst = v =>
    Array.isArray(v) && v.length ? tryVal(v[0]) : '';

  // 1. Wert aus der Original-Tabelle (gridIndex.serviceByPsn) – bleibt wie gehabt
  try {
    const pidClean = String(parcelId(r) || '').replace(/\D+/g,'');
    if (pidClean && typeof gridIndex !== 'undefined' && gridIndex.serviceByPsn instanceof Map){
      const domVal =
        gridIndex.serviceByPsn.get(pidClean) ||
        gridIndex.serviceByPsn.get(pidClean.padStart(14,'0'));
      if (domVal) return String(domVal).trim();
    }
  } catch(e){
    // ignorieren, weiter zu API-Feldern
  }

  // 2. direkte API-Felder (+ neu: serviceCodes-Array)
  let v = '';
  v = tryVal(r.serviceCode) ||
      tryVal(r.servicecode) ||
      tryVal(r.service_code) ||
      tryArrFirst(r.serviceCodes);       // <--- NEU
  if (v) return v;

  // 3. verschachteltes service-Objekt
  if (r.service && typeof r.service === 'object'){
    v = tryVal(r.service.code) ||
        tryVal(r.service.serviceCode) ||
        tryVal(r.service.id) ||
        tryArrFirst(r.service.serviceCodes);    // falls dort auch Arrays kommen
    if (v) return v;
  }

  // 4. verschachteltes product-Objekt
  if (r.product && typeof r.product === 'object'){
    v = tryVal(r.product.serviceCode) ||
        tryVal(r.product.code) ||
        tryVal(r.product.id) ||
        tryArrFirst(r.product.serviceCodes);    // fallback, wenn dort arrays liegen
    if (v) return v;
  }

  // kein Fallback auf additionalCodes
  return '';
}

  function statusClass(text){
    const t=String(text||'').toUpperCase();
    if(/PROBLEM|FAIL|NICHT/.test(t)) return NS+'badge-status-problem';
    if(/ZUGESTELLT|DELIVERED/.test(t)) return NS+'badge-status-ok';
    if(/ZUSTELLUNG|OUT_FOR_DELIVERY|IN_DELIVERY/.test(t)) return NS+'badge-status-run';
    return '';
  }

  function normRow(r){
    const pid=parcelId(r) || '';
    const {pfTs,ptTs,range}=buildPredictMeta(r);
    return {
      ...r,
      __pid: pid,
      __addr: [addrOf(r), placeOf(r)].filter(Boolean).join(' · ') || '—',
      __driver: driverOf(r),
      __tourNum: Number(tourOf(r) || 0),
      __status: statusOf(r) || '',
      __delivTs: deliveredTime(r) ? deliveredTime(r).getTime() : 0,
      __predFromTs: pfTs,
      __predToTs:   ptTs,
      __predRangeStr: range,
      __codesStr: (addCodes(r)||[]).join(', ') || '—',
      __expType: expressTypeOf(r),
      __serviceCode: serviceOf(r)
    };
  }

  function expandAndNorm(rows){
    const out=[];
    for(const r of rows){
      const list = Array.isArray(r.parcelNumbers) && r.parcelNumbers.length
        ? r.parcelNumbers
        : [ parcelId(r) ];
      const seen=new Set();
      for(const raw of list){
        const psn = String(raw||'').replace(/\D+/g,'');
        if(!psn || seen.has(psn)) continue;
        seen.add(psn);
        const rr={...r, __pidOverride: psn.length===13 ? '0'+psn : psn};
        out.push(normRow(rr));
      }
    }
    return out;
  }



  /* ====== TEIL 8/10 – Tabellen-UI, Sortierung, Modal (inkl. Predict immer) ====== */

  function buildHeaderHtml(){
    const ths=['Paketscheinnummer','Adresse','Fahrer','Tour','Status','Zustellzeit','Zusatzcode','Servicecode','Predict'];
    return `<tr>${ths.map((h,i)=>`<th data-col="${i}">${h}</th>`).join('')}</tr>`;
  }

  function buildTableShell(){
    return `
      <div id="${NS}vt-wrap" style="position:relative;height:min(70vh,720px);overflow:auto">
        <table class="${NS}tbl">
          <thead>${buildHeaderHtml()}</thead>
          <tbody id="${NS}vt-body"></tbody>
        </table>
      </div>`;
  }

  function rowHtml(r){
    const pLink = r.__pid
      ? `<a class="${NS}plink" href=https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${r.__pid} target="_blank" rel="noopener">${r.__pid}</a>`
      : '—';
    const eye = r.__pid ? `<button class="${NS}eye" title="Scanserver öffnen" data-psn="${esc(r.__pid)}">👁</button>` : '';
    const dtime = r.__delivTs ? formatHHMM(new Date(r.__delivTs)) : '—';

    const statusText = r.__status || '';
    const statusCls  = statusClass(statusText);
    const statusCell = statusText
      ? `<span class="${NS}badge ${statusCls}">${esc(statusText)}</span>`
      : '';

    const serviceText = r.__serviceCode || '';
    const serviceCell = serviceText
      ? `<span class="${NS}badge">${esc(serviceText)}</span>`
      : '';

    const pred = r.__predRangeStr || '—';

    const cells = [
      `${eye}${pLink}`,
      esc(r.__addr),
      esc(r.__driver),
      esc(String(r.__tourNum||'—')),
      statusCell,
      esc(dtime),
      esc(r.__codesStr),
      serviceCell,
      esc(pred)
    ];

    return `<tr>${cells.map(v=>`<td>${v}</td>`).join('')}</tr>`;
  }

  function openModal(title,rowsOrHtml){
    const m=document.getElementById(NS+'modal');
    const t=document.getElementById(NS+'modal-title');
    const b=document.getElementById(NS+'modal-body');
    if(t) t.textContent=title||'';

    if(Array.isArray(rowsOrHtml)){
      const rows=rowsOrHtml.slice();
      state._modal={rows,opts:{showPredict:true},title:title||''};

      if(b) b.innerHTML=buildTableShell();
      const tbody=document.getElementById(NS+'vt-body');
      const wrap =document.getElementById(NS+'vt-wrap');

      function renderAll(){
        if(!tbody) return;
        tbody.innerHTML = rows.map(r=>rowHtml(r)).join('');
      }

      const thead = wrap.querySelector('thead');
      if(thead){
        thead.addEventListener('click',ev=>{
          const th=ev.target.closest('th'); if(!th) return;
          const col=Number(th.dataset.col||0);
          Array.from(thead.querySelectorAll('th')).forEach(x=>x.classList.remove(NS+'sort-asc',NS+'sort-desc'));
          const asc=!(th.dataset.dir==='asc'); th.dataset.dir=asc?'asc':'desc';
          th.classList.add(asc?NS+'sort-asc':NS+'sort-desc');

          const getKey=r=>{
            switch(col){
              case 0: return r.__pid;
              case 1: return r.__addr;
              case 2: return r.__driver;
              case 3: return r.__tourNum;
              case 4: return r.__status || '';
              case 5: return r.__delivTs;
              case 6: return r.__codesStr;
              case 7: return r.__serviceCode || '';
              case 8: return r.__predFromTs || 0;
              default: return '';
            }
          };
          rows.sort((a,b)=>{
            const A=getKey(a),B=getKey(b);
            if(typeof A==='number' && typeof B==='number') return asc?(A-B):(B-A);
            return asc ? collator.compare(String(A),String(B)) : collator.compare(String(B),String(A));
          });
          state._modal.rows=rows;
          renderAll();
        });
      }

      renderAll();
    } else {
      if(b) b.innerHTML=rowsOrHtml||'';
    }

    if(m) m.style.display='flex';
  }

  function hideModal(){ const m=document.getElementById(NS+'modal'); if(m) m.style.display='none'; }

  /* ====== TEIL 9/10 – KPIs, Filter, Refresh-Logik ====== */

  function setLoading(on){ isLoading=!!on; const el=document.getElementById(NS+'loading'); if(el) el.classList.toggle('on',on); }
  function dimButtons(on){ document.querySelectorAll('.'+NS+'btn-sm').forEach(b=>b.classList.toggle(NS+'dim',!!on)); }
  function render(){
    const list=document.getElementById(NS+'list'); if(!list) return;
    list.innerHTML='';
    const d=document.createElement('div'); d.className=NS+'empty';
    d.textContent=isLoading?'Lade Daten …':'Aktualisiert.';
    list.appendChild(d);
  }

  function setKpis({prioAll,prioOpen,expAll,expOpen}){
    const set=(id,val,chipId)=>{
      const el=document.getElementById(id); const chip=document.getElementById(chipId);
      if(el) el.textContent=String(Number(val||0));
      if(chip) chip.style.display=Number(val||0)===0?'none':'';
    };
    set(NS+'kpi-prio-all',  prioAll, NS+'chip-prio-all');
    set(NS+'kpi-prio-open', prioOpen,NS+'chip-prio-open');
    set(NS+'kpi-exp-all',   expAll,  NS+'chip-exp-all');
    set(NS+'kpi-exp-open',  expOpen, NS+'chip-exp-open');
  }

  function addEvent(ev){
    const e={
      id:state.nextId++,
      title:ev.title||'Ereignis',
      meta:ev.meta||'',
      sev:ev.sev||'info',
      ts:ev.ts||Date.now(),
      read:!!ev.read,
      comment:ev.comment||'',
      hasComment:!!(ev.comment&&ev.comment.trim()),
      parcel:ev.parcel||'',
      kind:ev.kind||''
    };
    state.events.push(e);
    render();
  }

  function buildTableRowsAndCounts(prioRows,exp12Rows,exp18Rows){
    const prioDeliveries = prioRows.filter(isDelivery).filter(isPRIO);
    const prioAll  = prioDeliveries;
    const prioOpen = prioDeliveries.filter(r=>!delivered(r));

    const expRows=[...exp12Rows,...exp18Rows];
    const seen=new Set();
    const expDeliveries=expRows
      .filter(isDelivery)
      .filter(hasExpressAny)
      .filter(r=>{
        const id=parcelId(r);
        if(!id || seen.has(id)) return false;
        seen.add(id); return true;
      });
    const expAll  = expDeliveries;
    const expOpen = expDeliveries.filter(r=>!delivered(r));
    const expLate11 = expDeliveries.filter(hasExpress12).filter(r=>{
      const ft=fromTime(r); if(!ft) return false;
      return (ft.getHours()>11) || (ft.getHours()===11 && ft.getMinutes()>=1);
    });
    return {prioAll,prioOpen,expAll,expOpen,expLate11};
  }

  function filterByExpressSelection(rows){
    if(state.filterExpress==='12') return rows.filter(hasExpress12);
    if(state.filterExpress==='18') return rows.filter(hasExpress18);
    return rows;
  }

  function getFilteredExpressCounts(){
    const f=state.filterExpress;
    const filt = f==='12' ? hasExpress12 : f==='18' ? hasExpress18 : null;
    const expAllList  = filt ? state._expAllList.filter(filt)  : state._expAllList;
    const expOpenList = filt ? state._expOpenList.filter(filt) : state._expOpenList;
    return {expAllCount:expAllList.length, expOpenCount:expOpenList.length};
  }

  function updateKpisForCurrentState(){
    const {expAllCount,expOpenCount}=getFilteredExpressCounts();
    setKpis({
      prioAll: state._prioAllList.length,
      prioOpen:state._prioOpenList.length,
      expAll:expAllCount,
      expOpen:expOpenCount
    });
  }

  function showPrioAll(){
    const rows=state._prioAllList;
    openModal(`PRIO – in Ausrollung (alle) · ${rows.length}`,rows);
  }
  function showPrioOpen(){
    const rows=state._prioOpenList;
    openModal(`PRIO – noch nicht zugestellt · ${rows.length}`,rows);
  }
  function showExpAll(){
    const src=state._expAllList;
    const rows=filterByExpressSelection(src);
    const sel=state.filterExpress==='12'?' (12)':state.filterExpress==='18'?' (18)':'';
    openModal(`Express${sel} – in Ausrollung (alle) · ${rows.length}`,rows);
  }
  function showExpOpen(){
    const src=state._expOpenList;
    const rows=filterByExpressSelection(src);
    const sel=state.filterExpress==='12'?' (12)':state.filterExpress==='18'?' (18)':'';
    openModal(`Express${sel} – noch nicht zugestellt · ${rows.length}`,rows);
  }
  function showExpLate11(){
    const rows=state._expLate11List.slice();
    openModal(`Express 12 – falsch einsortiert (>11:01 geplant) · ${rows.length}`,rows);
  }

  async function refreshViaApi_FAST(){
    const {prioRows,exp12Rows,exp18Rows}=await (async()=>{
      const [p,e12,e18]=await Promise.all([
        fetchPagedFast(buildUrlPrio),
        fetchPagedFast((b,p)=>buildUrlElements(b,p,'023')),
        fetchPagedFast((b,p)=>buildUrlElements(b,p,'010'))
      ]);
      return {prioRows:p, exp12Rows:e12, exp18Rows:e18};
    })();

    const prioN  = expandAndNorm(prioRows);
    const exp12N = expandAndNorm(exp12Rows);
    const exp18N = expandAndNorm(exp18Rows);

    const {prioAll,prioOpen,expAll,expOpen,expLate11}=buildTableRowsAndCounts(prioN,exp12N,exp18N);

    state._prioAllList  = prioAll.slice();
    state._prioOpenList = prioOpen.slice();
    state._expAllList   = expAll.slice();
    state._expOpenList  = expOpen.slice();
    state._expLate11List= expLate11.slice();

    const {expAllCount,expOpenCount}=getFilteredExpressCounts();
    setKpis({prioAll:prioAll.length, prioOpen:prioOpen.length, expAll:expAllCount, expOpen:expOpenCount});

    state.events=[{
      id:++state.nextId,
      title:'Aktualisiert (FAST)',
      meta:`PRIO: in Ausrollung ${prioAll.length} • offen ${prioOpen.length} • EXPRESS: in Ausrollung ${expAll.length} • offen ${expOpenCount} • „>11:01“ (12er): ${expLate11.length}`,
      sev:'info',read:true,ts:Date.now()
    }];
  }

  async function refreshViaApi_Legacy(){
    const [prioRes,exp12Rows,exp18Rows]=await Promise.all([
      fetchPagedWithTotal(buildUrlPrio),
      fetchPaged((b,p)=>buildUrlElements(b,p,'023')),
      fetchPaged((b,p)=>buildUrlElements(b,p,'010'))
    ]);
    const prioN  = expandAndNorm(prioRes.rows);
    const exp12N = expandAndNorm(exp12Rows);
    const exp18N = expandAndNorm(exp18Rows);

    const {prioAll,prioOpen,expAll,expOpen,expLate11}=buildTableRowsAndCounts(prioN,exp12N,exp18N);

    state._prioAllList  = prioAll.slice();
    state._prioOpenList = prioOpen.slice();
    state._expAllList   = expAll.slice();
    state._expOpenList  = expOpen.slice();
    state._expLate11List= expLate11.slice();

    const {expAllCount,expOpenCount}=getFilteredExpressCounts();
    setKpis({prioAll:prioAll.length, prioOpen:prioOpen.length, expAll:expAllCount, expOpen:expOpenCount});

    state.events=[{
      id:++state.nextId,
      title:'Aktualisiert',
      meta:`PRIO: in Ausrollung ${prioAll.length} • offen ${prioOpen.length} • EXPRESS: in Ausrollung ${expAll.length} • offen ${expOpenCount} • „>11:01“ (12er): ${expLate11.length}`,
      sev:'info',read:true,ts:Date.now()
    }];
  }

  async function refreshViaApi_SAFE(){
    try{
      await refreshViaApi_FAST();
      const nothing=!state._prioAllList.length && !state._expAllList.length &&
                    !state._prioOpenList.length && !state._expOpenList.length;
      if(nothing) await refreshViaApi_Legacy();
    }catch(e){
      await refreshViaApi_Legacy();
    }
  }

  async function fullRefresh(){
    if(isBusy) return;
    try{
      if(!lastOkRequest){
        addEvent({
          title:'Hinweis',
          meta:'Kein API-Request erkannt. Bitte einmal die normale Pickup-Suche ausführen, danach funktioniert Auto 60s.',
          sev:'info',read:true
        });
        render(); return;
      }
      isBusy=true; setLoading(true); dimButtons(true);
      addEvent({
        title:'Aktualisiere (API)…',
        meta:'FAST-Paging • Fahrer aus Fahrzeugübersicht • Status & Servicecode direkt aus API',
        sev:'info',read:true
      });
      render();

      await buildTourDriverMap();
      await refreshViaApi_SAFE();
    }catch(e){
      console.error(e);
      addEvent({title:'Fehler (API)', meta:String(e && e.message || e), sev:'warn'});
    }finally{
      setLoading(false); dimButtons(false); isBusy=false; render();
    }
  }

  function scheduleAuto(){
    try{
      if(autoTimer){ clearInterval(autoTimer); autoTimer=null; }
      if(!autoEnabled) return;
      if(document.hidden) return;
      if(!lastOkRequest) return;
      autoTimer=setInterval(()=>{
        if(!lastOkRequest) return;
        fullRefresh().catch(()=>{});
      },60_000);
    }catch{}
  }
  document.addEventListener('visibilitychange',()=>scheduleAuto());

  /* ====== TEIL 10/10 – Boot ====== */

  function boot(){
    mountUI();
    scheduleAuto();
    setTimeout(()=>{
      if(!(getSetting('depotSuffix')||'')){
        const g=guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
        if(g){
          setSetting('depotSuffix',g);
          addEvent({title:'Einstellungen', meta:`Depotkennung aus Fahrzeugübersicht: ${g}`, sev:'info',read:true});
        }
      }
    },1200);
  }

})();
