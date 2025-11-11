// ==UserScript==
// @name         DPD Dispatcher – Prio/Express12 Monitoring
// @namespace    bodo.dpd.custom
// @version      6.3.1
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @description  PRIO/EXPRESS12: KPIs & Listen. Status DE (DOM bevorzugt), sortierbare Tabellen, Zusatzcode, Predict, Zustellzeit, Button „EXPRESS12 >11:01“. Panel bleibt offen; PSN mit Auge-Button öffnet Scanserver.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  /* ----------------------- Registrierung im Loader ----------------------- */
  const moduleDef = {
    id: 'prio-express-monitor',
    label: 'Prio/Express Monitoring',
    run: () => startModuleOnce()
  };
  if (window.TM && typeof window.TM.register === 'function') {
    window.TM.register(moduleDef);
  } else {
    window.__tmQueue = window.__tmQueue || [];
    window.__tmQueue.push(moduleDef);
  }

  /* =================== ab hier: dein Modul, lazy gestartet =================== */
  let started = false;
  function startModuleOnce() {
    if (started) { togglePanel(true); return; }
    started = true;
    boot();
    togglePanel(true);
  }

  /* ------------------- (Originalcode – minimal angepasst) ------------------- */

  const NS = 'pm-';
  const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
  const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const fmt = ts => { try { return new Date(ts||Date.now()).toLocaleString('de-DE'); } catch { return String(ts||''); } };

  const state={
    filterExpress: 'all',
    events:[], nextId:1, _bootShown:false,
    _prioAllList:[], _prioOpenList:[],
    _expAllList:[],  _expOpenList:[],
    _expLate11List:[],
    _modal:{rows:[], opts:{}, title:''}
  };
  let lastOkRequest=null;
  let autoEnabled = true, autoTimer = null, isBusy = false, isLoading = false;
  const commentCache = new Map(), statusCache = new Map();

  /* ---------- Styles ---------- */
  function ensureStyles(){
    if (document.getElementById(NS+'style')) return;
    const style=document.createElement('style'); style.id=NS+'style';
    style.textContent=`
    .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:8px 14px;border-radius:999px;font:600 13px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
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
    .${NS}modal{position:fixed;inset:0;display:none;align-items:center;justify-content:center;background:rgba(0,0,0,.35);z-index:100001}
    .${NS}modal-inner{background:#fff;max-width:min(1200px,95vw);max-height:78vh;overflow:auto;border-radius:12px;box-shadow:0 12px 28px rgba(0,0,0,.2);border:1px solid rgba(0,0,0,.12)}
    .${NS}modal-head{display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
    .${NS}modal-body{padding:8px 12px}
    .${NS}tbl{width:100%;border-collapse:collapse;font:12px system-ui}
    .${NS}tbl th, .${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top}
    .${NS}tbl th{text-align:left;background:#fafafa;position:sticky;top:0;cursor:pointer;user-select:none}
    .${NS}sort-asc::after{content:" ▲";font-size:11px}
    .${NS}sort-desc::after{content:" ▼";font-size:11px}
    .${NS}eye{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border:1px solid rgba(0,0,0,.2);border-radius:6px;background:#fff;margin-right:6px;cursor:pointer;font-size:12px;line-height:1}
    .${NS}eye:hover{background:#f3f4f6}
    .pm-tbl{table-layout:fixed; width:100%}
    .pm-tbl th, .pm-tbl td{white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
    .pm-tbl th{ z-index:1 }
    .pm-modal-inner{ max-width:min(1400px,95vw) }
    .pm-modal{align-items:flex-start}
    .pm-modal-inner{ width:min(1600px,96vw); height:min(88vh,1000px) }
    .pm-modal.pm-full .pm-modal-inner{ width:96vw; height:92vh }
    .pm-modal-body{padding:8px 12px; overflow:auto; max-height:calc(100% - 46px)}
    .pm-modal-head{position:sticky; top:0; background:#fff; z-index:2}
    .pm-resize{ position:absolute; right:6px; bottom:6px; width:16px; height:16px; cursor:se-resize; opacity:.6 }
    .pm-resize::after{ content:""; position:absolute; inset:0;
      background:linear-gradient(135deg, transparent 0 50%, rgba(0,0,0,.25) 50% 100%); border-radius:3px; }
    `;
    document.head.appendChild(style);
  }

  /* ---------- Settings (Passwort/Depot) ---------- */
  const LSKEY = 'pmSettings';
  function loadSettings(){ try { return Object.assign({ scanserverPass:'', depotSuffix:'' }, JSON.parse(localStorage.getItem(LSKEY)||'{}')); } catch { return { scanserverPass:'', depotSuffix:'' }; } }
  function saveSettings(s){ try { localStorage.setItem(LSKEY, JSON.stringify(s)); } catch{} }
  function setSetting(k,v){ const s = loadSettings(); s[k]=v; saveSettings(s); }
  function getSetting(k){ return loadSettings()[k]; }

  /* ---------- UI ---------- */
  function mountUI(){
    ensureStyles();
    if (document.getElementById(NS+'panel')) return;

    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel';
    panel.innerHTML=`
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
      <div id="${NS}note-capture" style="padding:8px 12px;opacity:.7">Hinweis: einmal die normale Liste laden – der letzte Request wird geklont.</div>
    `;
    document.body.appendChild(panel);

    const modal=document.createElement('div'); modal.id=NS+'modal'; modal.className=NS+'modal';
    modal.innerHTML=`<div class="${NS}modal-inner"><div class="${NS}modal-head"><div id="${NS}modal-title">Liste</div><button class="${NS}btn-sm" data-action="closeModal">Schließen</button></div><div class="${NS}modal-body" id="${NS}modal-body"></div></div>`;
    document.body.appendChild(modal);

    panel.addEventListener('click', async (e)=>{
      const k1=e.target.closest('#'+NS+'chip-prio-all');  if(k1){showPrioAll(); return;}
      const k2=e.target.closest('#'+NS+'chip-prio-open'); if(k2){showPrioOpen(); return;}
      const k3=e.target.closest('#'+NS+'chip-exp-all');   if(k3){showExpAll(); return;}
      const k4=e.target.closest('#'+NS+'chip-exp-open');  if(k4){showExpOpen(); return;}

      const b = e.target.closest('.'+NS+'btn-sm');
      if(!b) return;
      const a = b.dataset.action;
      if (a === 'openSettings') { openSettingsModal(); return; }
      if (a === 'refreshApi')   { await fullRefresh().catch(console.error); return; }
      if (a === 'showExpLate11'){ showExpLate11(); return; }
    });

    const expSel = document.getElementById(NS+'filter-express');
    if (expSel) {
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

    modal.addEventListener('click', (e)=>{
      if (e.target.dataset.action === 'closeModal' || e.target === modal) { hideModal(); return; }
      const eye = e.target.closest('button.'+NS+'eye[data-psn]'); if (eye){ openScanserver(String(eye.dataset.psn||'')); return; }
      const btn = e.target.closest('button[data-action]'); if (!btn) return;
      const a = btn.dataset.action;
      if (a === 'guessDepot') {
        const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
        if (g) {
          const inp = document.getElementById(NS+'inp-depot'); if (inp) inp.value = g;
          addEvent({title:'Einstellungen', meta:`Depotkennung erkannt: ${g}`, sev:'info', read:true});
        } else {
          addEvent({title:'Einstellungen', meta:'Depotkennung konnte im aktiven Fahrzeug-Grid nicht ermittelt werden.', sev:'warn', read:true});
        }
        return;
      }
      if (a === 'saveSettings') {
        const pass = (document.getElementById(NS+'inp-pass')?.value || '');
        const dep  = (document.getElementById(NS+'inp-depot')?.value || '').replace(/\D+/g,'').slice(-3);
        saveSettings({ ...loadSettings(), scanserverPass: pass, depotSuffix: dep });
        addEvent({title:'Einstellungen', meta:'Gespeichert.', sev:'info', read:true});
        hideModal();
        return;
      }
    });

    const autoDot = document.getElementById(NS+'auto-dot');
    const autoChip = document.getElementById(NS+'auto-chip');
    function setAutoUI(){ autoDot.classList.toggle('off', !autoEnabled); }
    autoChip.addEventListener('click', ()=>{ autoEnabled = !autoEnabled; setAutoUI(); scheduleAuto(); });
    setAutoUI();

    if (!state._bootShown) { addEvent({title:'Bereit', meta:'Status DE aus Tabelle • Sortierbare Überschriften • Zusatzcode • Predict • Zustellzeit • EXPRESS12 >11:01 • Auge→Scanserver', sev:'info', read:true}); state._bootShown=true; }
    render();
  }

  function togglePanel(force){
    const panel=document.getElementById(NS+'panel'); if(!panel) { mountUI(); return; }
    const isHidden=getComputedStyle(panel).display==='none';
    const show = force!=null ? !!force : isHidden;
    panel.style.setProperty('display', show?'block':'none', 'important');
  }

  /* ---------- Einstellungen: Modal ---------- */
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

  /* ---------- Grid/Depot-Helfer ---------- */
  const wait = (ms)=>new Promise(r=>setTimeout(r,ms));
  const norm = s => String(s||'').replace(/\s+/g,' ').trim();

  function* nodesDeep(root=document){
    const stack=[root];
    while(stack.length){
      const n=stack.pop();
      yield n;
      const kids = n instanceof ShadowRoot ? n.children : (n.querySelectorAll ? n.children : []);
      for (let i=kids.length-1;i>=0;i--) stack.push(kids[i]);
      if (n.shadowRoot) stack.push(n.shadowRoot);
      if (n.tagName==='IFRAME' && n.contentDocument) stack.push(n.contentDocument);
    }
  }
  function qsaDeep(selector, root=document){
    const out=[]; for (const n of nodesDeep(root)) if (n.querySelectorAll) out.push(...n.querySelectorAll(selector)); return out;
  }

  function findVehicleGridContainer(){
    const cands = qsaDeep(
      '.MuiDataGrid-virtualScroller, .MuiDataGrid-main, .ReactVirtualized__Grid, ' +
      '.ag-center-cols-viewport, .ag-body-viewport, [role="grid"], .rt-table, .rt-tbody, [data-rttable="true"]'
    ).filter(el => el.offsetParent !== null);
    if (!cands.length) return null;
    const scrollable = cands.filter(el => {
      const cs = getComputedStyle(el); const ov = cs.overflowY || cs.overflow;
      return (ov && /auto|scroll/i.test(ov)) && (el.scrollHeight > el.clientHeight);
    });
    const pool = scrollable.length ? scrollable : cands;
    return pool.reduce((best, el) => ((el.scrollHeight||0) > ((best?.scrollHeight)||0) ? el : best), null);
  }

  function detectColIndexes(root){
    const ths = qsaDeep('thead th,[role="columnheader"]').map(el => ({ el, txt: norm(el.textContent||el.title||'') }));
    let startIdx = -1; for (let i=ths.length-1; i>=0; i--){ if (ths[i].txt === 'Depot') { startIdx = i; break; } }
    let hdrSlice = startIdx>=0 ? ths.slice(startIdx, startIdx+9) : ths;
    const iTourFromHead = hdrSlice.findIndex(h => /\bTour(\s*nr|nummer)?\b/i.test(h.txt));
    const iDrvFromHead  = hdrSlice.findIndex(h => /^(Zusteller(\s*name)?|Fahrer)$/i.test(h.txt));
    let iTour = iTourFromHead >= 0 ? iTourFromHead : -1;
    let iDrv  = iDrvFromHead  >= 0 ? iDrvFromHead  : -1;
    if (iTour < 0 || iDrv < 0){
      const firstRow = qsaDeep('[role="row"], tbody tr', root).find(r => qsaDeep('[role="gridcell"], td', r).length);
      if (firstRow){
        const cells = qsaDeep('[role="gridcell"], td', firstRow);
        cells.forEach((c,i)=>{
          const label = norm(
            c.getAttribute('aria-label') || c.getAttribute('data-title') ||
            c.querySelector('[title]')?.getAttribute('title') || c.innerText || c.textContent || ''
          );
          if (iTour < 0 && /\bTour(\s*nr|nummer)?\b/i.test(label)) iTour=i;
          if (iDrv  < 0 && /(Zusteller(\s*name)?|Fahrer)/i.test(label))    iDrv =i;
        });
      }
    }
    return { iTour, iDrv };
  }

  function collectMapFromTable(tbl, map){
    const { iTour, iDrv } = detectColIndexes(tbl);
    if (iTour < 0 || iDrv < 0) return 0;
    const rows = qsaDeep('[role="row"], tbody tr', tbl).filter(r => qsaDeep('[role="gridcell"], td', r).length);
    let added = 0;
    for (const tr of rows){
      const tds = qsaDeep('[role="gridcell"], td', tr);
      const get = (i)=> norm(
        tds[i]?.getAttribute?.('aria-label') || tds[i]?.getAttribute?.('data-title') ||
        tds[i]?.querySelector?.('[title]')?.getAttribute('title') ||
        tds[i]?.innerText || tds[i]?.textContent || ''
      );
      const tour = get(iTour).replace(/[^\dA-Za-z]/g,'');
      const drv  = get(iDrv);
      if (tour && drv && !map.has(tour)){ map.set(tour, drv); added++; }
    }
    return added;
  }

  async function scrollGridAndCollect(map, {timeoutMs=18000, step=250}={}){
    let grid = findVehicleGridContainer(); if (!grid) return 0;
    const start = Date.now(); let totalAdded = 0, prev = -1;
    while (Date.now()-start < timeoutMs){
      const scopes = [grid, grid.parentElement, document];
      for (const scope of scopes){
        const tablesHere = qsaDeep('table,[role="grid"]', scope).filter(el => el.offsetParent !== null);
        for (const tbl of tablesHere) totalAdded += collectMapFromTable(tbl, map);
      }
      const maxY = grid.scrollHeight - grid.clientHeight;
      const next = Math.min((grid.scrollTop||0) + Math.max(80, grid.clientHeight*0.9), maxY);
      grid.scrollTop = next;
      if (next >= maxY || next === prev) break;
      prev = next; await new Promise(r=>setTimeout(r, step));
    }
    return totalAdded;
  }

  async function tryBuildTourDriverMapFromDom(){
    if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size) return;
    const visibleTables = qsaDeep('table,[role="grid"]').filter(el => el.offsetParent !== null);
    for (const tbl of visibleTables) {
      const map = new Map();
      const rows = qsaDeep('tbody tr,[role="row"]', tbl).filter(r => qsaDeep('td,[role="gridcell"]', r).length);
      const { iTour, iDrv } = detectColIndexes(tbl);
      if (iTour < 0 || iDrv < 0) continue;
      for (const tr of rows){
        const tds  = qsaDeep('td,[role="gridcell"]', tr);
        const get  = (i)=> norm(
          tds[i]?.getAttribute?.('aria-label') || tds[i]?.getAttribute?.('data-title') ||
          tds[i]?.querySelector?.('[title]')?.getAttribute('title') || tds[i]?.textContent || ''
        );
        const tour = get(iTour).replace(/[^\dA-Za-z]/g,'');
        const drv  = get(iDrv);
        if (tour && drv) map.set(tour, drv);
      }
      if (map.size){ window.__pmTour2Driver = map; return; }
    }
  }

  async function buildTourDriverMapAutoload(){
    try{
      if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size) return;
      const map = new Map();
      qsaDeep('table,[role="grid"]').filter(el => el.offsetParent !== null).forEach(tbl => collectMapFromTable(tbl, map));
      await scrollGridAndCollect(map);
      if (map.size){ window.__pmTour2Driver = map; }
    }catch(e){ console.warn('[PM] AutoLoad Fehler:', e); }
  }

  // ==== DOM-Index: Status / Kommentar / Tour→Fahrer ====
  const _gridIndex = { ts: 0, statusByPsn: new Map(), commentByPsn: new Map(), tour2driver: new Map() };

  function buildActiveGridIndex(){
    _gridIndex.statusByPsn.clear();
    _gridIndex.commentByPsn.clear();
    try { if (window.__pmTour2Driver instanceof Map) _gridIndex.tour2driver = window.__pmTour2Driver; } catch {}

    const grid = findVehicleGridContainer();
    if (!grid) { _gridIndex.ts = Date.now(); return; }

    const rows = qsaDeep('tbody tr, [role="row"]', grid);
    let psnIdx=-1, statIdx=-1, commIdx=-1, tourIdx=-1, drvIdx=-1;

    const hdrs = qsaDeep('thead th, [role="columnheader"]', grid);
    hdrs.forEach((h,i)=>{
      const t = (h.textContent||h.title||'').trim().toLowerCase();
      if (psnIdx < 0 && /(paket|psn|parcel|scheinnummer|paketscheinnummer)/.test(t)) psnIdx = i;
      if (statIdx< 0 && /(status|state)/.test(t)) statIdx = i;
      if (commIdx< 0 && /(kommentar|frei.?text|comment|note|notiz)/.test(t)) commIdx = i;
      if (tourIdx< 0 && /\btour(\s*nr|nummer)?\b/i.test(t)) tourIdx = i;
      if (drvIdx < 0 && /(zusteller|fahrer|courier|rider)/.test(t)) drvIdx = i;
    });

    rows.forEach(tr=>{
      const cells = qsaDeep('td, [role="gridcell"]', tr);
      const get = (i)=> i>=0 && cells[i] ? (cells[i].getAttribute('aria-label') || cells[i].textContent || '').trim() : '';
      const psn = (get(psnIdx)||'').replace(/\D+/g,'');
      if (!psn) return;
      const stat = get(statIdx) || (tr.querySelector('[title]')?.getAttribute('title')||'');
      const comm = get(commIdx) || '';
      if (stat) _gridIndex.statusByPsn.set(psn, stat);
      if (comm) _gridIndex.commentByPsn.set(psn, comm);

      const tour = (get(tourIdx)||'').replace(/[^\dA-Za-z]/g,'');
      const drv  = (get(drvIdx)||'').trim();
      if (tour && drv && !_gridIndex.tour2driver.has(tour)) _gridIndex.tour2driver.set(tour, drv);
    });

    try { window.__pmTour2Driver = _gridIndex.tour2driver; } catch {}
    _gridIndex.ts = Date.now();
  }

  function statusFromIndex(psn){ return _gridIndex.statusByPsn.get(String(psn||'')) || ''; }
  function commentFromIndex(psn){ return _gridIndex.commentByPsn.get(String(psn||'')) || ''; }

  // Platzhalter; gleich wird statusOf/ commentOf überschrieben
  const statusOf_OLD = r => '';
  const commentOf_OLD = r => '';

  function statusOf_FAST(r){ const s = statusFromIndex(parcelId(r)); return s || statusOf_OLD(r); }
  function commentOf_FAST(r){ const c = commentFromIndex(parcelId(r)); return c || commentOf_OLD(r); }

  // Debounce
  let _idxDebounce = 0;
  function scheduleGridIndexRebuild(){
    if (_idxDebounce) cancelAnimationFrame(_idxDebounce);
    _idxDebounce = requestAnimationFrame(()=>{ buildActiveGridIndex(); });
  }

  function startAutoloadTourDriverMap_FAST(){
    buildTourDriverMapAutoload();
    scheduleGridIndexRebuild();
    const mo = new MutationObserver(()=>{ scheduleGridIndexRebuild(); });
    mo.observe(document.documentElement, {subtree:true, childList:true, attributes:false});
  }

  /* ---------- Depot/Scanserver ---------- */
  function guessDepotSuffixFromVehicleTable(root){
    const grid = root || findVehicleGridContainer(); if (!grid) return '';
    const counts  = new Map();
    const pick = (t) => {
      t = String(t||'');
      let m = t.match(/d0010(\d{3})/i) || t.match(/0010(\d{3})/) || t.match(/010(\d{3})/);
      if (m) return m[1];
      m = t.match(/\b(\d{3})\b/); if (m) return m[1];
      m = t.match(/0{0,2}(\d{3})\b/); if (m) return m[1];
      return '';
    };
    const ths = qsaDeep('thead th,[role="columnheader"]', grid);
    const iDepot = ths.findIndex(th => /^Depot$/i.test((th.textContent||th.title||'').trim()));
    if (iDepot < 0) return '';
    const rows = qsaDeep('tbody tr,[role="row"]', grid).filter(r => qsaDeep('td,[role="gridcell"]', r).length).slice(0,1000);
    for (const tr of rows){
      const tds = qsaDeep('td,[role="gridcell"]', tr);
      const raw = (tds[iDepot]?.getAttribute?.('aria-label') || tds[iDepot]?.getAttribute?.('data-title') ||
                  tds[iDepot]?.querySelector?.('[title]')?.getAttribute('title') || tds[iDepot]?.innerText || tds[iDepot]?.textContent || '').trim();
      if (!raw) continue;
      const suf = pick(raw);
      if (suf && suf !== '000') counts.set(suf, (counts.get(suf)||0)+1);
    }
    if (!counts.size) return '';
    let best='', bestN=-1; for (const [k,n] of counts){ if (n>bestN){ best=k; bestN=n; } }
    return best;
  }

  function getScanserverBase(){
    let suf = String(getSetting('depotSuffix')||'').replace(/\D+/g,'').slice(-3);
    if (!suf) {
      const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
      if (g) { suf = g; setSetting('depotSuffix', g); }
    }
    if (!suf) suf = '157';
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
    if (psn.length === 13) psn = '0' + psn;
    const base = getScanserverBase();
    const params = new URLSearchParams();
    params.set('_url','file'); params.set('_passwd', pass);
    params.set('_disp','3'); params.set('_pivotxx','0'); params.set('_rastert','4');
    params.set('_rasteryt','0'); params.set('_rasterx','0'); params.set('_rastery','0');
    params.set('_pivot','0'); params.set('_pivotbp','0');
    params.set('_sortby','date|time'); params.set('_dca','0');
    params.set('_tabledef','psn|date|time|sa|tour|zc|sc|adr1|str|hno|plz1|city|dc|etafrom|etato');
    params.set('_arg59','dpd'); params.set('_arg0a', psn); params.set('_arg0b', psn); params.set('_arg0',  psn+','+psn);
    params.set('_csv','0');
    return `${base}?${params.toString()}`;
  }
  function openScanserver(psn){ const url = buildScanserverUrl(psn); if (url) window.open(url, '_blank', 'noopener'); }

  /* ---------- Network Hook ---------- */
  (function hook(){
    if (!window.__pm_fetch_hooked && window.fetch) {
      const orig = window.fetch;
      window.fetch = async function(i, init = {}) {
        const res = await orig(i, init);
        try {
          const uStr = typeof i === 'string' ? i : (i && i.url) || '';
          if (uStr.includes('/dispatcher/api/pickup-delivery') && res.ok) {
            const u = new URL(uStr, location.origin); const q = u.searchParams;
            if (!q.get('parcelNumber')) {
              const h = {};
              const src = (init && init.headers) || (i && i.headers);
              if (src){
                if (src.forEach) src.forEach((v,k)=>h[String(k).toLowerCase()]=String(v));
                else if (Array.isArray(src)) src.forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
                else Object.entries(src).forEach(([k,v])=>h[String(k).toLowerCase()]=String(v));
              }
              if (!h['authorization']) {
                const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                if (m) h['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
              }
              lastOkRequest = { url: u, headers: h };
              const n = document.getElementById(NS + 'note-capture'); if (n) n.style.display = 'none';
              try { scheduleAuto(); } catch {}
              if (autoEnabled && !isBusy && !document.hidden) { fullRefresh().catch(()=>{}); }
            }
          }
        } catch {}
        return res;
      };
      window.__pm_fetch_hooked = true;
    }
    if (!window.__pm_xhr_hooked && window.XMLHttpRequest) {
      const X = window.XMLHttpRequest;
      const open = X.prototype.open, send = X.prototype.send, setH = X.prototype.setRequestHeader;
      X.prototype.open = function(m, u) { this.__pm_url = (typeof u === 'string') ? new URL(u, location.origin) : null; this.__pm_headers = {}; return open.apply(this, arguments); };
      X.prototype.setRequestHeader = function(k, v) { try { this.__pm_headers[String(k).toLowerCase()] = String(v); } catch {} return setH.apply(this, arguments); };
      X.prototype.send = function() {
        const onload = () => {
          try {
            if (this.__pm_url && this.__pm_url.href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300) {
              const q = this.__pm_url.searchParams;
              if (!q.get('parcelNumber')) {
                if (!this.__pm_headers['authorization']) {
                  const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                  if (m) this.__pm_headers['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
                }
                lastOkRequest = { url: this.__pm_url, headers: this.__pm_headers };
                const n = document.getElementById(NS + 'note-capture'); if (n) n.style.display = 'none';
                try { scheduleAuto(); } catch {}
                try { if (autoEnabled && !isBusy && !document.hidden) { fullRefresh().catch(()=>{}); } } catch {}
              }
            }
          } catch {}
          this.removeEventListener('load', onload);
        };
        this.addEventListener('load', onload);
        return send.apply(this, arguments);
      };
      window.__pm_xhr_hooked = true;
    }
  })();

  /* ---------- API-Builder ---------- */
  function buildHeaders(h){ const H=new Headers(); try{
    if(h){ Object.entries(h).forEach(([k,v])=>{const key=k.toLowerCase(); if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)){ H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()), v);} }); }
    if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
  }catch{} return H; }

  function buildUrlPrio(base, page){
    const u = new URL(base.href); const q = u.searchParams;
    q.set('page', String(page)); q.set('pageSize','500'); q.set('priority','prio');
    q.delete('elements'); q.delete('parcelNumber'); u.search = q.toString(); return u;
  }
  function buildUrlElements(base, page, el){
    const u = new URL(base.href); const q = u.searchParams;
    q.set('page', String(page)); q.set('pageSize','500'); q.set('elements', String(el));
    q.delete('priority'); q.delete('parcelNumber'); u.search = q.toString(); return u;
  }

  // ---- Robust Response Helpers ----
  function pickArray(payload){
    if (Array.isArray(payload)) return payload;
    if (payload && Array.isArray(payload.items)) return payload.items;
    if (payload && Array.isArray(payload.content)) return payload.content;
    if (payload && payload.data){
      if (Array.isArray(payload.data)) return payload.data;
      if (Array.isArray(payload.data.items)) return payload.data.items;
      if (Array.isArray(payload.data.content)) return payload.data.content;
    }
    if (payload && Array.isArray(payload.results)) return payload.results;
    // manche HAL/embedded:
    if (payload && payload._embedded){
      const v = Object.values(payload._embedded).find(Array.isArray);
      if (Array.isArray(v)) return v;
    }
    return [];
  }
  function pickTotals(payload, sizeFallback){
    const p = payload || {};
    const pg = p.page || {};
    const totalElements =
      Number(p.totalElements ?? p.total ?? p.count ?? pg.totalElements ?? pg.total ?? 0);
    const totalPages =
      Number(p.totalPages ?? pg.totalPages ?? (totalElements ? Math.ceil(totalElements / (sizeFallback||500)) : 0));
    return {
      totalElements: (Number.isFinite(totalElements) ? totalElements : 0),
      totalPages:   (Number.isFinite(totalPages)   ? totalPages   : 0)
    };
  }

 // --- FIXED: FAST Pager mit Fallback-Probing, wenn Totals fehlen ---
async function fetchPagedFast(builder, {concurrency=6, size=500, hardMaxPages=200}={}) {
  if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
  const headers = buildHeaders(lastOkRequest.headers);

  // Seite 1
  const u1 = builder(lastOkRequest.url, 1);
  const r1 = await fetch(u1.toString(), { credentials:'include', headers });
  if (!r1.ok) throw new Error(`HTTP ${r1.status}`);
  const j1 = await r1.json();
  const chunk1 = pickArray(j1);

  // Totals lesen (können bei PRIO fehlen)
  const { totalElements, totalPages: tpRaw } = pickTotals(j1, size);
  let totalPages = tpRaw || 0;

  // Fall A: Totals bekannt → Parallel-Fetch wie gehabt
  if (totalPages > 1) {
    totalPages = Math.min(totalPages, hardMaxPages);
    const pages = [];
    for (let p=2; p<=totalPages; p++) pages.push(p);

    const out = chunk1.slice();
    let idx = 0;

    async function worker(){
      while (idx < pages.length){
        const p = pages[idx++];
        const u = builder(lastOkRequest.url, p);
        const res = await fetch(u.toString(), { credentials:'include', headers });
        if (!res.ok) continue;
        const j = await res.json();
        const arr = pickArray(j);
        if (arr.length) out.push(...arr);
        // Wenn unerwartet doch dünn → früh abbrechen
        if (arr.length < size){ idx = pages.length; break; }
      }
    }
    const workers = Array.from({length: Math.max(1, Math.min(concurrency, pages.length))}, worker);
    await Promise.all(workers);
    return out;
  }

  // Fall B: Totals unbekannt
  // - Wenn Seite 1 < size, gibt's nur 1 Seite → fertig
  if (chunk1.length < size) return chunk1;

  // - Wenn Seite 1 == size → sequentiell weiterprobieren bis erste "dünne" Page
  const out = chunk1.slice();
  let page = 2;
  while (page <= hardMaxPages) {
    const u = builder(lastOkRequest.url, page);
    const r = await fetch(u.toString(), { credentials:'include', headers });
    if (!r.ok) break;
    const j = await r.json();
    const arr = pickArray(j);
    if (!arr.length) break;
    out.push(...arr);
    if (arr.length < size) break; // letzte Seite erreicht
    page++;
  }
  return out;
}

  async function fetchAllFast(){
    const [prioRows, exp12Rows, exp18Rows] = await Promise.all([
      fetchPagedFast(buildUrlPrio),
      fetchPagedFast((base,p)=>buildUrlElements(base,p,'023')),
      fetchPagedFast((base,p)=>buildUrlElements(base,p,'010'))
    ]);
    return { prioRows, exp12Rows, exp18Rows };
  }

  // ---- Legacy (bewährter) Pager als Fallback ----
  async function fetchPaged(builder){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size = 500, maxPages = 50;
    let page = 1, rows = [];
    while (page <= maxPages) {
      const u = builder(lastOkRequest.url, page);
      const r = await fetch(u.toString(), { credentials:'include', headers });
      if (!r.ok) break;
      const j = await r.json();
      const chunk = pickArray(j);
      rows.push(...chunk);
      if (chunk.length < size) break;
      page++; await sleep(40);
    }
    return rows;
  }
  async function fetchPagedWithTotal(builder){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size = 500, maxPages = 50;
    let page = 1, rows = [], totalKnown = null;
    while (page <= maxPages) {
      const u = builder(lastOkRequest.url, page);
      const r = await fetch(u.toString(), { credentials:'include', headers });
      if (!r.ok) { if (page === 1) throw new Error(`HTTP ${r.status}`); break; }
      const j = await r.json();
      const chunk = pickArray(j);
      if (page === 1) {
        const { totalElements, totalPages } = pickTotals(j, size);
        if (Number.isFinite(totalElements) && totalElements >= 0) totalKnown = totalElements;
        if (Number.isFinite(totalPages) && totalPages > 0) totalKnown = totalKnown ?? totalPages * size;
      }
      rows.push(...chunk);
      if (chunk.length < size) break;
      page++; await sleep(40);
    }
    return { rows, total: totalKnown || rows.length };
  }

  // ---- Table / Modal ----
  const collator = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });

  function buildHeaderHtml(showPredict=false){
    const ths = ['Paketscheinnummer','Adresse','Fahrer','Tour','Status','Zustellzeit','Zusatzcode'];
    if (showPredict) ths.push('Predict');
    return `<tr>${ths.map((h,i)=>`<th data-col="${i}">${h}</th>`).join('')}</tr>`;
  }

  function buildTableShell(showPredict=false){
    return `
    <div id="${NS}vt-wrap" style="position:relative; height:min(70vh,720px); overflow:auto">
      <table class="${NS}tbl" style="position:absolute; top:0; left:0; right:0">
        <thead>${buildHeaderHtml(showPredict)}</thead>
        <tbody id="${NS}vt-body"></tbody>
      </table>
      <div id="${NS}vt-spacer" style="height:0px"></div>
    </div>`;
  }

  function rowHtml(r, showPredict){
    const pLink = r.__pid
      ? `<a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${r.__pid}" target="_blank" rel="noopener">${r.__pid}</a>`
      : '—';
    const eye = r.__pid ? `<button class="${NS}eye" title="Scanserver öffnen" data-psn="${esc(r.__pid)}">👁</button>` : '';
    const dtime = r.__delivTs ? formatHHMM(new Date(r.__delivTs)) : '—';
    const pred  = showPredict && !r.__delivTs ? (r.__predStartTs ? formatHHMM(new Date(r.__predStartTs)) : '—') : (showPredict?'—':'');

    const cells = [
      `${eye}${pLink}`,
      esc(r.__addr),
      esc(r.__driver),
      esc(String(r.__tourNum||'—')),
      esc(r.__status || '—'),
      esc(dtime),
      esc(r.__codesStr)
    ];
    if (showPredict) cells.push(esc(pred));

    return `<tr>${cells.map(v=>`<td>${v}</td>`).join('')}</tr>`;
  }

  function openModal(title, htmlOrRows, rowsOpt=null, optsOpt=null){
    const m=document.getElementById(NS+'modal');
    const t=document.getElementById(NS+'modal-title');
    const b=document.getElementById(NS+'modal-body');
    if (t) t.textContent = title || '';

    if (Array.isArray(htmlOrRows)) {
      const rows = htmlOrRows;
      const opts = Object.assign({showPredict:false}, optsOpt||{});
      state._modal = { rows: rows.slice(), opts, title: title||'' };

      if (b) b.innerHTML = buildTableShell(opts.showPredict);

      const wrap   = document.getElementById(NS+'vt-wrap');
      const tbody  = document.getElementById(NS+'vt-body');
      const spacer = document.getElementById(NS+'vt-spacer');

      const probe = document.createElement('tbody');
      probe.innerHTML = rowHtml(rows[0] || {}, opts.showPredict);
      const tmpTable = document.createElement('table');
      tmpTable.className = NS+'tbl';
      tmpTable.style.visibility = 'hidden';
      tmpTable.appendChild(probe);
      b.appendChild(tmpTable);
      const rowH = Math.max(28, tmpTable.querySelector('tr')?.getBoundingClientRect().height || 32);
      tmpTable.remove();

      const total = rows.length;
      const buffer = 20;

      function renderWindow(){
        if (!wrap || !tbody || !spacer) return;
        const vpH = wrap.clientHeight;
        const first = Math.max(0, Math.floor(wrap.scrollTop / rowH) - buffer);
        const visCount = Math.ceil(vpH / rowH) + 2*buffer;
        const last = Math.min(total, first + visCount);

        tbody.style.transform = `translateY(${first * rowH}px)`;

        let html = '';
        for (let i=first;i<last;i++) html += rowHtml(rows[i], opts.showPredict);
        tbody.innerHTML = html;

        spacer.style.height = `${total * rowH}px`;
      }

      const thead = wrap.querySelector('thead');
      if (thead){
        thead.addEventListener('click', (ev)=>{
          const th = ev.target.closest('th'); if (!th) return;
          const col = Number(th.dataset.col||0);
          Array.from(thead.querySelectorAll('th')).forEach(x=>x.classList.remove(NS+'sort-asc',NS+'sort-desc'));
          const asc = !(th.dataset.dir==='asc'); th.dataset.dir = asc ? 'asc' : 'desc';
          th.classList.add(asc?NS+'sort-asc':NS+'sort-desc');

          const getKey = (r)=>{
            switch(col){
              case 0: return r.__pid;
              case 1: return r.__addr;
              case 2: return r.__driver;
              case 3: return r.__tourNum;
              case 4: return r.__status;
              case 5: return r.__delivTs;
              case 6: return r.__codesStr;
              case 7: return r.__predStartTs;
              default: return '';
            }
          };
          rows.sort((a,b)=>{
            const A = getKey(a), B=getKey(b);
            if (typeof A==='number' && typeof B==='number') return asc ? (A-B) : (B-A);
            return asc ? collator.compare(String(A), String(B)) : collator.compare(String(B), String(A));
          });
          state._modal.rows = rows;
          renderWindow();
        });
      }

      let raf = 0;
      wrap.addEventListener('scroll', ()=>{
        if (raf) return;
        raf = requestAnimationFrame(()=>{ raf = 0; renderWindow(); });
      }, {passive:true});

      renderWindow();
    } else {
      if (b) b.innerHTML = htmlOrRows || '';
    }

    if (m) m.style.display='flex';
  }

  function hideModal(){ const m=document.getElementById(NS+'modal'); if(m) m.style.display='none'; }
  function formatHHMM(d){ return d ? d.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}) : ''; }

  const parcelId   = r => r.parcelNumber || (Array.isArray(r.parcelNumbers)&&r.parcelNumbers[0]) || r.id || '';
  const addrOf  = r => [r.street, r.houseno].filter(Boolean).join(' ');
  const placeOf = r => [r.postalCode, r.city].filter(Boolean).join(' ');
  const addCodes   = r => Array.isArray(r.additionalCodes)? r.additionalCodes.map(String): [];
  const isDelivery = r => String(r?.orderType || '').toUpperCase() === 'DELIVERY';
  const isPRIO     = r => String(r?.priority||r?.prio||'').toUpperCase() === 'PRIO';
  const hasExpress12 = r => { const el = r?.elements; if (Array.isArray(el)) return el.map(String).includes('023'); return typeof el === 'string' ? /\b023\b/.test(el) : false; };
  const hasExpress18 = r => { const el = r?.elements; if (Array.isArray(el)) return el.map(String).includes('010'); return typeof el === 'string' ? /\b010\b/.test(el) : false; };
  const hasExpressAny = r => hasExpress12(r) || hasExpress18(r);
  const expressTypeOf = r => hasExpress12(r) ? '12' : (hasExpress18(r) ? '18' : '');
  const composeDateTime = (dateStr, timeStr) => { if (!dateStr || !timeStr) return null; const s = `${dateStr}T${String(timeStr).slice(0,8)}`; const d = new Date(s); return isNaN(d) ? null : d; };
  const fromTime = r => r.from2 ? composeDateTime(r.date, r.from2) : null;
  const toTime   = r => r.to2   ? composeDateTime(r.date, r.to2)   : null;
  const deliveredTime = r => r.deliveredTime ? new Date(r.deliveredTime) : null;

  // Status-Mapping & Resolver (Legacy + DOM)
  let statusColIdx = null;
  const apiStatus = r => (r.statusName || r.statusText || r.stateText || r.status || r.deliveryStatus || r.parcelStatus || '').toString().trim();
  function findStatusColumnIndex(){
    if (statusColIdx != null) return statusColIdx;
    const headers = Array.from(document.querySelectorAll('thead th, [role="columnheader"]'));
    let idx = -1;
    headers.forEach((h,i)=>{ const t=(h.textContent||'').trim().toLowerCase(); if (idx===-1 && /(status|stat)/.test(t)) idx=i; });
    statusColIdx = idx >= 0 ? idx : null; return statusColIdx;
  }
  function statusFromDomByParcel(parcel){
    if (!parcel) return '';
    if (statusCache.has(parcel)) return statusCache.get(parcel)||'';
    const idx = findStatusColumnIndex();
    const rows = Array.from(document.querySelectorAll('tbody tr, [role="row"]')).filter(tr => (tr.textContent || '').includes(parcel));
    for (const tr of rows){
      let val = '';
      if (idx!=null){
        const cells = Array.from(document.querySelectorAll('td, [role="gridcell"]'));
        const cell = cells[idx];
        if (cell){
          const div = cell.querySelector('div[title]');
          val = (div?.getAttribute('title') || cell.textContent || '').trim();
        }
      }
      if (!val){
        const guess = Array.from(tr.querySelectorAll('*')).find(el=>/(ZUGESTELLT|ZUSTELLUNG|PROBLEM)/.test((el.textContent||'').trim().toUpperCase()));
        if (guess) val = (guess.textContent||'').trim();
      }
      if (val){ statusCache.set(parcel, val); return val; }
    }
    return '';
  }
  function mapEnToDeStatus(s){
    const u = String(s||'').toUpperCase();
    if (!u) return '';
    if (/DELIVERED|DELIVERED_TO_PUDO/.test(u)) return 'ZUGESTELLT';
    if (/OUT_FOR_DELIVERY|IN_DELIVERY|ON.?ROUTE|VEHICLE/.test(u)) return 'ZUSTELLUNG';
    if (/DELIVERY.*PROBLEM|PROBLEM|FAIL|NOT.?DELIVERED/.test(u)) return 'ZUSTELLUNG PROBLEM';
    if (/AT.?DEPOT|AT.?HUB/.test(u)) return 'IM DEPOT';
    if (/RETURN/.test(u)) return 'RETURNSENDUNG';
    if (/DELIVERY/.test(u)) return 'ZUSTELLUNG';
    return '';
  }
  const statusOf = r => { const p = parcelId(r); const fromDom = statusFromDomByParcel(p); if (fromDom) return fromDom; const mapped = mapEnToDeStatus(apiStatus(r)); return mapped || '—'; };

  // delivered-Logik nutzt FAST-Status
  const delivered = r => {
    if (r.deliveredTime) return true;
    const s = (statusOf_FAST(r) || statusOf(r) || '').toUpperCase();
    return /(ZUGESTELLT)/.test(s);
  };

  // Fahrer
  const tour2Driver = (() => window.__pmTour2Driver instanceof Map ? window.__pmTour2Driver : new Map());
  const tourOf  = r => r.tour ? String(r.tour) : '';
  const driverOf = r => {
    const direct = r.driverName || r.driver || r.courierName || r.riderName || r.tourDriver || '';
    if (direct && direct.trim()) return direct.trim();
    const key = String(tourOf(r) || '').replace(/[^\dA-Za-z]/g,'');
    return (tour2Driver().get(key) || '').trim();
  };

  function setLoading(on){ isLoading=!!on; const el=document.getElementById(NS+'loading'); if(el) el.classList.toggle('on',on); }
  function dimButtons(on){ document.querySelectorAll('.'+NS+'btn-sm').forEach(b=>b.classList.toggle(NS+'dim', !!on)); }
  function render(){
    const list = document.getElementById(NS+'list'); if (!list) return; list.innerHTML='';
    const d=document.createElement('div'); d.className=NS+'empty'; d.textContent= isLoading ? 'Lade Daten …' : 'Aktualisiert.';
    list.appendChild(d);
  }

  function setKpis({prioAll, prioOpen, expAll, expOpen}){
    const set = (id,val,chipId)=>{ const el=document.getElementById(id); const chip=document.getElementById(chipId); if(el) el.textContent=String(Number(val||0)); if(chip) chip.style.display=Number(val||0)===0?'none':''; };
    set(NS+'kpi-prio-all',  prioAll, NS+'chip-prio-all');
    set(NS+'kpi-prio-open', prioOpen, NS+'chip-prio-open');
    set(NS+'kpi-exp-all',   expAll,  NS+'chip-exp-all');
    set(NS+'kpi-exp-open',  expOpen, NS+'chip-exp-open');
  }
  function addEvent(ev){
    const e={ id:state.nextId++, title:ev.title||'Ereignis', meta:ev.meta||'', sev:ev.sev||'info', ts:ev.ts||Date.now(), read:!!ev.read, comment:ev.comment||'', hasComment:!!(ev.comment&&ev.comment.trim()), parcel:ev.parcel||'', kind:ev.kind||'' };
    state.events.push(e); render();
  }

  function buildUrlByParcel(base, parcel){
    const u = new URL(base.href); const q = u.searchParams;
    q.set('page','1'); q.set('pageSize','500'); q.set('parcelNumber', String(parcel));
    q.delete('priority'); q.delete('elements'); u.search = q.toString(); return u;
  }
  let commentColIdx = null;
  function findCommentColumnIndex() { if (commentColIdx != null) return commentColIdx; const headers = Array.from(document.querySelectorAll('thead th, [role="columnheader"]')); let idx = -1; headers.forEach((h, i) => { const t = (h.textContent || '').trim().toLowerCase(); if (idx === -1 && /(kommentar|frei.?text|free.?text|notiz)/.test(t)) idx = i; }); commentColIdx = idx >= 0 ? idx : null; return commentColIdx; }
  function commentFromDomByParcel(parcel) {
    if (!parcel) return '';
    const idx = findCommentColumnIndex();
    const rows = Array.from(document.querySelectorAll('tbody tr, [role="row"]')).filter(tr => (tr.textContent || '').includes(parcel));
    if (!rows.length) return '';
    for (const tr of rows) {
      if (idx != null) {
        const cells = Array.from(tr.querySelectorAll('td, [role="gridcell"]'));
        const cell = cells[idx];
        if (cell) {
          const div = cell.querySelector('div[title]');
          const val = (div?.getAttribute('title') || cell.textContent || '').trim();
          if (val) return val;
        }
      }
      const any = tr.querySelector('td div[title]'); if (any) { const val = (any.getAttribute('title') || any.textContent || '').trim(); if (val) return val; }
    }
    return '';
  }
  const pickComment = (r) => { const fields = [ r?.freeComment, r?.freeTextDc, r?.note, r?.comment ]; for (const f of fields) { if (Array.isArray(f) && f.length) return f.filter(Boolean).join(' | '); if (typeof f === 'string' && f.trim()) return f.trim(); } return ''; };
  const commentOf = (r) => { const fromApi = pickComment(r); if (fromApi) return fromApi; return commentFromDomByParcel(parcelId(r)); };

  async function fetchMissingComments(){
    if(!lastOkRequest) { addEvent({title:'Hinweis', meta:'Kein API-Request zum Klonen vorhanden.', sev:'info', read:true}); render(); return; }
    const headers = buildHeaders(lastOkRequest.headers);
    const toFill = state.events.filter(e => !e.hasComment && e.parcel);
    if (toFill.length === 0) return;
    dimButtons(true);
    let filled = 0;
    for (const ev of toFill){
      const p = ev.parcel;
      if (commentCache.has(p)) {
        const c = commentCache.get(p);
        if (c) { ev.comment = c; ev.hasComment = true; filled++; render(); }
        continue;
      }
      try{
        const url = buildUrlByParcel(lastOkRequest.url, p);
        const res = await fetch(url.toString(), { credentials:'include', headers });
        if (res.ok){
          const j = await res.json();
          const rows = pickArray(j);
          const r = rows.find(x => (parcelId(x) === p)) || rows[0];
          let c = '';
          if (r) c = pickComment(r) || commentFromDomByParcel(p);
          commentCache.set(p, c);
          if (c) { ev.comment = c; ev.hasComment = true; filled++; render(); }
        }
      }catch(e){ /* ignore */ }
      await sleep(80);
    }
    if (filled>0) addEvent({title:'Kommentare', meta:`Nachgeladen: ${filled}/${toFill.length}`, sev:'info', read:true});
    dimButtons(false); render();
  }

  function buildTableRowsAndCounts(prioRows, exp12Rows, exp18Rows){
    const prioDeliveries = prioRows.filter(isDelivery).filter(isPRIO);
    const prioAll  = prioDeliveries;
    const prioOpen = prioDeliveries.filter(r=>!delivered(r));

    const expRows   = [...exp12Rows, ...exp18Rows];
    const seen = new Set();
    const expDeliveries = expRows
      .filter(isDelivery)
      .filter(hasExpressAny)
      .filter(r=>{ const id = parcelId(r); if (!id || seen.has(id)) return false; seen.add(id); return true; });
    const expAll    = expDeliveries;
    const expOpen   = expDeliveries.filter(r=>!delivered(r));
    const expLate11 = expDeliveries.filter(hasExpress12).filter(r=>{
      const ft = fromTime(r); if (!ft) return false;
      return (ft.getHours() > 11) || (ft.getHours()===11 && ft.getMinutes()>=1);
    });
    return { prioAll, prioOpen, expAll, expOpen, expLate11 };
  }

  function showPrioAll(){
    const rows = state._prioAllList;
    openModal(`PRIO – in Ausrollung (alle) · ${rows.length}`, rows, rows, {showPredict:false});
  }
  function showPrioOpen(){
    const rows = state._prioOpenList;
    openModal(`PRIO – noch nicht zugestellt · ${rows.length}`, rows, rows, {showPredict:true});
  }
  function filterByExpressSelection(rows){ if (state.filterExpress === '12') return rows.filter(hasExpress12); if (state.filterExpress === '18') return rows.filter(hasExpress18); return rows; }
  function getFilteredExpressCounts(){
    const f = state.filterExpress;
    const filt = f==='12' ? hasExpress12 : f==='18' ? hasExpress18 : null;
    const expAllList  = filt ? state._expAllList.filter(filt)  : state._expAllList;
    const expOpenList = filt ? state._expOpenList.filter(filt) : state._expOpenList;
    return { expAllCount: expAllList.length, expOpenCount: expOpenList.length };
  }
  function updateKpisForCurrentState(){
    const { expAllCount, expOpenCount } = getFilteredExpressCounts();
    setKpis({ prioAll: state._prioAllList.length, prioOpen: state._prioOpenList.length, expAll: expAllCount, expOpen: expOpenCount });
  }
  function showExpAll(){
    const src = state._expAllList;
    const rows = filterByExpressSelection(src);
    const sel = state.filterExpress==='12'?' (12)': state.filterExpress==='18'?' (18)':'';
    openModal(`Express${sel} – in Ausrollung (alle) · ${rows.length}`, rows, rows, {showPredict:false});
  }
  function showExpOpen(){
    const src = state._expOpenList;
    const rows = filterByExpressSelection(src);
    const sel = state.filterExpress==='12'?' (12)': state.filterExpress==='18'?' (18)':'';
    openModal(`Express${sel} – noch nicht zugestellt · ${rows.length}`, rows, rows, {showPredict:true});
  }
  function showExpLate11(){
    const rows = state._expLate11List.slice();
    openModal(`Express 12 – falsch einsortiert (>11:01 geplant) · ${rows.length}`, rows, rows, {showPredict:true});
  }

  // ---- FAST Normalisierung ----
  function normRowFAST(r){
    const getPred = ()=>{ const f=fromTime(r); if (f) return +f; const t=toTime(r); if (t) return +t; return 0; };
    return {
      ...r,
      __pid: parcelId(r) || '',
      __addr: [addrOf(r), placeOf(r)].filter(Boolean).join(' · ') || '—',
      __driver: (driverOf(r) || _gridIndex.tour2driver.get(String(tourOf(r)||'').replace(/[^\dA-Za-z]/g,'')) || '—').trim(),
      __tourNum: Number(tourOf(r) || 0),
      __status: statusOf_FAST(r) || statusOf(r) || '—',
      __delivTs: deliveredTime(r) ? deliveredTime(r).getTime() : 0,
      __predStartTs: getPred(),
      __codesStr: (addCodes(r)||[]).join(', ') || '—',
      __expType: expressTypeOf(r)
    };
  }

  async function refreshViaApi_FAST(){
    const { prioRows, exp12Rows, exp18Rows } = await fetchAllFast();

    const prioN = prioRows.map(normRowFAST);
    const exp12N = exp12Rows.map(normRowFAST);
    const exp18N = exp18Rows.map(normRowFAST);

    const { prioAll, prioOpen, expAll, expOpen, expLate11 } = buildTableRowsAndCounts(prioN, exp12N, exp18N);

    state._prioAllList = prioAll.slice();
    state._prioOpenList = prioOpen.slice();
    state._expAllList = expAll.slice();
    state._expOpenList = expOpen.slice();
    state._expLate11List = expLate11.slice();

    const { expAllCount, expOpenCount } = getFilteredExpressCounts();
    setKpis({
      prioAll: prioAll.length,
      prioOpen: prioOpen.length,
      expAll: expAllCount,
      expOpen: expOpenCount
    });

    state.events = [{
      id: ++state.nextId,
      title:'Aktualisiert (FAST)',
      meta:`PRIO: in Ausrollung ${prioAll.length} • offen ${prioOpen.length} • EXPRESS: in Ausrollung ${expAll.length} • offen ${expOpenCount} • „>11:01“ (12er): ${expLate11.length}`,
      sev:'info', read:true, ts:Date.now()
    }];
  }

  // ---- Legacy-Normalisierung/Fallback (aus 6.0.0) ----
  function normRowLegacy(r){
    const getPred = ()=>{ const f=fromTime(r); if (f) return +f; const t=toTime(r); if (t) return +t; return 0; };
    return {
      ...r,
      __pid:        parcelId(r) || '',
      __addr:       [addrOf(r), placeOf(r)].filter(Boolean).join(' · ') || '—',
      __driver:     (driverOf(r) || '—').trim(),
      __tourNum:    Number(tourOf(r) || 0),
      __status:     statusOf(r) || '—',
      __delivTs:    deliveredTime(r) ? deliveredTime(r).getTime() : 0,
      __predStartTs: getPred(),
      __codesStr:   (addCodes(r)||[]).join(', ') || '—',
      __expType:    expressTypeOf(r)
    };
  }

  async function refreshViaApi_Legacy(){
    const [prioRes, exp12Rows, exp18Rows] = await Promise.all([
      fetchPagedWithTotal(buildUrlPrio),
      fetchPaged((base,p)=>buildUrlElements(base,p,'023')),
      fetchPaged((base,p)=>buildUrlElements(base,p,'010'))
    ]);
    const prioRows   = prioRes.rows.map(normRowLegacy);
    const exp12N     = exp12Rows.map(normRowLegacy);
    const exp18N     = exp18Rows.map(normRowLegacy);

    const { prioAll, prioOpen, expAll, expOpen, expLate11 } =
          buildTableRowsAndCounts(prioRows, exp12N, exp18N);

    state._prioAllList   = prioAll.slice();
    state._prioOpenList  = prioOpen.slice();
    state._expAllList    = expAll.slice();
    state._expOpenList   = expOpen.slice();
    state._expLate11List = expLate11.slice();

    const { expAllCount, expOpenCount } = getFilteredExpressCounts();
    setKpis({
      prioAll: prioAll.length,
      prioOpen: prioOpen.length,
      expAll: expAllCount,
      expOpen: expOpenCount
    });

    state.events = [{
      id: ++state.nextId,
      title:'Aktualisiert',
      meta:`PRIO: in Ausrollung ${prioAll.length} • offen ${prioOpen.length} • EXPRESS: in Ausrollung ${expAll.length} • offen ${expOpenCount} • „>11:01“ (12er): ${expLate11.length}`,
      sev:'info', read:true, ts:Date.now()
    }];
  }

  // ---- SAFE Wrapper: FAST zuerst, sonst Legacy ----
  async function refreshViaApi_SAFE(){
    try {
      await refreshViaApi_FAST();
      const nothing =
        !state._prioAllList.length && !state._expAllList.length &&
        !state._prioOpenList.length && !state._expOpenList.length;
      if (nothing) {
        await refreshViaApi_Legacy();
      }
    } catch(e){
      await refreshViaApi_Legacy();
    }
  }

  // ---- Full Refresh ----
  async function fullRefresh(){
    if (isBusy) return;
    try{
      if (!lastOkRequest) { addEvent({title:'Hinweis', meta:'Kein API-Request erkannt. Bitte einmal die normale Suche ausführen, danach funktioniert Auto 60s.', sev:'info', read:true}); render(); return; }
      isBusy = true; setLoading(true); dimButtons(true);
      addEvent({title:'Aktualisiere (API)…', meta:'FAST (Parallel-Paging) + DOM-Indizes; Fallback auf Legacy bei Bedarf', sev:'info', read:true}); render();

      // DOM-Index vorziehen
      scheduleGridIndexRebuild();

      await refreshViaApi_SAFE();

      // Kommentare ggf. nachladen
      await fetchMissingComments();
    } catch(e){
      console.error(e);
      addEvent({title:'Fehler (API)', meta:String(e && e.message || e), sev:'warn'});
    } finally{
      setLoading(false); dimButtons(false); isBusy = false; render();
    }
  }

  function scheduleAuto(){
    try{
      if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
      if (!autoEnabled) return;
      if (document.hidden) return;
      if (!lastOkRequest) return;
      autoTimer = setInterval(()=>{ if (!lastOkRequest) return; fullRefresh().catch(()=>{}); }, 60_000);
    }catch{}
  }
  document.addEventListener('visibilitychange', ()=>scheduleAuto());

  /* ---------- Boot ---------- */
  function boot(){
    mountUI();
    startAutoloadTourDriverMap_FAST();
    scheduleAuto();
    // Depotkennung initial raten
    setTimeout(()=> {
      if (!(getSetting('depotSuffix')||'')) {
        const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
        if (g) { setSetting('depotSuffix', g); addEvent({title:'Einstellungen', meta:`Depotkennung aus Fahrzeugübersicht: ${g}`, sev:'info', read:true}); }
      }
    }, 1200);
  }
})();
