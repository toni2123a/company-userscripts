// ==UserScript==
// @name         DPD Dispatcher – Prio / Express
// @namespace    bodo.dpd.custom
// @version      7.3.0
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @description  PRIO / EXPRESS: automatische Abholung-und-Zustellung-Erkennung, direkte API-Auswertung, sortierbare Listen, Predict-Zeitfenster, EXPRESS12 >11:01, Fahrer aus Fahrzeugübersicht, Systempartner aus lokaler TourMap.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const moduleDef = {
    id: 'prio-express-monitor',
    label: 'Prio / Express',
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
    if (started) {
      togglePanel(true);
      initialOpenRefresh().catch(console.error);
      return;
    }
    started = true;
    boot();
    togglePanel(true);
    initialOpenRefresh().catch(console.error);
  }

  const NS  = 'pm-';
  const esc = s => String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim();
  const tourKey = t => String(t || '').replace(/[^\dA-Za-z]/g, '').trim();
  const collator = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });

  const state = {
    filterExpress: 'all', // all | 12 | 18
    events: [],
    nextId: 1,
    _bootShown: false,
    _prioAllList: [],
    _prioOpenList: [],
    _expAllList: [],
    _expOpenList: [],
    _expLate11List: [],
    _modal: { rows: [], opts: {}, title: '', selected: new Set() },
    lastRefreshAt: 0
  };

  let lastOkRequest = null;
  let isBusy = false;
  let isLoading = false;
  let sourceRefreshTimer = null;
  let suppressSourceRefreshUntil = 0;

  const TP_IDB_NAME = 'fvpr_db';
  const TP_STORE = 'tourMap';
  const TP_PARTNERS_STORE = 'partners';

  let tourPartnerMap = new Map();

  function tpIdbOpen() {
    return new Promise((res, rej) => {
      const req = indexedDB.open(TP_IDB_NAME);
      req.onsuccess = () => res(req.result);
      req.onerror = () => rej(req.error);
    });
  }

  async function tpIdbAll(store) {
    try {
      const db = await tpIdbOpen();
      if (!db.objectStoreNames.contains(store)) return [];
      return new Promise((res) => {
        const r = db.transaction(store, 'readonly').objectStore(store).getAll();
        r.onsuccess = () => res(r.result || []);
        r.onerror = () => res([]);
      });
    } catch {
      return [];
    }
  }

  async function tpIdbGet(store, key) {
    try {
      const db = await tpIdbOpen();
      if (!db.objectStoreNames.contains(store)) return null;
      return new Promise((res) => {
        const r = db.transaction(store, 'readonly').objectStore(store).get(key);
        r.onsuccess = () => res(r.result || null);
        r.onerror = () => res(null);
      });
    } catch {
      return null;
    }
  }

  async function loadTourPartnerMap() {
    try {
      const rows = await tpIdbAll(TP_STORE);
      const m = new Map();
      for (const r of rows) {
        const t = tourKey(r.tour || '');
        const p = norm(r.partner || '');
        if (t && p) m.set(t, p);
      }
      tourPartnerMap = m;
    } catch {
      tourPartnerMap = new Map();
    }
  }

  async function getPartnerMailRecord(partnerName) {
    const key = norm(partnerName || '');
    if (!key) return null;
    return await tpIdbGet(TP_PARTNERS_STORE, key);
  }

  function splitEmails(raw) { return (raw || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean); }
  function isEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }
  function normalizeEmailList(raw) {
    const arr = splitEmails(raw);
    const valid = [], invalid = [], seen = new Set();
    for (const a of arr) {
      const low = a.toLowerCase();
      if (seen.has(low)) continue;
      seen.add(low);
      (isEmail(a) ? valid : invalid).push(a);
    }
    return { valid, invalid };
  }

  function openMailto(subject, to, cc) {
    const params = new URLSearchParams();
    if (subject) params.set('subject', subject);
    if (cc) params.set('cc', cc);
    window.location.href = `mailto:${encodeURIComponent(to || '')}?${params.toString()}`;
  }

  async function copyHtmlToClipboard(html) {
    try {
      if (!navigator.clipboard || !window.ClipboardItem) throw new Error('ClipboardItem nicht verfügbar');
      const blobHtml = new Blob([html], { type: 'text/html' });
      const blobTxt  = new Blob([html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()], { type: 'text/plain' });
      const item = new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobTxt });
      await navigator.clipboard.write([item]);
      return true;
    } catch {
      try {
        await navigator.clipboard.writeText(html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim());
        return true;
      } catch {
        return false;
      }
    }
  }

  function ensureStyles() {
    if (document.getElementById(NS + 'style')) return;
    const style = document.createElement('style');
    style.id = NS + 'style';
    style.textContent = `
      .${NS}panel{position:fixed;top:72px;left:50%;transform:translateX(-50%);width:min(1100px,95vw);max-height:78vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100000;display:none}
      .${NS}header{display:grid;grid-template-columns:1fr;gap:10px;align-items:start;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
      .${NS}toolbar{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:8px}
      .${NS}group{display:flex;flex-wrap:wrap;align-items:center;gap:8px;background:#f9fafb;border:1px solid rgba(0,0,0,.08);border-radius:12px;padding:6px 8px}
      .${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer}
      .${NS}btn-sm.active{background:#c1121f;color:#fff;border-color:#9b0d18}
      .${NS}btn-sm.dim{opacity:.6;pointer-events:none}
      .${NS}kpis{display:flex;flex-wrap:wrap;gap:8px;margin-top:4px}
      .${NS}kpi{flex:1 1 auto;background:#f5f5f5;border:1px solid rgba(0,0,0,.08);padding:6px 10px;border-radius:999px;font:600 12px system-ui;white-space:nowrap;cursor:pointer}
      .${NS}list{list-style:none;margin:0;padding:0}
      .${NS}empty{padding:14px 12px;opacity:.75;text-align:center;font:500 12px system-ui}
      .${NS}loading{display:none;padding:8px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:600 12px system-ui;background:#fffbe6}
      .${NS}loading.on{display:block}
      .${NS}foot{padding:8px 12px;border-top:1px solid rgba(0,0,0,.08);font:600 12px system-ui;color:#475569;background:#fafafa}
      .${NS}modal{position:fixed;inset:0;display:none;align-items:flex-start;justify-content:center;background:rgba(0,0,0,.35);z-index:100001}
      .${NS}modal-inner{background:#fff;width:min(1600px,96vw);height:min(88vh,1000px);overflow:auto;border-radius:12px;box-shadow:0 12px 28px rgba(0,0,0,.2);border:1px solid rgba(0,0,0,.12)}
      .${NS}modal-head{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui;position:sticky;top:0;background:#fff;z-index:2}
      .${NS}modal-body{padding:8px 12px;max-height:calc(100% - 46px);overflow:auto}
      .${NS}tbl{width:100%;border-collapse:collapse;font:12px system-ui}
      .${NS}tbl th,.${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .${NS}tbl th{text-align:left;background:#fafafa;position:sticky;top:0;cursor:pointer;user-select:none;z-index:1}
      .${NS}sort-asc::after{content:" ▲";font-size:11px}
      .${NS}sort-desc::after{content:" ▼";font-size:11px}
      .${NS}eye{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border:1px solid rgba(0,0,0,.2);border-radius:6px;background:#fff;margin-right:6px;cursor:pointer;font-size:12px;line-height:1}
      .${NS}eye:hover{background:#f3f4f6}
      .${NS}badge{display:inline-block;padding:2px 6px;border-radius:999px;font-size:11px;border:1px solid rgba(0,0,0,.15);background:#f3f4f6}
      .${NS}badge-status-ok{background:#16a34a;color:#fff;border-color:#15803d}
      .${NS}badge-status-problem{background:#dc2626;color:#fff;border-color:#b91c1c}
      .${NS}badge-status-run{background:#eab308;color:#111827;border-color:#ca8a04}
      .${NS}detail-row > td{background:#f9fafb;padding:6px 8px}
      .${NS}detail-inner{border-top:1px solid rgba(0,0,0,.08);margin-top:4px;padding-top:4px}
      .${NS}detail-inner table{width:100%;border-collapse:collapse;font-size:11px}
      .${NS}detail-inner th,.${NS}detail-inner td{border-bottom:1px solid rgba(0,0,0,.06);padding:3px 4px;white-space:nowrap}
      .${NS}row-express{background:#dcfce7}
    `;
    document.head.appendChild(style);
  }

  function formatDateTime(ts) {
    if (!ts) return '—';
    const d = new Date(ts);
    return d.toLocaleString('de-DE', { hour12: false });
  }

  function setLastRefreshNow() {
    state.lastRefreshAt = Date.now();
    const el = document.getElementById(NS + 'last-refresh');
    if (el) el.textContent = `Aktualisiert: ${formatDateTime(state.lastRefreshAt)}`;
  }

  function mountUI() {
    ensureStyles();
    if (document.getElementById(NS + 'panel')) return;

    const panel = document.createElement('div');
    panel.id = NS + 'panel';
    panel.className = NS + 'panel';
    panel.innerHTML = `
      <div class="${NS}header">
        <div class="${NS}toolbar">
          <div class="${NS}group">
            <button class="${NS}btn-sm active" id="${NS}btn-filter-all" data-action="filterAll">Alles</button>
            <button class="${NS}btn-sm" id="${NS}btn-filter-12" data-action="filter12">Express 12</button>
            <button class="${NS}btn-sm" id="${NS}btn-filter-18" data-action="filter18">Express 18</button>
          </div>
          <div class="${NS}group">
            <button class="${NS}btn-sm" data-action="openSettings">Einstellungen</button>
            <button class="${NS}btn-sm" data-action="refreshApi">Aktualisieren</button>
            <button class="${NS}btn-sm" data-action="showExpLate11">EXPRESS12 >11:01</button>
          </div>
        </div>
        <div class="${NS}kpis">
          <span class="${NS}kpi" id="${NS}chip-prio-all" data-action="showPrioAll">PRIO in Ausrollung: <b id="${NS}kpi-prio-all">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-prio-open" data-action="showPrioOpen">PRIO noch nicht zugestellt: <b id="${NS}kpi-prio-open">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-exp-all" data-action="showExpAll">EXPRESS in Ausrollung: <b id="${NS}kpi-exp-all">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-exp-open" data-action="showExpOpen">EXPRESS noch nicht zugestellt: <b id="${NS}kpi-exp-open">0</b></span>
        </div>
      </div>
      <div id="${NS}loading" class="${NS}loading">Lade Daten …</div>
      <ul id="${NS}list" class="${NS}list"></ul>
      <div class="${NS}foot" id="${NS}last-refresh">Aktualisiert: —</div>
    `;
    document.body.appendChild(panel);

    const modal = document.createElement('div');
    modal.id = NS + 'modal';
    modal.className = NS + 'modal';
    modal.innerHTML = `
      <div class="${NS}modal-inner">
        <div class="${NS}modal-head">
          <div id="${NS}modal-title">Liste</div>
          <div style="display:flex;gap:8px;align-items:center">
            <button class="${NS}btn-sm" data-action="mailSelected" style="display:none" id="${NS}mail-selected">Mail an Systempartner</button>
            <button class="${NS}btn-sm" data-action="closeModal">Schließen</button>
          </div>
        </div>
        <div class="${NS}modal-body" id="${NS}modal-body"></div>
      </div>
    `;
    document.body.appendChild(modal);

    panel.addEventListener('click', async (e) => {
      const chip = e.target.closest('.' + NS + 'kpi');
      if (chip) {
        const a = chip.dataset.action;
        if (a === 'showPrioAll') showPrioAll();
        if (a === 'showPrioOpen') showPrioOpen();
        if (a === 'showExpAll') showExpAll();
        if (a === 'showExpOpen') showExpOpen();
        return;
      }

      const b = e.target.closest('.' + NS + 'btn-sm');
      if (!b) return;
      const a = b.dataset.action;
      if (a === 'openSettings') { openSettingsModal(); return; }
      if (a === 'refreshApi') { await fullRefresh(true).catch(console.error); return; }
      if (a === 'showExpLate11') { showExpLate11(); return; }
      if (a === 'filterAll') { setExpressFilter('all'); return; }
      if (a === 'filter12') { setExpressFilter('12'); return; }
      if (a === 'filter18') { setExpressFilter('18'); return; }
    });

    modal.addEventListener('click', e => {
      if (e.target.dataset.action === 'closeModal' || e.target === modal) { hideModal(); return; }
      const eye = e.target.closest('button.' + NS + 'eye[data-psn]');
      if (eye) { openScanserver(String(eye.dataset.psn || '')); return; }
      const btn = e.target.closest('button[data-action]');
      if (!btn) return;
      const a = btn.dataset.action;
      if (a === 'guessDepot') { guessDepotFromVehicles(); return; }
      if (a === 'saveSettings') { saveSettingsFromModal(); return; }
      if (a === 'mailSelected') { mailSelectedLate11().catch(console.error); return; }
    });

    if (!state._bootShown) {
      addEvent({
        title: 'Bereit',
        meta: 'Quelle: /dispatcher/api/pickup-delivery · Fahrer aus Fahrzeugübersicht · Systempartner aus lokaler TourMap',
        sev: 'info',
        read: true
      });
      state._bootShown = true;
    }

    render();
    updateFilterButtons();

    // Klick außerhalb schließt Modal bzw. Hauptfenster
    document.addEventListener('mousedown', function(e) {

      const modal = document.getElementById(NS + 'modal');
      const modalInner = modal?.querySelector('.' + NS + 'modal-inner');

      if (modal && modal.style.display === 'flex') {
        if (modalInner && !modalInner.contains(e.target)) {
          hideModal();
        }
        return;
      }

      const panel = document.getElementById(NS + 'panel');

      if (
        panel &&
        getComputedStyle(panel).display !== 'none' &&
        !panel.contains(e.target)
      ) {
        panel.style.setProperty('display', 'none', 'important');
      }

    }, true);

  }

  function setExpressFilter(v) {
    state.filterExpress = v;
    updateFilterButtons();
    updateKpisForCurrentState();

    if (document.getElementById(NS + 'modal')?.style.display === 'flex') {
      if (/noch nicht zugestellt/i.test(state._modal.title)) showExpOpen();
      else if (/falsch einsortiert/i.test(state._modal.title)) showExpLate11();
      else if (/Express/i.test(state._modal.title)) showExpAll();
    }
  }

  function updateFilterButtons() {
    const all = document.getElementById(NS + 'btn-filter-all');
    const b12 = document.getElementById(NS + 'btn-filter-12');
    const b18 = document.getElementById(NS + 'btn-filter-18');
    if (all) all.classList.toggle('active', state.filterExpress === 'all');
    if (b12) b12.classList.toggle('active', state.filterExpress === '12');
    if (b18) b18.classList.toggle('active', state.filterExpress === '18');
  }

  function togglePanel(force) {
    const panel = document.getElementById(NS + 'panel');
    if (!panel) { mountUI(); return; }
    const isHidden = getComputedStyle(panel).display === 'none';
    const show = force != null ? !!force : isHidden;
    panel.style.setProperty('display', show ? 'block' : 'none', 'important');
  }

  const LSKEY = 'pmSettings';
  function loadSettings() {
    try { return Object.assign({ scanserverPass: '', depotSuffix: '' }, JSON.parse(localStorage.getItem(LSKEY) || '{}')); }
    catch { return { scanserverPass: '', depotSuffix: '' }; }
  }
  function saveSettingsObj(s) { try { localStorage.setItem(LSKEY, JSON.stringify(s)); } catch {} }
  function setSetting(k, v) { const s = loadSettings(); s[k] = v; saveSettingsObj(s); }
  function getSetting(k) { return loadSettings()[k]; }

  function openSettingsModal() {
    const s = loadSettings();
    const html = `
      <div style="display:grid;gap:10px;max-width:520px">
        <label style="display:grid;gap:6px;font:600 12px system-ui">
          Scanserver-Passwort
          <input id="${NS}inp-pass" type="password" placeholder="••••••••" value="${esc(s.scanserverPass || '')}"
                 style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px"/>
        </label>
        <label style="display:grid;gap:6px;font:600 12px system-ui">
          Depotkennung (3-stellig, z. B. 157)
          <div style="display:flex;gap:8px;align-items:center">
            <input id="${NS}inp-depot" type="text" pattern="\\d{3}" maxlength="3" placeholder="157" value="${esc(String(s.depotSuffix || '').slice(-3))}"
                   style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px;width:100px;text-align:center;font-weight:700;letter-spacing:.5px"/>
            <button class="${NS}btn-sm" data-action="guessDepot">Auto erkennen</button>
          </div>
          <div style="opacity:.7;font:12px system-ui">Host: <code>scanserver-d0010<strong>${esc(String(s.depotSuffix || '').slice(-3) || '157')}</strong>.ssw.dpdit.de</code></div>
        </label>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button class="${NS}btn-sm" data-action="saveSettings">Speichern</button>
        </div>
      </div>`;
    openModal('Einstellungen', html);
  }

  function findVehicleGridContainer() {
    const cands = Array.from(document.querySelectorAll(
      '.MuiDataGrid-virtualScroller, .MuiDataGrid-main, [data-rttable="true"], [role="grid"], table'
    )).filter(el => el.offsetParent !== null);
    if (!cands.length) return null;
    const scrollable = cands.filter(el => {
      const cs = getComputedStyle(el);
      const ov = cs.overflowY || cs.overflow;
      return /auto|scroll/i.test(ov || '') && el.scrollHeight > el.clientHeight;
    });
    const pool = scrollable.length ? scrollable : cands;
    return pool.reduce((best, el) => ((el.scrollHeight || 0) > ((best?.scrollHeight) || 0) ? el : best), null);
  }

  function guessDepotSuffixFromVehicleTable(root) {
    const grid = root || findVehicleGridContainer();
    if (!grid) return '';
    const counts = new Map();

    const pick = (t) => {
      t = String(t || '');
      let m = t.match(/d0010(\d{3})/i) || t.match(/0010(\d{3})/) || t.match(/010(\d{3})/);
      if (m) return m[1];
      m = t.match(/\b(\d{3})\b/);
      if (m) return m[1];
      m = t.match(/0{0,2}(\d{3})\b/);
      return m ? m[1] : '';
    };

    const ths = Array.from(grid.querySelectorAll('thead th,[role="columnheader"]'));
    const iDepot = ths.findIndex(th => /^Depot$/i.test((th.textContent || th.title || '').trim()));
    if (iDepot < 0) return '';

    const rows = Array.from(grid.querySelectorAll('tbody tr,[role="row"]'))
      .filter(r => r.querySelector('td,[role="gridcell"]'))
      .slice(0, 800);

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const raw = (tds[iDepot]?.getAttribute?.('aria-label') ||
                   tds[iDepot]?.getAttribute?.('data-title') ||
                   tds[iDepot]?.querySelector?.('[title]')?.getAttribute('title') ||
                   tds[iDepot]?.innerText || tds[iDepot]?.textContent || '').trim();
      if (!raw) continue;
      const suf = pick(raw);
      if (suf && suf !== '000') counts.set(suf, (counts.get(suf) || 0) + 1);
    }

    if (!counts.size) return '';
    let best = '', bestN = -1;
    for (const [k, n] of counts) {
      if (n > bestN) { best = k; bestN = n; }
    }
    return best;
  }

  function guessDepotFromVehicles() {
    const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
    if (g) {
      const inp = document.getElementById(NS + 'inp-depot');
      if (inp) inp.value = g;
      addEvent({ title: 'Einstellungen', meta: `Depotkennung erkannt: ${g}`, sev: 'info', read: true });
    } else {
      addEvent({ title: 'Einstellungen', meta: 'Depotkennung konnte nicht ermittelt werden.', sev: 'warn', read: true });
    }
  }

  function saveSettingsFromModal() {
    const pass = (document.getElementById(NS + 'inp-pass')?.value || '');
    const dep  = (document.getElementById(NS + 'inp-depot')?.value || '').replace(/\D+/g, '').slice(-3);
    saveSettingsObj({ ...loadSettings(), scanserverPass: pass, depotSuffix: dep });
    addEvent({ title: 'Einstellungen', meta: 'Gespeichert.', sev: 'info', read: true });
    hideModal();
  }

  function getScanserverBase() {
    let suf = String(getSetting('depotSuffix') || '').replace(/\D+/g, '').slice(-3);
    if (!suf) {
      const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
      if (g) { suf = g; setSetting('depotSuffix', g); }
    }
    if (!suf) suf = '157';
    return `https://scanserver-d0010${suf}.ssw.dpdit.de/cgi-bin/pa.cgi`;
  }

  function buildScanserverUrl(psnRaw) {
    const pass = getSetting('scanserverPass') || '';
    if (!pass) {
      addEvent({ title: 'Scanserver', meta: 'Kein Passwort hinterlegt. Bitte in den Einstellungen setzen.', sev: 'warn', read: true });
      openSettingsModal();
      return '';
    }
    let psn = String(psnRaw || '').replace(/\D+/g, '');
    if (psn.length === 13) psn = '0' + psn;

    const base = getScanserverBase();
    const params = new URLSearchParams();
    params.set('_url', 'file');
    params.set('_passwd', pass);
    params.set('_disp', '3');
    params.set('_pivotxx', '0');
    params.set('_rastert', '4');
    params.set('_rasteryt', '0');
    params.set('_rasterx', '0');
    params.set('_rastery', '0');
    params.set('_pivot', '0');
    params.set('_pivotbp', '0');
    params.set('_sortby', 'date|time');
    params.set('_dca', '0');
    params.set('_tabledef', 'psn|date|time|sa|tour|zc|sc|adr1|str|hno|plz1|city|dc|etafrom|etato');
    params.set('_arg59', 'dpd');
    params.set('_arg0a', psn);
    params.set('_arg0b', psn);
    params.set('_arg0', psn + ',' + psn);
    params.set('_csv', '0');
    return `${base}?${params.toString()}`;
  }

  function openScanserver(psn) {
    const url = buildScanserverUrl(psn);
    if (url) window.open(url, '_blank', 'noopener');
  }

  function storeCapturedPickupRequest(urlString, headersMaybe) {
    try {
      const u = new URL(urlString, location.origin);
      if (!u.href.includes('/dispatcher/api/pickup-delivery')) return;

      const q = u.searchParams;
      if (q.get('parcelNumber')) return;

      const h = {};
      const src = headersMaybe || {};

      if (src instanceof Headers) {
        src.forEach((v, k) => h[String(k).toLowerCase()] = String(v));
      } else if (Array.isArray(src)) {
        src.forEach(([k, v]) => h[String(k).toLowerCase()] = String(v));
      } else if (src && typeof src === 'object') {
        Object.entries(src).forEach(([k, v]) => h[String(k).toLowerCase()] = String(v));
      }

      if (!h['authorization']) {
        const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
        if (m) h['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
      }

      lastOkRequest = { url: u, headers: h };

      if (Date.now() > suppressSourceRefreshUntil && !isBusy) {
        queueRefreshFromSource();
      }
    } catch {}
  }

  function queueRefreshFromSource() {
    if (sourceRefreshTimer) clearTimeout(sourceRefreshTimer);
    sourceRefreshTimer = setTimeout(() => {
      sourceRefreshTimer = null;
      if (!document.hidden) fullRefresh(false).catch(console.error);
    }, 700);
  }

  (function hookNetwork() {
    if (!window.__pm_fetch_hooked && window.fetch) {
      const orig = window.fetch;
      window.fetch = async function (input, init = {}) {
        const res = await orig.apply(this, arguments);
        try {
          const uStr = typeof input === 'string' ? input : (input && input.url) || '';
          if (res.ok) storeCapturedPickupRequest(uStr, init?.headers || input?.headers);
        } catch {}
        return res;
      };
      window.__pm_fetch_hooked = true;
    }

    if (!window.__pm_xhr_hooked && window.XMLHttpRequest) {
      const X = window.XMLHttpRequest;
      const open = X.prototype.open, send = X.prototype.send, setH = X.prototype.setRequestHeader;

      X.prototype.open = function (m, u) {
        this.__pm_url = typeof u === 'string' ? new URL(u, location.origin) : null;
        this.__pm_headers = {};
        return open.apply(this, arguments);
      };

      X.prototype.setRequestHeader = function (k, v) {
        try { this.__pm_headers[String(k).toLowerCase()] = String(v); } catch {}
        return setH.apply(this, arguments);
      };

      X.prototype.send = function () {
        const onload = () => {
          try {
            if (this.__pm_url && this.status >= 200 && this.status < 300) {
              storeCapturedPickupRequest(this.__pm_url.href, this.__pm_headers);
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

  const gridIndex = { tour2driver: new Map() };
  const deliveryDetailsCache = new Map();

  function detectTourDriverCols(tbl) {
    const ths = Array.from(tbl.querySelectorAll('thead th,[role="columnheader"]'))
      .map(el => ({ el, txt: norm(el.textContent || el.title || '') }));
    let iTour = -1, iDrv = -1;
    ths.forEach((h, i) => {
      if (iTour < 0 && /\bTour(\s*nr|nummer)?\b/i.test(h.txt)) iTour = i;
      if (iDrv < 0 && /(Zusteller(\s*name)?|Fahrer)/i.test(h.txt)) iDrv = i;
    });
    return { iTour, iDrv };
  }

  function collectTourDriverFromTable(tbl, map) {
    const { iTour, iDrv } = detectTourDriverCols(tbl);
    if (iTour < 0 || iDrv < 0) return 0;
    const rows = Array.from(tbl.querySelectorAll('tbody tr,[role="row"]'))
      .filter(r => r.querySelector('td,[role="gridcell"]'));
    let added = 0;

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const get = i => norm(
        tds[i]?.getAttribute?.('aria-label') ||
        tds[i]?.getAttribute?.('data-title') ||
        tds[i]?.querySelector?.('[title]')?.getAttribute('title') ||
        tds[i]?.innerText || tds[i]?.textContent || ''
      );
      const tour = tourKey(get(iTour));
      const drv  = get(iDrv);
      if (tour && drv && !map.has(tour)) { map.set(tour, drv); added++; }
    }
    return added;
  }

  async function buildTourDriverMap() {
    try {
      const map = new Map();
      Array.from(document.querySelectorAll('table,[role="grid"]'))
        .filter(el => el.offsetParent !== null)
        .forEach(tbl => collectTourDriverFromTable(tbl, map));
      if (map.size) {
        gridIndex.tour2driver = map;
        window.__pmTour2Driver = map;
      }
    } catch {}
  }

  function getJwtAuthHeader() {
    const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
    return m ? 'Bearer ' + decodeURIComponent(m[1]) : '';
  }

  function buildHeaders(h) {
    const H = new Headers();
    try {
      if (h) {
        Object.entries(h).forEach(([k, v]) => {
          const key = k.toLowerCase();
          if (['authorization', 'accept', 'x-xsrf-token', 'x-csrf-token'].includes(key)) {
            H.set(key === 'accept' ? 'Accept' : key.replace(/(^.|-.)/g, s => s.toUpperCase()), v);
          }
        });
      }
      if (!H.has('Authorization')) {
        const auth = getJwtAuthHeader();
        if (auth) H.set('Authorization', auth);
      }
      if (!H.has('Accept')) H.set('Accept', 'application/json, text/plain, */*');
    } catch {}
    return H;
  }

  function buildBasePickupUrlFromToday() {
    const d = new Date();
    const ds = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const u = new URL(location.origin + '/dispatcher/api/pickup-delivery');
    u.searchParams.set('page', '1');
    u.searchParams.set('pageSize', '500');
    u.searchParams.set('dateFrom', ds);
    u.searchParams.set('dateTo', ds);
    u.searchParams.set('active', 'true');
    return u;
  }

  function buildUrlPrio(base, page) {
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize', '500');
    q.set('priority', 'prio');
    q.delete('elements');
    q.delete('parcelNumber');
    return u;
  }

  function buildUrlElements(base, page, el) {
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize', '500');
    q.set('elements', String(el));
    q.delete('priority');
    q.delete('parcelNumber');
    return u;
  }

  function pickArray(payload) {
    if (Array.isArray(payload)) return payload;
    if (payload && Array.isArray(payload.items)) return payload.items;
    if (payload && Array.isArray(payload.content)) return payload.content;
    if (payload && payload.data) {
      if (Array.isArray(payload.data)) return payload.data;
      if (Array.isArray(payload.data.items)) return payload.data.items;
      if (Array.isArray(payload.data.content)) return payload.data.content;
    }
    if (payload && Array.isArray(payload.results)) return payload.results;
    if (payload && payload._embedded) {
      const v = Object.values(payload._embedded).find(Array.isArray);
      if (Array.isArray(v)) return v;
    }
    return [];
  }

  function pickTotals(payload, sizeFallback) {
    const p = payload || {};
    const pg = p.page || {};
    const totalElements = Number(p.totalElements ?? p.total ?? p.count ?? pg.totalElements ?? pg.total ?? 0);
    const totalPages = Number(p.totalPages ?? pg.totalPages ?? (totalElements ? Math.ceil(totalElements / (sizeFallback || 500)) : 0));
    return {
      totalElements: Number.isFinite(totalElements) ? totalElements : 0,
      totalPages: Number.isFinite(totalPages) ? totalPages : 0
    };
  }

  async function fetchPagedFast(builder, { concurrency = 6, size = 500, hardMaxPages = 200 } = {}) {
    const baseUrl = lastOkRequest?.url ? new URL(lastOkRequest.url.href) : buildBasePickupUrlFromToday();
    const headers = buildHeaders(lastOkRequest?.headers || {});
    suppressSourceRefreshUntil = Date.now() + 5000;

    const u1 = builder(baseUrl, 1);
    const r1 = await fetch(u1.toString(), { credentials: 'include', headers });
    if (!r1.ok) throw new Error(`HTTP ${r1.status}`);
    const j1 = await r1.json();
    const chunk1 = pickArray(j1);

    const { totalPages: tpRaw } = pickTotals(j1, size);
    let totalPages = tpRaw || 0;

    if (totalPages > 1) {
      totalPages = Math.min(totalPages, hardMaxPages);
      const pages = [];
      for (let p = 2; p <= totalPages; p++) pages.push(p);

      const out = chunk1.slice();
      let idx = 0;

      async function worker() {
        while (idx < pages.length) {
          const p = pages[idx++];
          const u = builder(baseUrl, p);
          const res = await fetch(u.toString(), { credentials: 'include', headers });
          if (!res.ok) continue;
          const j = await res.json();
          const arr = pickArray(j);
          if (arr.length) out.push(...arr);
          if (arr.length < size) { idx = pages.length; break; }
        }
      }

      const workers = Array.from({ length: Math.max(1, Math.min(concurrency, pages.length)) }, worker);
      await Promise.all(workers);
      return out;
    }

    if (chunk1.length < size) return chunk1;

    const out = chunk1.slice();
    let page = 2;
    while (page <= hardMaxPages) {
      const u = builder(baseUrl, page);
      const r = await fetch(u.toString(), { credentials: 'include', headers });
      if (!r.ok) break;
      const j = await r.json();
      const arr = pickArray(j);
      if (!arr.length) break;
      out.push(...arr);
      if (arr.length < size) break;
      page++;
    }
    return out;
  }

  const parcelId = r => r.__pidOverride || r.parcelNumber || (Array.isArray(r.parcelNumbers) && r.parcelNumbers[0]) || r.id || '';
  const addrOf = r => [r.street, r.houseno].filter(Boolean).join(' ');
  const placeOf = r => [r.postalCode, r.city].filter(Boolean).join(' ');
  const addCodes = r => Array.isArray(r.additionalCodes) ? r.additionalCodes.map(String) : [];
  const isDelivery = r => String(r?.orderType || '').toUpperCase() === 'DELIVERY';
  const isPRIO = r => String(r?.priority || r?.prio || '').toUpperCase() === 'PRIO';

  const composeDateTime = (dateStr, timeStr) => {
    if (!dateStr || !timeStr) return null;
    const s = `${dateStr}T${String(timeStr).slice(0, 8)}`;
    const d = new Date(s);
    return isNaN(d) ? null : d;
  };

  const fromTime = r => r.from2 ? composeDateTime(r.date, r.from2) : null;
  const toTime = r => r.to2 ? composeDateTime(r.date, r.to2) : null;
  const deliveredTime = r => r.deliveredTime ? new Date(r.deliveredTime) : null;

  function apiStatus(r) {
    const v =
      r?.statusDisplay || r?.statusLabel || r?.statusDescription ||
      r?.statusName || r?.statusText || r?.stateText ||
      r?.deliveryStatus || r?.parcelStatus || '';
    return String(v || '').trim();
  }

  const statusOf = r => apiStatus(r);

  const delivered = r => {
    if (r.deliveredTime) return true;
    const s = (statusOf(r) || '').toUpperCase();
    return /ZUGESTELLT|DELIVERED/.test(s);
  };

  const tourOf = r => r.tour ? String(r.tour) : '';

  function driverOf(r) {
    const direct = r.driverName || r.driver || r.courierName || r.riderName || r.tourDriver || '';
    if (direct && direct.trim()) return direct.trim();
    const key = tourKey(tourOf(r) || '');
    const viaGrid =
      (gridIndex.tour2driver && gridIndex.tour2driver.get(key)) ||
      (window.__pmTour2Driver instanceof Map ? window.__pmTour2Driver.get(key) : '');
    return (viaGrid || '—').trim() || '—';
  }

  function partnerOfTour(tour) {
    const k = tourKey(tour || '');
    return tourPartnerMap.get(k) || '—';
  }

  function formatHHMM(d) {
    return d ? d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '';
  }

  function buildPredictMeta(r) {
    const pf = fromTime(r);
    const pt = toTime(r);
    const pfTs = pf ? +pf : 0;
    const ptTs = pt ? +pt : 0;
    let range = '—';
    if (pf && pt) range = `${formatHHMM(pf)} - ${formatHHMM(pt)}`;
    else if (pf) range = formatHHMM(pf);
    else if (pt) range = formatHHMM(pt);
    return { pfTs, ptTs, range };
  }

  function serviceCodesOf(r) {
    if (!r) return [];
    const set = new Set();

    const addFromVal = v => {
      if (v == null) return;
      String(v)
        .split(/[^\dA-Za-z]+/)
        .map(s => s.trim())
        .filter(Boolean)
        .forEach(code => set.add(code));
    };

    const addFromArr = arr => {
      if (!Array.isArray(arr)) return;
      arr.forEach(addFromVal);
    };

    addFromVal(r.serviceCode);
    addFromVal(r.servicecode);
    addFromVal(r.service_code);
    addFromArr(r.serviceCodes);

    if (r.service && typeof r.service === 'object') {
      addFromVal(r.service.code);
      addFromVal(r.service.serviceCode);
      addFromVal(r.service.id);
      addFromArr(r.service.serviceCodes);
    }

    if (r.product && typeof r.product === 'object') {
      addFromVal(r.product.serviceCode);
      addFromVal(r.product.code);
      addFromVal(r.product.id);
      addFromArr(r.product.serviceCodes);
    }

    const arr = Array.from(set);
    arr.sort((a, b) => collator.compare(a, b));
    return arr;
  }

  const EXPRESS12_CODES = new Set([
    '104','107','135','196','210','225','226','227','228','231','232','234','237','238','239','240',
    '243','245','247','249','255','261','262','267','269','286','310','311','323','379','412','414',
    '452','453','458','459','488','490','503','505','530','531','537','538','542','547','567',
    '786','797','811'
  ]);

  const EXPRESS18_CODES = new Set([
    '155','157','158','161','163','164','166','168','171','174',
    '219','224','230','236','265','318','324','378','419','422','423',
    '483','499','500','511','512','535','562','564','662','664',
    '787','799','812'
  ]);

  const EXPRESS_SERVICE_WHITELIST = new Set([...EXPRESS12_CODES, ...EXPRESS18_CODES]);

  function rowHasExpress12BySvc(r) {
    const arr = r.__serviceCodes || serviceCodesOf(r);
    return arr.some(c => EXPRESS12_CODES.has(String(c)));
  }

  function rowHasExpress18BySvc(r) {
    const arr = r.__serviceCodes || serviceCodesOf(r);
    return arr.some(c => EXPRESS18_CODES.has(String(c)));
  }

  function statusClass(text) {
    const t = String(text || '').toUpperCase();
    if (/PROBLEM|FAIL|NICHT/.test(t)) return NS + 'badge-status-problem';
    if (/ZUGESTELLT|DELIVERED/.test(t)) return NS + 'badge-status-ok';
    if (/ZUSTELLUNG|OUT_FOR_DELIVERY|IN_DELIVERY/.test(t)) return NS + 'badge-status-run';
    return '';
  }

  function normRow(r) {
    const pid = parcelId(r) || '';
    const { pfTs, ptTs, range } = buildPredictMeta(r);
    const svcArr = serviceCodesOf(r);
    const isExpSvc = svcArr.some(c => EXPRESS_SERVICE_WHITELIST.has(String(c)));
    const isExp12 = svcArr.some(c => EXPRESS12_CODES.has(String(c)));
    const isExp18 = !isExp12 && svcArr.some(c => EXPRESS18_CODES.has(String(c)));
    const expType = isExp12 ? '12' : (isExp18 ? '18' : '');
    const isDel = delivered(r);

    let highlightLate12 = false;
    if (isExp12 && !isDel && (pfTs || ptTs) && r.date) {
      const cut = new Date(`${r.date}T12:00:00`);
      if (!isNaN(cut)) {
        const cutTs = +cut;
        if ((pfTs && pfTs > cutTs) || (ptTs && ptTs > cutTs)) highlightLate12 = true;
      }
    }

    const t = tourOf(r) || '';
    const sysPartner = partnerOfTour(t);

    return {
      ...r,
      __pid: pid,
      __addr: [addrOf(r), placeOf(r)].filter(Boolean).join(' · ') || '—',
      __driver: driverOf(r),
      __tourNum: Number(tourOf(r) || 0),
      __systempartner: sysPartner,
      __status: statusOf(r) || '',
      __delivTs: deliveredTime(r) ? deliveredTime(r).getTime() : 0,
      __predFromTs: pfTs,
      __predToTs: ptTs,
      __predRangeStr: range,
      __codesStr: (addCodes(r) || []).join(', ') || '—',
      __expType: expType,
      __serviceCode: svcArr[0] || '',
      __serviceCodes: svcArr,
      __highlightLatePredict12: highlightLate12,
      __isExpressSvc: isExpSvc
    };
  }

  function expandAndNorm(rows) {
    const out = [];
    for (const r of rows) {
      const list = Array.isArray(r.parcelNumbers) && r.parcelNumbers.length ? r.parcelNumbers : [parcelId(r)];
      const seen = new Set();
      for (const raw of list) {
        const psn = String(raw || '').replace(/\D+/g, '');
        if (!psn || seen.has(psn)) continue;
        seen.add(psn);
        const rr = { ...r, __pidOverride: psn.length === 13 ? '0' + psn : psn };
        out.push(normRow(rr));
      }
    }
    return out;
  }

  function buildHeaderHtml(selectable = false) {
    const base = ['Paketscheinnummer','Adresse','Fahrer','Tour','Systempartner','Status','Zustellzeit','Zusatzcode','Servicecode','Predict'];
    const ths = selectable ? ['✓', ...base] : base;
    return `<tr>${ths.map((h, i) => `<th data-col="${i}">${h}</th>`).join('')}</tr>`;
  }

  function buildTableShell(selectable = false) {
    return `
      <div id="${NS}vt-wrap" style="position:relative;height:min(70vh,720px);overflow:auto">
        <table class="${NS}tbl">
          <thead>${buildHeaderHtml(selectable)}</thead>
          <tbody id="${NS}vt-body"></tbody>
        </table>
      </div>`;
  }

  function rowHtml(r, selectable = false) {
    const selKey = String(r.stopId ?? r.id ?? r.__pid ?? '');
    const checked = selectable && state._modal.selected?.has(selKey) ? 'checked' : '';
    const selCell = selectable ? `<input type="checkbox" data-sel="1" data-key="${esc(selKey)}" ${checked} />` : null;
    const pkgCount = Number(r.__pkgCount || 1);
    const psnLabel = (r.__pid && pkgCount > 1) ? `${r.__pid} (+${pkgCount - 1})` : (r.__pid || '—');

    const pLink = r.__pid
      ? `<a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${r.__pid}" target="_blank" rel="noopener">${esc(psnLabel)}</a>`
      : '—';

    const eye = r.__pid ? `<button class="${NS}eye" title="Scanserver öffnen" data-psn="${esc(r.__pid)}">👁</button>` : '';
    const dtime = r.__delivTs ? formatHHMM(new Date(r.__delivTs)) : '—';

    const statusText = r.__status || '';
    const statusCls = statusClass(statusText);
    const statusCell = statusText ? `<span class="${NS}badge ${statusCls}">${esc(statusText)}</span>` : '';

    const serviceBadges = (r.__serviceCodes && r.__serviceCodes.length)
      ? r.__serviceCodes.map(c => `<span class="${NS}badge">${esc(c)}</span>`).join(' ')
      : '';

    const pred = r.__predRangeStr || '—';

    const cells = [
      `${eye}${pLink}`,
      esc(r.__addr),
      esc(r.__driver),
      esc(String(r.__tourNum || '—')),
      esc(String(r.__systempartner || '—')),
      statusCell,
      esc(dtime),
      esc(r.__codesStr),
      serviceBadges,
      esc(pred)
    ];

    const finalCells = selectable ? [selCell, ...cells] : cells;
    const dataAttrs = [];
    if (r.id != null) dataAttrs.push(`data-delivery="${esc(String(r.id))}"`);
    if (r.stopId != null) dataAttrs.push(`data-stop="${esc(String(r.stopId))}"`);

    return `<tr ${dataAttrs.join(' ')}>${finalCells.map(v => `<td>${v}</td>`).join('')}</tr>`;
  }

  function openModal(title, rowsOrHtml) {
    const m = document.getElementById(NS + 'modal');
    const t = document.getElementById(NS + 'modal-title');
    const b = document.getElementById(NS + 'modal-body');
    if (t) t.textContent = title || '';

    if (Array.isArray(rowsOrHtml)) {
      const rows = rowsOrHtml.slice();
      const selectable = /falsch einsortiert/i.test(title || '');

      state._modal = {
        rows,
        opts: { showPredict: true, selectable },
        title: title || '',
        selected: state._modal.selected instanceof Set ? state._modal.selected : new Set()
      };

      if (b) b.innerHTML = buildTableShell(selectable);

      const tbody = document.getElementById(NS + 'vt-body');
      const wrap = document.getElementById(NS + 'vt-wrap');

      function renderAll() {
        if (!tbody) return;
        tbody.innerHTML = rows.map(r => rowHtml(r, selectable)).join('');
      }

      const mailBtn = document.getElementById(NS + 'mail-selected');
      if (mailBtn) mailBtn.style.display = selectable ? '' : 'none';

      if (tbody && selectable) {
        tbody.addEventListener('change', ev => {
          const cb = ev.target.closest('input[type="checkbox"][data-sel="1"]');
          if (!cb) return;
          const key = String(cb.dataset.key || '');
          if (!key) return;
          if (cb.checked) state._modal.selected.add(key);
          else state._modal.selected.delete(key);
        }, { passive: true });
      }

      const thead = wrap?.querySelector('thead');
      if (thead) {
        thead.addEventListener('click', ev => {
          const th = ev.target.closest('th');
          if (!th) return;
          const col = Number(th.dataset.col || 0);

          Array.from(thead.querySelectorAll('th')).forEach(x => x.classList.remove(NS + 'sort-asc', NS + 'sort-desc'));
          const asc = !(th.dataset.dir === 'asc');
          th.dataset.dir = asc ? 'asc' : 'desc';
          th.classList.add(asc ? NS + 'sort-asc' : NS + 'sort-desc');

          const offset = selectable ? 1 : 0;
          const getKey = r => {
            switch (col - offset) {
              case 0: return r.__pid;
              case 1: return r.__addr;
              case 2: return r.__driver;
              case 3: return r.__tourNum;
              case 4: return r.__systempartner || '';
              case 5: return r.__status || '';
              case 6: return r.__delivTs;
              case 7: return r.__codesStr;
              case 8: return r.__serviceCode || '';
              case 9: return r.__predFromTs || 0;
              default: return '';
            }
          };

          rows.sort((a, b) => {
            const A = getKey(a), B = getKey(b);
            if (typeof A === 'number' && typeof B === 'number') return asc ? (A - B) : (B - A);
            return asc ? collator.compare(String(A), String(B)) : collator.compare(String(B), String(A));
          });

          state._modal.rows = rows;
          renderAll();
        }, { passive: true });
      }

      renderAll();

      if (tbody) {
        tbody.addEventListener('click', ev => {
          const tr = ev.target.closest('tr');
          if (!tr) return;
          if (ev.target.closest('.' + NS + 'eye')) return;
          toggleStopDetailInline(tr);
        }, { passive: true });
      }
    } else {
      if (b) b.innerHTML = rowsOrHtml || '';
    }

    if (m) m.style.display = 'flex';
  }

  function hideModal() {
    const m = document.getElementById(NS + 'modal');
    if (m) m.style.display = 'none';
  }

  function setLoading(on) {
    isLoading = !!on;
    const el = document.getElementById(NS + 'loading');
    if (el) el.classList.toggle('on', on);
  }

  function dimButtons(on) {
    document.querySelectorAll('.' + NS + 'btn-sm').forEach(b => b.classList.toggle(NS + 'dim', !!on));
  }

  function render() {
    const list = document.getElementById(NS + 'list');
    if (!list) return;
    list.innerHTML = '';
    const d = document.createElement('div');
    d.className = NS + 'empty';
    d.textContent = isLoading ? 'Lade Daten …' : 'Daten geladen.';
    list.appendChild(d);
    setLastRefreshNow();
  }

  function setKpis({ prioAll, prioOpen, expAll, expOpen }) {
    const set = (id, val, chipId) => {
      const el = document.getElementById(id);
      const chip = document.getElementById(chipId);
      if (el) el.textContent = String(Number(val || 0));
      if (chip) chip.style.display = Number(val || 0) === 0 ? 'none' : '';
    };
    set(NS + 'kpi-prio-all', prioAll, NS + 'chip-prio-all');
    set(NS + 'kpi-prio-open', prioOpen, NS + 'chip-prio-open');
    set(NS + 'kpi-exp-all', expAll, NS + 'chip-exp-all');
    set(NS + 'kpi-exp-open', expOpen, NS + 'chip-exp-open');
  }

  function addEvent(ev) {
    state.events.push({
      id: state.nextId++,
      title: ev.title || 'Ereignis',
      meta: ev.meta || '',
      sev: ev.sev || 'info',
      ts: ev.ts || Date.now(),
      read: !!ev.read
    });
  }

  function buildTableRowsAndCounts(prioRows, exp12Rows, exp18Rows) {
    const prioDeliveries = prioRows.filter(isDelivery).filter(isPRIO);
    const prioAll = prioDeliveries;
    const prioOpen = prioDeliveries.filter(r => !delivered(r));

    const expRows = [...exp12Rows, ...exp18Rows].filter(r => r.__isExpressSvc);

    const seen = new Set();
    const expDeliveries = expRows
      .filter(isDelivery)
      .filter(r => {
        const id = parcelId(r);
        if (!id || seen.has(id)) return false;
        seen.add(id);
        return true;
      });

    const expAll = expDeliveries;
    const expOpen = expDeliveries.filter(r => !delivered(r));

    const expLate11 = expDeliveries
      .filter(rowHasExpress12BySvc)
      .filter(r => {
        const ft = fromTime(r);
        if (!ft) return false;
        return (ft.getHours() > 11) || (ft.getHours() === 11 && ft.getMinutes() >= 1);
      });

    return { prioAll, prioOpen, expAll, expOpen, expLate11 };
  }

  function filterByExpressSelection(rows) {
    if (state.filterExpress === '12') return rows.filter(rowHasExpress12BySvc);
    if (state.filterExpress === '18') return rows.filter(rowHasExpress18BySvc);
    return rows;
  }

  function getFilteredExpressCounts() {
    const filt = state.filterExpress === '12'
      ? rowHasExpress12BySvc
      : state.filterExpress === '18'
      ? rowHasExpress18BySvc
      : null;

    const expAllList = filt ? state._expAllList.filter(filt) : state._expAllList;
    const expOpenList = filt ? state._expOpenList.filter(filt) : state._expOpenList;
    return { expAllCount: expAllList.length, expOpenCount: expOpenList.length };
  }

  function updateKpisForCurrentState() {
    const { expAllCount, expOpenCount } = getFilteredExpressCounts();
    setKpis({
      prioAll: state._prioAllList.length,
      prioOpen: state._prioOpenList.length,
      expAll: expAllCount,
      expOpen: expOpenCount
    });
  }

  function groupRowsByStop(rows) {
    const map = new Map();
    for (const r of rows) {
      const key =
        (r.stopId != null ? String(r.stopId) :
         (r.id != null ? String(r.id) :
          `${r.__addr}#${r.__tourNum || ''}`));

      let g = map.get(key);
      if (!g) {
        g = { ...r };
        g.__pkgCount = 1;
        g.__deliveryId = r.id != null ? r.id : null;
        map.set(key, g);
      } else {
        g.__pkgCount++;
      }
    }
    return Array.from(map.values());
  }

  async function toggleStopDetailInline(tr) {
    if (!tr || !(lastOkRequest || getJwtAuthHeader())) return;
    const delId = tr.getAttribute('data-delivery');
    if (!delId) return;

    const tbody = tr.parentNode;
    if (!tbody) return;

    const next = tr.nextElementSibling;
    if (next && next.classList.contains(NS + 'detail-row')) {
      next.remove();
      return;
    }

    const detailRow = document.createElement('tr');
    detailRow.className = NS + 'detail-row';
    const colSpan = tr.children.length || 1;
    const td = document.createElement('td');
    td.colSpan = colSpan;
    td.textContent = 'Lade Stopp-Details …';
    detailRow.appendChild(td);
    tbody.insertBefore(detailRow, tr.nextSibling);

    try {
      let detail = deliveryDetailsCache.get(delId);
      if (!detail) {
        const headers = buildHeaders(lastOkRequest?.headers || {});
        const origin = (lastOkRequest?.url?.origin) || location.origin;
        const url = `${origin}/dispatcher/api/delivery/${encodeURIComponent(delId)}`;
        const res = await fetch(url, { credentials: 'include', headers });
        if (!res.ok) {
          td.textContent = `Fehler beim Laden (HTTP ${res.status})`;
          return;
        }
        detail = await res.json();
        deliveryDetailsCache.set(delId, detail);
      }

      const parcels = Array.isArray(detail.parcels) ? detail.parcels : [];
      const rowsHtml = parcels.map(p => {
        const svc = p.serviceCode || '';
        const isExpSvc = EXPRESS_SERVICE_WHITELIST.has(String(svc));
        const els = Array.isArray(p.elements) ? p.elements.join(', ') : (p.elements || '');
        const prio = p.priority || '';
        let psn = String(p.parcelNumber || '').replace(/\D+/g, '');
        if (psn.length === 13) psn = '0' + psn;

        const psnCell = psn
          ? `<button class="${NS}eye" title="Scanserver öffnen" data-psn="${esc(psn)}">👁</button><a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${esc(psn)}" target="_blank" rel="noopener">${esc(psn)}</a>`
          : '—';

        return `
          <tr class="${isExpSvc ? NS + 'row-express' : ''}">
            <td>${psnCell}</td>
            <td>${esc(svc || '—')}</td>
            <td>${esc(prio || '—')}</td>
            <td>${esc(els || '—')}</td>
            <td>${isExpSvc ? 'EXPRESS' : 'Normal'}</td>
          </tr>`;
      }).join('');

      const addr = [
        detail.addressStreet,
        detail.addressHouseno,
        detail.addressPcode,
        detail.addressCity
      ].filter(Boolean).join(' ');

      td.innerHTML = `
        <div class="${NS}detail-inner">
          <div style="margin-bottom:4px;">
            <b>Adresse:</b> ${esc(addr || '—')} ·
            <b>Tour:</b> ${esc(detail.tour || '—')} ·
            <b>Pakete:</b> ${parcels.length}
          </div>
          <table>
            <thead>
              <tr>
                <th>PSN</th>
                <th>Servicecode</th>
                <th>Priority</th>
                <th>Elements</th>
                <th>Typ</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml || '<tr><td colspan="5">Keine Pakete gefunden.</td></tr>'}
            </tbody>
          </table>
        </div>`;
    } catch {
      td.textContent = 'Fehler beim Laden der Stopp-Details.';
    }
  }

  function showPrioAll() {
    openModal(`PRIO – in Ausrollung (alle) · ${state._prioAllList.length}`, state._prioAllList);
  }

  function showPrioOpen() {
    openModal(`PRIO – noch nicht zugestellt · ${state._prioOpenList.length}`, state._prioOpenList);
  }

  function showExpAll() {
    const rows = filterByExpressSelection(state._expAllList);
    const grouped = groupRowsByStop(rows);
    const sel = state.filterExpress === '12' ? ' (12)' : state.filterExpress === '18' ? ' (18)' : '';
    openModal(`Express${sel} – in Ausrollung (alle) · ${grouped.length}`, grouped);
  }

  function showExpOpen() {
    const rows = filterByExpressSelection(state._expOpenList);
    const grouped = groupRowsByStop(rows);
    const sel = state.filterExpress === '12' ? ' (12)' : state.filterExpress === '18' ? ' (18)' : '';
    openModal(`Express${sel} – noch nicht zugestellt · ${grouped.length}`, grouped);
  }

  function showExpLate11() {
    const rows = state._expLate11List.slice();
    const grouped = groupRowsByStop(rows);
    openModal(`Express 12 – falsch einsortiert (>11:01 geplant) · ${grouped.length}`, grouped);
  }

  async function refreshViaApi_SAFE() {
    const [prioRows, exp12Rows, exp18Rows] = await Promise.all([
      fetchPagedFast(buildUrlPrio),
      fetchPagedFast((b, p) => buildUrlElements(b, p, '023')),
      fetchPagedFast((b, p) => buildUrlElements(b, p, '010'))
    ]);

    const prioN = expandAndNorm(prioRows);
    const exp12N = expandAndNorm(exp12Rows);
    const exp18N = expandAndNorm(exp18Rows);

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

    addEvent({
      title: 'Aktualisiert',
      meta: `PRIO: ${prioAll.length}/${prioOpen.length} offen · EXPRESS: ${expAll.length}/${expOpen.length} offen · >11:01: ${expLate11.length}`,
      sev: 'info',
      read: true,
      ts: Date.now()
    });
  }

  function findPickupDeliveryTrigger() {
    const selectors = ['[role="tab"]', '.mat-mdc-tab', '.mat-tab-label', '.mat-mdc-tab-link', '.mat-tab-link', 'button', 'a', 'div'];
    for (const sel of selectors) {
      const nodes = Array.from(document.querySelectorAll(sel));
      for (const el of nodes) {
        const txt = norm(el.textContent || '');
        if (!/abholung.*zustellung|zustellung.*abholung/i.test(txt)) continue;
        return el.closest('[role="tab"], .mat-mdc-tab, .mat-tab-label, .mat-mdc-tab-link, .mat-tab-link, button, a') || el;
      }
    }
    return null;
  }

  function isLikelyActiveTab(el) {
    if (!el) return false;
    const aria = String(el.getAttribute?.('aria-selected') || '').toLowerCase();
    const cls = String(el.className || '');
    if (aria === 'true') return true;
    if (/active|selected|mdc-tab--active|mat-mdc-tab-active|mat-tab-label-active/i.test(cls)) return true;
    try {
      const cs = getComputedStyle(el);
      const bg = String(cs.backgroundColor || '');
      const color = String(cs.color || '');
      if (bg === 'rgb(225, 6, 50)' || bg === 'rgb(229, 0, 54)' || color === 'rgb(255, 255, 255)') return true;
    } catch {}
    return false;
  }

  function fireRealClick(el) {
    if (!el) return false;
    try { el.scrollIntoView({ block: 'center', inline: 'center' }); } catch {}
    const evOpts = { bubbles: true, cancelable: true, view: window };
    try { el.dispatchEvent(new PointerEvent('pointerdown', evOpts)); } catch {}
    try { el.dispatchEvent(new MouseEvent('mousedown', evOpts)); } catch {}
    try { el.dispatchEvent(new PointerEvent('pointerup', evOpts)); } catch {}
    try { el.dispatchEvent(new MouseEvent('mouseup', evOpts)); } catch {}
    try { el.dispatchEvent(new MouseEvent('click', evOpts)); } catch {}
    try { if (typeof el.click === 'function') el.click(); } catch {}
    return true;
  }

  async function ensurePickupDeliveryActive() {
    const trigger = findPickupDeliveryTrigger();
    if (!trigger) return false;
    if (isLikelyActiveTab(trigger)) return true;
    fireRealClick(trigger);
    await sleep(1200);
    return true;
  }

  async function waitForPickupRequest(timeoutMs = 12000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (lastOkRequest?.url?.href?.includes('/dispatcher/api/pickup-delivery')) return true;
      await sleep(250);
    }
    return false;
  }

  async function ensurePickupDeliverySourceReady() {
    if (lastOkRequest?.url?.href?.includes('/dispatcher/api/pickup-delivery')) return true;
    await ensurePickupDeliveryActive();
    const ok = await waitForPickupRequest(12000);
    if (ok) return true;

    const auth = getJwtAuthHeader();
    if (auth) {
      lastOkRequest = { url: buildBasePickupUrlFromToday(), headers: { authorization: auth } };
      return true;
    }
    return false;
  }

  async function initialOpenRefresh() {
    await ensurePickupDeliverySourceReady();
    await fullRefresh(false);
  }

  async function fullRefresh(manual) {
    if (isBusy) return;
    try {
      isBusy = true;
      setLoading(true);
      dimButtons(true);

      const ready = await ensurePickupDeliverySourceReady();
      if (!ready) {
        addEvent({
          title: 'Hinweis',
          meta: 'Abholung und Zustellung konnte nicht initialisiert werden.',
          sev: 'warn',
          read: true
        });
        return;
      }

      await Promise.all([
        buildTourDriverMap(),
        loadTourPartnerMap()
      ]);

      await refreshViaApi_SAFE();
      state.lastRefreshAt = Date.now();
      render();
    } catch (e) {
      console.error(e);
      addEvent({ title: 'Fehler', meta: String(e?.message || e), sev: 'warn', read: true });
    } finally {
      setLoading(false);
      dimButtons(false);
      isBusy = false;
      render();
    }
  }

  async function mailSelectedLate11() {
    const modal = state._modal || {};
    const rows = Array.isArray(modal.rows) ? modal.rows : [];
    const selectedKeys = modal.selected instanceof Set ? modal.selected : new Set();
    if (!rows.length) { alert('Keine Daten.'); return; }
    if (selectedKeys.size === 0) { alert('Keine Zeilen markiert.'); return; }

    const selected = rows.filter(r => {
      const key = String(r.stopId ?? r.id ?? r.__pid ?? '');
      return selectedKeys.has(key);
    });

    const byPartner = new Map();
    for (const r of selected) {
      const p = norm(r.__systempartner || '—');
      if (!byPartner.has(p)) byPartner.set(p, []);
      byPartner.get(p).push(r);
    }

    for (const [partner, list] of byPartner) {
      if (!partner || partner === '—') {
        alert('Mindestens eine Auswahl hat keinen Systempartner.');
        continue;
      }

      const rec = await getPartnerMailRecord(partner);
      const toRaw = rec?.to || '';
      const ccRaw = rec?.cc || '';
      const alias = rec?.alias || partner;

      const toL = normalizeEmailList(toRaw);
      const ccL = normalizeEmailList(ccRaw);

      if (toL.valid.length === 0) {
        alert(`Für Systempartner "${partner}" ist keine gültige E-Mail-Adresse hinterlegt.`);
        continue;
      }

      const subject = `Express 12 – falsch einsortiert (>11:01) – ${alias} – ${new Date().toLocaleDateString('de-DE')}`;
      const html = buildLate11MailHtml(alias, list);
      const ok = await copyHtmlToClipboard(html);
      openMailto(subject, toL.valid.join(','), ccL.valid.join(','));
      if (ok) {
        alert(`Mail-Entwurf für "${alias}" geöffnet.\nHTML ist in der Zwischenablage – im Mail-Body STRG+V.`);
      } else {
        alert(`Mail-Entwurf für "${alias}" geöffnet.\nKopieren fehlgeschlagen – bitte Tabelle manuell kopieren.`);
      }
    }
  }

  function buildLate11MailHtml(partner, rows) {
    const escH = s => String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    const stamp = new Date().toLocaleString('de-DE');

    const bodyRows = rows.map(r => {
      const psn = escH(r.__pid || '—');
      const tour = escH(String(r.__tourNum || r.tour || '—'));
      const addr = escH(r.__addr || '—');
      const driver = escH(r.__driver || '—');
      const pred = escH(r.__predRangeStr || '—');
      const svc = escH((r.__serviceCodes || []).join(' ') || r.__serviceCode || '—');
      const status = escH(r.__status || '—');
      return `
        <tr>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${psn}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${tour}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${addr}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${driver}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${status}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${svc}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${pred}</td>
        </tr>`;
    }).join('');

    return `
      <div style="font:13px/1.45 -apple-system,Segoe UI,Arial,sans-serif;color:#111;">
        <div style="margin:0 0 10px 0;color:#334155">
          <b>${escH(partner)}</b> – Express 12 „falsch einsortiert“ (geplant > 11:01)<br/>
          Stand: ${escH(stamp)}
        </div>
        <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;min-width:700px;">
          <thead>
            <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">PSN</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Tour</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Adresse</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Fahrer</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Status</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Service</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Predict</th>
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>`;
  }

  function boot() {
    mountUI();
    setTimeout(() => {
      if (!(getSetting('depotSuffix') || '')) {
        const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
        if (g) {
          setSetting('depotSuffix', g);
          addEvent({ title: 'Einstellungen', meta: `Depotkennung aus Fahrzeugübersicht: ${g}`, sev: 'info', read: true });
        }
      }
    }, 1200);

    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && document.getElementById(NS + 'panel') && getComputedStyle(document.getElementById(NS + 'panel')).display !== 'none') {
        queueRefreshFromSource();
      }
    });
  }

})();
