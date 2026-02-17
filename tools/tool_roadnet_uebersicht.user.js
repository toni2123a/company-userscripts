// ==UserScript==
// @name         Roadnet – Transporteinheiten (Zusammenfassung + LTS Lebenslauf Fix + Frachtführer)
// @namespace    bodo.dpd.custom
// @version      1.8.2
// @description  Roadnet transport_units: Beladung/Entladung automatisch. Panel mit Sortierung, Badges, Zebra, Scroll-Memory, Auto-Fit. Copy formatiert. Optionaler Bridge-Export an lokales DPD-Dashboard. LTS# Klick: öffnet LTS auf /index.aspx, wartet auf Login/Session und springt dann robust auf /(S(...))/WBLebenslauf.aspx. Nummer wird persistent per postMessage/window.name übergeben. LTS füllt txtWBNR1/txtWBNR4 und triggert „Suchen“ robust (requestSubmit + click + submit), max. 3 Versuche, stoppt sobald Ergebnis-Tabelle sichtbar ist. LTS-Hinweis: fehlende/unsichtbare Lebenslauf-Spaltenköpfe werden robust korrigiert. Frachtführer: Bezeichnung statt Nummer; Entladung mit „Frachtführer“ am Ende.
// @match        https://roadnet.dpdgroup.com/execution/transport_units*
// @match        https://roadnet.dpdgroup.com/execution/trips*
// @match        http://lts.dpdit.de/*
// @match        https://lts.dpdit.de/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';
  if (window.__RN_TU_SUMMARY_RUNNING) return;
  window.__RN_TU_SUMMARY_RUNNING = true;

  /* ===================== SHARED HELPERS ===================== */
  const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
  const pad2 = (n) => String(n).padStart(2, '0');

  const normKey = (s) =>
    norm(s)
      .toLowerCase()
      .replace(/\u00a0/g, ' ')
      .replace(/[^\p{L}\p{N}#%/ ]/gu, '')
      .replace(/\s+/g, ' ')
      .trim();

  function toast(msg, ok = true) {
    const el = document.createElement('div');
    el.textContent = msg;
    el.style.cssText =
      'position:fixed;right:16px;bottom:16px;z-index:2147483647;padding:7px 10px;border-radius:10px;' +
      'font:700 11px system-ui;color:#fff;box-shadow:0 10px 30px rgba(0,0,0,.25);' +
      (ok ? 'background:#16a34a;' : 'background:#b91c1c;');
    document.body.appendChild(el);
    setTimeout(() => {
      el.style.opacity = '0';
      el.style.transition = 'opacity .25s';
      setTimeout(() => el.remove(), 260);
    }, 1100);
  }

  function debounce(fn, ms) {
    let t = null;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }

  function isVisible(el) {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  }

  function parseDeDateTime(s) {
    const t = norm(s);
    if (!t) return null;
    const m = t.match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\s*,?\s*(\d{1,2}):(\d{2}))?/);
    if (m) {
      const dd = Number(m[1]), mm = Number(m[2]), yy = Number(m[3].length === 2 ? ('20' + m[3]) : m[3]);
      const hh = m[4] != null ? Number(m[4]) : 0;
      const mi = m[5] != null ? Number(m[5]) : 0;
      const d = new Date(yy, mm - 1, dd, hh, mi, 0, 0);
      return Number.isFinite(d.getTime()) ? d : null;
    }
    const d2 = new Date(t);
    return Number.isFinite(d2.getTime()) ? d2 : null;
  }

  function parsePercentToUnit(s) {
    const t = norm(s).replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, '');
    if (!t) return null;
    const v = parseFloat(t);
    if (!Number.isFinite(v)) return null;
    return v / 100;
  }

  function fmtSumUnit(v) {
    if (!Number.isFinite(v)) return '0,0';
    return String(v.toFixed(1)).replace('.', ',');
  }

  function escHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function looksNumericOnly(s) {
    const t = norm(s);
    return !!t && /^[\d.\-]+$/.test(t) && !/[A-Za-zÄÖÜäöü]/.test(t);
  }

  function bestCellText(td) {
    if (!td) return '';
    const main = norm(td.innerText ?? td.textContent ?? '');
    if (!looksNumericOnly(main)) return main;

    const attrs = [
      td.getAttribute('title'),
      td.getAttribute('aria-label'),
      td.getAttribute('data-original-title'),
      td.getAttribute('data-tooltip'),
      td.getAttribute('data-title')
    ].map(norm).filter(Boolean);

    for (const a of attrs) {
      if (a && /[A-Za-zÄÖÜäöü]/.test(a)) return a;
    }

    const titleEl = td.querySelector('[title]');
    if (titleEl) {
      const t = norm(titleEl.getAttribute('title'));
      if (t && /[A-Za-zÄÖÜäöü]/.test(t)) return t;
    }

    return main;
  }

  /* =======================================================================
     LTS-SEITE: Auto-Fill + robuste Suche (max 3 Versuche), stop wenn Daten da
     ======================================================================= */
  (function ltsMessageReceiver() {
    if (!location.hostname.includes('lts.dpdit.de')) return;

    const ALLOWED_ORIGINS = new Set(['https://roadnet.dpdgroup.com']);
    const PENDING_KEY = 'rn_pending_wbl';
    const SUBMIT_DONE_KEY = 'rn_pending_submit_done';
    const WNAME_PREFIX = 'RN_WBL_PENDING=';
    const SESSION_RE = /^\/\(S\([^)]+\)\)\//i;
    const WBL_RE = /\/WBLebenslauf\.aspx$/i;
    const INDEX_RE = /\/index\.aspx$/i;
    const WBL_HEADER_FALLBACK = [
      'WB Nr',
      'Scandepot:',
      'Scanart:',
      'Zieldepot:',
      'Datum:',
      'Uhrzeit:',
      'Plombennummer:',
      'Auslastung (%):',
      'Express',
      'Bemerkung:'
    ];
    const WBL_LABEL_STYLE_ID = 'rn-wbl-header-label-style';

    let current = { left: '', right: '' };
    let submitLoop = null;
    let pendingWatch = null;

    const parsePending = (v) => {
      const m = String(v || '').match(/^(\d{4})-(\d{4})$/);
      return m ? { left: m[1], right: m[2] } : null;
    };

    function getSubmitDoneKey() {
      try {
        const v = String(sessionStorage.getItem(SUBMIT_DONE_KEY) || '');
        if (v) return v;
      } catch {}
      try {
        return String(localStorage.getItem(SUBMIT_DONE_KEY) || '');
      } catch {}
      return '';
    }

    function setSubmitDoneKey(v) {
      try { sessionStorage.setItem(SUBMIT_DONE_KEY, v); } catch {}
      try { localStorage.setItem(SUBMIT_DONE_KEY, v); } catch {}
    }

    function clearSubmitDoneKey() {
      try { sessionStorage.removeItem(SUBMIT_DONE_KEY); } catch {}
      try { localStorage.removeItem(SUBMIT_DONE_KEY); } catch {}
    }

    function wasSubmitDispatched(left, right) {
      return getSubmitDoneKey() === `${left}-${right}`;
    }

    function markSubmitDispatched(left, right) {
      setSubmitDoneKey(`${left}-${right}`);
    }

    function readWindowNamePending() {
      try {
        const m = String(window.name || '').match(/RN_WBL_PENDING=(\d{4})-(\d{4})/);
        return m ? { left: m[1], right: m[2] } : null;
      } catch {
        return null;
      }
    }

    function writeWindowNamePending(left, right) {
      try {
        const keep = String(window.name || '').replace(/\|?RN_WBL_PENDING=\d{4}-\d{4}/g, '');
        window.name = `${keep}|${WNAME_PREFIX}${left}-${right}`;
      } catch {}
    }

    function clearWindowNamePending() {
      try {
        window.name = String(window.name || '').replace(/\|?RN_WBL_PENDING=\d{4}-\d{4}/g, '');
      } catch {}
    }

    function pendingFromQuery() {
      try {
        const qp = new URLSearchParams(location.search);
        const bridge = String(qp.get('BRUECKE') || '').replace(/[^\d]/g, '');
        if (bridge.length === 8) return { left: bridge.slice(0, 4), right: bridge.slice(4) };
      } catch {}
      return null;
    }

    function savePending(left, right, opts = {}) {
      const v = `${left}-${right}`;
      const resetSubmit = !!opts.resetSubmit;
      if (resetSubmit || getSubmitDoneKey() !== v) clearSubmitDoneKey();
      try { sessionStorage.setItem(PENDING_KEY, v); } catch {}
      try { localStorage.setItem(PENDING_KEY, v); } catch {}
      writeWindowNamePending(left, right);
    }

    function peekPending() {
      try {
        const p = parsePending(sessionStorage.getItem(PENDING_KEY));
        if (p) return p;
      } catch {}

      try {
        const p = parsePending(localStorage.getItem(PENDING_KEY));
        if (p) {
          try { sessionStorage.setItem(PENDING_KEY, `${p.left}-${p.right}`); } catch {}
          return p;
        }
      } catch {}

      const pWn = readWindowNamePending();
      if (pWn) {
        savePending(pWn.left, pWn.right);
        return pWn;
      }

      const pQ = pendingFromQuery();
      if (pQ) {
        savePending(pQ.left, pQ.right);
        return pQ;
      }

      return null;
    }

    function clearPending() {
      try { sessionStorage.removeItem(PENDING_KEY); } catch {}
      try { localStorage.removeItem(PENDING_KEY); } catch {}
      clearSubmitDoneKey();
      clearWindowNamePending();
    }

    function getSessionPrefix() {
      const m = String(location.pathname || '').match(/^\/\(S\([^)]+\)\)/i);
      return m ? m[0] : '';
    }

    function getSessionIdQuery() {
      try {
        const qp = new URLSearchParams(location.search);
        const sid = String(qp.get('sessionid') || '').trim();
        return sid ? `?sessionid=${encodeURIComponent(sid)}` : '';
      } catch {
        return '';
      }
    }

    function ensureOnWblPage() {
      const sessionPrefix = getSessionPrefix();
      const onWbl = WBL_RE.test(location.pathname);

      if (sessionPrefix && onWbl) return true;

      if (sessionPrefix && !onWbl) {
        location.href = location.origin + sessionPrefix + '/WBLebenslauf.aspx' + getSessionIdQuery();
        return false;
      }

      // Ohne Session nicht auf WBL zwingen; zuerst Login/Session über index.aspx zulassen.
      if (!sessionPrefix && onWbl) {
        location.href = location.origin + '/index.aspx' + getSessionIdQuery();
        return false;
      }

      if (!sessionPrefix && !onWbl && !INDEX_RE.test(location.pathname)) {
        return false;
      }

      return false;
    }

    function findWblResultTable() {
      const allTables = Array.from(document.querySelectorAll('table'));
      if (!allTables.length) return null;

      const textScore = (txt) => {
        let score = 0;
        const t = String(txt || '').toLowerCase();
        if (t.includes('wb nr')) score += 45;
        if (t.includes('plombennummer')) score += 24;
        if (t.includes('scanart')) score += 22;
        if (t.includes('scandepot')) score += 20;
        if (t.includes('zieldepot')) score += 20;
        if (t.includes('auslastung')) score += 20;
        if (t.includes('bemerkung')) score += 18;
        return score;
      };

      let best = null;
      let bestScore = -1;
      for (const t of allTables) {
        const rows = Array.from(t.querySelectorAll('tr'));
        if (rows.length < 2) continue;

        let maxCols = 0;
        for (const r of rows.slice(0, 8)) {
          const c = r.querySelectorAll('th,td').length;
          if (c > maxCols) maxCols = c;
        }
        if (maxCols < 6) continue;

        const hasScopeHeader = !!t.querySelector('th[scope="col"]');
        const hasAnyHeader = !!t.querySelector('th');
        const score =
          textScore(t.innerText || t.textContent || '') +
          (hasScopeHeader ? 60 : 0) +
          (hasAnyHeader ? 20 : 0) +
          Math.min(rows.length, 35) +
          Math.min(maxCols, 18);

        if (score > bestScore) {
          bestScore = score;
          best = t;
        }
      }
      return best;
    }

    function forceVisibleWblHeaders(table) {
      if (!table) return;
      const ths = Array.from(table.querySelectorAll('th[scope="col"], thead th, tr:first-child th'));
      if (!ths.length) return;

      const tr = ths[0].closest('tr');
      if (tr) {
        tr.style.setProperty('color', '#111', 'important');
        tr.style.setProperty('font-weight', '700', 'important');
      }
      for (const th of ths) {
        th.style.setProperty('color', '#111', 'important');
        th.style.setProperty('-webkit-text-fill-color', '#111', 'important');
        th.style.setProperty('font-weight', '700', 'important');
        th.style.setProperty('font-size', '12px', 'important');
        th.style.setProperty('line-height', '1.2', 'important');
        th.style.setProperty('background-color', '#e9f0fb', 'important');
        th.style.setProperty('text-shadow', 'none', 'important');
        th.style.setProperty('opacity', '1', 'important');
        th.style.setProperty('visibility', 'visible', 'important');
        th.style.setProperty('text-indent', '0', 'important');
        th.style.setProperty('white-space', 'nowrap', 'important');
        th.style.setProperty('overflow', 'visible', 'important');
        th.style.setProperty('min-height', '20px', 'important');
      }
    }

    function ensureWblLabelStyle() {
      if (document.getElementById(WBL_LABEL_STYLE_ID)) return;
      const styleEl = document.createElement('style');
      styleEl.id = WBL_LABEL_STYLE_ID;
      styleEl.textContent = [
        'table#lGridScan th[data-rn-label]{position:absolute !important;overflow:visible !important;}',
        'table#lGridScan th[data-rn-label]::after{',
        'content: attr(data-rn-label);',
        'position: relative;',
        'display: inline-block;',
        'margin-left: 2px;',
        'color:#111 !important;',
        '-webkit-text-fill-color:#111 !important;',
        'font:700 12px/1.2 Arial, sans-serif !important;',
        'text-shadow:none !important;',
        'white-space:nowrap;',
        'opacity:1 !important;',
        'visibility:visible !important;',
        '}'
      ].join('');
      (document.head || document.documentElement).appendChild(styleEl);
    }

    function fillExistingWblHeaderTexts(table) {
      if (!table) return false;
      const ths = Array.from(table.querySelectorAll('th[scope="col"], thead th, tr:first-child th'));
      if (!ths.length) return false;
      ensureWblLabelStyle();

      const clean = (s) => String(s ?? '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
      const nonEmptyCount = ths.reduce((cnt, th) => cnt + (clean(th.textContent).length ? 1 : 0), 0);
      const probablyBroken = (ths.length >= 8 && nonEmptyCount <= 2);
      const mapCount = Math.min(ths.length, WBL_HEADER_FALLBACK.length);

      let changed = false;
      for (let i = 0; i < mapCount; i++) {
        const th = ths[i];
        if (!th) continue;
        th.setAttribute('data-rn-label', WBL_HEADER_FALLBACK[i]);
        const current = clean(th.textContent);
        if (probablyBroken || current === '' || current === '-') {
          th.textContent = WBL_HEADER_FALLBACK[i];
          changed = true;
        }
      }
      return changed;
    }

    function injectFallbackWblHeader(table) {
      if (!table || table.querySelector('th')) return false;
      const firstBodyRow = table.querySelector('tbody tr') || table.querySelector('tr');
      if (!firstBodyRow) return false;
      const bodyCells = Array.from(firstBodyRow.querySelectorAll('td'));
      if (bodyCells.length < 6) return false;

      let thead = table.querySelector('thead');
      if (!thead) {
        thead = document.createElement('thead');
        const tbody = table.querySelector('tbody');
        if (tbody) table.insertBefore(thead, tbody);
        else table.insertBefore(thead, table.firstChild);
      }
      if (thead.querySelector('tr[data-rn-wbl-fallback="1"]')) return true;

      const tr = document.createElement('tr');
      tr.setAttribute('data-rn-wbl-fallback', '1');
      tr.style.setProperty('color', '#111', 'important');
      tr.style.setProperty('font-weight', '700', 'important');

      const count = Math.min(bodyCells.length, WBL_HEADER_FALLBACK.length);
      for (let i = 0; i < count; i++) {
        const th = document.createElement('th');
        th.setAttribute('scope', 'col');
        th.textContent = WBL_HEADER_FALLBACK[i];
        th.style.setProperty('color', '#111', 'important');
        th.style.setProperty('font-weight', '700', 'important');
        th.style.setProperty('background-color', '#e9f0fb', 'important');
        th.style.setProperty('text-shadow', 'none', 'important');
        tr.appendChild(th);
      }

      thead.insertBefore(tr, thead.firstChild);
      return true;
    }

    function ensureWblHeadersVisible() {
      if (!WBL_RE.test(location.pathname)) return;
      const table = findWblResultTable();
      if (!table) return;

      if (!table.querySelector('th')) {
        injectFallbackWblHeader(table);
      }

      fillExistingWblHeaderTexts(table);
      forceVisibleWblHeaders(table);
    }

    function hasResults() {
      const tables = Array.from(document.querySelectorAll('table'));
      for (const t of tables) {
        const txt = (t.innerText || '').toLowerCase();
        if (txt.includes('wb nr')) {
          const bodyRows = t.querySelectorAll('tbody tr').length;
          const allRows = t.querySelectorAll('tr').length;
          if (bodyRows >= 1 || allRows >= 2) {
            setTimeout(() => ensureWblHeadersVisible(), 0);
            return true;
          }
        }
      }
      return false;
    }

    function stopSubmitLoop() {
      if (submitLoop) clearInterval(submitLoop);
      submitLoop = null;
    }

    function fillAndSubmit(left, right) {
      if (!WBL_RE.test(location.pathname)) return false;
      if (!SESSION_RE.test(location.pathname)) return false;

      if (hasResults()) {
        clearPending();
        stopSubmitLoop();
        return true;
      }

      const form = document.querySelector('form#frmLebenslauf') || document.querySelector('form');
      if (!form) return false;

      const i1 = form.querySelector('input[name="txtWBNR1"]');
      const i4 = form.querySelector('input[name="txtWBNR4"]');
      if (!i1 || !i4) return false;

      i1.value = left;
      i4.value = right;

      i1.dispatchEvent(new Event('input', { bubbles: true }));
      i1.dispatchEvent(new Event('change', { bubbles: true }));
      i4.dispatchEvent(new Event('input', { bubbles: true }));
      i4.dispatchEvent(new Event('change', { bubbles: true }));

      const btn =
        Array.from(form.querySelectorAll('input[type="submit"],button,input[type="button"]'))
          .find(x => /suchen/i.test((x.value || x.textContent || ''))) ||
        form.querySelector('input[name="cmdSearch"]') ||
        form.querySelector('#cmdSearch');

      try {
        if (btn && typeof form.requestSubmit === 'function') {
          markSubmitDispatched(left, right);
          form.requestSubmit(btn);
          return true;
        }
      } catch {}

      try {
        if (btn) {
          markSubmitDispatched(left, right);
          btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
          btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
          btn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
          return true;
        }
      } catch {}

      try {
        markSubmitDispatched(left, right);
        form.submit();
        return true;
      } catch {}

      return false;
    }

    function startSubmitLoop(left, right) {
      const key = `${left}-${right}`;
      const curKey = `${current.left}-${current.right}`;
      if (submitLoop && key === curKey) return;

      current = { left, right };
      stopSubmitLoop();

      let tries = 0;
      const run = () => {
        if (hasResults()) {
          clearPending();
          stopSubmitLoop();
          return;
        }

        tries++;
        fillAndSubmit(left, right);

        if (tries >= 3) {
          stopSubmitLoop();
        }
      };

      run();
      submitLoop = setInterval(run, 700);
    }

    function processPending() {
      const p = peekPending();
      if (!p) return;
      if (hasResults()) {
        clearPending();
        stopSubmitLoop();
        return;
      }

      // Nach bereits ausgelöstem Such-Submit nicht erneut automatisch submitten.
      if (SESSION_RE.test(location.pathname) && WBL_RE.test(location.pathname) && wasSubmitDispatched(p.left, p.right)) {
        return;
      }

      if (!ensureOnWblPage()) return;

      setTimeout(() => startSubmitLoop(p.left, p.right), 350);
    }

    function startPendingWatcher() {
      if (pendingWatch) return;
      pendingWatch = setInterval(() => {
        const p = peekPending();
        if (!p) {
          clearInterval(pendingWatch);
          pendingWatch = null;
          return;
        }

        if (hasResults()) {
          clearPending();
          stopSubmitLoop();
          clearInterval(pendingWatch);
          pendingWatch = null;
          return;
        }

        processPending();
      }, 1000);
    }

    window.addEventListener('message', (ev) => {
      if (!ALLOWED_ORIGINS.has(ev.origin)) return;
      const data = ev.data || {};
      if (data.type !== 'RN_WBL') return;

      const left = String(data.a || '');
      const right = String(data.b || '');
      if (!/^\d{4}$/.test(left) || !/^\d{4}$/.test(right)) return;

      savePending(left, right, { resetSubmit: true });
      processPending();
      startPendingWatcher();

      try {
        if (ev.source && typeof ev.source.postMessage === 'function') {
          ev.source.postMessage({ type: 'RN_WBL_ACK', ok: true }, ev.origin);
        }
      } catch {}
    });

    window.addEventListener('load', () => {
      ensureWblHeadersVisible();
      processPending();
      if (peekPending()) startPendingWatcher();
    });

    setTimeout(() => ensureWblHeadersVisible(), 350);
    if (peekPending()) startPendingWatcher();

    const obs = new MutationObserver(() => {
      ensureWblHeadersVisible();
      const p = peekPending();
      if (!p) return;
      if (hasResults()) {
        clearPending();
        stopSubmitLoop();
        return;
      }
      processPending();
    });
    if (document.body) obs.observe(document.body, { childList: true, subtree: true });
  })();

  /* ===================== ROADNET-TEIL ===================== */
  if (!location.hostname.includes('roadnet.dpdgroup.com')) return;

  const NS = 'rn-tu-';
  const PANEL_ID = `${NS}panel`;
  const BTN_ID = `${NS}openbtn`;
  const BTN_OV_ID = `${NS}openbtn-uebentl`;
  const TABLE_WRAP_CLASS = `${NS}tw`;
  const SCROLL_BOX_CLASS = `${NS}scrollbox`;
  const DEBOUNCE_MS = 220;
  const AUTOFIT_TO_WIDTH = true;
  const BRIDGE_ENABLED = true;
  const BRIDGE_PRIMARY_ENDPOINT = 'http://10.14.7.169/DPD_Dashboard/roadnet_tu_push.php';
  const BRIDGE_FALLBACK_ENDPOINTS = [];
  const BRIDGE_ENDPOINT_STORAGE_KEY = 'rn_tu_bridge_endpoint';
  const BRIDGE_DEPOT_STORAGE_KEY = 'rn_tu_bridge_depot';
  const BRIDGE_CLIENT_ID = 'roadnet-default'; // Mit roadnet_bridge_clients.json abgestimmt
  const BRIDGE_DEPOT = ''; // Optional fest setzen, z.B. 'D0157'
  const BRIDGE_TOKEN = '12c87869d52fc4651843512b595fc79ba8251d450988cbf4199e5918b33ecbd0'; // Pflicht bei requireClient=true: lokal im Userscript setzen
  const BRIDGE_MIN_INTERVAL_MS = 15000;
  const BRIDGE_POLL_INTERVAL_MS = 15000;
  const BRIDGE_MAX_ROWS = 700;
  const ROADNET_SOFT_REFRESH_ENABLED = true;
  const ROADNET_SOFT_REFRESH_INTERVAL_MS = 60000;
  const ROADNET_HARD_REFRESH_ENABLED = true;
  const ROADNET_HARD_REFRESH_INTERVAL_MS = 300000;
  const ROADNET_HARD_REFRESH_IDLE_MS = 120000;
  const ROADNET_STALE_RELOAD_ENABLED = true;
  const ROADNET_STALE_RELOAD_AFTER_MS = 60000;

  const MODEL_UNLOAD = {
    modeKey: 'unload',
    title: 'Roadnet Zusammenfassung – Entladung',
    defaultSort: { key: 'arrAct', dir: 'desc' },
    cols: [
      { key: 'status',  title: 'Status' },
      { key: 'from',    title: 'Von' },
      { key: 'depAct',  title: 'Tats. Abfahrt' },
      { key: 'lts',     title: 'LTS #' },
      { key: 'arrAct',  title: 'Tats. Ankunft' },
      { key: 'unlBeg',  title: 'Entladebeginn' },
      { key: 'unlEnd',  title: 'Entladung Ende' },
      { key: 'load',    title: '%' },
      { key: 'type',    title: 'Art' },
      { key: 'seal',    title: 'Plombe' },
      { key: 'carrier', title: 'Frachtführer' }
    ],
    resolve: (_headerMap, pick) => ({
      status:  pick(['status', 'status transporteinheit', 'te status', 'tu status']),
      from:    pick([
        'abgangsort', 'name abgangsstandort', 'code abgangsstation', 'code abgangsstandort',
        'abgangsstandort', 'abgang', 'von', 'von bu', 'von business unit', 'from', 'origin'
      ]),
      depAct:  pick(['tatsächliche abfahrt', 'tatsaechliche abfahrt', 'actual departure', 'departure actual']),
      lts:     pick([
        'nummer transporteinheit', 'transporteinheit nr', 'transport unit number', 'transport unit no',
        'lts #', 'lts', 'nummer transporte', 'wb nr', 'wbnr', 'bruecke'
      ]),
      arrAct:  pick(['tatsächliche ankunft', 'tatsaechliche ankunft', 'ankunft', 'actual arrival', 'arrival actual']),
      unlBeg:  pick(['entladung beginn', 'entladebeginn', 'entladung start', 'unload start', 'unloading start']),
      unlEnd:  pick(['entladung ende', 'entladeende', 'entladung abgeschlossen', 'unload end', 'unloading end']),
      load:    pick(['auslastung', 'auslastung %', 'load', 'load %', 'fill level', '%']),
      type:    pick(['art transporteinheit', 'typ', 'transportart', 'art transporteinh', 'verkehr', 'traffic']),
      seal:    pick(['plombennummer', 'plombe', 'seal', 'seal number']),
      carrier: pick(['bezeichnung frachtführer', 'bezeichnung frachtfuehrer', 'frachtführer', 'frachtfuehrer', 'carrier'])
    }),
    special: { statusErfasstKey: 'status', statusErfasstText: 'erfasst' }
  };

  const MODEL_LOAD = {
    modeKey: 'load',
    title: 'Roadnet Zusammenfassung – Beladung',
    defaultSort: { key: 'depPlan', dir: 'desc' },
    cols: [
      { key: 'status',   title: 'Status' },
      { key: 'toName',   title: 'An' },
      { key: 'depPlan',  title: 'Gepl. Abfahrt' },
      { key: 'depAct',   title: 'Tats. Abfahrt' },
      { key: 'loadBeg',  title: 'Beladung Beginn' },
      { key: 'loadEnd',  title: 'Beladung Ende' },
      { key: 'lts',      title: 'LTS #' },
      { key: 'sealDep',  title: 'Plombe Abfahrt' },
      { key: 'loadPct',  title: '%' },
      { key: 'carrier',  title: 'Frachtführer' },
      { key: 'type',     title: 'Art' }
    ],
    resolve: (_headerMap, pick) => ({
      status:  pick(['status', 'status transporteinheit', 'te status', 'tu status']),
      toName:  pick([
        'name empfangsstandort', 'name empfangsstation', 'empfangsort', 'empfangsstandort',
        'an', 'nach', 'code empfangsstation', 'code empfangsstandort', 'empfangsstation',
        'to', 'destination'
      ]),
      depPlan: pick(['geplante abfahrt', 'geplant abfahrt', 'planned departure', 'departure planned']),
      depAct:  pick(['tatsächliche abfahrt', 'tatsaechliche abfahrt', 'actual departure', 'departure actual']),
      loadBeg: pick(['beladung beginn', 'beladung start', 'loading start']),
      loadEnd: pick(['beladung ende', 'beladung abgeschlossen', 'loading end']),
      lts:     pick([
        'nummer transporteinheit', 'transporteinheit nr', 'transport unit number', 'transport unit no',
        'lts #', 'lts', 'nummer transporte', 'wb nr', 'wbnr', 'bruecke'
      ]),
      sealDep: pick(['plombennummer abfahrt', 'plombe abfahrt', 'seal departure', 'seal']),
      loadPct: pick(['auslastung', 'auslastung %', 'load', 'load %', 'fill level', '%']),
      carrier: pick(['bezeichnung frachtführer', 'bezeichnung frachtfuehrer', 'frachtführer', 'frachtfuehrer', 'carrier']),
      type:    pick(['art transporteinheit', 'typ', 'transportart', 'art transporteinh', 'verkehr', 'traffic'])
    }),
    special: { statusErfasstKey: null }
  };

  const state = {
    model: MODEL_UNLOAD,
    sort: { ...MODEL_UNLOAD.defaultSort },
    rows: [],
    lastStamp: '',
    inferredDepot: '',
    manualDepot: '',
    depotPromptShown: false,
    scrollLeft: 0,
    scrollTop: 0,
    scale: 1,
    lastUserInteractionAt: Date.now(),
    lastSoftRefreshAt: 0,
    lastHardRefreshAt: Date.now(),
    lastExtractSignature: '',
    lastDataChangeAt: Date.now()
  };
  const BRIDGE_LOG_MAX = 120;
  const bridgeState = {
    lastSig: '',
    lastSentAt: 0,
    inFlight: false,
    statusLabel: 'Idle',
    statusColor: '#4b5563',
    logs: []
  };

  function bridgeTriggerLabel(trigger) {
    return trigger === 'manual' ? 'MANUELL' : 'AUTO';
  }

  function timeStampShort(d = new Date()) {
    return `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  }

  function setBridgeStatus(label, color = '#4b5563') {
    bridgeState.statusLabel = norm(label) || 'Idle';
    bridgeState.statusColor = color || '#4b5563';
    renderBridgeLog();
  }

  function addBridgeLog(level, msg) {
    const entry = {
      ts: timeStampShort(new Date()),
      level: String(level || 'info').toLowerCase(),
      msg: norm(msg) || '-'
    };
    bridgeState.logs.push(entry);
    if (bridgeState.logs.length > BRIDGE_LOG_MAX) {
      bridgeState.logs.splice(0, bridgeState.logs.length - BRIDGE_LOG_MAX);
    }

    const line = `[RN-BRIDGE ${entry.level.toUpperCase()} ${entry.ts}] ${entry.msg}`;
    if (entry.level === 'error') console.error(line);
    else if (entry.level === 'warn') console.warn(line);
    else console.log(line);

    renderBridgeLog();
  }

  function clearBridgeLog() {
    bridgeState.logs.length = 0;
    renderBridgeLog();
  }

  function renderBridgeLog() {
    const panel = document.getElementById(PANEL_ID);
    if (!panel) return;

    const statusEl = panel.querySelector(`.${NS}bridge-status`);
    if (statusEl) {
      statusEl.textContent = bridgeState.statusLabel || 'Idle';
      statusEl.style.color = bridgeState.statusColor || '#4b5563';
    }

    const logEl = panel.querySelector(`.${NS}bridge-log`);
    if (!logEl) return;
    const lines = bridgeState.logs.map((e) => `${e.ts} [${String(e.level || '').toUpperCase()}] ${e.msg}`);
    logEl.textContent = lines.length ? lines.join('\n') : 'Noch keine Bridge-Logs.';
    logEl.scrollTop = logEl.scrollHeight;
  }

  function normalizeBridgeEndpoint(raw) {
    const t = norm(raw);
    if (!t) return '';
    if (!/^https?:\/\//i.test(t)) return '';
    return t.replace(/\/+$/, '');
  }

  function buildBridgeEndpoints() {
    const out = [];
    const seen = new Set();
    const push = (url) => {
      const cleaned = normalizeBridgeEndpoint(url);
      if (!cleaned || seen.has(cleaned)) return;
      seen.add(cleaned);
      out.push(cleaned);
    };

    push(BRIDGE_PRIMARY_ENDPOINT);
    (Array.isArray(BRIDGE_FALLBACK_ENDPOINTS) ? BRIDGE_FALLBACK_ENDPOINTS : []).forEach(push);
    try {
      push(localStorage.getItem(BRIDGE_ENDPOINT_STORAGE_KEY) || '');
    } catch {}
    return out;
  }

  function postBridgePayload(payload, endpoints, idx = 0, trigger = 'auto') {
    if (!Array.isArray(endpoints) || idx >= endpoints.length) {
      const err = new Error('no_endpoint_available');
      addBridgeLog('error', `${bridgeTriggerLabel(trigger)}: kein erreichbarer Endpoint.`);
      return Promise.reject(err);
    }

    const endpoint = endpoints[idx];
    const attempt = idx + 1;
    const total = endpoints.length;
    const startedAt = Date.now();
    addBridgeLog('info', `${bridgeTriggerLabel(trigger)}: Versuch ${attempt}/${total} -> ${endpoint}`);

    return fetch(endpoint, {
      method: 'POST',
      mode: 'cors',
      credentials: 'omit',
      headers: {
        'Content-Type': 'application/json',
        'X-Roadnet-Token': BRIDGE_TOKEN
      },
      body: JSON.stringify(payload)
    })
      .then((r) => r.text().then((bodyText) => ({ r, bodyText })))
      .then(({ r, bodyText }) => {
        const durationMs = Date.now() - startedAt;
        const bodyPreview = norm(bodyText).slice(0, 220);
        if (!r.ok) {
          addBridgeLog(
            'warn',
            `${bridgeTriggerLabel(trigger)}: ${endpoint} -> HTTP ${r.status} (${durationMs}ms)` +
            (bodyPreview ? ` | ${bodyPreview}` : '')
          );
          return postBridgePayload(payload, endpoints, idx + 1, trigger);
        }
        addBridgeLog(
          'ok',
          `${bridgeTriggerLabel(trigger)}: ${endpoint} -> HTTP ${r.status} (${durationMs}ms)` +
          (bodyPreview ? ` | ${bodyPreview}` : '')
        );
        return endpoint;
      })
      .catch((err) => {
        const durationMs = Date.now() - startedAt;
        const message = norm(err && err.message ? err.message : String(err || 'network_error'));
        addBridgeLog('error', `${bridgeTriggerLabel(trigger)}: ${endpoint} Fehler nach ${durationMs}ms: ${message}`);
        return postBridgePayload(payload, endpoints, idx + 1, trigger);
      });
  }

  function normalizeBridgeDepot(raw) {
    const t = norm(raw).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (/^D\d{4}$/.test(t)) return t;
    if (/^D\d{3}$/.test(t)) return `D0${t.slice(1)}`;
    if (/^\d{4}$/.test(t)) return `D${t}`;
    if (/^\d{3}$/.test(t)) return `D0${t}`;
    return '';
  }

  function readStoredBridgeDepot() {
    try {
      return normalizeBridgeDepot(localStorage.getItem(BRIDGE_DEPOT_STORAGE_KEY) || '');
    } catch {
      return '';
    }
  }

  function writeStoredBridgeDepot(raw) {
    const depot = normalizeBridgeDepot(raw);
    try {
      if (depot) localStorage.setItem(BRIDGE_DEPOT_STORAGE_KEY, depot);
      else localStorage.removeItem(BRIDGE_DEPOT_STORAGE_KEY);
    } catch {}
    return depot;
  }

  function getConfiguredBridgeDepot() {
    const fromConfig = normalizeBridgeDepot(BRIDGE_DEPOT);
    if (fromConfig) return fromConfig;

    const fromState = normalizeBridgeDepot(state.manualDepot);
    if (fromState) return fromState;

    const fromStorage = readStoredBridgeDepot();
    if (fromStorage) {
      state.manualDepot = fromStorage;
      return fromStorage;
    }
    return '';
  }

  function updateDepotUI() {
    const panel = document.getElementById(PANEL_ID);
    if (!panel) return;
    const depotEl = panel.querySelector(`.${NS}depotvalue`);
    if (!depotEl) return;
    const depot = getConfiguredBridgeDepot();
    depotEl.textContent = depot || 'nicht gesetzt';
    depotEl.style.color = depot ? '#065f46' : '#991b1b';
  }

  function promptBridgeDepot(force = false) {
    const fixed = normalizeBridgeDepot(BRIDGE_DEPOT);
    if (fixed) {
      state.manualDepot = fixed;
      updateDepotUI();
      return fixed;
    }

    let seed = getConfiguredBridgeDepot() || '';
    while (true) {
      const input = window.prompt('Bitte Standort/Depot eingeben (z.B. D0157, 0157 oder 157).', seed);
      if (input == null) {
        if (force) toast('Depot bleibt erforderlich fuer den Export', false);
        updateDepotUI();
        return '';
      }

      const depot = normalizeBridgeDepot(input);
      if (!depot) {
        alert('Ungueltiges Depot. Erlaubt sind z.B. D0157, 0157 oder 157.');
        seed = norm(input);
        continue;
      }

      state.manualDepot = writeStoredBridgeDepot(depot);
      state.depotPromptShown = true;
      updateDepotUI();
      toast(`Depot ${state.manualDepot} gespeichert`, true);
      return state.manualDepot;
    }
  }

  function ensureBridgeDepotConfigured(forcePrompt = false) {
    const configured = getConfiguredBridgeDepot();
    if (configured) {
      updateDepotUI();
      return configured;
    }

    if (!forcePrompt && state.depotPromptShown) {
      updateDepotUI();
      return '';
    }

    state.depotPromptShown = true;
    return promptBridgeDepot(forcePrompt);
  }

  function parseBridgeDepotCandidates(raw) {
    const out = new Set();
    const text = String(raw ?? '').toUpperCase();
    if (!text) return [];

    const direct = normalizeBridgeDepot(text);
    if (direct) out.add(direct);

    const addDigits = (digits) => {
      const depot = normalizeBridgeDepot(String(digits || '').replace(/[^\d]/g, ''));
      if (depot) out.add(depot);
    };

    for (const m of text.matchAll(/\bDE(\d{4})\b/g)) addDigits(m[1]);
    for (const m of text.matchAll(/\bDE(\d{4})[A-Z0-9]*/g)) addDigits(m[1]);
    for (const m of text.matchAll(/\b0010(\d{3,4})\b/g)) addDigits(m[1]);
    for (const m of text.matchAll(/\bD(\d{4})\b/g)) addDigits(m[1]);
    for (const m of text.matchAll(/\bD(\d{3})\b/g)) addDigits(m[1]);

    return Array.from(out);
  }

  function getBridgeDepotCandidatesFromCell(td) {
    if (!td) return [];
    const out = new Set();
    const texts = [
      bestCellText(td),
      td.innerText,
      td.textContent,
      td.getAttribute('title'),
      td.getAttribute('aria-label'),
      td.getAttribute('data-original-title'),
      td.getAttribute('data-tooltip'),
      td.getAttribute('data-title')
    ];

    for (const t of texts) {
      for (const cand of parseBridgeDepotCandidates(t)) out.add(cand);
    }
    return Array.from(out);
  }

  function inferBridgeDepotFromTable(headerMap, bodyRows) {
    if (!Array.isArray(bodyRows) || !bodyRows.length) return '';

    const counts = new Map();
    const add = (depot, weight) => {
      if (!depot) return;
      counts.set(depot, (counts.get(depot) || 0) + weight);
    };
    const addFromCell = (td, weight) => {
      for (const cand of getBridgeDepotCandidatesFromCell(td)) add(cand, weight);
    };

    const idxFromHeaders = (candidates) => {
      for (const cand of candidates) {
        const key = normKey(cand);
        if (headerMap.has(key)) return headerMap.get(key);
        for (const [hk, idx] of headerMap.entries()) {
          if (hk.includes(key)) return idx;
        }
      }
      return null;
    };

    const preferredCols = [
      { idx: idxFromHeaders(['code empfangsstandort', 'empfangsstandort', 'nach']), weight: 6 },
      { idx: idxFromHeaders(['name empfangsstandort', 'bezeichnung nach']), weight: 4 },
      { idx: idxFromHeaders(['code abgangsstandort', 'abgangsstandort', 'von']), weight: 3 },
      { idx: idxFromHeaders(['von bu', 'von business unit']), weight: 3 }
    ].filter(x => x.idx != null);

    for (const tr of bodyRows) {
      const tds = getRowCells(tr);
      if (!tds.length) continue;

      for (const pref of preferredCols) {
        if (pref.idx >= 0 && pref.idx < tds.length) addFromCell(tds[pref.idx], pref.weight);
      }

      for (const td of tds) addFromCell(td, 1);
    }

    let bestDepot = '';
    let bestScore = 0;
    for (const [depot, score] of counts.entries()) {
      if (score > bestScore) {
        bestDepot = depot;
        bestScore = score;
      }
    }
    return bestScore >= 2 ? bestDepot : '';
  }

  function detectBridgeDepot() {
    const configured = getConfiguredBridgeDepot();
    if (configured) return configured;

    try {
      const qp = new URLSearchParams(location.search);
      const fromQuery =
        normalizeBridgeDepot(qp.get('depot')) ||
        normalizeBridgeDepot(qp.get('depotId')) ||
        normalizeBridgeDepot(qp.get('Depot'));
      if (fromQuery) return fromQuery;
    } catch {}

    if (state.inferredDepot) return state.inferredDepot;

    const titleMatch = norm(document.title).toUpperCase().match(/\bD\d{4}\b/);
    if (titleMatch) return titleMatch[0];

    return '';
  }

  function buildBridgeSignature(model, rows) {
    let hash = 2166136261 >>> 0;
    const push = (value) => {
      const text = norm(value);
      for (let i = 0; i < text.length; i++) {
        hash ^= text.charCodeAt(i);
        hash = Math.imul(hash, 16777619) >>> 0;
      }
      hash ^= 124;
      hash = Math.imul(hash, 16777619) >>> 0;
    };

    push(model.modeKey);
    push(String(rows.length));
    for (const row of rows) {
      for (const col of model.cols) {
        push(row && row[col.key] != null ? row[col.key] : '');
      }
    }
    return `${model.modeKey}#${rows.length}#${hash.toString(16)}`;
  }

  function syncBridgeSnapshot(force = false, trigger = 'auto') {
    const triggerLabel = bridgeTriggerLabel(trigger);
    if (!BRIDGE_ENABLED) {
      if (trigger === 'manual') {
        setBridgeStatus('Bridge deaktiviert', '#991b1b');
        addBridgeLog('warn', `${triggerLabel}: Bridge ist deaktiviert.`);
      }
      return;
    }

    if (bridgeState.inFlight) {
      if (trigger === 'manual') {
        setBridgeStatus('Senden laeuft bereits', '#b45309');
        addBridgeLog('warn', `${triggerLabel}: Export laeuft bereits, bitte kurz warten.`);
      }
      return;
    }

    if (!Array.isArray(state.rows) || !state.rows.length) {
      if (trigger === 'manual') {
        setBridgeStatus('Keine Daten erkannt', '#b45309');
        addBridgeLog('warn', `${triggerLabel}: keine Tabellenzeilen erkannt.`);
      }
      return;
    }

    const nowMs = Date.now();
    const sig = buildBridgeSignature(state.model, state.rows);
    if (!force && sig === bridgeState.lastSig && (nowMs - bridgeState.lastSentAt) < BRIDGE_MIN_INTERVAL_MS) {
      if (trigger === 'manual') {
        addBridgeLog('info', `${triggerLabel}: Daten unveraendert, manueller Send wurde trotzdem erzwungen.`);
      }
      return;
    }

    const depot = ensureBridgeDepotConfigured(false);
    if (!depot) {
      setBridgeStatus('Depot fehlt', '#991b1b');
      addBridgeLog('warn', `${triggerLabel}: kein Depot gesetzt, Export abgebrochen.`);
      return;
    }
    const columns = state.model.cols.map(c => ({ key: c.key, title: c.title }));
    const rows = state.rows.slice(0, BRIDGE_MAX_ROWS).map((r) => {
      const row = {};
      for (const c of state.model.cols) row[c.key] = norm(r[c.key]);
      return row;
    });
    const payload = {
      source: 'roadnet_transport_units',
      clientId: norm(BRIDGE_CLIENT_ID),
      mode: state.model.modeKey,
      title: state.model.title,
      stamp: state.lastStamp,
      depot,
      columns,
      rows,
      sourceUrl: location.href,
      exportedAt: new Date().toISOString()
    };
    const endpoints = buildBridgeEndpoints();
    if (!endpoints.length) {
      setBridgeStatus('Kein Endpoint', '#991b1b');
      addBridgeLog('error', `${triggerLabel}: kein Endpoint konfiguriert.`);
      return;
    }

    setBridgeStatus(`${triggerLabel}: Sende...`, '#1d4ed8');
    addBridgeLog(
      'info',
      `${triggerLabel}: Start mode=${state.model.modeKey} depot=${depot} rows=${rows.length} endpoints=${endpoints.length}`
    );

    bridgeState.inFlight = true;
    postBridgePayload(payload, endpoints, 0, trigger)
      .then((usedEndpoint) => {
        bridgeState.lastSig = sig;
        bridgeState.lastSentAt = nowMs;
        try {
          localStorage.setItem(BRIDGE_ENDPOINT_STORAGE_KEY, usedEndpoint);
        } catch {}
        setBridgeStatus(`${triggerLabel}: OK ${timeStampShort(new Date())}`, '#065f46');
        addBridgeLog('ok', `${triggerLabel}: Export erfolgreich ueber ${usedEndpoint}`);
        if (trigger === 'manual') {
          toast(`Test-Senden erfolgreich (${rows.length} Zeilen)`, true);
        }
      })
      .catch((err) => {
        const msg = norm(err && err.message ? err.message : String(err || 'send_failed'));
        setBridgeStatus(`${triggerLabel}: Fehler`, '#991b1b');
        addBridgeLog('error', `${triggerLabel}: Export fehlgeschlagen (${msg})`);
        if (trigger === 'manual') {
          toast('Test-Senden fehlgeschlagen', false);
        }
      })
      .finally(() => {
        bridgeState.inFlight = false;
      });
  }

  function findActiveMode() {
    const cands = Array.from(document.querySelectorAll('a,button,li,span,div'))
      .map(el => ({ el, t: norm(el.textContent) }))
      .filter(x => x.t === 'Beladung' || x.t === 'Entladung')
      .filter(x => isVisible(x.el));
    if (!cands.length) return null;

    const scoreEl = (el) => {
      const cls = String(el.className || '');
      let s = 0;
      if (/active|selected|ui-state-active|ui-tabs-active|current/i.test(cls)) s += 1000;
      const p = el.parentElement;
      if (p) {
        const clsP = String(p.className || '');
        if (/active|selected|ui-state-active|ui-tabs-active|current/i.test(clsP)) s += 800;
      }
      return s;
    };

    const best = cands.map(x => ({ ...x, score: scoreEl(x.el) })).sort((a, b) => b.score - a.score)[0];
    if (!best) return null;
    return best.t === 'Beladung' ? 'load' : 'unload';
  }

  function syncModelFromUI() {
    const mode = findActiveMode();
    const next = (mode === 'load') ? MODEL_LOAD : MODEL_UNLOAD;
    if (state.model.modeKey !== next.modeKey) {
      state.model = next;
      state.sort = { ...next.defaultSort };
    }
  }

  function getRowCells(rowEl) {
    if (!rowEl) return [];
    const cells = Array.from(rowEl.querySelectorAll('td,[role="gridcell"],[role="cell"]'));
    return cells.filter(isVisible);
  }

  function scoreRoadnetRoot(el) {
    if (!el || !isVisible(el) || el.closest(`#${PANEL_ID}`)) return -1;
    const bodyRows = Array.from(el.querySelectorAll('tbody tr,[role="row"]'))
      .filter((row) => getRowCells(row).length);
    const rowCount = bodyRows.length;
    if (!rowCount) return -1;

    const headerCells = Array.from(el.querySelectorAll('thead th,[role="columnheader"]'));
    const headerText = headerCells.map((h) => norm(h.innerText || h.textContent || '')).join(' | ').toLowerCase();
    const modeHint = norm(el.innerText || el.textContent || '').slice(0, 1800).toLowerCase();

    let score = rowCount;
    if (String(el.className || '').includes('ui-datatable')) score += 40;
    if (headerText.includes('status')) score += 50;
    if (headerText.includes('lts') || headerText.includes('transporteinheit') || headerText.includes('wb')) score += 50;
    if (headerText.includes('auslast') || headerText.includes('%') || headerText.includes('load')) score += 20;
    if (modeHint.includes('entladung') || modeHint.includes('beladung')) score += 10;
    return score;
  }

  function findRoadnetDataTableRoot() {
    const direct = document.getElementById('frm_transport_units:tbl');
    if (direct && isVisible(direct)) return direct;

    const selectors = [
      'div.ui-datatable[id$=":tbl"]',
      'div.ui-datatable',
      'table[role="grid"]',
      'table[aria-label*="Transport"]',
      'table[aria-label*="transport"]',
      'table[id*="transport"]',
      'table'
    ];

    const seen = new Set();
    const cands = [];
    for (const sel of selectors) {
      for (const el of Array.from(document.querySelectorAll(sel))) {
        if (seen.has(el)) continue;
        seen.add(el);
        cands.push(el);
      }
    }

    let best = null;
    let bestScore = -1;
    for (const el of cands) {
      const score = scoreRoadnetRoot(el);
      if (score > bestScore) {
        bestScore = score;
        best = el;
      }
    }
    return bestScore > 0 ? best : null;
  }

  function getPrimefacesTableParts(root) {
    const headerTh = Array.from(root.querySelectorAll('.ui-datatable-scrollable-header thead th')).filter(isVisible);
    let headerCells = headerTh.length
      ? headerTh
      : Array.from(root.querySelectorAll('thead th,[role="columnheader"]')).filter(isVisible);
    if (!headerCells.length && root.tagName && String(root.tagName).toLowerCase() === 'table') {
      const firstRow = root.querySelector('tr');
      if (firstRow) {
        headerCells = Array.from(firstRow.querySelectorAll('th,td')).filter(isVisible);
      }
    }

    const bodyRows1 = Array.from(root.querySelectorAll('.ui-datatable-scrollable-body tbody tr'))
      .filter(tr => getRowCells(tr).length);
    const bodyRows2 = Array.from(root.querySelectorAll('tbody tr,[role="row"]'))
      .filter(tr => getRowCells(tr).length);

    const combined = [...bodyRows1, ...bodyRows2];
    const uniq = [];
    const seenKeys = new Set();
    for (const r of combined) {
      const cells = getRowCells(r);
      const sig = (r.getAttribute('data-rk') || '') + '|' + cells.map((c) => norm(c.innerText || c.textContent || '')).join('|');
      if (seenKeys.has(sig)) continue;
      seenKeys.add(sig);
      uniq.push(r);
    }
    return { headerCells, bodyRows: uniq };
  }

  function buildHeaderIndexMap(headerCells) {
    const map = new Map();
    headerCells.forEach((h, idx) => {
      const k = normKey(norm(h?.innerText ?? h?.textContent ?? ''));
      if (k && !map.has(k)) map.set(k, idx);
    });
    return map;
  }

  function makePicker(headerMap) {
    return (candidates) => {
      for (const cand of candidates) {
        const k = normKey(cand);
        if (headerMap.has(k)) return headerMap.get(k);
        for (const [hk, idx] of headerMap.entries()) {
          if (hk.includes(k)) return idx;
        }
      }
      return null;
    };
  }

  function badgeHTML(text, opts = {}) {
    const bg = opts.bg || '#111';
    const fg = opts.fg || '#fff';
    const border = opts.border ? `border:1px solid ${opts.border};` : '';
    return `<span style="display:inline-flex;align-items:center;justify-content:center;min-width:16px;height:15px;padding:0 6px;margin-left:6px;border-radius:999px;background:${bg};color:${fg};${border}font:800 9.5px system-ui;line-height:15px;">${escHtml(text)}</span>`;
  }

  function computeBadges(rows) {
    const total = rows.length;
    const cols = state.model.cols;
    const res = {};

    const typeCounts = { N: 0, J: 0, T: 0 };
    for (const r of rows) {
      const t = normKey(r.type);
      if (!t) continue;
      if (t.includes('normal')) typeCounts.N++;
      if (t.includes('jumbo')) typeCounts.J++;
      if (t.includes('trailer')) typeCounts.T++;
    }
    const typeBadge = `N ${typeCounts.N}/J ${typeCounts.J}/T ${typeCounts.T}`;

    if (state.model.modeKey === 'unload') {
      const erfasstCnt = rows.reduce((a, r) => a + (normKey(r.status).includes('erfasst') ? 1 : 0), 0);
      const sumAll = rows.reduce((acc, r) => acc + (parsePercentToUnit(r.load) ?? 0), 0);
      const sumMissingEnd = rows.reduce((acc, r) => norm(r.unlEnd) ? acc : acc + (parsePercentToUnit(r.load) ?? 0), 0);
      const sumMinusMissingEnd = sumAll - sumMissingEnd;

      for (const c of cols) {
        if (c.key === 'status') {
          res[c.key] = { text: `${total}/${erfasstCnt}`, style: (erfasstCnt > 0 ? { bg: '#fee2e2', fg: '#991b1b', border: '#fecaca' } : null) };
          continue;
        }
        if (c.key === 'from') { res[c.key] = { text: String(total), style: null }; continue; }
        if (c.key === 'type') { res[c.key] = { text: typeBadge, style: null }; continue; }
        if (c.key === 'load') { res[c.key] = { text: `${fmtSumUnit(sumAll)}/${fmtSumUnit(sumMinusMissingEnd)}`, style: null }; continue; }

        const filled = rows.reduce((acc, r) => acc + (norm(r[c.key]) ? 1 : 0), 0);
        const empty = total - filled;
        res[c.key] = { text: `${filled}/${empty}`, style: null };
      }
      return res;
    }

    const sumAll = rows.reduce((acc, r) => acc + (parsePercentToUnit(r.loadPct) ?? 0), 0);

    for (const c of cols) {
      if (c.key === 'status')  { res[c.key] = { text: String(total), style: null }; continue; }
      if (c.key === 'toName')  { res[c.key] = { text: String(total), style: null }; continue; }
      if (c.key === 'type')    { res[c.key] = { text: typeBadge, style: null }; continue; }
      if (c.key === 'loadPct') { res[c.key] = { text: fmtSumUnit(sumAll), style: null }; continue; }

      const filled = rows.reduce((acc, r) => acc + (norm(r[c.key]) ? 1 : 0), 0);
      const empty = total - filled;
      res[c.key] = { text: `${filled}/${empty}`, style: null };
    }
    return res;
  }

  function sortRows(rows) {
    const { key, dir } = state.sort;
    const mul = dir === 'asc' ? 1 : -1;

    const getSortVal = (r) => {
      if (key === 'load' || key === 'loadPct') {
        const u = parsePercentToUnit(r[key]);
        return u == null ? -Infinity : u;
      }
      if (['depAct', 'arrAct', 'unlBeg', 'unlEnd', 'depPlan', 'loadBeg', 'loadEnd'].includes(key)) {
        const d = parseDeDateTime(r[key]);
        return d ? d.getTime() : -Infinity;
      }
      return norm(r[key]).toLowerCase();
    };

    return [...rows].sort((a, b) => {
      const va = getSortVal(a), vb = getSortVal(b);
      if (va < vb) return -1 * mul;
      if (va > vb) return 1 * mul;
      return 0;
    });
  }

  function computeScale(tableEl, containerEl) {
    if (!AUTOFIT_TO_WIDTH) return 1;
    if (!tableEl || !containerEl) return 1;

    const tableW = tableEl.scrollWidth || tableEl.getBoundingClientRect().width || 1;
    const contW = containerEl.getBoundingClientRect().width || 1;
    const maxW = Math.max(1, contW - 6);

    if (tableW <= maxW) return 1;
    const s = maxW / tableW;
    return Math.max(0.78, Math.min(1, s));
  }

  function buildClipboardText(rows) {
    const header1 = `Stand: ${state.lastStamp || ''}`;
    const header2 = state.model.title;
    const header = state.model.cols.map(c => c.title).join('\t');
    const lines = rows.map(r => state.model.cols.map(c => String(r[c.key] ?? '')).join('\t'));
    return [header1, header2, header, ...lines].join('\n');
  }

  function buildClipboardHTML(rows, badges) {
    const cols = state.model.cols;
    const thBg = '#eaf2ff';
    const border = '1px solid #d6dbe3';

    const badge = (t, style) => {
      const isRed = !!(style && style.bg && String(style.bg).includes('fee2e2'));
      const bg = isRed ? '#fee2e2' : '#111';
      const fg = isRed ? '#991b1b' : '#fff';
      const br = isRed ? '1px solid #fecaca' : 'none';
      return `<span style="display:inline-block;margin-left:6px;padding:0 6px;height:15px;line-height:15px;border-radius:999px;font:800 9.5px system-ui;background:${bg};color:${fg};border:${br};">${escHtml(t)}</span>`;
    };

    const thead = `
      <tr>
        ${cols.map(c => {
          const b = badges[c.key];
          const arrow = (state.sort.key === c.key) ? (state.sort.dir === 'asc' ? ' ▲' : ' ▼') : '';
          const bHtml = b ? badge(b.text, b.style) : '';
          return `<th style="background:${thBg};border:1px solid #d6dbe3;padding:3px 6px;white-space:nowrap;text-align:left;font:600 10px system-ui;">
                    ${escHtml(c.title)}${escHtml(arrow)}${bHtml}
                  </th>`;
        }).join('')}
      </tr>`;

    const tbody = rows.map((r, idx) => {
      const zebra = idx % 2 ? 'background:#fafafa;' : 'background:#fff;';
      const stIsErfasst =
        state.model.special.statusErfasstKey
          ? normKey(r[state.model.special.statusErfasstKey]).includes(state.model.special.statusErfasstText)
          : false;

      return `<tr style="${zebra}">
        ${cols.map(c => {
          let cellStyle = `border:1px solid #d6dbe3;padding:2px 6px;white-space:nowrap;font:500 10px system-ui;line-height:1.05;`;
          if (c.key === state.model.special.statusErfasstKey && stIsErfasst) cellStyle += 'background:rgba(220,38,38,.08);';
          const val = r[c.key];
          const inner = norm(val) ? escHtml(val) : '&mdash;';
          return `<td style="${cellStyle}">${inner}</td>`;
        }).join('')}
      </tr>`;
    }).join('');

    const meta = `
      <div style="font:800 11px system-ui;color:#111;margin:0 0 4px 0;">Stand: ${escHtml(state.lastStamp || '')}</div>
      <div style="font:800 11px system-ui;color:#111;margin:0 0 8px 0;">${escHtml(state.model.title)}</div>
    `;

    return `<div>${meta}<table style="border-collapse:collapse;"><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
  }

  async function copyRowsFormatted(rows) {
    const badges = computeBadges(rows);
    const html = buildClipboardHTML(rows, badges);
    const text = buildClipboardText(rows);

    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([text], { type: 'text/plain' })
      })
    ]);
  }

  /* ===================== LTS OPEN (Roadnet) ===================== */
  let __ltsWin = null;
  let __ltsPostTimer = null;
  let __ltsAckHandler = null;

  function openLtsLebenslaufFromLts(ltsRaw) {
    const lts = String(ltsRaw || '').replace(/[^\d]/g, '');
    if (lts.length !== 8) { toast('LTS # muss 8-stellig sein', false); return; }

    const a = lts.slice(0, 4);
    const b = lts.slice(4);

    const openUrl = 'https://lts.dpdit.de/index.aspx';
    try {
      if (!__ltsWin || __ltsWin.closed) {
        __ltsWin = window.open(openUrl, 'lts_lebenslauf');
      } else {
        __ltsWin.focus();
        try { __ltsWin.location.href = openUrl; } catch {}
      }
    } catch {
      __ltsWin = window.open(openUrl, 'lts_lebenslauf');
    }

    if (__ltsPostTimer) {
      clearInterval(__ltsPostTimer);
      __ltsPostTimer = null;
    }
    if (__ltsAckHandler) {
      window.removeEventListener('message', __ltsAckHandler);
      __ltsAckHandler = null;
    }

    let acked = false;
    const onAck = (ev) => {
      if (ev.origin !== 'https://lts.dpdit.de' && ev.origin !== 'http://lts.dpdit.de') return;
      const d = ev.data || {};
      if (d.type !== 'RN_WBL_ACK') return;
      acked = true;
      window.removeEventListener('message', onAck);
      __ltsAckHandler = null;
    };
    __ltsAckHandler = onAck;
    window.addEventListener('message', onAck);

    // Bis zu 10 Minuten erneut senden (wichtig bei manuellem Login + Sessionaufbau).
    let tries = 0;
    const maxTries = 600;
    const pump = () => {
      tries++;
      try {
        if (!__ltsWin || __ltsWin.closed) {
          if (__ltsPostTimer) clearInterval(__ltsPostTimer);
          __ltsPostTimer = null;
          window.removeEventListener('message', onAck);
          __ltsAckHandler = null;
          return;
        }
        if (acked) {
          if (__ltsPostTimer) clearInterval(__ltsPostTimer);
          __ltsPostTimer = null;
          window.removeEventListener('message', onAck);
          __ltsAckHandler = null;
          return;
        }
        __ltsWin.postMessage({ type: 'RN_WBL', a, b }, 'https://lts.dpdit.de');
        __ltsWin.postMessage({ type: 'RN_WBL', a, b }, 'http://lts.dpdit.de');
      } catch {}

      if (tries >= maxTries) {
        if (__ltsPostTimer) clearInterval(__ltsPostTimer);
        __ltsPostTimer = null;
        window.removeEventListener('message', onAck);
        __ltsAckHandler = null;
      }
    };

    pump();
    __ltsPostTimer = setInterval(pump, 1000);
  }

  /* ===================== UI ===================== */
  function ensureOpenButton() {
    let btn = document.getElementById(BTN_ID);
    if (btn) return btn;

    btn = document.createElement('button');
    btn.id = BTN_ID;
    btn.type = 'button';
    btn.textContent = 'Transporteinheiten';
    btn.style.cssText = [
      'position:fixed',
      'right:14px',
      'top:110px',
      'z-index:2147483001',
      'cursor:pointer',
      'border:1px solid rgba(0,0,0,.14)',
      'background:#fff',
      'border-radius:12px',
      'padding:6px 10px',
      'font:800 11px system-ui',
      'box-shadow:0 14px 40px rgba(0,0,0,.18)'
    ].join(';');

    btn.addEventListener('click', () => {
      const p = ensurePanel();
      p.style.display = (p.style.display === 'none') ? 'block' : 'none';
      if (p.style.display === 'block') extractData();
    });

    document.body.appendChild(btn);
    return btn;
  }

  function clickTabByExactText(targetText) {
    const target = norm(targetText);
    if (!target) return false;

    const candidates = Array.from(document.querySelectorAll('a,button,li,span,div,label'))
      .filter((el) => norm(el.textContent) === target && isVisible(el));
    if (!candidates.length) return false;

    const preferred =
      candidates.find((el) => el.tagName === 'A' || el.tagName === 'BUTTON') ||
      candidates[0];

    try {
      preferred.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      preferred.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      preferred.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      return true;
    } catch {
      try {
        preferred.click();
        return true;
      } catch {
        return false;
      }
    }
  }

  function ensureEntladungShortcutButton() {
    if (!location.href.includes('/execution/transport_units')) return null;
    if (window.__RN_OV_ENTL_RUNNING) return document.getElementById(BTN_OV_ID);
    window.__RN_OV_ENTL_RUNNING = true;

    const OVERLAY_ID = 'rnOvEntlOverlay';
    const STYLE_ID = 'rnOvEntlStyle';
    const OPEN_KEY = 'rnOvEntlOverlayOpen';

    const n = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
    const nk = (s) =>
      n(s)
        .toLowerCase()
        .replace(/\u00a0/g, ' ')
        .replace(/[^\p{L}\p{N}%# ]/gu, '')
        .replace(/\s+/g, ' ')
        .trim();

    const esc = (s) =>
      String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');

    const fmtStamp = () => {
      const d = new Date();
      const p = (x) => String(x).padStart(2, '0');
      return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
    };

    const fmtWB = (raw) => {
      const digits = n(raw).replace(/\D/g, '');
      if (digits.length === 8) return `${digits.slice(0, 4)}-${digits.slice(4)}`;
      return n(raw);
    };

    const fmtPct = (raw) => {
      const t = n(raw).replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, '');
      if (!t) return '';
      const v = parseFloat(t);
      if (!Number.isFinite(v)) return '';
      return `${Number.isInteger(v) ? String(v) : v.toFixed(1).replace('.', ',')}%`;
    };

    const isVisibleEl = (el) => {
      if (!el) return false;
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0;
    };

    const isOverlayVisible = () => {
      const el = document.getElementById(OVERLAY_ID);
      return !!el && getComputedStyle(el).display !== 'none';
    };

    const setOverlayOpen = (open) => {
      try { sessionStorage.setItem(OPEN_KEY, open ? '1' : '0'); } catch {}
    };
    const getOverlayOpen = () => {
      try { return sessionStorage.getItem(OPEN_KEY) === '1'; } catch { return false; }
    };

    function ensureEntladungTab() {
      const tabs = Array.from(document.querySelectorAll('a,button,li,span,div'))
        .filter((el) => n(el.textContent) === 'Entladung' && isVisibleEl(el));
      if (!tabs.length) return false;

      const activeRe = /ui-state-active|ui-tabs-active|ui-tabs-selected|active|selected|current/i;
      const active = tabs.some((el) =>
        activeRe.test(String(el.className || '')) ||
        activeRe.test(String(el.parentElement?.className || '')) ||
        !!el.closest?.('.ui-state-active, .ui-tabs-active, .ui-tabs-selected, .active, .selected, .current')
      );
      if (active) return true;

      const el = tabs.find((x) => x.tagName === 'A' || x.tagName === 'BUTTON') || tabs[0];
      try {
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        return true;
      } catch {
        try { el.click(); return true; } catch { return false; }
      }
    }

    function clickRefreshButton() {
      const cands = Array.from(document.querySelectorAll('button,a,span'))
        .filter(isVisibleEl)
        .filter((el) => {
          const t = n(el.textContent);
          const ti = n(el.getAttribute('title'));
          const ar = n(el.getAttribute('aria-label'));
          const cl = String(el.className || '');
          return /aktualisieren/i.test(t) || /refresh|reload|aktualisieren/i.test(ti) || /refresh|reload|aktualisieren/i.test(ar) || /refresh/i.test(cl);
        });

      const btn = cands.find((el) => el.tagName === 'BUTTON') || cands[0];
      if (!btn) return false;
      try {
        btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
        btn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        return true;
      } catch {
        try { btn.click(); return true; } catch { return false; }
      }
    }

    function tableRoot() {
      const direct = document.getElementById('frm_transport_units:tbl');
      if (direct && direct.classList.contains('ui-datatable')) return direct;

      const cands = Array.from(document.querySelectorAll('div.ui-datatable[id$=":tbl"]')).filter(isVisibleEl);
      if (!cands.length) return null;
      if (cands.length === 1) return cands[0];

      let best = cands[0];
      let bestRows = -1;
      for (const c of cands) {
        const rows = c.querySelectorAll('tbody tr').length;
        if (rows > bestRows) {
          bestRows = rows;
          best = c;
        }
      }
      return best;
    }

    function headerMap(root) {
      const heads = Array.from(root.querySelectorAll('.ui-datatable-scrollable-header thead th')).filter(isVisibleEl);
      const h = heads.length ? heads : Array.from(root.querySelectorAll('thead th'));
      const map = new Map();
      h.forEach((th, idx) => {
        const key = nk(n(th.innerText ?? th.textContent ?? ''));
        if (key && !map.has(key)) map.set(key, idx);
      });
      return map;
    }

    function pickIdx(hm, keys) {
      for (const k of keys) {
        const key = nk(k);
        if (hm.has(key)) return hm.get(key);
        for (const [hk, idx] of hm.entries()) {
          if (hk.includes(key)) return idx;
        }
      }
      return null;
    }

    function cellText(td) {
      if (!td) return '';
      const main = n(td.innerText ?? td.textContent ?? '');
      if (/[A-Za-zÄÖÜäöü]/.test(main)) return main;
      const attrs = [
        td.getAttribute('title'),
        td.getAttribute('aria-label'),
        td.getAttribute('data-original-title'),
        td.getAttribute('data-tooltip'),
        td.getAttribute('data-title')
      ].map(n).filter(Boolean);
      return attrs.find((x) => /[A-Za-zÄÖÜäöü]/.test(x)) || main;
    }

    function extractRows() {
      const root = tableRoot();
      if (!root) return [];

      const hm = headerMap(root);
      const idx = {
        status: pickIdx(hm, ['status']),
        carrier: pickIdx(hm, ['frachtführer', 'frachtfuehrer', 'carrier', 'dienstleister', 'transporteur', 'unternehmer', 'spedition']),
        from: pickIdx(hm, ['herkunft', 'abgangsort', 'abgang', 'von', 'name abgangsstandort', 'code abgangsstation']),
        wb: pickIdx(hm, ['wb nr', 'nummer transporteinheit', 'lts #', 'lts', 'nummer transporte', 'transporteinheit']),
        fuel: pickIdx(hm, ['auslastung', 'auslastung (%)', 'fuellstand', 'füllstand']),
        unlBeg: pickIdx(hm, ['entladung beginn', 'entladebeginn']),
        unlEnd: pickIdx(hm, ['entladung ende', 'entladeende']),
        traffic: pickIdx(hm, ['verkehrsart', 'art', 'transportart'])
      };

      const rowsA = Array.from(root.querySelectorAll('.ui-datatable-scrollable-body tbody tr')).filter((tr) => tr.querySelectorAll('td').length);
      const rowsB = Array.from(root.querySelectorAll('tbody tr')).filter((tr) => tr.querySelectorAll('td').length);
      const merged = [...rowsA, ...rowsB];

      const uniq = [];
      const seen = new Set();
      for (const tr of merged) {
        const sig = (tr.getAttribute('data-rk') || '') + '|' + n(tr.innerText).slice(0, 180);
        if (seen.has(sig)) continue;
        seen.add(sig);
        uniq.push(tr);
      }

      const out = [];
      for (const tr of uniq) {
        const tds = Array.from(tr.querySelectorAll('td'));
        if (!tds.length) continue;
        const get = (k) => {
          const i = idx[k];
          if (i == null || i < 0 || i >= tds.length) return '';
          return cellText(tds[i]);
        };
        const row = {
          status: get('status'),
          carrier: get('carrier'),
          from: get('from'),
          wb: get('wb'),
          fuel: get('fuel'),
          unlBeg: get('unlBeg'),
          unlEnd: get('unlEnd'),
          traffic: get('traffic')
        };
        if (Object.values(row).some((v) => n(v))) out.push(row);
      }
      return out;
    }

    const isPickup = (row) => {
      const t = nk(row.traffic);
      return t.includes('pickup') || t === 'pickup';
    };
    const isUnloaded = (row) => {
      const s = nk(row.status);
      return s.includes('entladen') || !!n(row.unlEnd);
    };
    const isAtGate = (row) => !!n(row.unlBeg) && !n(row.unlEnd);

    function ensureUI() {
      if (!document.getElementById(BTN_OV_ID)) {
        const btn = document.createElement('button');
        btn.id = BTN_OV_ID;
        btn.type = 'button';
        btn.textContent = 'ÜbEntl';
        btn.style.cssText = 'position:fixed;right:14px;top:168px;z-index:2147483647;cursor:pointer;border:1px solid rgba(255,255,255,.14);background:rgba(10,15,25,.62);color:#eaf2ff;border-radius:9px;padding:2px 5px;font:900 8.5px system-ui;line-height:1;letter-spacing:.2px;box-shadow:0 8px 22px rgba(0,0,0,.20)';
        btn.addEventListener('click', () => toggleOverlay());
        document.body.appendChild(btn);
      }

      if (!document.getElementById(STYLE_ID)) {
        const style = document.createElement('style');
        style.id = STYLE_ID;
        style.textContent = `
          #${OVERLAY_ID}{position:fixed;inset:0;z-index:2147483646;display:none;color:#eef6ff;background:radial-gradient(1000px 600px at 20% -10%,rgba(71,120,220,.27),rgba(8,12,20,0) 60%),linear-gradient(180deg,#0b1220,#070c16 55%,#060a13);font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;}
          .ovTop{height:64px;display:flex;align-items:center;justify-content:space-between;padding:10px 18px;border-bottom:1px solid rgba(255,255,255,.08);background:rgba(5,10,18,.35);}
          .ovTitle{font-weight:950;font-size:clamp(18px,1.6vw,34px)}
          .ovMeta{display:flex;gap:16px;align-items:center;color:rgba(234,242,255,.78);font-weight:800;font-size:clamp(12px,1vw,18px)}
          .ovClose{cursor:pointer;border:1px solid rgba(255,255,255,.14);background:rgba(255,255,255,.06);color:#fff;border-radius:14px;padding:8px 12px;font-weight:950;font-size:clamp(12px,1vw,18px)}
          .ovBody{height:calc(100vh - 64px);padding:16px 18px 18px;box-sizing:border-box;display:grid;grid-template-columns:1.05fr 2.2fr 2.2fr;grid-template-rows:auto auto 1fr;gap:14px;}
          .ovCard{border:1px solid rgba(255,255,255,.08);background:rgba(9,15,30,.55);border-radius:16px;box-shadow:0 22px 70px rgba(0,0,0,.32);overflow:hidden;min-width:0;}
          .ovKpi{grid-column:1 / 4;display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
          .ovKInner{padding:16px 16px 12px}
          .ovLbl{color:rgba(234,242,255,.78);font-weight:900;font-size:clamp(12px,1.3vw,25px)}
          .ovVal{margin-top:6px;font-weight:1000;line-height:1;font-size:clamp(42px,6.2vw,140px)}
          .ovSub{margin-top:6px;font-weight:850;font-size:clamp(10px,.85vw,14px);color:rgba(234,242,255,.70)}
          .ovProg{grid-column:1 / 4;padding:12px 14px}
          .ovBar{height:16px;border-radius:999px;background:rgba(255,255,255,.10);overflow:hidden;border:1px solid rgba(255,255,255,.10)}
          .ovBarIn{height:100%;width:0;background:linear-gradient(90deg,rgba(34,197,94,.95),rgba(59,130,246,.95));}
          .ovProgText{display:flex;justify-content:space-between;margin-top:8px;font-weight:900;color:rgba(234,242,255,.78);font-size:clamp(11px,.95vw,16px)}
          .ovSide{grid-column:1;display:flex;flex-direction:column;min-height:0;}
          .ovMain1{grid-column:2;min-height:0;}
          .ovMain2{grid-column:3;min-height:0;}
          .ovHead{padding:14px 14px 12px;border-bottom:1px solid rgba(255,255,255,.08);display:flex;justify-content:space-between;align-items:center}
          .ovHTitle{font-weight:1000;font-size:clamp(16px,1.7vw,30px)}
          .ovHCount{font-weight:1000;font-size:clamp(16px,1.7vw,30px);color:rgba(234,242,255,.88)}
          .ovList{padding:14px;display:flex;flex-direction:column;gap:10px;overflow:auto;max-height:100%}
          .ovRow{display:flex;align-items:center;justify-content:space-between;padding:10px 12px;border-radius:14px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.08)}
          .ovRowN{font-weight:950;font-size:clamp(13px,1.2vw,22px);word-break:break-word;padding-right:10px}
          .ovRowC{font-weight:1000;font-size:clamp(16px,1.8vw,30px)}
          .ovTableWrap{height:100%;display:flex;flex-direction:column;min-height:0}
          .ovTableBody{overflow:auto;min-height:0;background:rgba(9,15,30,.10)}
          .ovTable{width:100%;border-collapse:separate;border-spacing:0;table-layout:fixed}
          .ovTable th{position:sticky;top:0;z-index:2;background:rgba(12,18,32,.96);border-bottom:1px solid rgba(255,255,255,.12);text-align:left;padding:10px 12px;font-weight:950;font-size:clamp(13px,1.2vw,22px)}
          .ovTable td{border-bottom:1px solid rgba(255,255,255,.08);padding:10px 12px;font-weight:900;font-size:clamp(12px,1.2vw,20px);line-height:1.06;word-break:break-word}
          .ovZ0{background:rgba(255,255,255,0)} .ovZ1{background:rgba(255,255,255,.03)}
          .ovWB{white-space:nowrap;font-variant-numeric:tabular-nums}
          .ovFuel{white-space:nowrap}
          .ovBadge{display:inline-flex;align-items:center;justify-content:center;padding:4px 8px;border-radius:999px;font-weight:1000;font-size:clamp(10px,.95vw,16px);border:1px solid rgba(255,255,255,.12);word-break:break-word}
          .ovBRed{background:rgba(239,68,68,.22);color:#ffd7d7;border-color:rgba(239,68,68,.40)}
          .ovBGreen{background:rgba(34,197,94,.20);color:#d6ffe5;border-color:rgba(34,197,94,.40)}
          .ovBGray{background:rgba(148,163,184,.16);color:rgba(234,242,255,.90);border-color:rgba(148,163,184,.28)}
        `;
        document.head.appendChild(style);
      }

      if (!document.getElementById(OVERLAY_ID)) {
        const wrap = document.createElement('div');
        wrap.id = OVERLAY_ID;
        wrap.innerHTML = `
          <div class="ovTop">
            <div class="ovTitle">Frühschicht – Übersicht Entladung</div>
            <div class="ovMeta">
              <div>Stand: <span id="ovStamp">—</span></div>
              <button class="ovClose" id="ovClose" type="button">Schließen</button>
            </div>
          </div>
          <div class="ovBody">
            <div class="ovKpi">
              <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Avisierte TE</div><div class="ovVal" id="ovKpiTotal">0</div><div class="ovSub">Pickup ignoriert</div></div></div>
              <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Entladen</div><div class="ovVal" id="ovKpiDone">0</div><div class="ovSub">Status „Entladen“ oder Entladeende ✓</div></div></div>
              <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Noch offen</div><div class="ovVal" id="ovKpiOpen">0</div><div class="ovSub">Ohne Entladeende</div></div></div>
              <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Am Tor</div><div class="ovVal" id="ovKpiGate">0</div><div class="ovSub">Entladebeginn ✓ / Entladeende ✗</div></div></div>
            </div>
            <div class="ovCard ovProg">
              <div class="ovBar"><div class="ovBarIn" id="ovBarIn"></div></div>
              <div class="ovProgText"><span id="ovProgText">Fortschritt: 0%</span><span></span></div>
            </div>
            <div class="ovCard ovSide">
              <div class="ovHead"><div class="ovHTitle">Status-Zählung</div><div class="ovHCount" id="ovStatusTotal">—</div></div>
              <div class="ovList" id="ovStatusList"></div>
            </div>
            <div class="ovCard ovMain1">
              <div class="ovTableWrap">
                <div class="ovHead"><div class="ovHTitle">Noch zu entladen</div><div class="ovHCount" id="ovOpenCount">0</div></div>
                <div class="ovTableBody" id="ovOpenBody"></div>
              </div>
            </div>
            <div class="ovCard ovMain2">
              <div class="ovTableWrap">
                <div class="ovHead"><div class="ovHTitle">Am Tor</div><div class="ovHCount" id="ovGateCount">0</div></div>
                <div class="ovTableBody" id="ovGateBody"></div>
              </div>
            </div>
          </div>
        `;
        document.body.appendChild(wrap);

        document.getElementById('ovClose').addEventListener('click', () => {
          const el = document.getElementById(OVERLAY_ID);
          if (el) el.style.display = 'none';
          setOverlayOpen(false);
          stopAutoScroll();
        });
      }
    }

    const badgeHtml = (status) => {
      const s = nk(status);
      if (s.includes('abgefahren')) return `<span class="ovBadge ovBRed">Abgefahren</span>`;
      if (s.includes('angekommen')) return `<span class="ovBadge ovBGreen">Angekommen</span>`;
      return `<span class="ovBadge ovBGray">${esc(status || '—')}</span>`;
    };

    function buildTable(rows, mode) {
      const cols = mode === 'gate'
        ? [
            { key: 'wb', label: 'WB Nr', cls: 'ovWB' },
            { key: 'from', label: 'Herkunft', cls: '' },
            { key: 'fuel', label: 'Ausl…', cls: 'ovFuel' },
            { key: 'status', label: 'Status', cls: '' }
          ]
        : [
            { key: 'wb', label: 'WB Nr', cls: 'ovWB' },
            { key: 'from', label: 'Herkunft', cls: '' },
            { key: 'fuel', label: 'Ausl…', cls: 'ovFuel' },
            { key: 'status', label: 'Status', cls: '' },
            { key: 'carrier', label: 'Frachtf.', cls: '' }
          ];

      const colStyle = mode === 'gate'
        ? '<colgroup><col style="width:20%"><col style="width:44%"><col style="width:16%"><col style="width:20%"></colgroup>'
        : '<colgroup><col style="width:18%"><col style="width:36%"><col style="width:14%"><col style="width:20%"><col style="width:12%"></colgroup>';

      return `
        <table class="ovTable">
          ${colStyle}
          <thead><tr>${cols.map((c) => `<th>${esc(c.label)}</th>`).join('')}</tr></thead>
          <tbody>
            ${rows.map((r, i) => {
              const z = i % 2 ? 'ovZ1' : 'ovZ0';
              return `<tr class="${z}">${cols.map((c) => {
                if (c.key === 'status') return `<td class="${c.cls}">${badgeHtml(r.status)}</td>`;
                if (c.key === 'wb') return `<td class="${c.cls}">${esc(fmtWB(r.wb) || '—')}</td>`;
                if (c.key === 'fuel') return `<td class="${c.cls}">${esc(fmtPct(r.fuel) || '—')}</td>`;
                return `<td class="${c.cls}">${esc(n(r[c.key]) || '—')}</td>`;
              }).join('')}</tr>`;
            }).join('')}
          </tbody>
        </table>
      `;
    }

    let rendering = false;
    function renderOverlay() {
      if (rendering) return;
      rendering = true;
      try {
        ensureUI();
        ensureEntladungTab();

        const stamp = document.getElementById('ovStamp');
        if (stamp) stamp.textContent = fmtStamp();

        const all = extractRows();
        const eligible = all.filter((r) => !isPickup(r));
        const total = eligible.length;
        const done = eligible.filter(isUnloaded).length;
        const gate = eligible.filter(isAtGate);
        const open = eligible.filter((r) => !isUnloaded(r));
        const openOnly = open.filter((r) => !isAtGate(r));

        document.getElementById('ovKpiTotal').textContent = String(total);
        document.getElementById('ovKpiDone').textContent = String(done);
        document.getElementById('ovKpiOpen').textContent = String(open.length);
        document.getElementById('ovKpiGate').textContent = String(gate.length);

        const pct = total > 0 ? Math.round((done / total) * 100) : 0;
        document.getElementById('ovBarIn').style.width = `${pct}%`;
        document.getElementById('ovProgText').textContent = `Fortschritt: ${pct}% (${done}/${total})`;

        const sc = new Map();
        for (const r of eligible) {
          const key = n(r.status) || '—';
          sc.set(key, (sc.get(key) || 0) + 1);
        }
        const statusEntries = Array.from(sc.entries()).sort((a, b) => b[1] - a[1]);
        document.getElementById('ovStatusList').innerHTML = statusEntries.length
          ? statusEntries.map(([k, v]) => `<div class="ovRow"><div class="ovRowN">${esc(k)}</div><div class="ovRowC">${v}</div></div>`).join('')
          : '<div style="padding:8px 6px;color:rgba(234,242,255,.75);font-weight:900;">Keine Daten.</div>';
        document.getElementById('ovStatusTotal').textContent = String(eligible.length);

        document.getElementById('ovOpenBody').innerHTML = buildTable(openOnly, 'open');
        document.getElementById('ovGateBody').innerHTML = buildTable(gate, 'gate');
        document.getElementById('ovOpenCount').textContent = String(openOnly.length);
        document.getElementById('ovGateCount').textContent = String(gate.length);
      } finally {
        rendering = false;
      }
    }


    const AUTO_SCROLL = { enabled: true, step: 0.6, intervalMs: 25, pauseMs: 1200 };
    let autoScrollTimer = null;
    let autoScrollState = new Map();

    function getAutoScrollHosts() {
      const hosts = [];
      const a = document.getElementById('ovStatusList');
      const b = document.getElementById('ovOpenBody');
      const c = document.getElementById('ovGateBody');
      if (a) hosts.push(a);
      if (b) hosts.push(b);
      if (c) hosts.push(c);
      return hosts;
    }

    function isScrollableHost(el) {
      return !!el && (el.scrollHeight - el.clientHeight) > 6;
    }

    function tickAutoScroll() {
      if (!AUTO_SCROLL.enabled) return;
      if (!isOverlayVisible()) return;

      const hosts = getAutoScrollHosts();
      const now = Date.now();

      for (const el of hosts) {
        if (!isScrollableHost(el)) {
          autoScrollState.delete(el);
          continue;
        }

        let st = autoScrollState.get(el);
        if (!st) {
          st = { dir: 1, nextMoveAt: now + AUTO_SCROLL.pauseMs };
          autoScrollState.set(el, st);
        }

        if (now < st.nextMoveAt) continue;

        const maxScroll = el.scrollHeight - el.clientHeight;
        let nextTop = el.scrollTop + st.dir * AUTO_SCROLL.step;

        if (nextTop <= 0) {
          nextTop = 0;
          st.dir = 1;
          st.nextMoveAt = now + AUTO_SCROLL.pauseMs;
        } else if (nextTop >= maxScroll) {
          nextTop = maxScroll;
          st.dir = -1;
          st.nextMoveAt = now + AUTO_SCROLL.pauseMs;
        }

        el.scrollTop = nextTop;
      }
    }

    function startAutoScroll() {
      stopAutoScroll();
      autoScrollTimer = setInterval(tickAutoScroll, AUTO_SCROLL.intervalMs);
    }

    function stopAutoScroll() {
      if (autoScrollTimer) clearInterval(autoScrollTimer);
      autoScrollTimer = null;
      autoScrollState = new Map();
    }
    function toggleOverlay() {
      ensureUI();
      ensureEntladungTab();
      const el = document.getElementById(OVERLAY_ID);
      if (!el) return;
      const vis = getComputedStyle(el).display !== 'none';
      el.style.display = vis ? 'none' : 'block';
      setOverlayOpen(!vis);
      if (!vis) {
        renderOverlay();
        startAutoScroll();
      } else {
        stopAutoScroll();
      }
    }

    const renderDebounced = (() => {
      const fn = () => { if (isOverlayVisible()) renderOverlay(); };
      return debounce(fn, 250);
    })();

    function installObserver() {
      const root = tableRoot();
      if (!root) return false;
      const mo = new MutationObserver(() => renderDebounced());
      mo.observe(root, { childList: true, subtree: true, characterData: true, attributes: true });
      return true;
    }

    ensureUI();

    setTimeout(() => {
      ensureEntladungTab();
      if (getOverlayOpen()) {
        const el = document.getElementById(OVERLAY_ID);
        if (el) el.style.display = 'block';
        renderOverlay();
        startAutoScroll();
      }
    }, 800);

    let tries = 0;
    const obsTry = setInterval(() => {
      tries++;
      if (installObserver()) clearInterval(obsTry);
      if (tries >= 40) clearInterval(obsTry);
    }, 500);

    setInterval(() => {
      if (!isOverlayVisible()) return;
      setOverlayOpen(true);
      if (!ensureEntladungTab()) return;
      clickRefreshButton();
      setTimeout(() => { if (isOverlayVisible()) renderOverlay(); }, 1400);
    }, 20000);

    setInterval(() => { if (isOverlayVisible()) renderOverlay(); }, 4000);

    return document.getElementById(BTN_OV_ID);
  }

  function ensurePanel() {
    let panel = document.getElementById(PANEL_ID);
    if (panel) return panel;

    panel = document.createElement('div');
    panel.id = PANEL_ID;
    panel.style.cssText = [
      'position:fixed',
      'left:50%',
      'transform:translateX(-50%)',
      'top:150px',
      'width:min(1180px, calc(100vw - 18px))',
      'max-height:calc(100vh - 185px)',
      'overflow:hidden',
      'z-index:2147483000',
      'background:#fff',
      'border:1px solid rgba(0,0,0,.12)',
      'border-radius:14px',
      'box-shadow:0 20px 60px rgba(0,0,0,.22)',
      'font:11px system-ui',
      'display:none'
    ].join(';');

    panel.innerHTML = `
      <div class="${NS}head" style="position:sticky;top:0;z-index:2;background:#fff;padding:6px 8px;border-bottom:1px solid rgba(0,0,0,.08);">
        <div style="display:flex;align-items:center;gap:10px;justify-content:space-between;">
          <div>
            <div class="${NS}stand" style="font:800 11px system-ui;color:#111;">
              Stand: <span class="${NS}stamp" style="font:800 11px system-ui;color:#111;"></span>
            </div>
            <div class="${NS}sub" style="font:800 11px system-ui;color:#111;margin-top:2px;"></div>
            <div class="${NS}depotline" style="font:700 10px system-ui;color:#374151;margin-top:2px;">
              Depot: <span class="${NS}depotvalue" style="font:800 10px system-ui;color:#991b1b;">nicht gesetzt</span>
              <button class="${NS}btn-depot" type="button" style="cursor:pointer;margin-left:6px;border:1px solid rgba(0,0,0,.12);background:#fff;border-radius:8px;padding:1px 6px;font:700 10px system-ui;">
                Depot setzen
              </button>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:8px;">
            <button class="${NS}btn-send-test" type="button" style="cursor:pointer;border:1px solid rgba(0,0,0,.12);background:#ecfeff;border-radius:10px;padding:4px 8px;font:800 10px system-ui;color:#0f766e;">
              Senden (Test)
            </button>
            <button class="${NS}btn-copy" type="button" style="cursor:pointer;border:1px solid rgba(0,0,0,.12);background:#fff;border-radius:10px;padding:4px 8px;font:800 10px system-ui;">
              In Zwischenablage kopieren
            </button>
            <button class="${NS}btn-close" type="button" title="Schließen" style="cursor:pointer;border:1px solid rgba(0,0,0,.12);background:#fff;border-radius:10px;padding:4px 8px;font:900 10px system-ui;">
              ✕
            </button>
          </div>
        </div>
      </div>
      <div class="${NS}body" style="padding:6px 8px;">
        <div class="${NS}hint" style="color:#666;font:700 10px system-ui;margin:2px 0 6px 0;"></div>
        <div class="${SCROLL_BOX_CLASS}" style="overflow:auto;max-height:calc(100vh - 265px);border:1px solid rgba(0,0,0,.10);border-radius:12px;"></div>
      </div>
    `;

    document.body.appendChild(panel);

    panel.querySelector(`.${NS}btn-close`).addEventListener('click', () => (panel.style.display = 'none'));

    panel.querySelector(`.${NS}btn-copy`).addEventListener('click', async () => {
      try {
        const rowsSorted = sortRows(state.rows);
        await copyRowsFormatted(rowsSorted);
        toast('Tabelle formatiert kopiert', true);
      } catch {
        try {
          await navigator.clipboard.writeText(buildClipboardText(sortRows(state.rows)));
          toast('Als Text kopiert (Fallback)', true);
        } catch {
          toast('Kopieren fehlgeschlagen', false);
        }
      }
    });

    panel.querySelector(`.${NS}btn-send-test`).addEventListener('click', () => {
      extractData({ skipBridgeSync: true });
      syncBridgeSnapshot(true, 'manual');
    });

    panel.querySelector(`.${NS}btn-depot`).addEventListener('click', () => {
      ensureBridgeDepotConfigured(true);
      render();
    });

    const sb = panel.querySelector(`.${SCROLL_BOX_CLASS}`);
    sb.addEventListener('scroll', () => {
      state.scrollLeft = sb.scrollLeft;
      state.scrollTop = sb.scrollTop;
    });

    return panel;
  }

  function render() {
    const panel = ensurePanel();
    const sb = panel.querySelector(`.${SCROLL_BOX_CLASS}`);
    const hint = panel.querySelector(`.${NS}hint`);
    const sub = panel.querySelector(`.${NS}sub`);
    const stamp = panel.querySelector(`.${NS}stamp`);

    stamp.textContent = state.lastStamp || '';
    sub.textContent = state.model.title;
    updateDepotUI();
    const depot = getConfiguredBridgeDepot();

    const rows = sortRows(state.rows);
    const badges = computeBadges(rows);
    const cols = state.model.cols;

    hint.textContent = rows.length
      ? `Zeilen: ${rows.length} | Depot: ${depot || 'nicht gesetzt'} | Sortierung: ${cols.find(x => x.key === state.sort.key)?.title ?? state.sort.key} (${state.sort.dir})`
      : `Keine Daten erkannt.${depot ? ` Depot: ${depot}` : ' Depot nicht gesetzt.'}`;

    const thStyle =
      'position:sticky;top:0;background:linear-gradient(#eef5ff,#eaf2ff);' +
      'border-bottom:1px solid rgba(0,0,0,.10);border-right:1px solid rgba(0,0,0,.06);' +
      'padding:3px 4px;text-align:left;white-space:nowrap;font:600 10px system-ui;color:#111;cursor:pointer;';
    const tdBase =
      'border-bottom:1px solid rgba(0,0,0,.06);border-right:1px solid rgba(0,0,0,.05);' +
      'padding:2px 4px;white-space:nowrap;vertical-align:middle;font:500 10px system-ui;color:#111;line-height:1.05;';
    const tdDim = 'color:#777;font:500 10px system-ui;';
    const rowCopyBtnStyle =
      'cursor:pointer;border:1px solid rgba(0,0,0,.10);background:#fff;border-radius:9px;padding:1px 6px;font:700 10px system-ui;';
    const ltsLinkStyle =
      'cursor:pointer;color:#2563eb;text-decoration:underline;font:700 10px system-ui;';

    const oldLeft = state.scrollLeft;
    const oldTop = state.scrollTop;

    const tableHTML = `
      <div class="${TABLE_WRAP_CLASS}" style="transform-origin:top left;">
        <table class="${NS}tbl" style="border-collapse:separate;border-spacing:0;width:100%;">
          <thead>
            <tr>
              ${cols.map(c => {
                const arrow = state.sort.key === c.key ? (state.sort.dir === 'asc' ? ' ▲' : ' ▼') : '';
                const b = badges[c.key];
                const badge = b ? badgeHTML(b.text, b.style || {}) : '';
                return `<th data-key="${c.key}" style="${thStyle}">${c.title}${arrow}${badge}</th>`;
              }).join('')}
              <th style="${thStyle};cursor:default;border-right:0;">Aktion</th>
            </tr>
          </thead>
          <tbody>
            ${rows.map((r, i) => {
              const zebra = i % 2 ? 'background:#fafafa;' : 'background:#fff;';
              const stIsErfasst =
                state.model.special.statusErfasstKey
                  ? normKey(r[state.model.special.statusErfasstKey]).includes(state.model.special.statusErfasstText)
                  : false;

              return `
                <tr style="${zebra}">
                  ${cols.map(c => {
                    let extra = '';
                    if (c.key === state.model.special.statusErfasstKey && stIsErfasst) extra += 'background:rgba(220,38,38,.08);';

                    const val = r[c.key];
                    let cell = norm(val) ? val : `<span style="${tdDim}">—</span>`;

                    if (c.key === 'lts' && norm(val)) {
                      const clean = String(val).replace(/[^\d]/g, '');
                      cell = `<span class="${NS}ltslink" data-lts="${clean}" style="${ltsLinkStyle}">${val}</span>`;
                    }

                    return `<td style="${tdBase}${extra}">${cell}</td>`;
                  }).join('')}
                  <td style="${tdBase};border-right:0;">
                    <button class="${NS}rowcopy" data-i="${i}" type="button" style="${rowCopyBtnStyle}">Kopieren</button>
                  </td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      </div>
    `;

    sb.innerHTML = tableHTML;

    const tw = sb.querySelector(`.${TABLE_WRAP_CLASS}`);
    const tbl = sb.querySelector(`table.${NS}tbl`);
    const scale = computeScale(tbl, sb);
    state.scale = scale;
    if (tw) tw.style.transform = `scale(${scale})`;

    if (scale < 1 && tw) {
      const naturalH = tw.getBoundingClientRect().height / scale;
      const scaledH = naturalH * scale;
      const diff = naturalH - scaledH;
      tw.style.paddingBottom = `${Math.max(0, diff)}px`;
    }

    requestAnimationFrame(() => {
      sb.scrollLeft = oldLeft;
      sb.scrollTop = oldTop;
    });

    sb.querySelectorAll('th[data-key]').forEach(th => {
      th.addEventListener('click', () => {
        const key = th.getAttribute('data-key');
        if (!key) return;
        if (state.sort.key === key) state.sort.dir = (state.sort.dir === 'asc' ? 'desc' : 'asc');
        else { state.sort.key = key; state.sort.dir = 'asc'; }
        render();
      });
    });

    sb.querySelectorAll(`.${NS}ltslink`).forEach(el => {
      el.addEventListener('click', (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        const lts = el.getAttribute('data-lts') || '';
        openLtsLebenslaufFromLts(lts);
      });
    });

    sb.querySelectorAll(`button.${NS}rowcopy`).forEach(btn => {
      btn.addEventListener('click', async () => {
        const idx = Number(btn.getAttribute('data-i'));
        const r = rows[idx];
        if (!r) return;
        try {
          await copyRowsFormatted([r]);
          toast('Zeile formatiert kopiert', true);
        } catch {
          try {
            await navigator.clipboard.writeText(cols.map(c => String(r[c.key] ?? '')).join('\t'));
            toast('Zeile als Text kopiert (Fallback)', true);
          } catch {
            toast('Kopieren fehlgeschlagen', false);
          }
        }
      });
    });
  }

  function extractData(opts = {}) {
    const skipBridgeSync = !!(opts && opts.skipBridgeSync);
    const panel = document.getElementById(PANEL_ID);
    const shouldRender = !!(panel && panel.style.display !== 'none');

    syncModelFromUI();

    const root = findRoadnetDataTableRoot();
    if (!root) {
      state.rows = [];
      state.lastStamp = '';
      if (shouldRender) render();
      return;
    }

    const { headerCells, bodyRows } = getPrimefacesTableParts(root);
    const headerMap = buildHeaderIndexMap(headerCells);
    const pick = makePicker(headerMap);
    const idx = state.model.resolve(headerMap, pick);
    const inferredDepot = inferBridgeDepotFromTable(headerMap, bodyRows);
    if (inferredDepot) state.inferredDepot = inferredDepot;

    const out = [];
    for (const tr of bodyRows) {
      const tds = getRowCells(tr);
      if (!tds.length) continue;

      const get = (k) => {
        const i = idx[k];
        if (i == null || i < 0 || i >= tds.length) return '';
        return bestCellText(tds[i]);
      };

      const item = {};
      for (const c of state.model.cols) item[c.key] = get(c.key);

      if (state.model.cols.some(c => norm(item[c.key]))) out.push(item);
    }

    const now = new Date();
    state.lastStamp = `${pad2(now.getDate())}.${pad2(now.getMonth() + 1)}.${now.getFullYear()} ${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
    state.rows = out;
    const currentExtractSignature = buildBridgeSignature(state.model, out);
    const nowMs = Date.now();
    if (currentExtractSignature !== state.lastExtractSignature) {
      state.lastExtractSignature = currentExtractSignature;
      state.lastDataChangeAt = nowMs;
    }
    if (!skipBridgeSync) {
      syncBridgeSnapshot(false, 'auto');
    }

    if (shouldRender) render();
  }

  const extractDebounced = debounce(extractData, DEBOUNCE_MS);

  function installObserver() {
    const obs = new MutationObserver(() => {
      const panel = document.getElementById(PANEL_ID);
      if ((!panel || panel.style.display === 'none') && !BRIDGE_ENABLED) return;
      extractDebounced();
    });
    obs.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class', 'style'] });
  }

  /* ===================== AUTO-CLICK: Entladung + Fernverkehr ===================== */
  const AUTOCLICK_ENABLED = true;
  const AUTOCLICK_TARGETS = ['Entladung', 'Fernverkehr'];
  const AUTOCLICK_MAX_WAIT_MS = 30000;
  const AUTOCLICK_POLL_MS = 500;

  function autoClickTargets() {
    if (!AUTOCLICK_ENABLED) return;
    if (!location.href.includes('/execution/transport_units')) return;

    const started = Date.now();
    const remaining = [...AUTOCLICK_TARGETS];

    const poll = setInterval(() => {
      if (!remaining.length || (Date.now() - started) > AUTOCLICK_MAX_WAIT_MS) {
        clearInterval(poll);
        return;
      }

      const target = remaining[0];
      const candidates = Array.from(document.querySelectorAll('a,button,li,span,div,label'))
        .filter(el => {
          const t = norm(el.textContent);
          return t === target && isVisible(el);
        });

      if (!candidates.length) return;

      const isAlreadyActive = (el) => {
        const cls = String(el.className || '');
        const pCls = el.parentElement ? String(el.parentElement.className || '') : '';
        return /active|selected|ui-state-active|ui-tabs-active|current/i.test(cls) ||
               /active|selected|ui-state-active|ui-tabs-active|current/i.test(pCls);
      };

      const el = candidates[0];
      if (isAlreadyActive(el)) {
        remaining.shift();
        return;
      }

      el.click();
      remaining.shift();
    }, AUTOCLICK_POLL_MS);
  }

  function isTransportUnitsPage() {
    return location.href.includes('/execution/transport_units');
  }

  function registerUserInteractionTracking() {
    const mark = () => { state.lastUserInteractionAt = Date.now(); };
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'pointerdown'].forEach((eventName) => {
      window.addEventListener(eventName, mark, { passive: true });
    });
  }

  function findRoadnetRefreshTrigger(root) {
    const refreshRe = /(aktualis|refresh|neu laden|reload|synchroni|sync)/i;
    const iconRe = /(refresh|reload|sync|pi-refresh|fa-refresh|icon-refresh)/i;
    const scopes = [];
    if (root) scopes.push(root);
    if (root && root.closest('form')) scopes.push(root.closest('form'));
    scopes.push(document);

    const seen = new Set();
    const candidates = [];
    const selectors = [
      'button',
      'a',
      '[role="button"]',
      '[title*="Aktual"]',
      '[title*="aktual"]',
      '[title*="Refresh"]',
      '[title*="refresh"]',
      '[aria-label*="Aktual"]',
      '[aria-label*="aktual"]',
      '[aria-label*="Refresh"]',
      '[aria-label*="refresh"]',
      '.ui-icon-refresh',
      '.pi-refresh',
      '.fa-refresh',
      '.fa-sync'
    ];

    for (const scope of scopes) {
      for (const sel of selectors) {
        const nodes = Array.from(scope.querySelectorAll(sel));
        for (const node of nodes) {
          const el = node.closest ? (node.closest('button,a,[role="button"],.ui-button,.ui-commandlink') || node) : node;
          if (seen.has(el)) continue;
          seen.add(el);
          if (!isVisible(el) || el.closest(`#${PANEL_ID}`)) continue;

          const text = norm(el.textContent || '');
          const title = norm(el.getAttribute('title') || '');
          const aria = norm(el.getAttribute('aria-label') || '');
          const cls = String((el.className || '') + ' ' + (node.className || ''));
          const blob = `${text} ${title} ${aria}`;
          if (!refreshRe.test(blob) && !iconRe.test(cls)) continue;

          let score = 0;
          if (refreshRe.test(text)) score += 35;
          if (refreshRe.test(title) || refreshRe.test(aria)) score += 40;
          if (iconRe.test(cls)) score += 25;
          if (root && root.closest('form') && root.closest('form').contains(el)) score += 25;
          if (root && root.contains(el)) score += 50;
          candidates.push({ el, score });
        }
      }
    }

    if (!candidates.length) return null;
    candidates.sort((a, b) => b.score - a.score);
    return candidates[0].el || null;
  }

  function triggerRoadnetSoftRefresh() {
    if (!ROADNET_SOFT_REFRESH_ENABLED || !isTransportUnitsPage()) return false;
    const now = Date.now();
    if ((now - state.lastSoftRefreshAt) < ROADNET_SOFT_REFRESH_INTERVAL_MS) return false;

    const root = findRoadnetDataTableRoot();
    if (!root) return false;
    const trigger = findRoadnetRefreshTrigger(root);
    if (!trigger) return false;

    state.lastSoftRefreshAt = now;
    try {
      trigger.click();
      addBridgeLog('info', 'AUTO: Roadnet-Datenaktualisierung ausgelöst.');
      setTimeout(() => autoClickTargets(), 1200);
      setTimeout(() => extractData(), 1400);
      return true;
    } catch {
      return false;
    }
  }

  function maybeHardReloadRoadnetPage() {
    if (!ROADNET_HARD_REFRESH_ENABLED || !isTransportUnitsPage()) return;
    const now = Date.now();
    const hasAnyDataContext = (Array.isArray(state.rows) && state.rows.length > 0) || bridgeState.lastSentAt > 0;
    const staleDue = ROADNET_STALE_RELOAD_ENABLED
      && hasAnyDataContext
      && (now - state.lastDataChangeAt) >= ROADNET_STALE_RELOAD_AFTER_MS;
    const periodicDue = (now - state.lastHardRefreshAt) >= ROADNET_HARD_REFRESH_INTERVAL_MS
      && (now - state.lastUserInteractionAt) >= ROADNET_HARD_REFRESH_IDLE_MS;

    if (!staleDue && !periodicDue) return;

    state.lastHardRefreshAt = now;
    if (staleDue) {
      addBridgeLog('warn', 'AUTO: Daten unverändert, Seite wird fuer frische Roadnet-Daten neu geladen.');
    } else {
      addBridgeLog('warn', 'AUTO: Seite wird neu geladen, um aktuelle Roadnet-Daten zu erfassen.');
    }
    location.reload();
  }

  function installRoadnetAutoRefresh() {
    if (!isTransportUnitsPage()) return;
    setInterval(() => {
      triggerRoadnetSoftRefresh();
      maybeHardReloadRoadnetPage();
    }, 10000);
  }

  /* ===================== BOOTSTRAP ===================== */
  registerUserInteractionTracking();
  ensureOpenButton();
  try {
    ensureEntladungShortcutButton();
  } catch (e) {
    console.error('[RN-TU] UebEntl init failed', e);
  }
  ensurePanel();
  if (BRIDGE_ENABLED) ensureBridgeDepotConfigured(false);
  installObserver();
  autoClickTargets();
  installRoadnetAutoRefresh();
  if (BRIDGE_ENABLED) {
    setTimeout(() => extractData(), 1500);
    setInterval(() => extractData(), BRIDGE_POLL_INTERVAL_MS);
  }
})();










