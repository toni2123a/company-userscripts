// ==UserScript==
// @name         Roadnet Transporteinheiten
// @namespace    bodo.dpd.custom
// @version      2.2
// @description  Roadnet transport_units: Zusammenfassung mit altem Design und neuen Funktionen
// @match        https://roadnet.dpdgroup.com/execution/transport_units*
// @match        https://roadnet.dpdgroup.com/execution/trips*
// @match        http://lts.dpdit.de/*
// @match        https://lts.dpdit.de/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==
(function () {
  'use strict';
  if (window.__RN_TU_PERFEKT_RUNNING) return;
  window.__RN_TU_PERFEKT_RUNNING = true;
  const norm = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
  const pad2 = (n) => String(n).padStart(2, '0');
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const normKey = (s) =>
    norm(s)
      .toLowerCase()
      .replace(/\u00a0/g, ' ')
      .replace(/[^\p{L}\p{N}#%/_\- ]/gu, '')
      .replace(/\s+/g, ' ')
      .trim();
  function escHtml(s) {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
  function toast(msg, ok = true) {
    const el = document.createElement('div');
    el.textContent = msg;
    el.style.cssText =
      'position:fixed;right:16px;bottom:16px;z-index:2147483647;padding:7px 10px;border-radius:10px;' +
      'font:700 11px Arial,sans-serif;color:#fff;box-shadow:0 10px 30px rgba(0,0,0,.25);' +
      (ok ? 'background:#16a34a;' : 'background:#b91c1c;');
    document.body.appendChild(el);
    setTimeout(() => {
      el.style.opacity = '0';
      el.style.transition = 'opacity .25s';
      setTimeout(() => el.remove(), 260);
    }, 1200);
  }
  function isVisible(el) {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    const cs = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && cs.visibility !== 'hidden' && cs.display !== 'none';
  }
  function debounce(fn, ms) {
    let t = null;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }
  function parseDeDateTime(s) {
    const t = norm(s);
    if (!t) return null;
    const m = t.match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\s*,?\s*(\d{1,2}):(\d{2}))?/);
    if (m) {
      const dd = Number(m[1]);
      const mm = Number(m[2]);
      const yy = Number(m[3].length === 2 ? ('20' + m[3]) : m[3]);
      const hh = m[4] != null ? Number(m[4]) : 0;
      const mi = m[5] != null ? Number(m[5]) : 0;
      const d = new Date(yy, mm - 1, dd, hh, mi, 0, 0);
      return Number.isFinite(d.getTime()) ? d : null;
    }
    const d2 = new Date(t);
    return Number.isFinite(d2.getTime()) ? d2 : null;
  }
  function fmtDeDateTime(d) {
    if (!(d instanceof Date) || !Number.isFinite(d.getTime())) return '';
    return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
  }
  function fmtDateOnlyForInput(d) {
    if (!(d instanceof Date) || !Number.isFinite(d.getTime())) return '';
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
  }
  function parseDateOnlyInput(s) {
    const m = norm(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
    return Number.isFinite(d.getTime()) ? d : null;
  }
  function fmtDateRangeText(from, to) {
    return `${fmtDeDateTime(from)} → ${fmtDeDateTime(to)}`;
  }
  function parseDateRangeText(s) {
    const parts = norm(s).split(/\s*(?:->|→)\s*/);
    if (parts.length !== 2) return null;
    const from = parseDeDateTime(parts[0]);
    const to = parseDeDateTime(parts[1]);
    if (!from || !to) return null;
    return { from, to };
  }
  function getTodayStart() {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return d;
  }
  function getNowRoundedToMinute() {
    const d = new Date();
    d.setSeconds(0, 0);
    return d;
  }
  function parsePercentToUnit(s) {
    const t = norm(s).replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, '');
    if (!t) return null;
    const v = parseFloat(t);
    return Number.isFinite(v) ? v / 100 : null;
  }
  function fmtSumUnit(v) {
    if (!Number.isFinite(v)) return '0,0';
    return String(v.toFixed(1)).replace('.', ',');
  }
  function formatLtsForDisplay(raw) {
    const digits = String(raw || '').replace(/[^\d]/g, '');
    if (digits.length === 8) return `${digits.slice(0, 4)}-${digits.slice(4)}`;
    return norm(raw);
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
      if (/[A-Za-zÄÖÜäöü]/.test(a)) return a;
    }
    const titleEl = td.querySelector('[title]');
    if (titleEl) {
      const t = norm(titleEl.getAttribute('title'));
      if (/[A-Za-zÄÖÜäöü]/.test(t)) return t;
    }
    return main;
  }
  function short10(s) {
    return norm(s).slice(0, 10);
  }
  function isTrailerValue(s) {
    return normKey(s).includes('trailer');
  }
  function getClickableText(el) {
    return norm(el?.textContent || el?.innerText || el?.value || '');
  }
  function getAllClickableCandidates() {
    return Array.from(document.querySelectorAll('a,button,span,div,li,input[type="button"],input[type="submit"]'))
      .filter(isVisible);
  }
  function findClickableByExactText(text) {
    const target = norm(text);
    return getAllClickableCandidates().find(el => getClickableText(el) === target) || null;
  }
  function findClickableByIncludesText(text) {
    const target = norm(text).toLowerCase();
    return getAllClickableCandidates().find(el => getClickableText(el).toLowerCase().includes(target)) || null;
  }
  async function clickIfNeededExact(text) {
    const el = findClickableByExactText(text);
    if (!el) return false;
    const cls = String(el.className || '');
    const pcls = String(el.parentElement?.className || '');
    if (/active|selected|ui-state-active|ui-tabs-active|current/i.test(cls + ' ' + pcls)) return true;
    el.click();
    await sleep(450);
    return true;
  }
  function closeRoadnetDatePopup(input) {
    try { if (input) input.blur(); } catch {}
    try { document.body.click(); } catch {}
    try {
      document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true }));
      document.dispatchEvent(new KeyboardEvent('keyup', { key: 'Tab', code: 'Tab', bubbles: true }));
    } catch {}
  }
  function safeClick(el) {
    if (!el) return false;
    try {
      el.click();
      return true;
    } catch {
      try {
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
        el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        return true;
      } catch {
        return false;
      }
    }
  }
function findReloadPageButton() {
  const candidates = Array.from(document.querySelectorAll(
    'a,button,li,span,div,input[type="button"],input[type="submit"]'
  ))
    .filter(isVisible)
    .filter(el => !el.closest(`#${PANEL_ID}`));

  for (const el of candidates) {
    const text = normKey(
      [
        el.textContent,
        el.innerText,
        el.value,
        el.getAttribute('title'),
        el.getAttribute('aria-label'),
        el.getAttribute('data-label')
      ].filter(Boolean).join(' ')
    );

    if (
      text === 'reload_page' ||
      text.includes('reload_page') ||
      text.includes('reload page') ||
      text.includes('neu laden') ||
      text.includes('aktualisieren')
    ) {
      return (
        el.closest('a,button,li.ui-menuitem,li,span.ui-menuitem-text') ||
        el
      );
    }
  }

  return null;
}

async function clickReloadPageIfVisible(timeoutMs = 9000) {
  const start = Date.now();

  while (Date.now() - start < timeoutMs) {
    const btn = findReloadPageButton();

    if (btn) {
      const clickable =
        btn.querySelector?.('a,button,input[type="button"],input[type="submit"]') ||
        btn.closest?.('a,button') ||
        btn;

      clickable.scrollIntoView?.({ block: 'center', inline: 'center' });
      await sleep(80);

      if (safeClick(clickable)) {
        await sleep(1800);
        return true;
      }
    }

    await sleep(250);
  }

  return false;
}  /* ===================== LTS ROBUST ===================== */
  (function ltsRobustReceiver() {
    if (!location.hostname.includes('lts.dpdit.de')) return;
    const ALLOWED_ORIGINS = new Set(['https://roadnet.dpdgroup.com']);
    const PENDING_KEY = 'rn_pending_wbl';
    const SUBMIT_DONE_KEY = 'rn_pending_submit_done';
    const WNAME_PREFIX = 'RN_WBL_PENDING=';
    const SESSION_RE = /^\/\(S\([^)]+\)\)\//i;
    const WBL_RE = /\/WBLebenslauf\.aspx$/i;
    const INDEX_RE = /\/index\.aspx$/i;
    let submitLoop = null;
    let pendingWatch = null;
    let current = { left: '', right: '' };
    const parsePending = (v) => {
      const m = String(v || '').match(/^(\d{4})-(\d{4})$/);
      return m ? { left: m[1], right: m[2] } : null;
    };
    function getSubmitDoneKey() {
      try {
        return String(sessionStorage.getItem(SUBMIT_DONE_KEY) || localStorage.getItem(SUBMIT_DONE_KEY) || '');
      } catch {
        return '';
      }
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
        const bridge = String(qp.get('BRUECKE') || qp.get('wbl') || '').replace(/[^\d]/g, '');
        if (bridge.length === 8) return { left: bridge.slice(0, 4), right: bridge.slice(4) };
      } catch {}
      return null;
    }
    function savePending(left, right, resetSubmit = false) {
      const v = `${left}-${right}`;
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
    function ensureOnWblPage() {
      const sessionPrefix = getSessionPrefix();
      const onWbl = WBL_RE.test(location.pathname);
      if (sessionPrefix && onWbl) return true;
      if (sessionPrefix && !onWbl) {
        location.href = location.origin + sessionPrefix + '/WBLebenslauf.aspx';
        return false;
      }
      if (!sessionPrefix && onWbl) {
        location.href = location.origin + '/index.aspx';
        return false;
      }
      if (!sessionPrefix && !onWbl && !INDEX_RE.test(location.pathname)) return false;
      return false;
    }
    function hasResults() {
      const tables = Array.from(document.querySelectorAll('table'));
      for (const t of tables) {
        const txt = (t.innerText || '').toLowerCase();
        if (txt.includes('wb nr') || txt.includes('plombennummer') || txt.includes('scanart')) {
          const rows = t.querySelectorAll('tr').length;
          if (rows >= 2) return true;
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
      if (submitLoop && key === `${current.left}-${current.right}`) return;
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
        if (tries >= 3) stopSubmitLoop();
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
      if (SESSION_RE.test(location.pathname) && WBL_RE.test(location.pathname) && wasSubmitDispatched(p.left, p.right)) return;
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
      savePending(left, right, true);
      processPending();
      startPendingWatcher();
      try {
        ev.source?.postMessage?.({ type: 'RN_WBL_ACK', ok: true }, ev.origin);
      } catch {}
    });
    window.addEventListener('load', () => {
      processPending();
      if (peekPending()) startPendingWatcher();
    });
    if (peekPending()) startPendingWatcher();
    const obs = new MutationObserver(() => {
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
  if (!location.hostname.includes('roadnet.dpdgroup.com')) return;
  /* ===================== ROADNET ===================== */
  const NS = 'rn-tu-';
  const PANEL_ID = `${NS}panel`;
  const BTN_ID = `${NS}openbtn`;
  const TABLE_WRAP_CLASS = `${NS}tw`;
  const SCROLL_BOX_CLASS = `${NS}scrollbox`;
  const AUTO_REFRESH_SECONDS = 60;
  const MODEL_UNLOAD = {
    modeKey: 'unload',
    tabText: 'Entladung',
    title: 'Roadnet Zusammenfassung – Entladung',
    defaultSort: { key: 'arrAct', dir: 'desc' },
    cols: [
      { key: 'status', title: 'Status' },
      { key: 'from', title: 'Von' },
      { key: 'depAct', title: 'Tats. Abfahrt' },
      { key: 'lts', title: 'LTS #' },
      { key: 'arrAct', title: 'Tats. Ankunft' },
      { key: 'unlBeg', title: 'Entladebeginn' },
      { key: 'unlEnd', title: 'Entladung Ende' },
      { key: 'load', title: '%' },
      { key: 'type', title: 'Art' },
      { key: 'carrier', title: 'Frachtführer' },
      { key: 'seal', title: 'Plombe' }
    ],
    resolve: (_headerMap, pick) => ({
      status: pick(['status', 'status transporteinheit', 'te status', 'tu status']),
      from: pick(['abgangsort', 'name abgangsstandort', 'abgangsstandort', 'abgang', 'von', 'from', 'origin']),
      fromCode: pick(['code abgangsstandort', 'code abgangsstation', 'abgangsstandort']),
      depAct: pick(['tatsächliche abfahrt', 'tatsaechliche abfahrt', 'actual departure', 'departure actual']),
      lts: pick(['nummer transporteinheit', 'transporteinheit nr', 'transport unit number', 'transport unit no', 'lts #', 'lts', 'nummer transporte', 'wb nr', 'wbnr', 'bruecke']),
      arrAct: pick(['tatsächliche ankunft', 'tatsaechliche ankunft', 'ankunft', 'actual arrival', 'arrival actual']),
      unlBeg: pick(['entladung beginn', 'entladebeginn', 'entladung start', 'unload start', 'unloading start']),
      unlEnd: pick(['entladung ende', 'entladeende', 'entladung abgeschlossen', 'unload end', 'unloading end']),
      load: pick(['auslastung', 'auslastung %', 'auslastung (%)', 'load', 'load %', 'fill level', '%']),
      type: pick(['art transporteinheit', 'typ', 'transportart', 'art transporteinh', 'verkehr', 'traffic']),
      carrier: pick(['bezeichnung frachtführer', 'bezeichnung frachtfuehrer', 'frachtführer', 'frachtfuehrer', 'carrier']),
      seal: pick(['plombennummer', 'plombe', 'seal', 'seal number'])
    })
  };
  const MODEL_LOAD = {
    modeKey: 'load',
    tabText: 'Beladung',
    title: 'Roadnet Zusammenfassung – Beladung',
    defaultSort: { key: 'depPlan', dir: 'desc' },
    cols: [
      { key: 'status', title: 'Status' },
      { key: 'toName', title: 'An' },
      { key: 'depPlan', title: 'Gepl. Abfahrt' },
      { key: 'depAct', title: 'Tats. Abfahrt' },
      { key: 'loadBeg', title: 'Beladung Beginn' },
      { key: 'loadEnd', title: 'Beladung Ende' },
      { key: 'lts', title: 'LTS #' },
      { key: 'sealDep', title: 'Plombe Abfahrt' },
      { key: 'loadPct', title: '%' },
      { key: 'carrier', title: 'Frachtführer' },
      { key: 'type', title: 'Art' }
    ],
    resolve: (_headerMap, pick) => ({
      status: pick(['status', 'status transporteinheit', 'te status', 'tu status']),
      toName: pick(['name empfangsstandort', 'name empfangsstation', 'empfangsort', 'empfangsstandort', 'an', 'nach', 'to', 'destination']),
      toCode: pick(['code empfangsstation', 'code empfangsstandort', 'empfangsstation', 'empfangsstandort']),
      depPlan: pick(['geplante abfahrt', 'geplant abfahrt', 'planned departure', 'departure planned']),
      depAct: pick(['tatsächliche abfahrt', 'tatsaechliche abfahrt', 'actual departure', 'departure actual']),
      loadBeg: pick(['beladung beginn', 'beladung start', 'loading start']),
      loadEnd: pick(['beladung ende', 'beladung abgeschlossen', 'loading end']),
      lts: pick(['nummer transporteinheit', 'transporteinheit nr', 'transport unit number', 'transport unit no', 'lts #', 'lts', 'nummer transporte', 'wb nr', 'wbnr', 'bruecke']),
      sealDep: pick(['plombennummer abfahrt', 'plombe abfahrt', 'seal departure', 'seal']),
      loadPct: pick(['auslastung', 'auslastung %', 'auslastung (%)', 'load', 'load %', 'fill level', '%']),
      carrier: pick(['bezeichnung frachtführer', 'bezeichnung frachtfuehrer', 'frachtführer', 'frachtfuehrer', 'carrier']),
      type: pick(['art transporteinheit', 'typ', 'transportart', 'art transporteinh', 'verkehr', 'traffic'])
    })
  };
  const state = {
    model: MODEL_UNLOAD,
    sort: { ...MODEL_UNLOAD.defaultSort },
    rowsMain: [],
    rows: [],
    selectedRows: new Set(),
    lastStamp: '',
    scrollLeft: 0,
    scrollTop: 0,
    panelOpen: false,
    panelRange: { from: null, to: null },
    autoRefreshEnabled: false,
    autoRefreshEverySec: AUTO_REFRESH_SECONDS,
    autoRefreshLastRun: 0,
    refreshRunning: false
  };
  function rowSelectKey(row) {
    const lts = String(row.lts || '').replace(/[^\d]/g, '');
    return [
      lts,
      norm(row.from || row.toName),
      norm(row.depAct),
      norm(row.arrAct),
      norm(row.unlEnd || row.loadEnd),
      norm(row.carrier)
    ].join('|');
  }
  function dedupeKey(row) {
    const lts = String(row.lts || '').replace(/[^\d]/g, '');
    return [
      lts,
      norm(row.depAct),
      norm(row.arrAct),
      norm(row.unlEnd),
      norm(row.loadEnd),
      norm(row.toName),
      norm(row.from),
      norm(row.carrier)
    ].join('|');
  }
  function mergeRows(rows) {
    const map = new Map();
    for (const r of rows) {
      const k = dedupeKey(r);
      if (!map.has(k)) map.set(k, { ...r });
    }
    return Array.from(map.values());
  }
  function findActiveMode() {
    const cands = Array.from(document.querySelectorAll('a,button,li,span,div'))
      .map(el => ({ el, t: norm(el.textContent) }))
      .filter(x => x.t === 'Beladung' || x.t === 'Entladung')
      .filter(x => isVisible(x.el));
    if (!cands.length) return null;
    const scoreEl = (el) => {
      const cls = String(el.className || '');
      const pcls = String(el.parentElement?.className || '');
      let s = 0;
      if (/active|selected|ui-state-active|ui-tabs-active|current/i.test(cls)) s += 1000;
      if (/active|selected|ui-state-active|ui-tabs-active|current/i.test(pcls)) s += 800;
      return s;
    };
    const best = cands.map(x => ({ ...x, score: scoreEl(x.el) })).sort((a, b) => b.score - a.score)[0];
    return best?.t === 'Beladung' ? 'load' : 'unload';
  }
  function syncModelFromUI() {
    const mode = findActiveMode();
    const next = mode === 'load' ? MODEL_LOAD : MODEL_UNLOAD;
    if (state.model.modeKey !== next.modeKey) {
      state.model = next;
      state.sort = { ...next.defaultSort };
      state.selectedRows.clear();
    }
  }
  function findRoadnetDataTableRoot() {
    const direct = document.getElementById('frm_transport_units:tbl');
    if (direct && direct.classList.contains('ui-datatable')) return direct;
    const cands = Array.from(document.querySelectorAll('div.ui-datatable[id$=":tbl"]'))
      .filter(isVisible)
      .filter(el => !el.closest(`#${PANEL_ID}`));
    if (!cands.length) return null;
    if (cands.length === 1) return cands[0];
    let best = null;
    let bestScore = -1;
    for (const el of cands) {
      const txt = normKey(el.innerText || '');
      const rows = el.querySelectorAll('tbody tr').length;
      const score =
        rows +
        (txt.includes('tatsächliche ankunft') ? 60 : 0) +
        (txt.includes('beladung beginn') ? 60 : 0) +
        (txt.includes('entladung ende') ? 30 : 0) +
        (txt.includes('auslastung') ? 20 : 0) +
        (txt.includes('nummer transport') ? 20 : 0);
      if (score > bestScore) {
        bestScore = score;
        best = el;
      }
    }
    return best;
  }
  function getPrimefacesTableParts(root) {
    const headerTh = Array.from(root.querySelectorAll('.ui-datatable-scrollable-header thead th')).filter(isVisible);
    const headerCells = headerTh.length ? headerTh : Array.from(root.querySelectorAll('thead th'));
    const bodyRows1 = Array.from(root.querySelectorAll('.ui-datatable-scrollable-body tbody tr'))
      .filter(tr => tr.querySelectorAll('td').length);
    const bodyRows2 = Array.from(root.querySelectorAll('tbody tr'))
      .filter(tr => tr.querySelectorAll('td').length);
    const combined = [...bodyRows1, ...bodyRows2];
    const uniq = [];
    const seen = new Set();
    for (const r of combined) {
      const sig = (r.getAttribute('data-rk') || '') + '|' + norm(r.innerText).slice(0, 160);
      if (seen.has(sig)) continue;
      seen.add(sig);
      uniq.push(r);
    }
    return { headerCells, bodyRows: uniq };
  }
  function buildHeaderIndexMap(headerCells) {
    const map = new Map();
    headerCells.forEach((h, idx) => {
      const k = normKey(h.innerText ?? h.textContent ?? '');
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
          if (hk.includes(k) || k.includes(hk)) return idx;
        }
      }
      return null;
    };
  }
  function isExtractedHeaderLikeRow(item) {
    const vals = Object.values(item || {}).map(v => normKey(v)).filter(Boolean);
    if (!vals.length) return false;

    const titleSet = new Set(state.model.cols.map(c => normKey(c.title)));
    const keySet = new Set(state.model.cols.map(c => normKey(c.key)));
    const matches = vals.reduce((acc, v) => acc + (titleSet.has(v) || keySet.has(v) ? 1 : 0), 0);

    if (matches >= 2) return true;
    if (vals.includes('status') && (vals.includes('art') || vals.includes('lts') || vals.includes('lts #'))) return true;
    if (vals.includes('status') && vals.some(v => v.includes('abfahrt') || v.includes('ankunft') || v.includes('beladung') || v.includes('entladung'))) return true;

    return false;
  }

  function appendThreeDigitPlaceCode(name, codeRaw) {
    const base = norm(name);
    if (!base) return base;
    const codeDigits = String(codeRaw || '').replace(/[^\d]/g, '');
    const last3 = codeDigits.length >= 3 ? codeDigits.slice(-3) : '';
    if (!last3) return base;
    if (new RegExp(`\\s-\\s${last3}$`).test(base)) return base;
    if (/\s-\s\d{3}$/.test(base)) return base;
    return `${base} - ${last3}`;
  }
  function extractRowsFromCurrentTable() {
    syncModelFromUI();
    const root = findRoadnetDataTableRoot();
    if (!root) return [];
    const { headerCells, bodyRows } = getPrimefacesTableParts(root);
    const headerMap = buildHeaderIndexMap(headerCells);
    const pick = makePicker(headerMap);
    const idx = state.model.resolve(headerMap, pick);
    const out = [];
    for (const tr of bodyRows) {
      const tds = Array.from(tr.querySelectorAll('td'));
      if (!tds.length) continue;
      const get = (k) => {
        const i = idx[k];
        if (i == null || i < 0 || i >= tds.length) return '';
        return bestCellText(tds[i]);
      };
      const item = {};
      for (const c of state.model.cols) item[c.key] = get(c.key);
      if (isExtractedHeaderLikeRow(item)) continue;
      if ('carrier' in item) item.carrier = short10(item.carrier);
      if (state.model.modeKey === 'unload') {
        item.from = appendThreeDigitPlaceCode(item.from, get('fromCode'));
      }
      if (state.model.modeKey === 'load') {
        item.toName = appendThreeDigitPlaceCode(item.toName, get('toCode'));
      }
      const hasAny = state.model.cols.some(c => norm(item[c.key]));
      if (hasAny) out.push(item);
    }
    return out;
  }
  function getTableSignature() {
    const rows = extractRowsFromCurrentTable();
    return `${state.model.modeKey}|${rows.length}|${rows.slice(0, 6).map(r => `${r.lts}|${r.depAct}|${r.arrAct}|${r.unlEnd}|${r.loadEnd}|${r.carrier}`).join('||')}`;
  }
  async function waitForTableChange(oldSig, timeout = 12000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      await sleep(250);
      const sig = getTableSignature();
      if (sig !== oldSig) return true;
    }
    return false;
  }
  function findDateRangeInput() {
    const inputs = Array.from(document.querySelectorAll('input[type="text"], input'))
      .filter(isVisible)
      .filter(el => !el.closest(`#${PANEL_ID}`));
    for (const el of inputs) {
      const v = norm(el.value || el.getAttribute('value') || '');
      if (/\d{1,2}\.\d{1,2}\.\d{2,4}\s+\d{1,2}:\d{2}\s*(?:->|→)\s*\d{1,2}\.\d{1,2}\.\d{2,4}\s+\d{1,2}:\d{2}/.test(v)) {
        return el;
      }
    }
    return null;
  }
  function getCurrentRoadnetRange() {
    const input = findDateRangeInput();
    const parsed = parseDateRangeText(input ? input.value : '');
    if (parsed?.from && parsed?.to) return parsed;
    return { from: getTodayStart(), to: getNowRoundedToMinute() };
  }
  function setPanelRange(from, to) {
    state.panelRange.from = from instanceof Date ? new Date(from.getTime()) : null;
    state.panelRange.to = to instanceof Date ? new Date(to.getTime()) : null;
    const panel = document.getElementById(PANEL_ID);
    if (!panel) return;
    const fromEl = panel.querySelector(`.${NS}range-from`);
    const toEl = panel.querySelector(`.${NS}range-to`);
    if (fromEl) fromEl.value = state.panelRange.from ? fmtDateOnlyForInput(state.panelRange.from) : '';
    if (toEl) toEl.value = state.panelRange.to ? fmtDateOnlyForInput(state.panelRange.to) : '';
  }
  function syncPanelRangeFromRoadnet() {
    const { from, to } = getCurrentRoadnetRange();
    setPanelRange(from, to);
  }
function getPanelRangeOrFallback() {
  const panel = document.getElementById(PANEL_ID);
  const fallback = getCurrentRoadnetRange();
  if (!panel) return fallback;
  const fromEl = panel.querySelector(`.${NS}range-from`);
  const toEl = panel.querySelector(`.${NS}range-to`);
  const fromDate = parseDateOnlyInput(fromEl?.value);
  const toDate = parseDateOnlyInput(toEl?.value);
  const from = fromDate || new Date(fallback.from.getTime());
  const to = toDate || new Date(fallback.to.getTime());
  // Uhrzeit IMMER aus Roadnet übernehmen, nicht aus dem Datumfeld ableiten.
  // Damit bleibt z. B. 12:00 bis 12:00 erhalten.
  if (fallback.from instanceof Date && Number.isFinite(fallback.from.getTime())) {
    from.setHours(
      fallback.from.getHours(),
      fallback.from.getMinutes(),
      0,
      0
    );
  }
  if (fallback.to instanceof Date && Number.isFinite(fallback.to.getTime())) {
    to.setHours(
      fallback.to.getHours(),
      fallback.to.getMinutes(),
      0,
      0
    );
  }
  return { from, to };
}
  async function setDateRangeValue(input, fromDate, toDate) {
    if (!input) return false;
    const newVal = fmtDateRangeText(fromDate, toDate);
    try {
      input.focus();
      input.click();
    } catch {}
    await sleep(100);
    const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
    if (nativeSetter) nativeSetter.call(input, newVal);
    else input.value = newVal;
    input.setAttribute('value', newVal);
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(150);
    closeRoadnetDatePopup(input);
    await sleep(180);
    try { input.blur(); } catch {}
    input.dispatchEvent(new Event('blur', { bubbles: true }));
    await sleep(220);
    const maybeApply =
      findClickableByExactText('Anzeigen') ||
      findClickableByExactText('Suchen') ||
      findClickableByExactText('Aktualisieren') ||
      findClickableByIncludesText('anzeigen') ||
      findClickableByIncludesText('suchen') ||
      findClickableByIncludesText('aktualisieren');
    if (maybeApply) {
      maybeApply.click();
    } else {
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
      input.dispatchEvent(new KeyboardEvent('keypress', { key: 'Enter', code: 'Enter', bubbles: true }));
      input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', bubbles: true }));
    }
    await sleep(1200);
    return true;
  }
  async function applyPanelRangeToRoadnet() {
    const input = findDateRangeInput();
    if (!input) {
      toast('Datumsbereich im Roadnet nicht gefunden', false);
      return false;
    }
    const { from, to } = getPanelRangeOrFallback();
    if (!(from instanceof Date) || !(to instanceof Date) || !Number.isFinite(from.getTime()) || !Number.isFinite(to.getTime())) {
      toast('Datumsbereich ungültig', false);
      return false;
    }
    if (from.getTime() > to.getTime()) {
      toast('Von darf nicht größer als Bis sein', false);
      return false;
    }
    const oldSig = getTableSignature();
    setPanelRange(from, to);
    const ok = await setDateRangeValue(input, from, to);
    if (!ok) return false;
    const reloaded = await clickReloadPageIfVisible(9000);
    if (reloaded) await sleep(2500);
    await waitForTableChange(oldSig, 12000);
refreshFromPage();
await sleep(300);
refreshFromPage();
      return true;
  }
  async function ensureRoadnetContext(modeKey) {
    await clickIfNeededExact('Transporteinheiten');
    await clickIfNeededExact('Fernverkehr');
    if (modeKey === 'load') await clickIfNeededExact('Beladung');
    else await clickIfNeededExact('Entladung');
    await sleep(350);
  }
  async function activatePanelMode(modeKey) {
    const next = modeKey === 'load' ? MODEL_LOAD : MODEL_UNLOAD;
    state.model = next;
    state.sort = { ...next.defaultSort };
    state.selectedRows.clear();
    updateModeButtons();
    await ensureRoadnetContext(modeKey);
    await applyPanelRangeToRoadnet();
    refreshFromPage();
  }
  function ensureOpenButton() {
    let btn = document.getElementById(BTN_ID);
    if (btn) return btn;
    btn = document.createElement('button');
    btn.id = BTN_ID;
    btn.type = 'button';
    btn.textContent = 'TU';
    btn.title = 'Roadnet Transporteinheiten öffnen';
    btn.style.cssText = [
      'position:fixed',
      'right:14px',
      'top:110px',
      'z-index:2147483001',
      'cursor:pointer',
      'border:1px solid #cfd6df',
      'background:#fff',
      'border-radius:10px',
      'padding:5px 8px',
      'font:900 11px Arial,sans-serif',
      'color:#111',
      'box-shadow:0 8px 22px rgba(0,0,0,.18)'
    ].join(';');
    btn.addEventListener('click', async () => {
      const p = ensurePanel();
      const open = p.style.display === 'none';
      p.style.display = open ? 'block' : 'none';
      state.panelOpen = open;
      if (open) {
        syncPanelRangeFromRoadnet();
        await activatePanelMode(state.model.modeKey);
      }
    });
    document.body.appendChild(btn);
    return btn;
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
      'top:clamp(6px, 3vh, 90px)',
      'width:min(1320px, calc(100vw - 12px))',
      'max-width:calc(100vw - 12px)',
      'max-height:calc(100vh - 12px)',
      'overflow:hidden',
      'z-index:2147483000',
      'background:#fff',
      'border:1px solid #222',
      'border-radius:8px',
      'box-shadow:0 18px 60px rgba(0,0,0,.28)',
      'display:none',
      'font-family:Arial,sans-serif'
    ].join(';');
    panel.innerHTML = `
      <style>
        #${PANEL_ID} * { box-sizing:border-box; }
        #${PANEL_ID} .${NS}head {
          display:flex;
          align-items:flex-start;
          justify-content:space-between;
          gap:10px;
          padding:6px 8px;
          border-bottom:1px solid #cfd6df;
          background:#fff;
        }
        #${PANEL_ID} .${NS}title {
          font:800 13px Arial,sans-serif;
          color:#000;
          margin-top:3px;
        }
        #${PANEL_ID} .${NS}meta {
          font:700 11px Arial,sans-serif;
          color:#000;
          line-height:1.25;
        }
        #${PANEL_ID} .${NS}actions {
          display:flex;
          gap:6px;
          align-items:center;
          flex-wrap:wrap;
          justify-content:flex-end;
        }
        #${PANEL_ID} button {
          cursor:pointer;
          border:1px solid #d8dee8;
          background:#fff;
          border-radius:10px;
          padding:5px 10px;
          font:700 11px Arial,sans-serif;
          color:#000;
        }
        #${PANEL_ID} button.${NS}active {
          background:#e60032;
          color:#fff;
          border-color:#e60032;
          font-weight:800;
        }
        #${PANEL_ID} button.${NS}close {
          border-radius:50%;
          width:28px;
          height:28px;
          padding:0;
          font:800 16px Arial,sans-serif;
          color:#111;
          background:#fff;
        }
        #${PANEL_ID} .${NS}range {
          display:flex;
          gap:6px;
          align-items:center;
          background:#fff;
          border:1px solid #d8dee8;
          border-radius:10px;
          padding:3px 6px;
          font:700 10px Arial,sans-serif;
          color:#000;
        }
        #${PANEL_ID} .${NS}rangebox {
          display:flex;
          flex-direction:column;
          gap:2px;
        }
        #${PANEL_ID} input[type="date"] {
          border:0;
          border-radius:6px;
          padding:4px 4px;
          font:700 11px Arial,sans-serif;
          color:#000;
          background:#fff;
        }
        #${PANEL_ID} .${NS}check {
          display:flex;
          gap:5px;
          align-items:center;
          border:1px solid #d8dee8;
          border-radius:10px;
          padding:6px 8px;
          font:700 11px Arial,sans-serif;
          background:#fff;
        }
        #${PANEL_ID} .${NS}body {
          padding:6px 8px;
          background:#fff;
        }
        #${PANEL_ID} .${NS}summary {
          display:flex;
          gap:5px;
          align-items:center;
          justify-content:flex-start;
          margin-bottom:4px;
          padding:3px 6px;
          background:#eef3f8;
          border:1px solid #d4dde8;
          border-radius:5px;
          font:700 11px Arial,sans-serif;
          color:#4b5563;
          white-space:nowrap;
          overflow:hidden;
          text-overflow:ellipsis;
        }
        #${PANEL_ID} .${NS}chip {
          background:transparent;
          border:0;
          border-radius:0;
          padding:0;
          font:700 11px Arial,sans-serif;
          color:#4b5563;
        }
        #${PANEL_ID} .${SCROLL_BOX_CLASS} {
          overflow:auto;
          border:1px solid #cfd6df;
          border-radius:6px;
          max-height:calc(100vh - 155px);
          background:#fff;
        }
        #${PANEL_ID} table {
          border-collapse:collapse;
          width:100%;
          font:600 11px Arial,sans-serif;
          color:#000;
          table-layout:auto;
        }
        #${PANEL_ID} th {
          position:sticky;
          top:0;
          background:#edf3fa;
          z-index:2;
          text-align:left;
          padding:4px 6px;
          border-right:1px solid #d4dde8;
          border-bottom:1px solid #c1cad6;
          white-space:nowrap;
          cursor:pointer;
          font:700 11px Arial,sans-serif;
          color:#000;
        }
        #${PANEL_ID} td {
          padding:4px 6px;
          border-right:1px solid #e5e7eb;
          border-bottom:1px solid #e5e7eb;
          white-space:nowrap;
          overflow:hidden;
          text-overflow:ellipsis;
          max-width:260px;
          font:600 11px Arial,sans-serif;
          color:#000;
        }
        #${PANEL_ID} tr:nth-child(even) td {
          background:#fafafa;
        }
        #${PANEL_ID} .${NS}trailercell {
          color:#991b1b !important;
          font-weight:800 !important;
        }
        #${PANEL_ID} .${NS}num {
          text-align:right;
          font-variant-numeric:tabular-nums;
        }
        #${PANEL_ID} .${NS}lts {
          color:#0057ff;
          text-decoration:underline;
          cursor:pointer;
          font-weight:700;
        }
        #${PANEL_ID} .${NS}bubble {
          display:inline-flex;
          align-items:center;
          justify-content:center;
          min-width:16px;
          height:15px;
          padding:0 6px;
          margin-left:6px;
          border-radius:999px;
          background:#111;
          color:#fff;
          font:800 9.5px system-ui;
          line-height:15px;
          vertical-align:middle;
        }
        #${PANEL_ID} .${NS}bubble-red {
          background:#fee2e2;
          color:#991b1b;
          border:1px solid #fecaca;
        }
        #${PANEL_ID} .${NS}redpill {
          display:inline-block;
          padding:1px 7px;
          border-radius:999px;
          background:#fee2e2;
          color:#991b1b;
          border:1px solid #fecaca;
          font:800 10.5px Arial,sans-serif;
          line-height:1.25;
        }
        #${PANEL_ID} .${NS}thtxt {
          vertical-align:middle;
        }
        #${PANEL_ID} .${NS}selectcol {
          width:34px;
          min-width:34px;
          max-width:34px;
          text-align:center;
        }
        #${PANEL_ID} .${NS}rowcheck,
        #${PANEL_ID} .${NS}checkall {
          width:14px;
          height:14px;
          cursor:pointer;
        }
        @media (max-width:900px) {
          #${PANEL_ID} {
            left:6px !important;
            right:6px !important;
            top:6px !important;
            transform:none !important;
            width:auto !important;
            max-width:none !important;
            max-height:calc(100vh - 12px) !important;
          }
          #${PANEL_ID} .${NS}head {
            flex-direction:column;
            align-items:stretch;
          }
          #${PANEL_ID} .${NS}actions {
            justify-content:flex-start;
            gap:5px;
          }
          #${PANEL_ID} button {
            padding:5px 7px;
            font-size:10px;
          }
          #${PANEL_ID} .${NS}summary {
            overflow:auto;
            white-space:nowrap;
          }
          #${PANEL_ID} .${SCROLL_BOX_CLASS} {
            max-height:calc(100vh - 215px);
          }
          #${PANEL_ID} table {
            font-size:10px;
          }
          #${PANEL_ID} th,
          #${PANEL_ID} td {
            padding:3px 5px;
            font-size:10px;
          }
        }
      </style>
      <div class="${NS}head">
        <div>
          <div class="${NS}meta">Stand: <span class="${NS}stamp">—</span></div>
          <div class="${NS}title">Roadnet Zusammenfassung – Entladung</div>
        </div>
        <div class="${NS}actions">
          <button type="button" class="${NS}mode ${NS}mode-load">Beladung</button>
          <button type="button" class="${NS}mode ${NS}mode-unload">Entladung</button>
          <div class="${NS}range">
            <div class="${NS}rangebox">
              <span>Von</span>
              <input type="date" class="${NS}range-from">
            </div>
            <div class="${NS}rangebox">
              <span>Bis</span>
              <input type="date" class="${NS}range-to">
            </div>
          </div>
          <label class="${NS}check">
            <input type="checkbox" class="${NS}auto-refresh">
            <span>Auto-Refresh 60s</span>
          </label>
          <button type="button" class="${NS}apply-range">Neu laden</button>
          <button type="button" class="${NS}copy-selected">Auswahl kopieren</button>
          <button type="button" class="${NS}copy">Alle kopieren</button>
          <button type="button" class="${NS}close">×</button>
        </div>
      </div>
      <div class="${NS}body">
        <div class="${NS}summary">
          <span class="${NS}chip">Zeilen: <b class="${NS}count">0</b></span>
          <span>|</span>
          <span class="${NS}chip">Bereich: <b class="${NS}range-label">—</b></span>
          <span>|</span>
          <span class="${NS}chip">Sortierung: <b class="${NS}sortlabel">—</b></span>
          <span>|</span>
          <span class="${NS}chip">Auto-Refresh: <b class="${NS}autoreflabel">aus</b></span>
        </div>
        <div class="${SCROLL_BOX_CLASS}">
          <div class="${TABLE_WRAP_CLASS}"></div>
        </div>
      </div>
    `;
    document.body.appendChild(panel);
    panel.querySelector(`.${NS}close`).addEventListener('click', () => {
      panel.style.display = 'none';
      state.panelOpen = false;
    });
    panel.querySelector(`.${NS}mode-unload`).addEventListener('click', () => activatePanelMode('unload'));
    panel.querySelector(`.${NS}mode-load`).addEventListener('click', () => activatePanelMode('load'));
    panel.querySelector(`.${NS}apply-range`).addEventListener('click', () => applyPanelRangeToRoadnet());
    panel.querySelector(`.${NS}copy`).addEventListener('click', async () => {
  await copyRowsFormatted(state.rows);
});
    panel.querySelector(`.${NS}copy-selected`).addEventListener('click', () => copySelectedRows());
    panel.querySelector(`.${NS}auto-refresh`).addEventListener('change', (e) => {
      state.autoRefreshEnabled = !!e.target.checked;
      updateSummaryLine();
    });
    const scrollBox = panel.querySelector(`.${SCROLL_BOX_CLASS}`);
    scrollBox.addEventListener('scroll', () => {
      state.scrollLeft = scrollBox.scrollLeft;
      state.scrollTop = scrollBox.scrollTop;
    });
    updateModeButtons();
    return panel;
  }
  function updateModeButtons() {
    const panel = document.getElementById(PANEL_ID);
    if (!panel) return;
    panel.querySelector(`.${NS}mode-unload`)?.classList.toggle(`${NS}active`, state.model.modeKey === 'unload');
    panel.querySelector(`.${NS}mode-load`)?.classList.toggle(`${NS}active`, state.model.modeKey === 'load');
    const title = panel.querySelector(`.${NS}title`);
    if (title) title.textContent = state.model.title;
  }
  function updateSummaryLine() {
    const panel = document.getElementById(PANEL_ID);
    if (!panel) return;
    const range = getPanelRangeOrFallback();
    const rangeText = `${fmtDeDateTime(range.from)} – ${fmtDeDateTime(range.to)}`;
    panel.querySelector(`.${NS}range-label`).textContent = rangeText;
    panel.querySelector(`.${NS}autoreflabel`).textContent = state.autoRefreshEnabled ? 'ein' : 'aus';
  }
  function sortRows(rows) {
    const { key, dir } = state.sort;
    const factor = dir === 'asc' ? 1 : -1;
    return [...rows].sort((a, b) => {
      let av = a[key] ?? '';
      let bv = b[key] ?? '';
      if (key === 'load' || key === 'loadPct') {
        av = parsePercentToUnit(av) ?? -999;
        bv = parsePercentToUnit(bv) ?? -999;
      } else {
        const ad = parseDeDateTime(av);
        const bd = parseDeDateTime(bv);
        if (ad && bd) {
          av = ad.getTime();
          bv = bd.getTime();
        } else {
          av = norm(av).toLowerCase();
          bv = norm(bv).toLowerCase();
        }
      }
      if (av < bv) return -1 * factor;
      if (av > bv) return 1 * factor;
      return 0;
    });
  }
  function refreshFromPage() {
    syncModelFromUI();
    const rows = extractRowsFromCurrentTable();
    state.rowsMain = rows;
    state.rows = sortRows(mergeRows(rows));
    state.lastStamp = fmtDeDateTime(new Date());
    renderTable();
  }
  function buildHeaderBubble(key, rows) {
    const total = rows.length;
    if (!total) return '';

    if (state.model.modeKey === 'unload' && key === 'status') {
      const erfasstCnt = rows.reduce((a, r) => a + (normKey(r.status).includes('erfasst') ? 1 : 0), 0);
      return `<span class="${NS}bubble">${total}/${erfasstCnt}</span>`;
    }

    if (key === 'type') {
      const typeCounts = { N: 0, J: 0, T: 0 };
      for (const r of rows) {
        const t = normKey(r.type);
        if (!t) continue;
        if (t.includes('normal')) typeCounts.N++;
        if (t.includes('jumbo')) typeCounts.J++;
        if (t.includes('trailer')) typeCounts.T++;
      }
      return `<span class="${NS}bubble">N ${typeCounts.N}/J ${typeCounts.J}/T ${typeCounts.T}</span>`;
    }

    if ((state.model.modeKey === 'unload' && key === 'load') || (state.model.modeKey === 'load' && key === 'loadPct')) {
      const sumAll = rows.reduce((acc, r) => acc + (parsePercentToUnit(r[key]) ?? 0), 0);
      const endKey = state.model.modeKey === 'unload' ? 'unlEnd' : 'loadEnd';
      const sumMissingEnd = rows.reduce((acc, r) => {
        if (norm(r[endKey])) return acc;
        return acc + (parsePercentToUnit(r[key]) ?? 0);
      }, 0);
      return `<span class="${NS}bubble">${fmtSumUnit(sumAll)}/${fmtSumUnit(sumAll - sumMissingEnd)}</span>`;
    }

    if ((state.model.modeKey === 'unload' && key === 'from') || (state.model.modeKey === 'load' && key === 'toName')) {
      return `<span class="${NS}bubble">${total}</span>`;
    }

    const filled = rows.reduce((acc, r) => acc + (norm(r[key]) ? 1 : 0), 0);
    const empty = total - filled;
    return `<span class="${NS}bubble">${filled}/${empty}</span>`;
  }

  function renderTable() {
    const panel = ensurePanel();
    updateModeButtons();
    const wrap = panel.querySelector(`.${TABLE_WRAP_CLASS}`);
    const scrollBox = panel.querySelector(`.${SCROLL_BOX_CLASS}`);
    const cols = state.model.cols;
    const rows = state.rows;
    panel.querySelector(`.${NS}stamp`).textContent = state.lastStamp || '—';
    panel.querySelector(`.${NS}count`).textContent = String(rows.length);
    panel.querySelector(`.${NS}sortlabel`).textContent =
      `${cols.find(c => c.key === state.sort.key)?.title || state.sort.key} (${state.sort.dir === 'asc' ? 'asc' : 'desc'})`;
    updateSummaryLine();
    const html = `
      <table>
        <thead>
          <tr>
            <th class="${NS}selectcol">
              <input type="checkbox" class="${NS}checkall" title="Alle auswählen">
            </th>
            ${cols.map(c => {
              const mark = state.sort.key === c.key ? (state.sort.dir === 'asc' ? ' ▲' : ' ▼') : '';
              const bubble = buildHeaderBubble(c.key, rows);
              return `
                <th data-key="${escHtml(c.key)}">
                  <span class="${NS}thtxt">${escHtml(c.title)}${mark}</span>
                  ${bubble}
                </th>
              `;
            }).join('')}
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => {
            const trailer = isTrailerValue(r.type);
            const key = rowSelectKey(r);
            const checked = state.selectedRows.has(key) ? 'checked' : '';
            return `
              <tr>
                <td class="${NS}selectcol">
                  <input type="checkbox" class="${NS}rowcheck" data-key="${escHtml(key)}" ${checked}>
                </td>
                ${cols.map(c => {
                  const val = r[c.key] ?? '';
                  if (c.key === 'lts') {
                    const display = formatLtsForDisplay(val);
                    const digits = String(val || '').replace(/[^\d]/g, '');
                    return `<td><span class="${NS}lts" data-lts="${escHtml(digits)}">${escHtml(display)}</span></td>`;
                  }
                  if (c.key === 'status' && normKey(val).includes('erfasst')) {
                    return `<td title="${escHtml(val)}"><span class="${NS}redpill">${escHtml(val)}</span></td>`;
                  }
                  if (c.key === 'type' && isTrailerValue(val)) {
                    return `<td title="${escHtml(val)}"><span class="${NS}redpill">${escHtml(val)}</span></td>`;
                  }
                  const cls = (c.key === 'load' || c.key === 'loadPct') ? `${NS}num` : '';
                  return `<td class="${cls}" title="${escHtml(val)}">${escHtml(val)}</td>`;
                }).join('')}
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    `;
    wrap.innerHTML = html;
    wrap.querySelectorAll('th[data-key]').forEach(th => {
      th.addEventListener('click', () => {
        const key = th.getAttribute('data-key');
        if (state.sort.key === key) {
          state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
        } else {
          state.sort.key = key;
          state.sort.dir = 'asc';
        }
        state.rows = sortRows(state.rows);
        renderTable();
      });
    });
    wrap.querySelectorAll(`.${NS}lts[data-lts]`).forEach(el => {
      el.addEventListener('click', () => openLts(el.getAttribute('data-lts')));
    });
    const checkAll = wrap.querySelector(`.${NS}checkall`);
    if (checkAll) {
      checkAll.checked = rows.length > 0 && rows.every(r => state.selectedRows.has(rowSelectKey(r)));
      checkAll.addEventListener('change', () => {
        if (checkAll.checked) {
          rows.forEach(r => state.selectedRows.add(rowSelectKey(r)));
        } else {
          rows.forEach(r => state.selectedRows.delete(rowSelectKey(r)));
        }
        renderTable();
      });
    }
    wrap.querySelectorAll(`.${NS}rowcheck`).forEach(cb => {
      cb.addEventListener('change', () => {
        const key = cb.getAttribute('data-key');
        if (!key) return;
        if (cb.checked) state.selectedRows.add(key);
        else state.selectedRows.delete(key);
      });
    });
    setTimeout(() => {
      scrollBox.scrollLeft = state.scrollLeft || 0;
      scrollBox.scrollTop = state.scrollTop || 0;
    }, 0);
  }
  function openLts(raw) {
    const digits = String(raw || '').replace(/[^\d]/g, '');
    if (digits.length !== 8) {
      toast('Ungültige LTS/WB-Nummer', false);
      return;
    }
    const left = digits.slice(0, 4);
    const right = digits.slice(4);
    const w = window.open('https://lts.dpdit.de/index.aspx', 'rn_lts_wbl');
    if (!w) {
      toast('Popup blockiert', false);
      return;
    }
    try {
      w.name = `${String(w.name || '').replace(/\|?RN_WBL_PENDING=\d{4}-\d{4}/g, '')}|RN_WBL_PENDING=${left}-${right}`;
    } catch {}
    let tries = 0;
    const timer = setInterval(() => {
      tries++;
      try {
        w.postMessage({ type: 'RN_WBL', a: left, b: right }, 'https://lts.dpdit.de');
      } catch {}
      if (tries >= 30 || w.closed) clearInterval(timer);
    }, 500);
    toast(`LTS ${left}-${right} geöffnet`, true);
  }
function rowsToClipboardText(rows) {
  const cols = state.model.cols;
  const lines = [];
  lines.push(state.model.title);
  lines.push(`Stand: ${state.lastStamp || fmtDeDateTime(new Date())}`);
  lines.push('');
  lines.push(cols.map(c => c.title).join('\t'));
  for (const r of rows) {
    lines.push(cols.map(c => {
      if (c.key === 'lts') return formatLtsForDisplay(r[c.key]);
      return norm(r[c.key]);
    }).join('\t'));
  }
  return lines.join('\n');
}
function rowsToClipboardHtml(rows) {
  const cols = state.model.cols;
  const head = cols.map(c => `
    <th style="
      background:#edf3fa;
      color:#000;
      font-family:Arial,sans-serif;
      font-size:11px;
      font-weight:700;
      border:1px solid #c1cad6;
      padding:4px 6px;
      text-align:left;
      white-space:nowrap;
    ">${escHtml(c.title)}</th>
  `).join('');
  const body = rows.map((r, idx) => {
    const trailer = isTrailerValue(r.type);
    return `
      <tr>
        ${cols.map(c => {
          let val = r[c.key] ?? '';
          if (c.key === 'lts') val = formatLtsForDisplay(val);
          const marked = (c.key === 'status' && normKey(val).includes('erfasst')) || (c.key === 'type' && isTrailerValue(val));
          const bg = idx % 2 ? '#fafafa' : '#ffffff';
          const color = '#000000';
          const inner = marked
            ? `<span style="display:inline-block;padding:1px 7px;border-radius:999px;background:#fee2e2;color:#991b1b;border:1px solid #fecaca;font-weight:800;">${escHtml(val)}</span>`
            : escHtml(val);
          return `
            <td style="
              background:${bg};
              color:${color};
              font-family:Arial,sans-serif;
              font-size:11px;
              font-weight:600;
              border:1px solid #e5e7eb;
              padding:4px 6px;
              white-space:nowrap;
              mso-number-format:'\\@';
            ">${inner}</td>
          `;
        }).join('')}
      </tr>
    `;
  }).join('');
  return `
    <html>
      <body>
        <div style="font-family:Arial,sans-serif;font-size:11px;font-weight:700;color:#000;">
          Stand: ${escHtml(state.lastStamp || fmtDeDateTime(new Date()))}
        </div>
        <div style="font-family:Arial,sans-serif;font-size:13px;font-weight:800;color:#000;margin-bottom:6px;">
          ${escHtml(state.model.title)}
        </div>
        <table style="
          border-collapse:collapse;
          font-family:Arial,sans-serif;
          font-size:11px;
          color:#000;
        ">
          <thead>
            <tr>${head}</tr>
          </thead>
          <tbody>
            ${body}
          </tbody>
        </table>
      </body>
    </html>
  `;
}
async function copyRowsFormatted(rows) {
  if (!rows.length) {
    toast('Keine Zeilen vorhanden', false);
    return;
  }
  const plain = rowsToClipboardText(rows);
  const html = rowsToClipboardHtml(rows);
  try {
    await navigator.clipboard.write([
      new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([plain], { type: 'text/plain' })
      })
    ]);
    toast(`${rows.length} Zeile(n) formatiert kopiert`, true);
  } catch (err) {
    try {
      await navigator.clipboard.writeText(plain);
      toast(`${rows.length} Zeile(n) kopiert, aber ohne Formatierung`, false);
    } catch {
      toast('Kopieren fehlgeschlagen', false);
    }
  }
}
async function copySelectedRows() {
  const selected = state.rows.filter(r => state.selectedRows.has(rowSelectKey(r)));
  if (!selected.length) {
    toast('Keine Zeilen ausgewählt', false);
    return;
  }
  await copyRowsFormatted(selected);
}
async function copyTable() {
  await copyRowsFormatted(state.rows);
}
const refreshDebounced = debounce(() => {
  if (state.panelOpen) refreshFromPage();
}, 350);
let rnTableObserver = null;
let rnObservedRoot = null;

function installObserver() {
  const root = findRoadnetDataTableRoot();
  if (!root) return false;

  if (rnObservedRoot === root && rnTableObserver) return true;

  if (rnTableObserver) {
    try { rnTableObserver.disconnect(); } catch {}
  }

  rnObservedRoot = root;
  rnTableObserver = new MutationObserver(() => {
    if (!state.panelOpen) return;

    // Sofort lesen, sobald Roadnet neue Zeilen einsetzt
    requestAnimationFrame(() => {
      refreshFromPage();
    });
  });

  rnTableObserver.observe(root, {
    childList: true,
    subtree: true,
    characterData: true,
    attributes: true
  });

  return true;
}setInterval(() => {
  if (!state.panelOpen) return;
  if (!state.autoRefreshEnabled) return;
  if (state.refreshRunning) return;
  const now = Date.now();
  if (now - state.autoRefreshLastRun < state.autoRefreshEverySec * 1000) return;
  state.autoRefreshLastRun = now;
  applyPanelRangeToRoadnet();
}, 1000);
ensureOpenButton();
setInterval(() => {
  installObserver();

  if (state.panelOpen) {
    refreshFromPage();
  }
}, 1000);
})();
