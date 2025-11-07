// ==UserScript==
// @name         Dispatcher Site Tool Loader (hellgrau, flat)
// @namespace    bodo.tools
// @version      1.3.7
// @description  Schlankes Panel (ohne Kopfzeile) im hellgrauen Tab-Stil der Seite
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @noframes
// @grant        GM_addStyle
// @grant        unsafeWindow
// ==/UserScript==

(function () {
  'use strict';

  const QUEUE_KEY = '__tmQueue';
  const global = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;

  const DEFAULT_PANEL_SELECTORS = [
    '#pm-panel', '#kn-panel', '#tools-panel',
    '[id$="-panel"]:not(#tm-tools-panel)'
  ];

  const $all = (sel, root=document) => Array.from(root.querySelectorAll(sel));
  const isShown = (el) => !!el && el.offsetParent !== null && getComputedStyle(el).display !== 'none';
  const hideEls = (els) => els.forEach(el => el && el.style && el.style.setProperty('display','none','important'));
  const hideBySelectors = (selectors) => (selectors||[]).forEach(sel => { try { hideEls($all(sel)); } catch {} });

  function guessPanelsFor(id){
    const s = String(id||'').toLowerCase();
    const arr = [
      `#${id}-panel`,
      `[data-panel="${id}"]`,
      `.panel-${id}`
    ];
    if (/\b(pm|prio|express|exp)\b/.test(s)) arr.push('#pm-panel');
    if (/\b(kn|kunde|kunden|rekl|neu)\b/.test(s)) arr.push('#kn-panel');
    return arr;
  }

  /* ===== UI: hellgrau, wie Tabs der Website ===== */
  function ensurePanel() {
    if (document.getElementById('tm-tools-panel')) return;

    GM_addStyle(`
      #tm-tools-panel{
        position:fixed;
        top:8px;
        left:50%;
        transform:translateX(-50%);
        z-index:2147483648;
        font-family:ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial;
        --tm-btn-bg:#e5e7eb;       /* hellgrau */
        --tm-btn-bd:#d1d5db;       /* Rahmen leicht dunkler */
        --tm-btn-fg:#111827;       /* dunkelgraue Schrift */
        --tm-btn-hover:#f3f4f6;    /* leicht heller beim Hover */
        --tm-btn-active:#d1d5db;   /* leicht dunkler beim Klick */
        --tm-radius:8px;
      }
      #tm-tools-panel .tm-card{
        display:inline-flex;
        background:transparent;
        padding:0;
        width:max-content;
        max-width:96vw;
      }
      #tm-tools-panel .tm-row{
        display:flex;
        align-items:center;
        gap:6px;
        overflow-x:auto;
        overflow-y:hidden;
        scrollbar-width:thin;
        padding:2px;
      }
      #tm-tools-panel button.tm-btn{
        all:unset;
        padding:6px 12px;
        border-radius:var(--tm-radius);
        border:1px solid var(--tm-btn-bd);
        background:var(--tm-btn-bg);
        color:var(--tm-btn-fg);
        font-size:12px;
        line-height:1;
        white-space:nowrap;
        cursor:pointer;
        box-shadow:0 1px 2px rgba(0,0,0,.05);
        transition:background-color .12s ease, transform .02s ease, box-shadow .12s ease;
      }
      #tm-tools-panel button.tm-btn:hover{
        background:var(--tm-btn-hover);
        box-shadow:0 2px 4px rgba(0,0,0,.08);
      }
      #tm-tools-panel button.tm-btn:active{
        background:var(--tm-btn-active);
        transform:translateY(1px);
        box-shadow:inset 0 1px 2px rgba(0,0,0,.1);
      }
      #tm-tools-panel button.tm-btn:focus-visible{
        outline:2px solid rgba(59,130,246,.4);
        outline-offset:2px;
      }
    `);

    const root = document.createElement('div');
    root.id = 'tm-tools-panel';
    root.innerHTML = `
      <div class="tm-card" id="tm-card">
        <div class="tm-row" id="tm-row"></div>
      </div>
    `;
    document.body.appendChild(root);
  }

  function addButton({ id, label, run }) {
    ensurePanel();
    const row = document.getElementById('tm-row');
    if (!row || document.getElementById(`tm-btn-${id}`)) return;

    const btn = document.createElement('button');
    btn.className = 'tm-btn';
    btn.id = `tm-btn-${id}`;
    btn.textContent = label;

    btn.addEventListener('click', async () => {
      try {
        if (TM._openFlag[id] || TM.isOpen(id)) {
          TM.close(id);
          return;
        }
        TM.closeAllExcept(id);
        await Promise.resolve(run());
        TM._openFlag[id] = true;
      } catch (err) {
        console.error(`[TM] Modul "${id}" Fehler:`, err);
        alert(`Modul "${label}" hat einen Fehler geworfen. Siehe Konsole.`);
      }
    });

    row.appendChild(btn);
  }

  /* ===== Registry / API ===== */
  const TM = {
    modules: new Map(),
    _openFlag: Object.create(null),

    register(def) {
      const { id, label, run } = def || {};
      if (!id || !label || typeof run !== 'function') return;
      if (TM.modules.has(id)) return;
      TM.modules.set(id, def);
      addButton({ id, label, run });
    },

    isOpen(id) {
      const mod = TM.modules.get(id);
      if (!mod) return false;
      if (typeof mod.isOpen === 'function') {
        try { return !!mod.isOpen(); } catch {}
      }
      const cand = (mod.panels || []).concat(guessPanelsFor(id));
      return cand.some(sel => $all(sel).some(isShown));
    },

    close(id) {
      const mod = TM.modules.get(id);
      try { if (mod && typeof mod.close === 'function') mod.close(); } catch {}
      if (mod && mod.panels && mod.panels.length) hideBySelectors(mod.panels);
      hideBySelectors(guessPanelsFor(id));
      hideBySelectors(DEFAULT_PANEL_SELECTORS);
      TM._openFlag[id] = false;
    },

    closeAllExcept(exceptId) {
      TM.modules.forEach((_, id) => { if (id !== exceptId) TM.close(id); });
    },
  };

  global.TM = TM;
  global.tm  = TM;

  if (!Array.isArray(global[QUEUE_KEY])) global[QUEUE_KEY] = [];
  setTimeout(() => {
    const queued = global[QUEUE_KEY];
    if (Array.isArray(queued) && queued.length) {
      queued.forEach(def => { try { TM.register(def); } catch (e) { console.error(e); } });
      queued.length = 0;
    }
  }, 0);

  ensurePanel();
})();
