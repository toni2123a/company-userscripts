// ==UserScript==
// @name         Dispatcher Site Tool Loader
// @namespace    bodo.tools
// @version      1.3.1
// @description  Zentrales Panel mit Buttons für Module (lazy run) • Close-Others • zuverlässiges Toggle beim erneuten Klick
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @noframes
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.dispatcher.site.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.dispatcher.site.user.js
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// ==/UserScript==

(function () {
  'use strict';

  const QUEUE_KEY = '__tmQueue';
  const global = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;

  // Bekannte Panels + generische Heuristik (Loader selbst NICHT schließen!)
  const DEFAULT_PANEL_SELECTORS = [
    '#pm-panel',        // Prio/Express
    '#kn-panel',        // Neu-/Rekla-Kunden
    '#tools-panel',     // evtl. weitere
    '[id$="-panel"]:not(#tm-tools-panel)' // generisch: alles mit -panel
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

  /* ================= Panel UI ================ */
  function ensurePanel() {
    if (document.getElementById('tm-tools-panel')) return;

    GM_addStyle(`
      #tm-tools-panel{position:fixed;bottom:16px;left:16px;display:flex;flex-direction:column;gap:8px;z-index:999999;font-family:ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial;}
      #tm-tools-panel .tm-card{background:#111827cc;backdrop-filter:blur(6px);color:#f9fafb;border:1px solid #374151;border-radius:12px;box-shadow:0 6px 18px rgba(0,0,0,.25);padding:8px;min-width:260px;max-width:90vw;}
      #tm-tools-panel .tm-header{display:flex;align-items:center;justify-content:space-between;gap:8px;cursor:move;user-select:none;font-size:12px;opacity:.9}
      #tm-tools-panel .tm-title{font-weight:600;letter-spacing:.2px}
      #tm-tools-panel .tm-row{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-top:6px}
      #tm-tools-panel button.tm-btn{all:unset;padding:6px 10px;border-radius:8px;border:1px solid #4b5563;background:#1f2937;font-size:12px;line-height:1;white-space:nowrap;cursor:pointer}
      #tm-tools-panel button.tm-btn:hover{background:#374151}
      #tm-tools-panel .tm-subtle{opacity:.75;font-size:11px}
    `);

    const root = document.createElement('div');
    root.id = 'tm-tools-panel';
    root.innerHTML = `
      <div class="tm-card" id="tm-card">
        <div class="tm-header" id="tm-drag">
          <div class="tm-title">Site Tools</div>
          <div class="tm-subtle" id="tm-count">0 Module</div>
        </div>
        <div class="tm-row" id="tm-row"></div>
      </div>
    `;
    document.body.appendChild(root);

    // Drag + persist
    const drag = document.getElementById('tm-drag');
    const pos = GM_getValue('tm_panel_pos', { left: 16, bottom: 16 });
    root.style.left = `${pos.left}px`;
    root.style.bottom = `${pos.bottom}px`;

    let dragging=false, startX=0, startY=0, startLeft=0, startBottom=0;
    function onDown(e){ dragging=true;
      const r=root.getBoundingClientRect(); startX=e.clientX; startY=e.clientY; startLeft=r.left; startBottom=window.innerHeight-r.bottom;
      document.addEventListener('mousemove',onMove); document.addEventListener('mouseup',onUp);
    }
    function onMove(e){ if(!dragging) return;
      const dx=e.clientX-startX, dy=e.clientY-startY;
      root.style.left=`${Math.max(0,Math.min(startLeft+dx,window.innerWidth-100))}px`;
      root.style.bottom=`${Math.max(0,Math.min(startBottom-dy,window.innerHeight-40))}px`;
    }
    function onUp(){ dragging=false;
      GM_setValue('tm_panel_pos',{left:parseInt(root.style.left||'16',10),bottom:parseInt(root.style.bottom||'16',10)});
      document.removeEventListener('mousemove',onMove); document.removeEventListener('mouseup',onUp);
    }
    drag.addEventListener('mousedown',onDown);
  }

  function updateCount() {
    const el = document.getElementById('tm-count');
    if (el) el.textContent = `${TM.modules.size} Modul${TM.modules.size===1?'':'e'}`;
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
        // Toggle: wenn offen (Flag ODER DOM) -> schließen
        if (TM._openFlag[id] || TM.isOpen(id)) {
          TM.close(id);
          return;
        }
        // Sonst andere schließen …
        TM.closeAllExcept(id);
        // … und starten
        await Promise.resolve(run());
        // Offenen-Flag setzen
        TM._openFlag[id] = true;
      } catch (err) {
        console.error(`[TM] Modul "${id}" Fehler:`, err);
        alert(`Modul "${label}" hat einen Fehler geworfen. Siehe Konsole.`);
      }
    });

    row.appendChild(btn);
    updateCount();
  }

  /* ================= Registry / API ================ */
  const TM = {
    modules: new Map(),
    _openFlag: Object.create(null),

    /**
     * register({ id, label, run, close?, isOpen?, panels? })
     */
    register(def) {
      const { id, label, run } = def || {};
      if (!id || !label || typeof run !== 'function') {
        console.warn('[TM] Ungültige Modul-Registrierung', def);
        return;
      }
      if (TM.modules.has(id)) {
        console.warn(`[TM] Modul "${id}" bereits registriert – übersprungen.`);
        return;
      }
      TM.modules.set(id, {
        id, label, run,
        close:  (typeof def.close  === 'function') ? def.close  : null,
        isOpen: (typeof def.isOpen === 'function') ? def.isOpen : null,
        panels: Array.isArray(def.panels) ? def.panels.slice() : []
      });
      addButton({ id, label, run });
      console.log('[TM] Modul registriert:', id);
    },

    // DOM-Open-Erkennung (optional + heuristisch)
    isOpen(id) {
      const mod = TM.modules.get(id);
      if (!mod) return false;

      // eigene isOpen
      if (typeof mod.isOpen === 'function') {
        try { return !!mod.isOpen(); } catch {}
      }

      // Panels aus Registrierung
      const cand = (mod.panels && mod.panels.length) ? mod.panels.slice() : [];

      // Heuristik für dieses Modul
      cand.push(...guessPanelsFor(id));

      // Sichtbar?
      return cand.some(sel => $all(sel).some(isShown));
    },

    // Modul schließen (robust)
    close(id) {
      const mod = TM.modules.get(id);
      // 1) modul-spezifische Close-Logik
      try { if (mod && typeof mod.close === 'function') mod.close(); } catch {}

      // 2) gezielte Panels des Moduls schließen
      if (mod && mod.panels && mod.panels.length) hideBySelectors(mod.panels);

      // 3) Heuristik-Panels des Moduls schließen
      hideBySelectors(guessPanelsFor(id));

      // 4) aggressiver Fallback: alle bekannten -panel schließen (außer Loader)
      hideBySelectors(DEFAULT_PANEL_SELECTORS);

      // Flag zurücksetzen
      TM._openFlag[id] = false;
    },

    closeAllExcept(exceptId) {
      TM.modules.forEach((_, id) => { if (id !== exceptId) TM.close(id); });
      // Nichts am exceptId ändern – das öffnet der Button gleich
    },
  };

  // global sichtbar
  global.TM = TM;
  global.tm  = TM;

  // Queue (Module, die vor dem Loader kamen)
  if (!Array.isArray(global[QUEUE_KEY])) global[QUEUE_KEY] = [];
  setTimeout(() => {
    const queued = global[QUEUE_KEY];
    if (Array.isArray(queued) && queued.length) {
      console.log('[TM] Queue gefunden, registriere', queued.length, 'Modul(e)…');
      queued.forEach(def => { try { TM.register(def); } catch (e) { console.error(e); } });
      queued.length = 0;
      updateCount();
    }
  }, 0);

  ensurePanel();
  console.log('[TM] Loader bereit. global.TM vorhanden?', !!global.TM);
})();
