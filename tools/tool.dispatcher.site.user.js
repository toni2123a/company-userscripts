// ==UserScript==
// @name         Site Tools Panel (Loader)
// @namespace    bodo.tools
// @version      1.0.2
// @description  Zentrales Panel mit Buttons für einzelne Module (lazy run)
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @noframes
// @grant        GM_addStyle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// ==/UserScript==

(function () {
  'use strict';

  const QUEUE_KEY = '__tmQueue';
  const global = (typeof unsafeWindow !== 'undefined') ? unsafeWindow : window;

  // Panel erstellen (einmalig)
  function ensurePanel() {
    if (document.getElementById('tm-tools-panel')) return;

    // Styles
    GM_addStyle(`
      #tm-tools-panel {
        position: fixed;
        bottom: 16px;
        left: 16px;
        display: flex;
        flex-direction: column;
        gap: 8px;
        z-index: 999999;
        font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
      }
      #tm-tools-panel .tm-card {
        background: #111827cc; /* semi transparent */
        backdrop-filter: blur(6px);
        color: #f9fafb;
        border: 1px solid #374151;
        border-radius: 12px;
        box-shadow: 0 6px 18px rgba(0,0,0,.25);
        padding: 8px;
        min-width: 260px;
        max-width: 90vw;
      }
      #tm-tools-panel .tm-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
        cursor: move;  /* drag handle */
        user-select: none;
        font-size: 12px;
        opacity: .9;
      }
      #tm-tools-panel .tm-title {
        font-weight: 600;
        letter-spacing: .2px;
      }
      #tm-tools-panel .tm-row {
        display: flex;
        flex-wrap: wrap;
        align-items: center;
        gap: 6px;      /* Buttons nebeneinander, kein Überlappen */
        margin-top: 6px;
      }
      #tm-tools-panel button.tm-btn {
        all: unset;
        padding: 6px 10px;
        border-radius: 8px;
        border: 1px solid #4b5563;
        background: #1f2937;
        font-size: 12px;
        line-height: 1;
        white-space: nowrap;
        cursor: pointer;
      }
      #tm-tools-panel button.tm-btn:hover { background: #374151; }
      #tm-tools-panel .tm-subtle {
        opacity: .75;
        font-size: 11px;
      }
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

    // Draggable + Position speichern
    const drag = document.getElementById('tm-drag');
    const pos = GM_getValue('tm_panel_pos', { left: 16, bottom: 16 });
    root.style.left = `${pos.left}px`;
    root.style.bottom = `${pos.bottom}px`;

    let dragging = false;
    let startX = 0, startY = 0, startLeft = 0, startBottom = 0;

    function onDown(e) {
      dragging = true;
      const rect = root.getBoundingClientRect();
      startX = e.clientX;
      startY = e.clientY;
      startLeft = rect.left;
      startBottom = window.innerHeight - rect.bottom;
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    }
    function onMove(e) {
      if (!dragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const newLeft = Math.max(0, Math.min(startLeft + dx, window.innerWidth - 100));
      const newBottom = Math.max(0, Math.min(startBottom - dy, window.innerHeight - 40));
      root.style.left = `${newLeft}px`;
      root.style.bottom = `${newBottom}px`;
    }
    function onUp() {
      dragging = false;
      const left = parseInt(root.style.left || '16', 10);
      const bottom = parseInt(root.style.bottom || '16', 10);
      GM_setValue('tm_panel_pos', { left, bottom });
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    }
    drag.addEventListener('mousedown', onDown);
  }

  function updateCount() {
    const count = TM.modules.size;
    const el = document.getElementById('tm-count');
    if (el) el.textContent = `${count} Modul${count === 1 ? '' : 'e'}`;
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
        await Promise.resolve(run());
      } catch (err) {
        console.error(`[TM] Modul "${id}" Fehler:`, err);
        alert(`Modul "${label}" hat einen Fehler geworfen. Siehe Konsole.`);
      }
    });

    row.appendChild(btn);
    updateCount();
  }

  // Registry
  const TM = {
    modules: new Map(),
    register({ id, label, run }) {
      if (!id || !label || typeof run !== 'function') {
        console.warn('[TM] Ungültige Modul-Registrierung', { id, label, run });
        return;
      }
      if (TM.modules.has(id)) {
        console.warn(`[TM] Modul "${id}" bereits registriert – übersprungen.`);
        return;
      }
      TM.modules.set(id, { id, label, run });
      addButton({ id, label, run });
      console.log('[TM] Modul registriert:', id);
    },
  };

  // global verfügbar machen (+ Alias für Konsole)
  global.TM = TM;
  global.tm = TM;

  // Queue initialisieren & abarbeiten
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
  console.log('[TM] Loader bereit. Sichtbares global.TM?', !!global.TM);
})();
