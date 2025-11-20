// ==UserScript==
// @name         DPD LEM
// @namespace    https://bodo.dpd
// @version      1.9
// @description  Belegnummer automatisch mit letzter Belegnummer + 1 vorbelegen und Spalte "Beleg-Nr." sortierbar machen. Suche über Benutzerdefinierten Kundennamen
// @match        https://dpd.lademittel.management/page/posting/postingOverview.xhtml*
// @match        https://dpd.lademittel.management/page/posting/postingCreate.xhtml*
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.lem.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.lem.user.js
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function() {
  'use strict';

  const href = window.location.href;
  const BELEG_INDEX = 5; // "Beleg-Nr."-Spalte in der Übersicht

  if (href.includes('postingOverview.xhtml')) {
    runOnOverview();
  } else if (href.includes('postingCreate.xhtml')) {
    runOnCreate();
  }

  // ---------- gemeinsame Helfer ----------

  function parseBelegNum(str) {
    const txt = (str || '').trim();
    if (!txt) return null;
    const m = txt.match(/^(.*?)(\d+)\s*$/);
    if (!m) return null;
    const prefix = m[1];
    const numStr = m[2];
    const num = Number(numStr);
    if (Number.isNaN(num)) return null;
    return { raw: txt, prefix, num, width: numStr.length };
  }

  function formatNextBeleg(parsed) {
    if (!parsed) return null;
    const next = parsed.num + 1;
    const numStr = String(next).padStart(parsed.width, '0');
    return parsed.prefix + numStr;
  }

  function sleep(ms) {
    return new Promise(r => setTimeout(r, ms));
  }

  function getTable() {
    return document.querySelector('table');
  }

  // ====================================================
  // 1. Übersicht: "Beleg-Nr." sortierbar + höchste merken
  // ====================================================

  function sortBelegColumn(descending) {
    const table = getTable();
    if (!table) return;
    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort((a, b) => {
      const ta = (a.children[BELEG_INDEX]?.textContent || '').trim();
      const tb = (b.children[BELEG_INDEX]?.textContent || '').trim();
      const pa = parseBelegNum(ta);
      const pb = parseBelegNum(tb);
      if (!pa && !pb) return 0;
      if (!pa) return 1;
      if (!pb) return -1;
      const diff = pa.num - pb.num;
      return descending ? -diff : diff;
    });
    rows.forEach(r => tbody.appendChild(r));
  }

  function enableBelegHeaderSorting() {
    const table = getTable();
    if (!table) return;

    const headers = table.querySelectorAll('thead th');
    let belegHeader = headers[BELEG_INDEX];
    if (!belegHeader) {
      belegHeader = Array.from(headers)
        .find(th => th.textContent.trim().startsWith('Beleg-Nr'));
    }
    if (!belegHeader) return;

    belegHeader.style.cursor = 'pointer';
    belegHeader.title = 'Client-Sortierung Beleg-Nr.';
    let indicator = belegHeader.querySelector('.dpd-beleg-sort-indicator');
    if (!indicator) {
      indicator = document.createElement('span');
      indicator.className = 'dpd-beleg-sort-indicator';
      indicator.style.fontSize = '0.8em';
      indicator.style.marginLeft = '4px';
      indicator.textContent = '⇅';
      belegHeader.appendChild(indicator);
    }

    let descending = true;
    belegHeader.addEventListener('click', () => {
      sortBelegColumn(descending);
      indicator.textContent = descending ? '↓' : '↑';
      descending = !descending;
    });

    // initial absteigend sortieren
    sortBelegColumn(true);
    indicator.textContent = '↓';
  }

  function runOnOverview() {
    setTimeout(async () => {
      enableBelegHeaderSorting();

      await sleep(200);
      const table = getTable();
      if (!table) return;

      const rows = table.querySelectorAll('tbody tr');
      let topParsed = null;
      for (const tr of rows) {
        const td = tr.children[BELEG_INDEX];
        if (!td) continue;
        const txt = td.textContent.trim();
        const parsed = parseBelegNum(txt);
        if (parsed) {
          topParsed = parsed;
          break;
        }
      }
      if (topParsed) {
        console.log('DPD-Belegscript: letzte Belegnummer (Übersicht):', topParsed.raw);
        localStorage.setItem('dpd_lastBeleg', topParsed.raw);
      }
    }, 800);
  }

  // ====================================================
  // 2. Create-Seite
  //    - Belegnummer +1
  //    - Unternehmen-Filter über beide Spalten (Client)
  // ====================================================

  function runOnCreate() {
    autoFillBelegnummer();
    enableUnternehmenSearch();
  }

  function autoFillBelegnummer() {
    const last = localStorage.getItem('dpd_lastBeleg');
    if (!last) return;

    const parsed = parseBelegNum(last);
    const nextBeleg = formatNextBeleg(parsed);
    if (!nextBeleg) return;

    const tryFill = () => {
      const input = document.getElementById('postingEditForm:palletNoteNumber:validatableInput');
      if (!input) return false;
      if (!input.value) {
        input.value = nextBeleg;
        input.dispatchEvent(new Event('input',  { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        console.log('DPD-Belegscript: Belegfeld mit', nextBeleg, 'befüllt.');
        localStorage.setItem('dpd_lastBeleg', nextBeleg);
      }
      return true;
    };

    if (!tryFill()) {
      const intId = setInterval(() => {
        if (tryFill()) clearInterval(intId);
      }, 500);
      setTimeout(() => clearInterval(intId), 10000);
    }
  }

  // -------- Unternehmen-Filter über beide Spalten (Client, global) --------

 function enableUnternehmenSearch() {
  const fieldId = 'postingEditForm:customerAccountAddressInput_input';
  const panelId = 'postingEditForm:customerAccountAddressInput_panel';
  const tourFieldId = 'postingEditForm:j_idt383:validatableInput';

  let panel = null;
  let items = null; // [{ row, search }]
  let observedPanel = null;
  let observer = null;

  function ensurePanelAndObserver() {
    const current = document.getElementById(panelId);
    if (!current) return false;

    if (current !== observedPanel) {
      observedPanel = current;
      panel = current;
      items = null; // neue DOM-Struktur -> neu einlesen

      if (observer) observer.disconnect();
      observer = new MutationObserver(() => {
        // Panel-Inhalt hat sich geändert -> Liste verwerfen
        items = null;
      });
      observer.observe(panel, { childList: true, subtree: true });
      console.log('DPD-Script: Unternehmen-Panel (re)verbunden.');
    }
    return true;
  }

  function attachRowClick(tr) {
    if (!tr || tr.dataset.dpdTourBound === '1') return;

    tr.dataset.dpdTourBound = '1';

    tr.addEventListener('click', function () {
      const tourInput = document.getElementById(tourFieldId);
      if (!tourInput) return;

      const tds = tr.querySelectorAll('td');
      const col2 = (tds[1]?.textContent || '').trim();

      // erste zusammenhängende Ziffernfolge aus Spalte 2 holen (z.B. "157" aus "157 Thiemo Test")
      const m = col2.match(/\d+/);
      const tourVal = m ? m[0] : '';

      tourInput.value = tourVal;
      tourInput.dispatchEvent(new Event('input',  { bubbles: true }));
      tourInput.dispatchEvent(new Event('change', { bubbles: true }));

      console.log('DPD-Script: Tour-Feld gesetzt auf:', tourVal, 'aus', col2);
    });
  }

  function snapshotItems() {
    if (!ensurePanelAndObserver()) return;
    const tbody = panel.querySelector('table.ui-autocomplete-items tbody');
    if (!tbody) return;
    const rows = Array.from(tbody.querySelectorAll('tr.ui-autocomplete-item'));
    if (!rows.length) return;

    items = rows.map(tr => {
      const tds = tr.querySelectorAll('td');
      const col1 = (tds[0]?.textContent || '').trim();
      const col2 = (tds[1]?.textContent || '').trim();
      const combined = (col1 + ' ' + col2).trim().toLowerCase();

      // Klick-Handler pro Zeile nur einmal anhängen
      attachRowClick(tr);

      return { row: tr, search: combined };
    });

    console.log('DPD-Script: Unternehmen-Liste gesichert, Einträge:', items.length);
  }

  function applyFilter(query) {
    if (!ensurePanelAndObserver()) return;

    if (!items || !items.length) {
      snapshotItems();
      if (!items || !items.length) return;
    }

    const q = (query || '').trim().toLowerCase();

    items.forEach(it => {
      if (!q || it.search.includes(q)) {
        it.row.style.display = '';
      } else {
        it.row.style.display = 'none';
      }
    });

    panel.scrollTop = 0;
  }

  let filterTimer = null;
  function scheduleFilter() {
    if (filterTimer) clearTimeout(filterTimer);
    filterTimer = setTimeout(() => {
      const inputEl = document.getElementById(fieldId);
      const val = inputEl ? inputEl.value : '';
      applyFilter(val);
    }, 200);
  }

  // Globale Listener in der Capture-Phase,
  // damit sie auch nach Neurendern des Inputs greifen.
  document.addEventListener('input', function(e) {
    const t = e.target;
    if (t && t.id === fieldId) {
      e.stopImmediatePropagation(); // PrimeFaces-Filter blocken
      scheduleFilter();
    }
  }, true);

  document.addEventListener('keyup', function(e) {
    const t = e.target;
    if (t && t.id === fieldId) {
      e.stopImmediatePropagation();
      scheduleFilter();
    }
  }, true);

  console.log('DPD-Script: Unternehmen-Suche (Clientfilter, global + Tour) aktiviert.');
}

})();
