// ==UserScript==
// @name         DPD LEM
// @namespace    https://bodo.dpd
// @version      2.4
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
  let fileSortDescending = true;

  // Seite erkennen
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
    // Nur die Haupt-Tabelle mit der Spalte "Beleg-Nr." verwenden
    const tables = Array.from(document.querySelectorAll('table'));
    for (const t of tables) {
      const ths = Array.from(t.querySelectorAll('thead th'));
      if (ths.some(th => th.textContent.trim().startsWith('Beleg-Nr'))) {
        return t;
      }
    }
    // Fallback (sollte eigentlich nicht mehr nötig sein)
    return document.querySelector('table');
  }

  // ====================================================
  // 1. Übersicht: Beleg-Spalte sortierbar + höchste merken
  //               + "Beleg?"-Spalte (hat Anhang?)
  // ====================================================

  function fileScore(txt) {
    const t = (txt || '').trim();
    if (t === '📎') return 3;     // Beleg vorhanden
    if (t === '⌛') return 2;     // wird geprüft
    if (t === '?')  return 1;    // Fehler
    return 0;                    // kein Beleg / Strich
  }

  function sortFileColumn(descending) {
    const table = getTable();
    if (!table) return;
    const tbody = table.querySelector('tbody');
    if (!tbody) return;

    const rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort((a, b) => {
      const ta = (a.children[BELEG_INDEX + 1]?.textContent || '').trim();
      const tb = (b.children[BELEG_INDEX + 1]?.textContent || '').trim();
      const va = fileScore(ta);
      const vb = fileScore(tb);
      const diff = va - vb;
      return descending ? -diff : diff;
    });

    rows.forEach(r => tbody.appendChild(r));
  }






  // pro Zeile prüfen, ob in der Detailansicht ein Beleg existiert
   // NEU: pro Zeile prüfen, ob in der Detailansicht ein Beleg (Filepreview) existiert

 // =====================================================
// Übersicht – wir machen NICHTS mehr, um PF nicht zu brechen
// =====================================================

function runOnOverview() {
  // Nur letzte Belegnummer auslesen (ohne DOM umzubauen)
  setTimeout(() => {
    const table = getTable();
    if (!table) return;
    const rows = table.querySelectorAll('tbody tr');
    for (const tr of rows) {
      const td = tr.children[BELEG_INDEX];
      if (!td) continue;
      const txt = td.textContent.trim();
      if (txt) {
        localStorage.setItem('dpd_lastBeleg', txt);
        console.log('DPD-Belegscript: letzte Belegnummer aktualisiert:', txt);
        break;
      }
    }
  }, 500);
}


  // PrimeFaces AJAX: wenn Datatable via Paginator/Filter neu gerendert wird,
  // nach dem Update unsere Logik erneut ausführen
  document.addEventListener('pfAjaxComplete', function(e) {
    try {
      const detail = e.detail || {};
      const updateIds = detail.updateIds || detail.updateId || [];
      const ids = Array.isArray(updateIds) ? updateIds : [updateIds];

      if (!ids.some(id => typeof id === 'string' &&
                          id.indexOf('postingOverviewForm:postingOverviewTable') !== -1)) {
        return;
      }

      setTimeout(() => {
        enableBelegHeaderSorting();
        markRowsWithBeleg();
        updateLastBelegFromOverview();
      }, 50);
    } catch (err) {
      console.error('DPD-Script: Fehler im pfAjaxComplete-Handler (Overview):', err);
    }
  }, true);

  // ====================================================
  // 2. Create-Seite
  //    - Belegnummer +1 (nur bei gespeicherter Buchung hochgezählt)
  //    - Unternehmen-Filter über beide Spalten (Client)
  //    - Tour aus 2. Spalte übernehmen
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
        // nur das Feld vorbelegen – NICHT dpd_lastBeleg erhöhen
        input.value = nextBeleg;
        input.dispatchEvent(new Event('input',  { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        console.log('DPD-Belegscript: Belegfeld mit', nextBeleg, 'befüllt.');
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

  // -------- Unternehmen-Filter + Tourfeld + Tastenkürzel --------
  function enableUnternehmenSearch() {
    const fieldId     = 'postingEditForm:customerAccountAddressInput_input';
    const panelId     = 'postingEditForm:customerAccountAddressInput_panel';
    const buttonId    = 'postingEditForm:customerAccountAddressInput_button';
    const tourFieldId = 'postingEditForm:j_idt383:validatableInput';

    let panel    = null;
    let items    = null;   // [{row, search, col1, col2}]
    let allReady = false;

    function ensurePanel() {
      if (!panel) {
        panel = document.getElementById(panelId);
      }
      return !!panel;
    }

    // Tour-Feld direkt aus Spalte B der angeklickten Zeile setzen
    function attachRowClick(tr) {
      if (!tr || tr.dataset.dpdTourBound === '1') return;
      tr.dataset.dpdTourBound = '1';

      tr.addEventListener('click', function () {
        const tourInput = document.getElementById(tourFieldId);
        if (!tourInput) return;

        const tds  = tr.querySelectorAll('td');
        const col2 = (tds[1]?.textContent || '').trim();
        const m    = col2.match(/\d+/);
        const tourVal = m ? m[0] : '';

        if (!tourVal) {
          console.log('DPD-Script: Tour-Feld nicht geändert (keine Zahl in Spalte B):', col2);
          return;
        }

        tourInput.value = tourVal;
        tourInput.dispatchEvent(new Event('input',  { bubbles: true }));
        tourInput.dispatchEvent(new Event('change', { bubbles: true }));
        console.log('DPD-Script: Tour-Feld (per Klick) gesetzt auf:', tourVal);
      });
    }

    // komplette Liste per PointerEvent laden
    function preloadAllItems() {
      if (allReady) return;

      const btn = document.getElementById(buttonId);
      if (!btn) {
        console.warn('DPD-Script: Dropdown-Button nicht gefunden.');
        return;
      }

      console.log('DPD-Script: Starte PointerEvent für Dropdown-Button...');

      function realClick(el) {
        ['pointerdown', 'mousedown', 'pointerup', 'mouseup', 'click'].forEach(type => {
          el.dispatchEvent(new PointerEvent(type, {
            bubbles: true,
            cancelable: true,
            composed: true,
            pointerId: 1,
            pointerType: 'mouse',
            isPrimary: true
          }));
        });
      }

      realClick(btn);

      const start = Date.now();
      const intId = setInterval(() => {
        if (!ensurePanel()) return;

        const tbody = panel.querySelector('table.ui-autocomplete-items tbody');
        if (tbody) {
          const rows = Array.from(tbody.querySelectorAll('tr.ui-autocomplete-item'));
          if (rows.length) {
            items = rows.map(tr => {
              const tds  = tr.querySelectorAll('td');
              const col1 = (tds[0]?.textContent || '').trim();
              const col2 = (tds[1]?.textContent || '').trim();
              const combined = (col1 + ' ' + col2).toLowerCase();

              attachRowClick(tr);
              return { row: tr, search: combined, col1, col2 };
            });

            allReady = true;
            console.log('DPD-Script: Unternehmen-Gesamtliste geladen, Einträge:', items.length);

            panel.style.display = 'none';
            panel.classList.add('ui-helper-hidden');

            clearInterval(intId);
          }
        }

        if (Date.now() - start > 5000) {
          console.warn('DPD-Script: Timeout beim Laden der Unternehmen-Liste.');
          clearInterval(intId);
        }
      }, 100);
    }

    function showPanel() {
      if (!ensurePanel()) return;
      panel.style.display = 'block';
      panel.classList.remove('ui-helper-hidden');
    }

    function applyFilter(query) {
      if (!allReady) preloadAllItems();
      if (!allReady || !items || !items.length) return;
      if (!ensurePanel()) return;

      const q = (query || '').trim().toLowerCase();
      const tbody = panel.querySelector('table.ui-autocomplete-items tbody');
      if (!tbody) return;

      items.forEach(it => {
        if (it.row.parentNode !== tbody) {
          tbody.appendChild(it.row);
        }
        if (!q || it.search.includes(q)) {
          it.row.style.display = '';
        } else {
          it.row.style.display = 'none';
        }
      });

      panel.scrollTop = 0;
    }

    // ersten sichtbaren Treffer wählen
    function chooseFirstVisible() {
      if (!allReady) preloadAllItems();
      if (!allReady || !items || !items.length) return;

      const first = items.find(it => it.row.style.display !== 'none');
      if (!first) {
        console.log('DPD-Script: kein sichtbarer Treffer.');
        return;
      }

      // Zeile anklicken, damit PF den Kunden auswählt
      first.row.click();

      // Tour explizit aus Spalte B setzen
      const tourInput = document.getElementById(tourFieldId);
      if (tourInput && first.col2) {
        const m = first.col2.match(/\d+/);
        const tourVal = m ? m[0] : '';
        if (tourVal) {
          tourInput.value = tourVal;
          tourInput.dispatchEvent(new Event('input',  { bubbles: true }));
          tourInput.dispatchEvent(new Event('change', { bubbles: true }));
          console.log('DPD-Script: Tour-Feld (per chooseFirstVisible) gesetzt auf:', tourVal);
        }
      }

      console.log('DPD-Script: erste gefilterte Zeile gewählt.');
    }

    // --- Events ---

    preloadAllItems();

    let filterTimer = null;
    function scheduleFilter() {
      if (filterTimer) clearTimeout(filterTimer);
      filterTimer = setTimeout(() => {
        const inputEl = document.getElementById(fieldId);
        const val = inputEl ? inputEl.value : '';
        applyFilter(val);
      }, 150);
    }

    // Eingabe -> nur Client-Filter
    document.addEventListener('input', function (e) {
      const t = e.target;
      if (t && t.id === fieldId) {
        e.stopImmediatePropagation();
        scheduleFilter();
      }
    }, true);

    document.addEventListener('keyup', function (e) {
      const t = e.target;
      if (t && t.id === fieldId) {
        e.stopImmediatePropagation();
        scheduleFilter();
      }
    }, true);

    // Tastenkürzel
    document.addEventListener('keydown', function (e) {
      const t = e.target;
      if (!t || t.id !== fieldId) return;

      if (e.key === 'F4' || (e.key === 'ArrowDown' && e.altKey)) {
        showPanel();
        applyFilter(t.value || '');
        e.preventDefault();
        return;
      }

      if (e.key === 'Enter' || e.key === 'Tab') {
        showPanel();
        applyFilter(t.value || '');
        chooseFirstVisible();
        if (e.key === 'Enter') {
          e.preventDefault();
        }
      }
    }, true);

    console.log('DPD-Script: Unternehmen-Suche (Spalte A+B, Vollcache, Tastatur) aktiviert.');
  }

})();
