// ==UserScript==
// @name         DPD LEM
// @namespace    https://bodo.dpd
// @version      2.8
// @description  Belegnummer automatisch mit letzter Belegnummer + 1 vorbelegen und Spalte "Beleg-Nr." sortierbar machen. Suche über Benutzerdefinierten Kundennamen
// @match        https://dpd.lademittel.management/page/posting/postingOverview.xhtml*
// @match        https://dpd.lademittel.management/page/posting/postingCreate.xhtml*
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.lem.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.lem.user.js
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-end
// ==/UserScript==
(function() {
  'use strict';

  const href = window.location.href;
  const BELEG_INDEX = 5; // "Beleg-Nr."-Spalte in der Übersicht
  let fileSortDescending = true;
 const LAST_BELEG_KEY = 'dpd_lastBeleg_v2'; // neuer Schlüssel im TM-Storage

  function getLastBeleg() {
    // 1. Neuer Speicher (Tampermonkey)
    let val = GM_getValue(LAST_BELEG_KEY, null);
    if (val) {
      console.log('DPD-Script: LAST_BELEG aus TM-Storage:', val);
      return val;
    }

    // 2. Fallback: alter localStorage-Wert (Migration)
    const legacy = localStorage.getItem('dpd_lastBeleg');
    if (legacy) {
      GM_setValue(LAST_BELEG_KEY, legacy);
      console.log('DPD-Script: legacy dpd_lastBeleg migriert nach TM-Storage:', legacy);
      return legacy;
    }

    // 3. gar nichts vorhanden
    console.log('DPD-Script: kein LAST_BELEG gefunden.');
    return null;
  }

  function setLastBeleg(value) {
    if (!value) return;
    const v = String(value).trim();
    GM_setValue(LAST_BELEG_KEY, v);
    // optional: alten Speicher weiter pflegen
    localStorage.setItem('dpd_lastBeleg', v);
    console.log('DPD-Script: LAST_BELEG gespeichert:', v);
  }

  // --- Einmalige Initialisierung (Startwert setzen) ---
  // Nur verwenden, wenn noch kein Wert existiert.
  (function initLastBelegOnce() {
    const existing = getLastBeleg();
    if (existing) {
      console.log('DPD-Script: LAST_BELEG existiert bereits:', existing);
      return;
    }

    // HIER deine aktuell letzte echte Belegnummer eintragen:
    const initial = '0815';   // Beispiel

    setLastBeleg(initial);
    console.log('DPD-Script: Initialer LAST_BELEG gesetzt auf:', initial);
  })();

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




function sortBelegColumnDescending() {
  const table = getTable();
  if (!table) {
    console.warn('DPD-Script: Tabelle nicht gefunden – Sortierung übersprungen.');
    return;
  }

  // Suche den TH-Header "Beleg-Nr."
  const ths = Array.from(table.querySelectorAll('thead th'));
  const belegTh = ths.find(th => th.textContent.trim().startsWith('Beleg-Nr'));
  if (!belegTh) {
    console.warn('DPD-Script: Beleg-Nr.-Spaltenheader nicht gefunden.');
    return;
  }

  // PrimeFaces benötigt 2 Klicks, um absteigend zu sortieren:
  // 1. Klick = aufsteigend
  // 2. Klick = absteigend (Z-A)
  belegTh.click(); // ascending
  setTimeout(() => belegTh.click(), 100); // descending

  console.log('DPD-Script: Beleg-Nr.-Spalte automatisch auf Z→A sortiert.');
}

function ensureBelegdatumDescending(callback) {
  const table = getTable();
  if (!table) {
    console.warn('DPD-Script: Tabelle nicht gefunden – kann Belegdatum nicht prüfen.');
    callback && callback();
    return;
  }

  const ths = Array.from(table.querySelectorAll('thead th'));
  const dateTh = ths.find(th => th.textContent.trim().startsWith('Belegdatum'));
  if (!dateTh) {
    console.warn('DPD-Script: Belegdatum-Spalte nicht gefunden.');
    callback && callback();
    return;
  }

  // prüfen, ob Belegdatum schon absteigend sortiert ist
  const ariaSort = dateTh.getAttribute('aria-sort');
  const icon = dateTh.querySelector(
    '.ui-sortable-column-icon, .pi, .fa'
  );
  const iconClass = icon ? icon.className : '';

  const isDescending =
    ariaSort === 'descending' ||
    /sort-amount-down|sort-down|pi-sort-amount-down/i.test(iconClass);

  if (isDescending) {
    console.log('DPD-Script: Belegdatum bereits absteigend sortiert.');
    callback && callback();
    return;
  }

  console.log('DPD-Script: sortiere Belegdatum absteigend ...');

  // einmalig auf pfAjaxComplete warten, wenn die Tabelle neu gerendert wurde
  function onPfAjaxComplete(e) {
    try {
      const detail = e.detail || {};
      const updateIds = detail.updateIds || detail.updateId || [];
      const ids = Array.isArray(updateIds) ? updateIds : [updateIds];

      // nur reagieren, wenn unsere Overview-Tabelle aktualisiert wurde
      if (!ids.some(id => typeof id === 'string' &&
                          id.indexOf('postingOverviewForm:postingOverviewTable') !== -1)) {
        return;
      }

      console.log('DPD-Script: Belegdatum-Sortierung fertig, Tabelle aktualisiert.');
      document.removeEventListener('pfAjaxComplete', onPfAjaxComplete, true);

      // ganz kurz warten, dann Callback
      setTimeout(() => callback && callback(), 100);
    } catch (err) {
      console.error('DPD-Script: Fehler im pfAjaxComplete-Handler (Belegdatum):', err);
      document.removeEventListener('pfAjaxComplete', onPfAjaxComplete, true);
      callback && callback();
    }
  }

  document.addEventListener('pfAjaxComplete', onPfAjaxComplete, true);

  // Klick auf den Belegdatum-Header → PrimeFaces sortiert serverseitig
  dateTh.click();
}

  // pro Zeile prüfen, ob in der Detailansicht ein Beleg existiert
   // NEU: pro Zeile prüfen, ob in der Detailansicht ein Beleg (Filepreview) existiert

 // =====================================================
// Übersicht – wir machen NICHTS mehr, um PF nicht zu brechen
// =====================================================

function runOnOverview() {
  // etwas warten, bis die Tabelle steht
  setTimeout(() => {
    ensureBelegdatumDescending(() => {
      updateLastBelegFromTopRow();
    });
  }, 500);
}

function updateLastBelegFromTopRow() {
  const table = getTable();
  if (!table) {
    console.warn('DPD-Script: Tabelle nicht gefunden – kein Beleg ermittelt.');
    return;
  }

  const firstRow = table.querySelector('tbody tr');
  if (!firstRow) {
    console.log('DPD-Script: keine Datenzeilen in der Tabelle.');
    return;
  }

  const td = firstRow.children[BELEG_INDEX];
  if (!td) {
    console.log('DPD-Script: Beleg-Nr.-Zelle nicht gefunden.');
    return;
  }

  const txt = td.textContent.trim();
  if (!txt) {
    console.log('DPD-Script: Beleg-Nr. in erster Zeile leer.');
    return;
  }

  setLastBeleg(txt);
console.log('DPD-Belegscript: letzte Belegnummer (über Belegdatum Z-A):', txt);

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

    function updateLastBelegFromSuccessMessage() {
  // Suche nach einem Element, das "Beleg-Nr." enthält
  const all = Array.from(document.querySelectorAll('div, span, li, p'));
  for (const el of all) {
    const txt = (el.textContent || '').trim();
    if (!txt || txt.indexOf('Beleg-Nr.') === -1) continue;

    // Zahl nach "Beleg-Nr." rausziehen (z.B. 5566 oder 0426-367)
    const m = txt.match(/Beleg-Nr\.\s*([0-9\-]+)/);
    if (!m) continue;

    const beleg = m[1];
    setLastBeleg(beleg);
    console.log('DPD-Script: LAST_BELEG aus Erfolgsmeldung gesetzt auf:', beleg);
    return; // ersten Treffer reicht
  }

  // keine Meldung gefunden → nichts tun
}

  // ====================================================
  // 2. Create-Seite
  //    - Belegnummer +1 (nur bei gespeicherter Buchung hochgezählt)
  //    - Unternehmen-Filter über beide Spalten (Client)
  //    - Tour aus 2. Spalte übernehmen
  // ====================================================

function runOnCreate() {
  // 1. Falls eine Erfolgsmeldung vorhanden ist, letzten Beleg aus der Meldung übernehmen
  updateLastBelegFromSuccessMessage();

  // 2. Neue Belegnummer aus dem gespeicherten Wert +1 vorbelegen
  autoFillBelegnummer();

  // 3. Rest wie gehabt
  enableUnternehmenSearch();
}



   function enableBelegTracking() {
  const belegInputId = 'postingEditForm:palletNoteNumber:validatableInput';
  const belegInput = document.getElementById(belegInputId);

  if (!belegInput) {
    console.warn('DPD-Script: Belegfeld für Tracking nicht gefunden.');
    return;
  }

  belegInput.addEventListener('change', function () {
    const val = (belegInput.value || '').trim();
    if (!val) {
      console.log('DPD-Script: Belegfeld leer beim change – Kandidat verworfen.');
      belegCandidate = null;
      return;
    }
    belegCandidate = val;
    console.log('DPD-Script: Beleg-Kandidat via change auf', val, 'gesetzt (noch nicht gespeichert).');
  }, true);
}

function attachSaveListener() {
  const form = document.getElementById('postingEditForm');
  if (!form) {
    console.warn('DPD-Script: postingEditForm nicht gefunden – Save-Listener nicht gesetzt.');
    return;
  }

  const belegInputId = 'postingEditForm:palletNoteNumber:validatableInput';

  form.addEventListener('submit', function () {
    const belegInput = document.getElementById(belegInputId);
    if (!belegInput) {
      console.warn('DPD-Script: Belegfeld beim Submit nicht gefunden.');
      return;
    }

    // zuerst Kandidaten nehmen, sonst aktuellen Feldwert
    const val = (belegCandidate || belegInput.value || '').trim();
    if (!val) {
      console.log('DPD-Script: kein Belegwert beim Speichern (submit), nichts gemerkt.');
      return;
    }

    setLastBeleg(val);
    console.log('DPD-Script: LAST_BELEG beim Speichern (submit) dauerhaft auf', val, 'gesetzt.');
  }, true);
}


  function autoFillBelegnummer() {
    const last = getLastBeleg();
      console.log('DPD-Script: autoFillBelegnummer – last =', last);
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
    let lastTypedTour = null;  // <- NEU: letzte vom Benutzer eingegebene Tournummer
    let panel    = null;
    let items    = null;   // [{row, search, col1, col2}]
    let allReady = false;
    let belegCandidate = null; // letzter geänderter Belegwert, noch nicht „fest“

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

    const tourVal = lastTypedTour;  // nur noch die gemerkte Benutzereingabe verwenden

    if (!tourVal) {
      console.log('DPD-Script: Tour-Feld nicht geändert (Benutzer hat keine Tournummer eingegeben).');
      return;
    }

    tourInput.value = tourVal;
    tourInput.dispatchEvent(new Event('input',  { bubbles: true }));
    tourInput.dispatchEvent(new Event('change', { bubbles: true }));
    console.log('DPD-Script: Tour-Feld gesetzt auf (Benutzereingabe):', tourVal);
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

  // Klick simulieren → PF wählt den Kunden,
  // attachRowClick setzt Tour anhand der Logik oben
  first.row.click();
  console.log('DPD-Script: erste gefilterte Zeile per Tastatur gewählt.');
}



    // --- Events ---

    preloadAllItems();

    let filterTimer = null;
 function scheduleFilter() {
  if (filterTimer) clearTimeout(filterTimer);
  filterTimer = setTimeout(() => {
    const inputEl = document.getElementById(fieldId);
    const val = inputEl ? inputEl.value : '';

    // NEU: Tournummer aus der Benutzereingabe merken (letzte Zahl)
    const nums = val.match(/\d+/g);
    lastTypedTour = nums ? nums[nums.length - 1] : null;

    applyFilter(val);
  }, 150);
}


    // Eingabe -> nur Client-Filter
   document.addEventListener('input', function (e) {
  const t = e.target;
  if (t && t.id === fieldId) {
    if (!e.isTrusted) return; // NEU: nur echte Tastatureingaben
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

  // Dropdown öffnen
  if (e.key === 'F4' || (e.key === 'ArrowDown' && e.altKey)) {
    showPanel();
    applyFilter(t.value || '');
    e.preventDefault();   // hier wollen wir im Feld bleiben
    return;
  }

  // Kunde per Enter/Tab auswählen
  if (e.key === 'Enter' || e.key === 'Tab') {
    showPanel();
    applyFilter(t.value || '');
    chooseFirstVisible();

    if (e.key === 'Enter') {
      // Enter: im Feld bleiben
      e.preventDefault();
    } else {
      // Tab: Fokus soll ins nächste Feld springen
      // -> kein preventDefault
    }
  }
}, true);


    console.log('DPD-Script: Unternehmen-Suche (Spalte A+B, Vollcache, Tastatur) aktiviert.');
  }

})();
