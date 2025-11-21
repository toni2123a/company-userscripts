// ==UserScript==
// @name         DPD LEM
// @namespace    https://bodo.dpd
// @version      2.3
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
  //               + NEU: "Beleg?"-Spalte (hat Anhang?)
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

  async function updateLastBelegFromOverview() {
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
  }

  // NEU: pro Zeile prüfen, ob in der Detailansicht ein Beleg (Filepreview) existiert
function markRowsWithBeleg() {
  const table = getTable();
  if (!table) return;

  const theadRow = table.querySelector('thead tr');
  if (!theadRow) return;

  // Referenz auf vorhandenen Beleg-Nr.-Header holen
  const refBelegTh = theadRow.children[BELEG_INDEX];
  if (!refBelegTh) return;

  // Header-Spalte "Beleg?" nach der Beleg-Nr. einfügen
  let fileTh = theadRow.querySelector('th.dpd-hasfile');
  if (!fileTh) {
    // bestehenden Header klonen (ohne Kinder), damit Klassen/Attribute erhalten bleiben
    fileTh = refBelegTh.cloneNode(false);
    fileTh.classList.add('dpd-hasfile');
    fileTh.textContent = 'Beleg?';

    const ref = theadRow.children[BELEG_INDEX + 1] || null;
    theadRow.insertBefore(fileTh, ref);
  }

  // Sortier-Handler wie gehabt
  if (!fileTh.dataset.dpdSortable) {
    fileTh.dataset.dpdSortable = '1';
    fileTh.style.cursor = 'pointer';

    let indicator = fileTh.querySelector('.dpd-file-sort-indicator');
    if (!indicator) {
      indicator = document.createElement('span');
      indicator.className = 'dpd-file-sort-indicator';
      indicator.style.fontSize = '0.8em';
      indicator.style.marginLeft = '4px';
      indicator.textContent = '⇅';
      fileTh.appendChild(indicator);
    }

    fileTh.addEventListener('click', () => {
      sortFileColumn(fileSortDescending);
      indicator.textContent = fileSortDescending ? '↓' : '↑';
      fileSortDescending = !fileSortDescending;
    });
  }

  // ... ab hier dein bisheriger Code in markRowsWithBeleg() unverändert lassen


  const bodyRows = table.querySelectorAll('tbody tr');

  bodyRows.forEach(tr => {
    if (tr.dataset.dpdFileChecked === '1') return;
    tr.dataset.dpdFileChecked = '1';

    const tds = tr.children;
    let fileTd = tr.querySelector('td.dpd-hasfile');
    if (!fileTd) {
      fileTd = document.createElement('td');
      fileTd.className = 'dpd-hasfile';
      fileTd.textContent = '…';
      const refTd = tds[BELEG_INDEX + 1] || null;
      tr.insertBefore(fileTd, refTd);
    }

    const detailLink = tr.querySelector('a[href*="postingDetail.xhtml"]');
    if (!detailLink) {
      fileTd.textContent = '-';
      fileTd.title = 'kein Detail-Link';
      return;
    }

    const url = detailLink.href;
    fileTd.textContent = '⌛';
    fileTd.title = 'prüfe…';

    fetch(url, { credentials: 'include' })
      .then(resp => resp.text())
      .then(html => {
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const hasFile =
          !!doc.querySelector('.sae-filepreview-container') ||
          !!doc.querySelector('img[id*="filePreviewImageFile"]');

        if (hasFile) {
          fileTd.textContent = '📎';
          fileTd.title = 'Beleg vorhanden';
        } else {
          fileTd.textContent = '–';
          fileTd.title = 'kein Beleg';
        }
      })
      .catch(err => {
        console.error('Fehler beim Prüfen des Belegs:', err);
        fileTd.textContent = '?';
        fileTd.title = 'Fehler beim Prüfen';
      });
  });

  const tbody = table.querySelector('tbody');
  if (tbody && !tbody.dataset.dpdFileObserver) {
    const obs = new MutationObserver(() => {
      markRowsWithBeleg();
    });
    obs.observe(tbody, { childList: true });
    tbody.dataset.dpdFileObserver = '1';
  }
}


  function runOnOverview() {
    setTimeout(async () => {
      enableBelegHeaderSorting();
      await sleep(200);
      await updateLastBelegFromOverview();
      markRowsWithBeleg();
    }, 800);
  }

  // ====================================================
  // 2. Create-Seite
  //    - Belegnummer +1
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

  // -------- Unternehmen-Filter + Tourfeld --------

   // -------- Unternehmen-Filter + Tourfeld --------

   // -------- Unternehmen-Filter + Tourfeld + Tastenkürzel --------

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

        // NICHT überschreiben, wenn keine Zahl gefunden wurde
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


    // 1) EINMALIG: komplette Liste per PointerEvent laden
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

            // Panel wieder verstecken
            panel.style.display = 'none';
            panel.classList.add('ui-helper-hidden');

            clearInterval(intId);
          }
        }

        // Sicherheits-Timeout
        if (Date.now() - start > 5000) {
          console.warn('DPD-Script: Timeout beim Laden der Unternehmen-Liste.');
          clearInterval(intId);
        }
      }, 100);
    }

    // Panel anzeigen (nur noch clientseitig)
    function showPanel() {
      if (!ensurePanel()) return;
      panel.style.display = 'block';
      panel.classList.remove('ui-helper-hidden');
    }

    // Client-Filter über Spalte A + B
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

    // 4) ersten sichtbaren Treffer wählen (per Klick auf die Zeile)
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

      // Tour explizit aus unserer gespeicherten Spalte B setzen
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

    // schon beim Start alles vorladen
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

    // Eingabe abfangen -> nur unser Client-Filter, keine Server-Anfrage mehr
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

      // F4 oder Alt+↓ -> Panel öffnen + filtern
      if (e.key === 'F4' || (e.key === 'ArrowDown' && e.altKey)) {
        showPanel();
        applyFilter(t.value || '');
        e.preventDefault();
        return;
      }

      // Enter/Tab -> ersten sichtbaren Treffer wählen (per Klick)
      if (e.key === 'Enter' || e.key === 'Tab') {
        showPanel();
        applyFilter(t.value || '');
        chooseFirstVisible();
        if (e.key === 'Enter') {
          e.preventDefault(); // kein Submit
        }
        // Tab lassen wir durch, damit der Fokus weitergehen kann
      }
    }, true);

    console.log('DPD-Script: Unternehmen-Suche (Spalte A+B, Vollcache, Tastatur) aktiviert.');
  }






})();
