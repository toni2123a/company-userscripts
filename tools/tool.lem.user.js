// ==UserScript==
// @name         DPD LEM
// @namespace    https://bodo.dpd
// @version      2.1
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

  // Header-Spalte "Beleg?" nach der Beleg-Nr. einfügen
  let fileTh = theadRow.querySelector('th.dpd-hasfile');
  if (!fileTh) {
    fileTh = document.createElement('th');
    fileTh.className = 'dpd-hasfile';
    fileTh.textContent = 'Beleg?';
    const ref = theadRow.children[BELEG_INDEX + 1] || null;
    theadRow.insertBefore(fileTh, ref);
  }

  // >>> NEU: Sortierung für "Beleg?"-Spalte
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

  // -------- Unternehmen-Filter + Tourfeld --------

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
        items = null;

        if (observer) observer.disconnect();
        observer = new MutationObserver(() => {
          items = null; // Panel-Inhalt geändert
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

    // globale Listener, damit es auch nach Neurendern funktioniert
    document.addEventListener('input', function(e) {
      const t = e.target;
      if (t && t.id === fieldId) {
        e.stopImmediatePropagation();
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
