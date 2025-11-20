// ==UserScript==
// @name         DPD Belegnummer Auto-Fill + Sortierbare Spalte
// @namespace    https://bodo.dpd
// @version      1.2
// @description  Belegnummer automatisch mit letzter Belegnummer + 1 vorbelegen und Spalte "Beleg-Nr." sortierbar machen
// @match        https://dpd.lademittel.management/page/posting/postingOverview.xhtml*
// @match        https://dpd.lademittel.management/page/posting/postingCreate.xhtml*
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function() {
  'use strict';

  const href = window.location.href;
  const BELEG_INDEX = 5; // Spaltenindex "Beleg-Nr."

  if (href.includes('postingOverview.xhtml')) {
    runOnOverview();
  } else if (href.includes('postingCreate.xhtml')) {
    runOnCreate();
  }

  // ---------- Hilfsfunktionen ----------

  // "0426-355" -> { raw, prefix:"0426-", num:355, width:3 }
  function parseBelegNum(str) {
    const txt = (str || '').trim();
    if (!txt) return null;

    const match = txt.match(/^(.*?)(\d+)\s*$/);
    if (!match) return null;

    const prefix = match[1];
    const numStr = match[2];
    const width = numStr.length;
    const num = Number(numStr);

    if (Number.isNaN(num)) return null;

    return { raw: txt, prefix, num, width };
  }

  function formatNextBeleg(parsed) {
    if (!parsed) return null;
    const next = parsed.num + 1;
    const numStr = String(next).padStart(parsed.width, '0');
    return parsed.prefix + numStr;
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function getTable() {
    return document.querySelector('table');
  }

  // ----------------- Client-Sortierung für "Beleg-Nr." -----------------

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

      // Nicht-numerische Werte ("Ausgleich") nach hinten
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

    // Fallback über Textsuche, falls sich etwas verschiebt
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

    let descending = true; // Start: absteigend (größte Nummer oben)

    belegHeader.addEventListener('click', () => {
      sortBelegColumn(descending);
      indicator.textContent = descending ? '↓' : '↑';
      descending = !descending;
    });

    // initial einmal absteigend sortieren (für Auto-Erkennung)
    sortBelegColumn(true);
    indicator.textContent = '↓';
  }

  // ====================================================
  // 1. OVERVIEW: sortierbar machen + oberste numerische
  //              Belegnummer als "letzte" speichern
  // ====================================================

  function runOnOverview() {
    setTimeout(async () => {
      enableBelegHeaderSorting();

      // kleine Pause, dann oberste numerische Belegnummer holen
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
        console.log('DPD-Belegscript: letzte Belegnummer (aus sortierter Übersicht):', topParsed.raw);
        localStorage.setItem('dpd_lastBeleg', topParsed.raw);
      } else {
        console.warn('DPD-Belegscript: keine numerische Belegnummer gefunden.');
      }
    }, 800);
  }

  // ====================================================
  // 2. CREATE: letzte Belegnummer +1 ins Feld schreiben
  //    und als neue letzte merken
  // ====================================================

  function runOnCreate() {
    const last = localStorage.getItem('dpd_lastBeleg');
    if (!last) {
      console.warn('DPD-Belegscript: keine letzte Belegnummer im localStorage gefunden.');
      return;
    }

    const parsed = parseBelegNum(last);
    const nextBeleg = formatNextBeleg(parsed);

    if (!nextBeleg) {
      console.warn('DPD-Belegscript: nächste Belegnummer konnte nicht berechnet werden.');
      return;
    }

    const tryFill = () => {
      // dein Feld: postingEditForm:palletNoteNumber:validatableInput
      const input = document.getElementById('postingEditForm:palletNoteNumber:validatableInput');
      if (!input) return false;

      if (!input.value) {
        input.value = nextBeleg;
        input.dispatchEvent(new Event('input',  { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        console.log('DPD-Belegscript: Belegfeld mit', nextBeleg, 'befüllt.');

        // gleich als neue letzte Belegnummer speichern
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

})();
