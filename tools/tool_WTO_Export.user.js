// ==UserScript==
// @name         WTO Shop – Artikel-Export (XLSX)
// @namespace    https://dpd.shop.wto-werbung.de/
// @version      1.5.0
// @description  Exportiert alle Artikel einer Kategorie inkl. aller Unterkategorien in festem Spaltenformat (Artikel-Nr., Artikel, Gruppe, Groessen, Laenge, Bereich, Preis (EUR), Materialnummer, Bestandspflege) als Excel- oder CSV-Datei
// @author       -
// @match        https://dpd.shop.wto-werbung.de/index.php?*controller=category*
// @match        https://dpd.shop.wto-werbung.de/*id_category=*
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ---------- Konfiguration ----------
  const PER_PAGE = 60;          // Produkte pro Kategorie-Abruf
  const CONCURRENCY = 5;        // parallele Detailseiten-Abrufe für die Referenz
  const REF_SELECTORS = [
    '#product_reference [itemprop="sku"]',
    '[itemprop="sku"]',
    '#product_reference .editable',
    '#product_reference span',
    '#product_reference',
  ];

  // ---------- Hilfsfunktionen ----------
  const qs = (sel, root = document) => root.querySelector(sel);
  const qsa = (sel, root = document) => Array.from(root.querySelectorAll(sel));
  const clean = (s) => (s || '').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();

  function getParam(name, url = location.href) {
    const m = url.match(new RegExp('[?&]' + name + '=([^&]+)'));
    return m ? decodeURIComponent(m[1]) : null;
  }

  async function fetchDoc(url) {
    const res = await fetch(url, { credentials: 'include' });
    if (!res.ok) throw new Error('HTTP ' + res.status + ' bei ' + url);
    const html = await res.text();
    return new DOMParser().parseFromString(html, 'text/html');
  }

  // Produkte aus einem Kategorie-Dokument auslesen
  function parseProducts(doc, catName) {
    return qsa('li.ajax_block_product', doc).map((li) => {
      const nameEl = qs('a.product-name', li);
      const name = clean(nameEl && (nameEl.getAttribute('title') || nameEl.textContent));
      const href = nameEl ? nameEl.href : '';
      const desc = clean(qs('p.product-desc', li) && qs('p.product-desc', li).textContent);
      const priceEl = qs('.right-block .product-price', li) || qs('.product-price', li);
      const price = clean(priceEl && priceEl.textContent);
      const id = (qs('[data-id-product]', li) && qs('[data-id-product]', li).getAttribute('data-id-product'))
        || getParam('id_product', href) || '';
      return { id, name, desc, price, href, reference: '', category: catName || '' };
    });
  }

  // Direkte Unterkategorie-IDs aus dem #subcategories-Block auslesen
  function getSubcategoryIds(doc) {
    const ids = [];
    qsa('#subcategories a.subcategory-name, #subcategories .subcategory-image > a', doc).forEach((a) => {
      const id = getParam('id_category', a.href);
      if (id && !ids.includes(id)) ids.push(id);
    });
    return ids;
  }

  // Alle (verschachtelten) Unterkategorie-IDs aus dem linken Kategorienbaum,
  // begrenzt auf den Teilbaum der übergebenen Kategorie. Fängt Ebenen ab,
  // die keinen #subcategories-Block anzeigen.
  function getTreeDescendantIds(doc, catId) {
    const cid = String(catId);
    const node = qsa('#categories_block_left a', doc).find((a) => getParam('id_category', a.href) === cid);
    if (!node || !node.closest) return [];
    const li = node.closest('li');
    if (!li) return [];
    const ids = [];
    qsa('a', li).forEach((a) => {
      const id = getParam('id_category', a.href);
      if (id && id !== cid && !ids.includes(id)) ids.push(id);
    });
    return ids;
  }

  function docCategoryName(doc, fallback) {
    return clean(qs('.cat-name', doc) && qs('.cat-name', doc).textContent) || fallback;
  }

  function getTotalCount(doc) {
    const el = qs('.heading-counter', doc) || qs('.product-count', doc);
    const m = el && el.textContent.match(/(\d+)\s*Artikel/);
    return m ? parseInt(m[1], 10) : null;
  }

  // Referenz/Artikelnummer aus einer Produkt-Detailseite auslesen
  function extractReference(doc) {
    for (const sel of REF_SELECTORS) {
      const el = qs(sel, doc);
      if (el) {
        const val = clean(el.getAttribute && el.getAttribute('content') ? el.getAttribute('content') : el.textContent);
        if (val) return val;
      }
    }
    return '';
  }

  // Aufgaben mit begrenzter Parallelität abarbeiten
  async function runPool(items, worker, onProgress) {
    let index = 0, done = 0;
    async function next() {
      while (index < items.length) {
        const i = index++;
        try { await worker(items[i], i); } catch (e) { console.warn('Fehler:', e); }
        done++;
        if (onProgress) onProgress(done, items.length);
      }
    }
    await Promise.all(Array.from({ length: Math.min(CONCURRENCY, items.length) }, next));
  }

  const COLUMNS = ['Artikel-Nr.', 'Artikel', 'Gruppe', 'Groessen', 'Laenge', 'Bereich', 'Preis (EUR)', 'Materialnummer', 'Bestandspflege'];

  function categoryName() {
    return clean(qs('.cat-name') && qs('.cat-name').textContent) || 'Kategorie';
  }

  // Preis auf reine Zahl reduzieren: "3,72 €" -> "3,72"
  function priceNumber(s) {
    return clean(s).replace(/[^\d.,]/g, '');
  }

  function productToRow(p) {
    return {
      'Artikel-Nr.': p.reference,
      'Artikel': p.name,
      'Gruppe': p.category,
      'Groessen': '',
      'Laenge': '',
      'Bereich': 'WTO',
      'Preis (EUR)': priceNumber(p.price),
      'Materialnummer': '',
      'Bestandspflege': '',
    };
  }

  // Alle Produkte einer einzelnen Kategorie (alle Seiten) einsammeln
  async function collectCategoryProducts(catId, lang, base, status) {
    const firstUrl = `${base}?id_category=${catId}&controller=category&id_lang=${lang}&n=${PER_PAGE}&p=1`;
    const firstDoc = await fetchDoc(firstUrl);
    const catName = docCategoryName(firstDoc, 'Kategorie ' + catId);
    const total = getTotalCount(firstDoc);
    const totalPages = total ? Math.ceil(total / PER_PAGE) : 1;

    let products = parseProducts(firstDoc, catName);
    for (let p = 2; p <= totalPages; p++) {
      status(`Lade „${catName}" – Seite ${p}/${totalPages} …`);
      const doc = await fetchDoc(`${base}?id_category=${catId}&controller=category&id_lang=${lang}&n=${PER_PAGE}&p=${p}`);
      products = products.concat(parseProducts(doc, catName));
    }
    // Unterkategorien aus beiden Quellen kombinieren (Vollständigkeit)
    const subCatIds = [];
    getSubcategoryIds(firstDoc).concat(getTreeDescendantIds(firstDoc, catId)).forEach((id) => {
      if (!subCatIds.includes(id)) subCatIds.push(id);
    });
    return { products, subCatIds };
  }

  // Sammelt alle Artikel der Kategorie UND aller (verschachtelten) Unterkategorien
  async function collectProducts(status) {
    const startId = getParam('id_category');
    const lang = getParam('id_lang') || '1';
    const base = location.origin + location.pathname;
    if (!startId) { alert('Keine Kategorie-ID in der URL gefunden.'); return null; }

    // 1) Kategoriebaum ab der aktuellen Kategorie durchlaufen
    const visited = new Set();
    const queue = [startId];
    let products = [];

    while (queue.length) {
      const catId = queue.shift();
      if (visited.has(catId)) continue;
      visited.add(catId);
      status(`Durchsuche Kategorien … (${visited.size} geladen, ${queue.length} offen)`);
      const { products: catProducts, subCatIds } = await collectCategoryProducts(catId, lang, base, status);
      products = products.concat(catProducts);
      subCatIds.forEach((id) => { if (!visited.has(id) && !queue.includes(id)) queue.push(id); });
    }
    console.log('[WTO-Export] Besuchte Kategorien:', visited.size, [...visited]);
    console.log('[WTO-Export] Gefundene Artikel (vor Dedup):', products.length);

    // 2) Duplikate anhand der id_product entfernen (Produkt kann in mehreren Kategorien liegen)
    const seen = new Set();
    products = products.filter((p) => {
      const key = p.id || p.href;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    if (!products.length) { alert('Keine Artikel gefunden.'); return null; }

    // 3) Referenz/Artikelnummer je Produkt von der Detailseite holen
    await runPool(
      products,
      async (prod) => {
        if (!prod.href) return;
        const doc = await fetchDoc(prod.href);
        prod.reference = extractReference(doc);
      },
      (done, all) => status(`Lese Artikelnummern … ${done}/${all}`)
    );

    return products;
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ---------- Export: XLSX ----------
  async function exportXlsx(status) {
    const products = await collectProducts(status);
    if (!products) return;

    status('Erstelle Excel-Datei …');
    const ws = XLSX.utils.json_to_sheet(products.map(productToRow));
    ws['!cols'] = [{ wch: 18 }, { wch: 45 }, { wch: 15 }, { wch: 12 }, { wch: 14 }, { wch: 15 }, { wch: 12 }, { wch: 15 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    const catName = categoryName();
    XLSX.utils.book_append_sheet(wb, ws, catName.substring(0, 31) || 'Artikel');
    XLSX.writeFile(wb, `WTO_${catName.replace(/[^\w\-]+/g, '_')}_${getParam('id_category')}.xlsx`);
    status(`Fertig: ${products.length} Artikel exportiert.`);
  }

  // ---------- Export: CSV (Fallback) ----------
  function csvEscape(val) {
    const s = String(val == null ? '' : val);
    return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }

  async function exportCsv(status) {
    const products = await collectProducts(status);
    if (!products) return;

    status('Erstelle CSV-Datei …');
    const lines = [COLUMNS.join(';')];
    for (const p of products) {
      const row = productToRow(p);
      lines.push(COLUMNS.map((c) => csvEscape(row[c])).join(';'));
    }
    // BOM für korrekte Umlaut-Darstellung in Excel, CRLF als Zeilenende
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const catName = categoryName();
    downloadBlob(blob, `WTO_${catName.replace(/[^\w\-]+/g, '_')}_${getParam('id_category')}.csv`);
    status(`Fertig: ${products.length} Artikel exportiert.`);
  }

  // ---------- UI ----------
  function makeButton(label, bg) {
    const btn = document.createElement('button');
    btn.textContent = label;
    Object.assign(btn.style, {
      padding: '12px 18px', background: bg, color: '#fff', border: 'none',
      borderRadius: '6px', cursor: 'pointer', fontSize: '14px', fontWeight: 'bold',
      boxShadow: '0 2px 8px rgba(0,0,0,.3)', display: 'block', width: '100%',
      marginTop: '8px',
    });
    return btn;
  }

  function createUI() {
    if (qs('#wto-export-box')) return;
    const box = document.createElement('div');
    box.id = 'wto-export-box';
    Object.assign(box.style, {
      position: 'fixed', bottom: '20px', right: '20px', zIndex: 99999, width: '260px',
    });

    const xlsxBtn = makeButton('⬇ Artikel als Excel (XLSX)', '#d0121a');
    const csvBtn = makeButton('⬇ Artikel als CSV', '#444');

    const buttons = [xlsxBtn, csvBtn];
    function bind(btn, label, fn) {
      const status = (msg) => { btn.textContent = msg; };
      btn.addEventListener('click', async () => {
        buttons.forEach((b) => { b.disabled = true; b.style.opacity = '.7'; });
        try {
          await fn(status);
        } catch (e) {
          console.error(e);
          status('Fehler – siehe Konsole');
          alert('Export fehlgeschlagen: ' + e.message);
        } finally {
          setTimeout(() => {
            buttons.forEach((b) => { b.disabled = false; b.style.opacity = '1'; });
            btn.textContent = label;
          }, 4000);
        }
      });
    }
    bind(xlsxBtn, '⬇ Artikel als Excel (XLSX)', exportXlsx);
    bind(csvBtn, '⬇ Artikel als CSV', exportCsv);

    box.appendChild(xlsxBtn);
    box.appendChild(csvBtn);
    document.body.appendChild(box);
  }

  createUI();
})();
