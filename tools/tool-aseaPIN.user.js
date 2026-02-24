// ==UserScript==
// @name         ASEA PIN Freigabe
// @namespace    http://tampermonkey.net/
// @version      6.11
// @description  Eingangsmengenabgleich: Tour-Bubbles + QR-Popup, Mehrfachauswahl + Liste kopieren (WhatsApp-Text) + Kopie (Sammelbild) + Kopie mit Code (Sammelbild inkl. Barcode je Zeile, Spaltenbreite automatisch) + Übersicht (Systempartner -> Anzahl, Zeitfenster aus aktueller Seite + Gesamtsumme).
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/scanmonitor\.cgi.*$/
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
  'use strict';

  const DEPOT = (window.location.hostname.match(/scanserver-d(\d{7})\.ssw\.dpdit\.de/i) || [])[1] || '0000000';
  const STORAGE_KEY = 'spTourConfigScanmonitor_v5';
  const COLLAPSE_KEY = STORAGE_KEY + '_collapsed';

  // =========================
  // QR / Barcode Helfer
  // =========================

  async function fetchQrContentForTour(tour) {
    const base = window.location.origin + '/lso/jcrp_ws/scanpocket/qrcode/clearance-granted?tour=';
    const url = base + encodeURIComponent(tour);
    const resp = await fetch(url, { credentials: 'include' });
    if (!resp.ok) throw new Error('HTTP-Status ' + resp.status + ' (' + resp.statusText + ')');

    let json;
    try { json = await resp.json(); } catch { throw new Error('Antwort ist kein gültiges JSON.'); }

    if (json && json.qrCode && typeof json.qrCode.code === 'string' && json.qrCode.code.trim() !== '') return json.qrCode.code;
    if (json && typeof json.code === 'string' && json.code.trim() !== '') return json.code;

    console.error('QR-API Antwort für Tour', tour, json);
    if (json && (json.error || json.message)) throw new Error('Server meldet: ' + (json.error || json.message));
    throw new Error('Server liefert keinen QR-Code für diese Tour (kein Feld "code" in der Antwort).');
  }

  function buildQrImageUrl(content) {
    return 'https://barcodeapi.org/api/qr/' + encodeURIComponent(content);
  }

  function buildCode128Url(content) {
    return 'https://barcodeapi.org/api/128/' + encodeURIComponent(content);
  }

  function buildFinalBarcodeFromRow(plz5Digits, paketNr, soCode) {
    return `%00${plz5Digits}${paketNr}${soCode}276`;
  }

  // =========================
  // Clipboard / Image
  // =========================

  async function copyCanvas(canvas) {
    if (!navigator.clipboard || !window.ClipboardItem) return false;
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    if (!blob) return false;
    try {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      return true;
    } catch {
      return false;
    }
  }

  async function copyTextToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', 'readonly');
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return ok;
  }

  async function fetchImageBitmap(url) {
    const resp = await fetch(url, { mode: 'cors' });
    if (!resp.ok) throw new Error('Bild nicht ladbar: ' + resp.status);
    const blob = await resp.blob();

    if (window.createImageBitmap) return await createImageBitmap(blob);

    return await new Promise((resolve, reject) => {
      const img = new Image();
      img.crossOrigin = 'anonymous';
      img.onload = () => resolve(img);
      img.onerror = () => reject(new Error('Bild decode fehlgeschlagen'));
      img.src = URL.createObjectURL(blob);
    });
  }

  // =========================
  // Text Wrap / Metrics
  // =========================

  function wrapTextLines(ctx, text, maxWidth, maxLines = 2) {
    const words = String(text || '').split(/\s+/).filter(Boolean);
    if (!words.length) return [''];
    const lines = [];
    let line = words[0];

    for (let i = 1; i < words.length; i++) {
      const test = line + ' ' + words[i];
      if (ctx.measureText(test).width <= maxWidth) {
        line = test;
      } else {
        lines.push(line);
        line = words[i];
        if (lines.length >= maxLines - 1) break;
      }
    }
    lines.push(line);

    if (lines.length === maxLines) {
      let t = lines[maxLines - 1];
      while (ctx.measureText(t + '…').width > maxWidth && t.length > 1) t = t.slice(0, -1);
      if (t !== lines[maxLines - 1]) lines[maxLines - 1] = t + '…';
    }

    return lines.slice(0, maxLines);
  }

  function plz5(s) {
    const m = String(s || '').match(/(\d{5})/);
    return m ? m[1] : '';
  }

  // =========================
  // Tabelle finden / Daten holen (frame-sicher)
  // =========================

  function collectAllDocs() {
    const docs = [];
    try { docs.push(document); } catch {}
    try {
      if (window.frames && window.frames.length) {
        for (let i = 0; i < window.frames.length; i++) {
          const f = window.frames[i];
          try { if (f && f.document) docs.push(f.document); } catch {}
        }
      }
    } catch {}
    return docs.filter(Boolean);
  }

  function normKey(s) {
    return String(s || '')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase()
      .replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue').replace(/ß/g, 'ss')
      .replace(/[^a-z0-9]+/g, '');
  }

  function getHeaderRow(table) {
    const rows = Array.from(table.querySelectorAll('tr'));
    if (!rows.length) return null;

    let best = null;
    let bestScore = -1;

    for (const r of rows.slice(0, 6)) {
      const cells = Array.from(r.querySelectorAll('th,td'));
      if (!cells.length) continue;

      const keys = cells.map(c => normKey(c.textContent)).filter(Boolean);
      const hasPsn = keys.includes('paketscheinnummer');
      const thCount = r.querySelectorAll('th').length;

      const score = (hasPsn ? 100 : 0) + thCount + keys.length * 0.2;
      if (score > bestScore) { bestScore = score; best = r; }
    }

    return best || rows[0];
  }

  function getColumnIndexMap(headerRow) {
    const cells = Array.from(headerRow.querySelectorAll('th,td'));
    const map = new Map();
    cells.forEach((cell, idx) => {
      const k = normKey(cell.textContent);
      if (k) map.set(k, idx);
    });
    return map;
  }

  function resolveIdx(colMap, want) {
    const k = normKey(want);
    if (colMap.has(k)) return colMap.get(k);

    const syn = {
      paketscheinnummer: ['paketscheinnummer', 'paketschein', 'sendungsnummer'],
      socode: ['socode', 'so'],
      zusatzcodes: ['zusatzcodes', 'zusatzcode', 'zusatz'],
      empfaengerplz: ['empfaengerplz', 'empfaenger plz', 'plz'],
      ort: ['ort', 'empfaengerort', 'empfaenger ort'],
      strasse: ['strasse', 'straße', 'str'],
      name1: ['name1', 'empfaengername', 'name'],
      name2: ['name2'],
      umverfuegung: ['umverfuegung', 'umverfugung']
    };

    const list = syn[k] || [];
    for (const cand of list) {
      const ck = normKey(cand);
      if (colMap.has(ck)) return colMap.get(ck);
    }
    return null;
  }

  function isRowVisibleInDoc(win, tr) {
    if (!tr) return false;
    if (tr.hidden) return false;

    const cs = win.getComputedStyle(tr);
    if (cs.display === 'none' || cs.visibility === 'hidden') return false;

    let el = tr;
    while (el && el !== tr.ownerDocument.body) {
      const s = win.getComputedStyle(el);
      if (s.display === 'none' || s.visibility === 'hidden') return false;
      el = el.parentElement;
    }
    return true;
  }

  function cellText(cell) {
    return (cell?.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function scoreTable(table) {
    const hr = getHeaderRow(table);
    if (!hr) return 0;
    const keys = new Set(Array.from(hr.querySelectorAll('th,td')).map(c => normKey(c.textContent)).filter(Boolean));
    const must = ['paketscheinnummer', 'socode', 'zusatzcodes', 'empfaengerplz', 'ort', 'strasse', 'name1', 'umverfuegung'];
    let score = 0;
    for (const m of must) if (keys.has(m)) score += 10;
    if (keys.has('name2')) score += 2;
    return score;
  }

  function findBestTableAnyDoc() {
    const docs = collectAllDocs();
    let best = null;

    for (const d of docs) {
      for (const t of Array.from(d.querySelectorAll('table'))) {
        const s = scoreTable(t);
        if (s <= 0) continue;
        if (!best || s > best.score) best = { doc: d, table: t, score: s };
      }
    }
    return best;
  }

  function extractVisibleRows_ANYWHERE() {
    const found = findBestTableAnyDoc();
    if (!found) throw new Error('Tabelle nicht gefunden.');

    const { doc: dataDoc, table } = found;
    const win = dataDoc.defaultView || window;

    const headerRow = getHeaderRow(table);
    if (!headerRow) throw new Error('Kopfzeile nicht erkannt.');

    const colMap = getColumnIndexMap(headerRow);

    const idx = {
      psn: resolveIdx(colMap, 'paketscheinnummer'),
      so:  resolveIdx(colMap, 'socode'),
      zc:  resolveIdx(colMap, 'zusatzcodes'),
      plz: resolveIdx(colMap, 'empfaengerplz'),
      ort: resolveIdx(colMap, 'ort'),
      str: resolveIdx(colMap, 'strasse'),
      n1:  resolveIdx(colMap, 'name1'),
      n2:  resolveIdx(colMap, 'name2'),
      umv: resolveIdx(colMap, 'umverfuegung')
    };

    const missing = Object.entries(idx).filter(([,v]) => v === null || v === undefined).map(([k]) => k);
    const missingReal = missing.filter(k => k !== 'n2');
    if (missingReal.length) throw new Error('Spalten nicht gefunden: ' + missingReal.join(', '));

    const allRows = Array.from(table.querySelectorAll('tr'));
    const headerIndex = allRows.indexOf(headerRow);

    const rows = [];
    for (let i = headerIndex + 1; i < allRows.length; i++) {
      const tr = allRows[i];
      if (!isRowVisibleInDoc(win, tr)) continue;

      const tds = Array.from(tr.querySelectorAll('td'));
      if (!tds.length) continue;

      const psn = cellText(tds[idx.psn]);
      if (!psn) continue;

      const so  = cellText(tds[idx.so]);
      const zc  = cellText(tds[idx.zc]);
      const plzDigits = plz5(cellText(tds[idx.plz]));
      const ort = cellText(tds[idx.ort]);
      const str = cellText(tds[idx.str]);
      const n1  = cellText(tds[idx.n1]);
      const n2  = (idx.n2 != null) ? cellText(tds[idx.n2]) : '';
      const umv = cellText(tds[idx.umv]);

      rows.push({
        psn, so, zc,
        plz: plzDigits,
        ort, str,
        name1: n1,
        name2: n2,
        name: (n1 + (n2 ? ' ' + n2 : '')).trim(),
        umv
      });
    }

    if (!rows.length) throw new Error('Keine sichtbaren Datenzeilen gefunden.');
    return rows;
  }

  // =========================
  // WhatsApp Text
  // =========================

  function padRight(s, w) {
    s = String(s ?? '');
    if (s.length >= w) return s;
    return s + ' '.repeat(w - s.length);
  }

  function clip(s, w) {
    s = String(s ?? '');
    if (s.length <= w) return s;
    if (w <= 1) return s.slice(0, w);
    return s.slice(0, w - 1) + '…';
  }

  function buildWhatsAppText(rows) {
    const list = rows.map(r => ({
      psn: r.psn,
      so: r.so,
      zc: r.zc,
      plzort: (r.plz + (r.ort ? ' ' + r.ort : '')).trim(),
      str: r.str,
      name: r.name,
      umv: r.umv
    }));

    const W = {
      psn:    Math.min(16, Math.max(12, ...list.map(r => r.psn.length))),
      so:     Math.min(6,  Math.max(2,  ...list.map(r => r.so.length))),
      zc:     Math.min(10, Math.max(6,  ...list.map(r => r.zc.length))),
      plzort: Math.min(22, Math.max(8,  ...list.map(r => r.plzort.length))),
      str:    Math.min(24, Math.max(10, ...list.map(r => r.str.length))),
      name:   Math.min(24, Math.max(10, ...list.map(r => r.name.length))),
      umv:    Math.min(12, Math.max(3,  ...list.map(r => r.umv.length)))
    };

    const header =
      padRight('Paketschein', W.psn) + '  ' +
      padRight('SO',         W.so)  + '  ' +
      padRight('Zusatz',     W.zc)  + '  ' +
      padRight('PLZ Ort',    W.plzort) + '  ' +
      padRight('Strasse',    W.str) + '  ' +
      padRight('Name',       W.name) + '  ' +
      padRight('Umv',        W.umv);

    const sep = '-'.repeat(header.length);

    const lines = [header, sep];
    for (const r of list) {
      lines.push(
        padRight(clip(r.psn,    W.psn),    W.psn)    + '  ' +
        padRight(clip(r.so,     W.so),     W.so)     + '  ' +
        padRight(clip(r.zc,     W.zc),     W.zc)     + '  ' +
        padRight(clip(r.plzort, W.plzort), W.plzort) + '  ' +
        padRight(clip(r.str,    W.str),    W.str)    + '  ' +
        padRight(clip(r.name,   W.name),   W.name)   + '  ' +
        padRight(clip(r.umv,    W.umv),    W.umv)
      );
    }

    return '```' + '\n' + lines.join('\n') + '\n' + '```';
  }

  // =========================
  // Sammelbild (wie zuvor)
  // =========================

  function computeAutoColumnWidths(ctx, rows) {
    const MIN = { psn: 130, so: 40, zc: 80, plzort: 120, str: 150, name: 160, umv: 70 };
    const MAX = { psn: 240, so: 70, zc: 160, plzort: 260, str: 360, name: 360, umv: 140 };

    const vals = rows.map(r => ({
      psn: r.psn || '',
      so: r.so || '',
      zc: r.zc || '',
      plzort: (r.plz + (r.ort ? ' ' + r.ort : '')).trim(),
      str: r.str || '',
      name: r.name || '',
      umv: r.umv || ''
    }));

    const keys = Object.keys(MIN);
    const W = {};
    for (const k of keys) {
      let w = ctx.measureText(k.toUpperCase()).width + 18;
      for (const v of vals) {
        const text = String(v[k] || '');
        w = Math.max(w, ctx.measureText(text).width + 18);
      }
      W[k] = Math.min(MAX[k], Math.max(MIN[k], Math.ceil(w)));
    }
    return W;
  }

  async function buildScreenshotCanvas(rows, { withBarcode }) {
    const scale = Math.min(2, Math.max(1, (window.devicePixelRatio || 1)));

    const pad = 12;
    const headerH = 34;
    const gap = 10;

    const barcodeW = 360;
    const barcodeH = 42;

    const lineH = 14;
    const rowPadY = 6;

    const measure = document.createElement('canvas');
    const mctx = measure.getContext('2d');
    mctx.font = '12px Arial';

    const col = computeAutoColumnWidths(mctx, rows);

    const textAreaW =
      col.psn + col.so + col.zc + col.plzort + col.str + col.name + col.umv + (6 * gap);

    const extraW = withBarcode ? (16 + barcodeW) : 0;

    const rowHeights = [];
    for (const r of rows) {
      const plzort = (r.plz + (r.ort ? ' ' + r.ort : '')).trim();
      const name = (r.name1 + (r.name2 ? ' ' + r.name2 : '')).trim();

      const counts = [
        wrapTextLines(mctx, r.psn, col.psn, 2).length,
        wrapTextLines(mctx, r.so, col.so, 2).length,
        wrapTextLines(mctx, r.zc, col.zc, 2).length,
        wrapTextLines(mctx, plzort, col.plzort, 2).length,
        wrapTextLines(mctx, r.str, col.str, 2).length,
        wrapTextLines(mctx, name, col.name, 2).length,
        wrapTextLines(mctx, r.umv, col.umv, 2).length
      ];

      const maxLines = Math.max(...counts, 1);
      const textHeight = rowPadY * 2 + (maxLines * lineH);
      const minHeight = withBarcode ? (rowPadY * 2 + barcodeH) : 0;

      rowHeights.push(Math.max(textHeight, minHeight, 26));
    }

    const totalRowsH = rowHeights.reduce((a, b) => a + b, 0);

    const widthCss = pad * 2 + textAreaW + extraW;
    const heightCss = pad * 2 + headerH + totalRowsH;

    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(widthCss * scale);
    canvas.height = Math.ceil(heightCss * scale);

    const ctx = canvas.getContext('2d');
    ctx.setTransform(scale, 0, 0, scale, 0, 0);

    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, widthCss, heightCss);

    ctx.fillStyle = '#666';
    ctx.font = '10px Arial';
    const ts = new Date().toLocaleString();
    const tw = ctx.measureText(ts).width;
    ctx.fillText(ts, widthCss - pad - tw, pad + 10);
    ctx.fillStyle = '#000';

    ctx.font = 'bold 12px Arial';
    let x = pad;
    const yHead = pad + 16;
    const drawHead = (label, w) => { ctx.fillText(label, x, yHead); x += w + gap; };

    drawHead('Paketschein', col.psn);
    drawHead('SO', col.so);
    drawHead('Zusatz', col.zc);
    drawHead('PLZ Ort', col.plzort);
    drawHead('Strasse', col.str);
    drawHead('Name', col.name);
    drawHead('Umv', col.umv);

    const barcodeX = pad + textAreaW;
    if (withBarcode) ctx.fillText('Barcode', barcodeX + 16, yHead);

    ctx.strokeStyle = '#ddd';
    ctx.beginPath();
    ctx.moveTo(pad, pad + headerH - 8);
    ctx.lineTo(widthCss - pad, pad + headerH - 8);
    ctx.stroke();

    ctx.font = '12px Arial';
    let yCursor = pad + headerH;

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const rowH = rowHeights[i];
      const y = yCursor;

      if (i % 2 === 1) {
        ctx.fillStyle = '#f7f7f7';
        ctx.fillRect(pad, y, widthCss - pad * 2, rowH);
        ctx.fillStyle = '#000';
      }

      const plzort = (r.plz + (r.ort ? ' ' + r.ort : '')).trim();
      const name = (r.name1 + (r.name2 ? ' ' + r.name2 : '')).trim();

      x = pad;
      const baseY = y + rowPadY + 12;

      const drawCell = (text, w) => {
        const lines = wrapTextLines(ctx, String(text || ''), w, 2);
        ctx.fillText(lines[0] || '', x, baseY);
        if (lines[1]) ctx.fillText(lines[1], x, baseY + lineH);
        x += w + gap;
      };

      drawCell(r.psn, col.psn);
      drawCell(r.so, col.so);
      drawCell(r.zc, col.zc);
      drawCell(plzort, col.plzort);
      drawCell(r.str, col.str);
      drawCell(name, col.name);
      drawCell(r.umv, col.umv);

      if (withBarcode) {
        const bx = barcodeX + 16;
        const by = y + Math.floor((rowH - barcodeH) / 2);

        const finalCode = buildFinalBarcodeFromRow(r.plz, r.psn, r.so);
        const url = buildCode128Url(finalCode);

        ctx.strokeStyle = '#ccc';
        ctx.strokeRect(bx, by, barcodeW, barcodeH);

        try {
          const bmp = await fetchImageBitmap(url);
          const iw = bmp.width || barcodeW;
          const ih = bmp.height || barcodeH;
          const s = Math.min(barcodeW / iw, barcodeH / ih);
          const dw = iw * s;
          const dh = ih * s;
          const dx = bx + (barcodeW - dw) / 2;
          const dy = by + (barcodeH - dh) / 2;
          ctx.drawImage(bmp, dx, dy, dw, dh);
        } catch {}
      }

      ctx.strokeStyle = '#eee';
      ctx.beginPath();
      ctx.moveTo(pad, y + rowH);
      ctx.lineTo(widthCss - pad, y + rowH);
      ctx.stroke();

      yCursor += rowH;
    }

    return canvas;
  }

  // =========================
  // Row Counter in Button Labels
  // =========================

  function getVisibleRowCountSafe() {
    try {
      const rows = extractVisibleRows_ANYWHERE();
      return Array.isArray(rows) ? rows.length : 0;
    } catch {
      return 0;
    }
  }

  function formatCopyLabel(base, n) {
    return `${n} ${base}`;
  }

  function startRowCountAutoUpdate(doc, containerEl) {
    if (!containerEl) return;

    const btnList = containerEl.querySelector('#tm-copy-list');
    const btnImg  = containerEl.querySelector('#tm-copy-img');
    const btnCode = containerEl.querySelector('#tm-copy-with-code');

    if (!btnList && !btnImg && !btnCode) return;

    if (btnList && !btnList.dataset.baseText) btnList.dataset.baseText = (btnList.textContent || 'Liste kopieren');
    if (btnImg  && !btnImg.dataset.baseText)  btnImg.dataset.baseText  = (btnImg.textContent  || 'Kopie');
    if (btnCode && !btnCode.dataset.baseText) btnCode.dataset.baseText = (btnCode.textContent || 'Kopie mit Code');

    if (containerEl.dataset.rowCountTimer === '1') return;
    containerEl.dataset.rowCountTimer = '1';

    let last = null;

    const tick = () => {
      const n = getVisibleRowCountSafe();
      if (n === last) return;
      last = n;

      if (btnList && !btnList.disabled) btnList.textContent = formatCopyLabel(btnList.dataset.baseText, n);
      if (btnImg  && !btnImg.disabled)  btnImg.textContent  = formatCopyLabel(btnImg.dataset.baseText,  n);
      if (btnCode && !btnCode.disabled) btnCode.textContent = formatCopyLabel(btnCode.dataset.baseText, n);
    };

    tick();
    window.setInterval(tick, 700);

    doc.addEventListener('visibilitychange', () => {
      if (!doc.hidden) tick();
    }, { passive: true });
  }

  // =========================
// =========================
// Übersicht: Systempartner -> Anzahl (0er ausblenden, Summe sticky, Zeitraum + alle aktuellen Filter aus Seite übernehmen)
// =========================

function isVisibleInput(el) {
  if (!el) return false;
  if (el.type === 'hidden') return false;
  if (el.disabled) return false;
  const r = el.getBoundingClientRect();
  if (!r || r.width <= 0 || r.height <= 0) return false;
  const cs = (el.ownerDocument.defaultView || window).getComputedStyle(el);
  if (cs.display === 'none' || cs.visibility === 'hidden') return false;
  return true;
}

// findet das "richtige" Dokument (auch wenn Frames) mit dem report-form
function findReportDoc() {
  const docs = collectAllDocs();
  let best = null;
  for (const d of docs) {
    const sel = d.querySelector('select[name="systempartner"]');
    const from = d.querySelector('input[name="stimestamp_from"]');
    const till = d.querySelector('input[name="stimestamp_till"]');
    if (sel && from && till) {
      if (isVisibleInput(from) && isVisibleInput(till)) return d;
      best = best || d;
    }
  }
  return best || document;
}

function readTimeRangeFromDoc(doc) {
  const re = /^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}$/;
  const fromEl = doc.querySelector('input[name="stimestamp_from"], input[id*="stimestamp_from"], input[name*="timestamp_from"], input[id*="timestamp_from"]');
  const tillEl = doc.querySelector('input[name="stimestamp_till"], input[id*="stimestamp_till"], input[name*="timestamp_till"], input[id*="timestamp_till"]');
  const fromV = (fromEl?.value || '').trim();
  const tillV = (tillEl?.value || '').trim();
  if (re.test(fromV) && re.test(tillV)) return { from: fromV, till: tillV };

  // fallback: URL params
  try {
    const p = new URLSearchParams(window.location.search);
    const from = (p.get('stimestamp_from') || '').trim();
    const till = (p.get('stimestamp_till') || '').trim();
    if (re.test(from) && re.test(till)) return { from, till };
  } catch {}
  return null;
}

// serialisiert ALLE aktuellen Filter aus dem Formular (inkl. Checkboxen/Selects)
// und gibt baseUrl + params zurück
function getBaseReportUrlAndParamsFromPage() {
  const doc = findReportDoc();

  // Ziel-URL: immer report_inbound_ofd.cgi
  const baseUrl = window.location.origin + '/cgi-bin/report_inbound_ofd.cgi';

  // bestes Formular finden: das, das systempartner enthält
  let form = null;
  const sel = doc.querySelector('select[name="systempartner"]');
  if (sel) form = sel.closest('form');
  if (!form) form = doc.querySelector('form') || null;

  const params = new URLSearchParams();

  if (form) {
    const fd = new FormData(form);

    // FormData -> URLSearchParams (inkl. mehrfach-Werte)
    for (const [k, v] of fd.entries()) {
      if (v === null || v === undefined) continue;
      const s = String(v).trim();
      // leere Werte (z.B. systempartner="") trotzdem zulassen? -> hier: ignorieren, außer systempartner/Zeitraum
      if (s === '' && !/^systempartner$/i.test(k) && !/^stimestamp_(from|till)$/i.test(k)) continue;
      params.append(k, s);
    }
  }

  // sicherstellen, dass notwendige "Standard"-Parameter vorhanden sind
  if (!params.has('doAction')) params.set('doAction', 'true');
  if (!params.has('reportType')) params.set('reportType', 'simple');
  if (!params.has('reportTypeSelect')) params.set('reportTypeSelect', params.get('reportType') || 'simple');
  if (!params.has('orderby')) params.set('orderby', 'PSN');
  if (!params.has('sortorder')) params.set('sortorder', 'ASC');

  // Zeitraum aus sichtbaren Feldern hart übernehmen (falls FormData ihn nicht hatte)
  const tr = readTimeRangeFromDoc(doc);
  if (tr) {
    params.set('stimestamp_from', tr.from);
    params.set('stimestamp_till', tr.till);
  }

  return { doc, baseUrl, params };
}

async function fetchHtml(url) {
  const resp = await fetch(url, { credentials: 'include' });
  if (!resp.ok) throw new Error('HTTP ' + resp.status + ' ' + resp.statusText);
  return await resp.text();
}

// robuster Count: findet PSN-Spalte und zählt Data-Rows
function countRowsInReportHtml(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');

  const tables = Array.from(doc.querySelectorAll('table'));
  if (!tables.length) return 0;

  let best = null;
  let bestScore = -1;
  for (const t of tables) {
    const s = scoreTable(t);
    if (s > bestScore) { bestScore = s; best = t; }
  }
  const table = best || tables[0];

  const headerRow = getHeaderRow(table);
  if (!headerRow) return 0;

  const allRows = Array.from(table.querySelectorAll('tr'));
  const headerIndex = allRows.indexOf(headerRow);

  const colMap = getColumnIndexMap(headerRow);
  const psnIdx = resolveIdx(colMap, 'paketscheinnummer') ?? 0;

  let cnt = 0;
  for (let i = headerIndex + 1; i < allRows.length; i++) {
    const tr = allRows[i];
    const tds = Array.from(tr.querySelectorAll('td'));
    if (!tds.length) continue;

    const psn = (tds[psnIdx]?.textContent || '').trim();
    if (!psn) continue;

    // minimal check: beginnt typischerweise mit Ziffer
    if (!/^\d/.test(psn)) continue;

    cnt++;
  }
  return cnt;
}

function getSystempartnerOptions(doc) {
  const sel = doc.querySelector('select[name="systempartner"]');
  if (!sel) return [];
  const opts = Array.from(sel.options || []);
  return opts
    .map(o => ({
      value: (o.value || '').toString().trim(),
      label: (o.textContent || '').toString().trim()
    }))
    .filter(o => o.label && !/^alle$/i.test(o.label));
}

function createOverviewOverlay(doc) {
  const overlay = doc.createElement('div');
  Object.assign(overlay.style, {
    position: 'fixed',
    inset: '0',
    background: 'rgba(0,0,0,0.45)',
    zIndex: '1000000',
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'center',
    paddingTop: '70px'
  });

  const box = doc.createElement('div');
  Object.assign(box.style, {
    background: '#fff',
    borderRadius: '6px',
    minWidth: '540px',
    maxWidth: '92vw',
    maxHeight: '80vh',
    overflow: 'auto',
    boxShadow: '0 10px 30px rgba(0,0,0,0.35)',
    fontFamily: 'Arial, sans-serif',
    fontSize: '12px',
    position: 'relative'
  });

  const head = doc.createElement('div');
  Object.assign(head.style, {
    padding: '10px 12px',
    borderBottom: '1px solid #e5e5e5',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '10px',
    position: 'sticky',
    top: '0',
    background: '#fff',
    zIndex: '2'
  });

  const title = doc.createElement('div');
  title.textContent = 'Übersicht – Systempartner (Anzahl angezeigter Zeilen)';
  title.style.fontWeight = 'bold';

  const sub = doc.createElement('div');
  sub.style.color = '#666';
  sub.style.marginTop = '2px';

  const left = doc.createElement('div');
  left.appendChild(title);
  left.appendChild(sub);

  const btnClose = doc.createElement('button');
  btnClose.type = 'button';
  btnClose.textContent = 'Schließen';
  btnClose.style.padding = '3px 8px';
  btnClose.style.cursor = 'pointer';

  head.appendChild(left);
  head.appendChild(btnClose);

  const body = doc.createElement('div');
  body.style.padding = '10px 12px 44px 12px';

  const table = doc.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';

  const thead = doc.createElement('thead');
  const thr = doc.createElement('tr');

  const th1 = doc.createElement('th');
  th1.textContent = 'Systempartner';
  th1.style.textAlign = 'left';
  th1.style.borderBottom = '1px solid #ddd';
  th1.style.padding = '6px 6px';

  const th2 = doc.createElement('th');
  th2.textContent = 'Anzahl';
  th2.style.textAlign = 'right';
  th2.style.borderBottom = '1px solid #ddd';
  th2.style.padding = '6px 6px';
  th2.style.width = '140px';

  thr.appendChild(th1);
  thr.appendChild(th2);
  thead.appendChild(thr);
  table.appendChild(thead);

  const tbody = doc.createElement('tbody');
  table.appendChild(tbody);

  body.appendChild(table);

  const sticky = doc.createElement('div');
  Object.assign(sticky.style, {
    position: 'sticky',
    bottom: '0',
    background: '#fff',
    borderTop: '2px solid #ddd',
    padding: '8px 12px',
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    fontWeight: 'bold',
    zIndex: '2'
  });

  const sumLabel = doc.createElement('div');
  sumLabel.textContent = 'Summe';

  const sumCell = doc.createElement('div');
  sumCell.textContent = '0';
  sumCell.style.fontVariantNumeric = 'tabular-nums';

  sticky.appendChild(sumLabel);
  sticky.appendChild(sumCell);

  box.appendChild(head);
  box.appendChild(body);
  box.appendChild(sticky);
  overlay.appendChild(box);

  const close = () => {
    try { overlay.remove(); } catch {}
    try { doc.removeEventListener('click', outsideClickCapture, true); } catch {}
    try { doc.removeEventListener('keydown', escClose, true); } catch {}
  };

  const outsideClickCapture = (e) => {
    if (!overlay.isConnected) return;
    if (!box.contains(e.target)) close();
  };

  const escClose = (e) => {
    if (e.key === 'Escape') close();
  };

  doc.addEventListener('click', outsideClickCapture, true);
  doc.addEventListener('keydown', escClose, true);

  overlay.addEventListener('click', (e) => {
    if (e.target === overlay) close();
  });

  box.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  btnClose.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    close();
  });

  return { overlay, sub, tbody, sumCell, close };
}

async function runOverview(doc) {
  // IMPORTANT: ab hier komplett aus dem "report-doc" lesen
  const { doc: rdoc, baseUrl, params: baseParams } = getBaseReportUrlAndParamsFromPage();
  const options = getSystempartnerOptions(rdoc);

  const ui = createOverviewOverlay(doc);
  doc.body.appendChild(ui.overlay);

  if (!options.length) {
    ui.sub.textContent = '(keine Systempartner gefunden)';
    return;
  }

  const from = (baseParams.get('stimestamp_from') || '').trim();
  const till = (baseParams.get('stimestamp_till') || '').trim();
  ui.sub.textContent = (from && till) ? `Zeitraum: ${from} – ${till}` : 'Zeitraum: (nicht gefunden)';

  ui.tbody.innerHTML = '';
  for (const o of options) {
    const tr = doc.createElement('tr');

    const td1 = doc.createElement('td');
    td1.textContent = o.label;
    td1.style.padding = '6px 6px';
    td1.style.borderBottom = '1px solid #f0f0f0';

    const td2 = doc.createElement('td');
    td2.textContent = '…';
    td2.style.padding = '6px 6px';
    td2.style.borderBottom = '1px solid #f0f0f0';
    td2.style.textAlign = 'right';
    td2.style.fontVariantNumeric = 'tabular-nums';

    tr.appendChild(td1);
    tr.appendChild(td2);
    ui.tbody.appendChild(tr);
  }

  const results = [];
  const CONCURRENCY = 4;
  let idx = 0;
  let done = 0;

  let runningSum = 0;

  function updateProgress() {
    const base = (from && till) ? `Zeitraum: ${from} – ${till}` : 'Zeitraum: (nicht gefunden)';
    ui.sub.textContent = `${base}   ${done}/${options.length}`;
    ui.sumCell.textContent = String(runningSum);
  }

  updateProgress();

  async function worker() {
    while (idx < options.length && ui.overlay.isConnected) {
      const opt = options[idx++];

      try {
        const p = new URLSearchParams(baseParams.toString());
        // systempartner MUSS gesetzt werden – einige Systeme erwarten auch leere value => dann label nutzen
        const sp = opt.value || opt.label;
        p.set('systempartner', sp);

        const url = baseUrl + '?' + p.toString();
        const html = await fetchHtml(url);
        const c = countRowsInReportHtml(html);

        const count = Number.isFinite(c) ? c : 0;
        results.push({ label: opt.label, count });
        runningSum += count;
      } catch (e) {
        results.push({ label: opt.label, count: 0, err: (e?.message || String(e)) });
      } finally {
        done++;
        updateProgress();
      }
    }
  }

  const workers = [];
  for (let i = 0; i < Math.min(CONCURRENCY, options.length); i++) workers.push(worker());
  await Promise.all(workers);

  if (!ui.overlay.isConnected) return;

  const nonZero = results.filter(r => (r.count || 0) > 0);
  nonZero.sort((a, b) => (b.count - a.count) || a.label.localeCompare(b.label));

  ui.tbody.innerHTML = '';

  if (!nonZero.length) {
    const tr = doc.createElement('tr');
    const td = doc.createElement('td');
    td.colSpan = 2;
    td.textContent = 'Keine Treffer (alle 0).';
    td.style.padding = '8px 6px';
    td.style.color = '#666';
    tr.appendChild(td);
    ui.tbody.appendChild(tr);
    ui.sumCell.textContent = '0';
  } else {
    let sum = 0;
    for (const it of nonZero) {
      const tr = doc.createElement('tr');

      const td1 = doc.createElement('td');
      td1.textContent = it.label;
      td1.style.padding = '6px 6px';
      td1.style.borderBottom = '1px solid #f0f0f0';

      const td2 = doc.createElement('td');
      td2.textContent = String(it.count);
      td2.style.padding = '6px 6px';
      td2.style.borderBottom = '1px solid #f0f0f0';
      td2.style.textAlign = 'right';
      td2.style.fontVariantNumeric = 'tabular-nums';

      tr.appendChild(td1);
      tr.appendChild(td2);
      ui.tbody.appendChild(tr);

      sum += it.count;
    }
    ui.sumCell.textContent = String(sum);
  }

  const base = (from && till) ? `Zeitraum: ${from} – ${till}` : 'Zeitraum: (nicht gefunden)';
  ui.sub.textContent = `${base}   fertig`;
}
  // =========================
  // Buttons unter Mehrfachauswahl
  // =========================

  function addCopyButtons(doc, containerEl) {
    if (!containerEl) return;

    if (false && !containerEl.querySelector('#tm-copy-list')) {
      const btn = doc.createElement('button');
      btn.id = 'tm-copy-list';
      btn.type = 'button';
      btn.textContent = 'Liste kopieren';

      btn.addEventListener('click', async () => {
        const old = btn.dataset.baseText || btn.textContent;
        btn.disabled = true;
        btn.textContent = 'Kopiere...';
        try {
          const rows = extractVisibleRows_ANYWHERE();
          const text = buildWhatsAppText(rows);
          const ok = await copyTextToClipboard(text);
          if (!ok) throw new Error('Clipboard nicht verfügbar.');
          btn.textContent = 'Kopiert ✓';
          setTimeout(() => { btn.textContent = old; btn.disabled = false; }, 1200);
        } catch (e) {
          console.error(e);
          btn.textContent = old;
          btn.disabled = false;
          alert('Konnte Liste nicht kopieren:\n' + (e?.message || e));
        }
      });

      containerEl.appendChild(btn);
    }

    if (!containerEl.querySelector('#tm-copy-img')) {
      const b = doc.createElement('button');
      b.id = 'tm-copy-img';
      b.type = 'button';
      b.textContent = 'Kopie';
      b.title = 'Kopiert ein Sammelbild wie Screenshot (ohne Barcode)';

      b.addEventListener('click', async () => {
        const old = b.dataset.baseText || b.textContent;
        b.disabled = true;
        b.textContent = 'Erzeuge Bild...';
        try {
          const rows = extractVisibleRows_ANYWHERE();
          const canvas = await buildScreenshotCanvas(rows, { withBarcode: false });
          b.textContent = 'Kopiere...';
          const ok = await copyCanvas(canvas);
          if (!ok) throw new Error('Browser blockiert Clipboard-Bild.');
          b.textContent = 'Kopiert ✓';
          setTimeout(() => { b.textContent = old; b.disabled = false; }, 1400);
        } catch (e) {
          console.error(e);
          b.textContent = old;
          b.disabled = false;
          alert('Kopie fehlgeschlagen:\n' + (e?.message || e));
        }
      });

      containerEl.appendChild(b);
    }

    if (false && !containerEl.querySelector('#tm-copy-with-code')) {
      const btn2 = doc.createElement('button');
      btn2.id = 'tm-copy-with-code';
      btn2.type = 'button';
      btn2.textContent = 'Kopie mit Code';
      btn2.title = 'Kopiert ein Sammelbild wie Screenshot inkl. Barcode je Zeile';

      btn2.addEventListener('click', async () => {
        const old = btn2.dataset.baseText || btn2.textContent;
        btn2.disabled = true;
        btn2.textContent = 'Erzeuge Bild...';
        try {
          const rows = extractVisibleRows_ANYWHERE();
          const bad = rows.find(r => !r.plz || !r.psn || !r.so);
          if (bad) throw new Error('Mindestens eine Zeile hat keine PLZ(5)/Paketschein/SOCode – Barcode kann nicht gebaut werden.');

          const canvas = await buildScreenshotCanvas(rows, { withBarcode: true });
          btn2.textContent = 'Kopiere...';
          const ok = await copyCanvas(canvas);
          if (!ok) throw new Error('Browser blockiert Clipboard-Bild.');

          btn2.textContent = 'Kopiert ✓';
          setTimeout(() => { btn2.textContent = old; btn2.disabled = false; }, 1400);
        } catch (e) {
          console.error(e);
          btn2.textContent = old;
          btn2.disabled = false;
          alert('Kopie mit Code fehlgeschlagen:\n' + (e?.message || e));
        }
      });

      containerEl.appendChild(btn2);
    }

    startRowCountAutoUpdate(doc, containerEl);
  }

  // =========================
  // QR Canvas + Popups (wie zuvor; unverändert)
  // =========================

  async function buildSingleCanvas(tour, imgUrl) {
    const pad = 18;
    const w = 420;
    const h = 520;

    const canvas = document.createElement('canvas');
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext('2d');

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);

    ctx.fillStyle = '#000000';
    ctx.font = 'bold 18px Arial';
    ctx.fillText('MDE Freigabe PIN (Zustellung)', pad, pad + 22);

    ctx.font = 'bold 16px Arial';
    ctx.fillText('Tour ' + tour, pad, pad + 52);

    ctx.font = '12px Arial';
    ctx.fillText('Depot: ' + DEPOT, pad, pad + 74);

    const qrSize = 320;
    const qrX = Math.floor((w - qrSize) / 2);
    const qrY = 110;

    try {
      const bmp = await fetchImageBitmap(imgUrl);
      ctx.drawImage(bmp, qrX, qrY, qrSize, qrSize);
    } catch {
      ctx.fillStyle = '#c00';
      ctx.font = '14px Arial';
      ctx.fillText('QR nicht ladbar', pad, qrY + 24);
    }

    return canvas;
  }

  async function buildMultiCanvas(tours, imgUrls) {
    const cols = 3;
    const cardW = 260;
    const cardH = 295;
    const pad = 16;
    const headerH = 52;

    const rows = Math.ceil(tours.length / cols);
    const width = pad * 2 + cols * cardW + (cols - 1) * pad;
    const height = pad * 2 + headerH + rows * cardH + (rows - 1) * pad;

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;

    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);

    ctx.fillStyle = '#000000';
    ctx.font = 'bold 18px Arial';
    ctx.fillText('MDE Freigabe PIN (Zustellung) – mehrere Touren', pad, pad + 22);
    ctx.font = '12px Arial';
    ctx.fillText('Depot: ' + DEPOT + ' | Touren: ' + tours.join(', '), pad, pad + 42);

    const qrSize = 200;
    const qrXoff = Math.floor((cardW - qrSize) / 2);
    const topBase = pad + headerH;

    for (let i = 0; i < tours.length; i++) {
      const t = tours[i];
      const col = i % cols;
      const row = Math.floor(i / cols);

      const x = pad + col * (cardW + pad);
      const y = topBase + row * (cardH + pad);

      ctx.fillStyle = '#f7f7f7';
      ctx.fillRect(x, y, cardW, cardH);
      ctx.strokeStyle = '#cccccc';
      ctx.strokeRect(x, y, cardW, cardH);

      ctx.fillStyle = '#000000';
      ctx.font = 'bold 14px Arial';
      ctx.fillText('Tour ' + t, x + 12, y + 22);

      const url = imgUrls[t];
      if (url) {
        try {
          const bmp = await fetchImageBitmap(url);
          ctx.drawImage(bmp, x + qrXoff, y + 40, qrSize, qrSize);
        } catch {
          ctx.fillStyle = '#c00';
          ctx.font = '12px Arial';
          ctx.fillText('QR nicht ladbar', x + 12, y + 60);
        }
      } else {
        ctx.fillStyle = '#c00';
        ctx.font = '12px Arial';
        ctx.fillText('QR fehlt', x + 12, y + 60);
      }

      ctx.fillStyle = '#000000';
      ctx.font = '12px Arial';
      drawWrappedText(ctx, 'Depot: ' + DEPOT, x + 12, y + 265, cardW - 24, 14);
    }

    return canvas;
  }

  function drawWrappedText(ctx, text, x, y, maxWidth, lineHeight) {
    const words = (text || '').split(' ');
    let line = '';
    for (let n = 0; n < words.length; n++) {
      const testLine = line + words[n] + ' ';
      const metrics = ctx.measureText(testLine);
      if (metrics.width > maxWidth && n > 0) {
        ctx.fillText(line.trim(), x, y);
        line = words[n] + ' ';
        y += lineHeight;
      } else {
        line = testLine;
      }
    }
    if (line.trim()) ctx.fillText(line.trim(), x, y);
    return y;
  }

  async function showQrPopup(doc, tour) {
    try {
      const content = await fetchQrContentForTour(tour);
      const imgUrl = buildQrImageUrl(content);

      const overlay = doc.createElement('div');
      Object.assign(overlay.style, {
        position: 'fixed', inset: '0', background: 'rgba(0,0,0,0.6)',
        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: '999999'
      });

      const box = doc.createElement('div');
      Object.assign(box.style, {
        background: '#fff', padding: '12px', borderRadius: '4px', textAlign: 'center',
        minWidth: '320px', boxShadow: '0 0 15px rgba(0,0,0,0.5)',
        fontFamily: 'Arial, sans-serif', fontSize: '12px'
      });

      const title = doc.createElement('div');
      title.textContent = 'MDE Freigabe PIN (Zustellung) – Tour ' + tour;
      title.style.marginBottom = '6px';
      box.appendChild(title);

      const img = doc.createElement('img');
      img.src = imgUrl;
      Object.assign(img.style, { maxWidth: '360px', maxHeight: '360px', marginBottom: '6px' });
      box.appendChild(img);

      const info = doc.createElement('div');
      info.textContent = 'Tour: ' + tour + ' | Depot: ' + DEPOT;
      info.style.marginBottom = '6px';
      box.appendChild(info);

      const btnClose = doc.createElement('button');
      btnClose.textContent = 'Schließen';
      btnClose.style.margin = '4px';
      btnClose.onclick = () => overlay.remove();
      box.appendChild(btnClose);

      const btnPrint = doc.createElement('button');
      btnPrint.textContent = 'Drucken';
      btnPrint.style.margin = '4px';
      btnPrint.onclick = () => {
        const w = window.open('', '_blank');
        if (!w) return;
        w.document.write(
          '<html><head><title>Tour ' + tour + '</title></head>' +
          '<body style="text-align:center;font-family:Arial, sans-serif;">' +
          '<h3>MDE Freigabe PIN (Zustellung)</h3>' +
          '<img src="' + imgUrl + '"><br>' +
          'Tour: ' + tour + ' | Depot: ' + DEPOT +
          '</body></html>'
        );
        w.document.close();
        w.focus();
        w.print();
      };
      box.appendChild(btnPrint);

      const btnCopy = doc.createElement('button');
      btnCopy.textContent = 'Kopieren';
      btnCopy.style.margin = '4px';
      btnCopy.onclick = async () => {
        btnCopy.disabled = true;
        const old = btnCopy.textContent;
        btnCopy.textContent = 'Kopiere...';
        try {
          const canvas = await buildSingleCanvas(tour, imgUrl);
          const ok = await copyCanvas(canvas);
          btnCopy.textContent = ok ? 'Kopiert ✓' : old;
          setTimeout(() => { btnCopy.textContent = old; btnCopy.disabled = false; }, 1200);
          if (!ok) alert('Kopieren nicht möglich.');
        } catch {
          btnCopy.textContent = old;
          btnCopy.disabled = false;
          alert('Kopieren fehlgeschlagen.');
        }
      };
      box.appendChild(btnCopy);

      overlay.appendChild(box);
      overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
      doc.body.appendChild(overlay);
    } catch (e) {
      alert('Fehler beim Laden des QR-Codes für Tour ' + tour + ':\n' + e.message);
    }
  }

  async function showMultiQrPopup(doc, tours) {
    if (!tours || !tours.length) return;

    const overlay = doc.createElement('div');
    Object.assign(overlay.style, {
      position: 'fixed', inset: '0', background: 'rgba(0,0,0,0.6)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: '999999'
    });

    const box = doc.createElement('div');
    Object.assign(box.style, {
      background: '#fff', padding: '12px', borderRadius: '4px', textAlign: 'center',
      minWidth: '340px', maxWidth: '90vw', maxHeight: '90vh', overflow: 'auto',
      boxShadow: '0 0 15px rgba(0,0,0,0.5)', fontFamily: 'Arial, sans-serif', fontSize: '12px'
    });

    const title = doc.createElement('div');
    title.textContent = 'MDE Freigabe PIN – mehrere Touren (' + tours.join(', ') + ')';
    title.style.marginBottom = '6px';
    box.appendChild(title);

    const grid = doc.createElement('div');
    Object.assign(grid.style, { display: 'flex', flexWrap: 'wrap', justifyContent: 'center', gap: '10px' });
    box.appendChild(grid);

    const info = doc.createElement('div');
    info.textContent = 'Touren: ' + tours.join(', ') + ' | Depot: ' + DEPOT;
    info.style.margin = '6px 0';
    box.appendChild(info);

    const btnClose = doc.createElement('button');
    btnClose.textContent = 'Schließen';
    btnClose.style.margin = '4px';
    btnClose.onclick = () => overlay.remove();
    box.appendChild(btnClose);

    const btnPrint = doc.createElement('button');
    btnPrint.textContent = 'Alle drucken';
    btnPrint.style.margin = '4px';
    box.appendChild(btnPrint);

    const btnCopy = doc.createElement('button');
    btnCopy.textContent = 'Kopieren';
    btnCopy.style.margin = '4px';
    box.appendChild(btnCopy);

    overlay.appendChild(box);
    overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
    doc.body.appendChild(overlay);

    const imgUrls = {};
    const loadedTours = [];

    for (const tour of tours) {
      const card = doc.createElement('div');
      Object.assign(card.style, {
        border: '1px solid #ccc', borderRadius: '4px', padding: '4px',
        textAlign: 'center', minWidth: '150px'
      });

      const h = doc.createElement('div');
      h.textContent = 'Tour ' + tour;
      h.style.marginBottom = '4px';
      card.appendChild(h);

      const status = doc.createElement('div');
      status.textContent = 'Lädt...';
      status.style.fontSize = '10px';
      status.style.color = '#666';
      card.appendChild(status);

      grid.appendChild(card);

      try {
        const content = await fetchQrContentForTour(tour);
        const url = buildQrImageUrl(content);
        imgUrls[tour] = url;
        loadedTours.push(tour);

        const img = doc.createElement('img');
        img.src = url;
        Object.assign(img.style, { maxWidth: '160px', maxHeight: '160px' });

        card.replaceChild(img, status);
      } catch (e) {
        status.textContent = 'Fehler: ' + e.message;
        status.style.color = '#c00';
      }
    }

    btnPrint.onclick = () => {
      const w = window.open('', '_blank');
      if (!w) return;
      let html = '<html><head><title>MDE Freigabe PIN – mehrere Touren</title></head><body style="text-align:center;font-family:Arial, sans-serif;">';
      html += '<h3>MDE Freigabe PIN (Zustellung) – mehrere Touren</h3>';
      tours.forEach(t => {
        const url = imgUrls[t];
        if (!url) return;
        html += '<div style="page-break-inside:avoid;margin-bottom:20px;">';
        html += '<div>Tour ' + t + ' | Depot: ' + DEPOT + '</div>';
        html += '<img src="' + url + '"><br>';
        html += '</div>';
      });
      html += '</body></html>';
      w.document.write(html);
      w.document.close();
      w.focus();
      w.print();
    };

    btnCopy.onclick = async () => {
      if (!loadedTours.length) {
        alert('Keine QR-Codes geladen – kein Kopieren möglich.');
        return;
      }

      btnCopy.disabled = true;
      const old = btnCopy.textContent;
      btnCopy.textContent = 'Kopiere...';

      try {
        const canvas = await buildMultiCanvas(loadedTours, imgUrls);
        const ok = await copyCanvas(canvas);

        btnCopy.textContent = ok ? 'Kopiert ✓' : old;
        setTimeout(() => { btnCopy.textContent = old; btnCopy.disabled = false; }, 1200);
        if (!ok) alert('Kopieren nicht möglich.');
      } catch {
        btnCopy.textContent = old;
        btnCopy.disabled = false;
        alert('Kopieren fehlgeschlagen.');
      }
    };
  }

  // =========================
  // Config / UI Panel
  // =========================

  function normalizeName(name) { return name.trim().toLowerCase(); }

  function loadConfig() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY) || '[]';
      const arr = JSON.parse(raw);
      if (Array.isArray(arr)) {
        return arr.map(x => ({
          name: (x.name || '').toString(),
          tours: Array.isArray(x.tours) ? x.tours.map(t => t.toString()) : []
        }));
      }
    } catch (e) {
      console.error('Konfig laden fehlgeschlagen:', e);
    }
    return [];
  }

  function initInDocument(doc) {
    if (!doc.body || doc.getElementById('tm-sp-panel')) return;

    const css = `
#tm-sp-panel{position:fixed;top:70px;right:10px;width:360px;max-height:80vh;overflow:auto;background:#f5f5f5;border:1px solid #ccc;padding:6px 8px;font-family:Arial,sans-serif;font-size:12px;z-index:9999;box-sizing:border-box;}
#tm-sp-panel h3{margin:0 0 4px 0;font-size:13px;padding-top:22px;}
#tm-sp-panel button{padding:2px 6px;font-size:11px;margin:2px 2px;cursor:pointer;}
#tm-sp-info{margin:4px 0;}
.tm-tour-bubble{display:inline-block;padding:3px 8px;margin:2px 4px 2px 0;border-radius:12px;background:#7d7d7d;color:#fff;cursor:pointer;white-space:nowrap;border:1px solid transparent;}
.tm-tour-bubble:hover{filter:brightness(1.1);}
.tm-tour-bubble.tm-selected{background:#4a4a4a;border-color:#000;}
#tm-collapse-btn{position:absolute;top:0;right:0;width:28px;height:28px;padding:0;margin:0;line-height:28px;text-align:center;font-weight:bold;}
#tm-overview-btn{position:absolute;top:0;left:0;height:28px;padding:0 8px;margin:0;line-height:28px;text-align:center;font-weight:bold;}
#tm-sp-panel.tm-collapsed{width:28px;height:28px;padding:0;overflow:hidden;background:transparent;border:none;}
#tm-sp-panel.tm-collapsed *{display:none !important;}
#tm-sp-panel.tm-collapsed #tm-collapse-btn{display:block !important;}
`;

    const style = doc.createElement('style');
    style.textContent = css;
    doc.head.appendChild(style);

    const panel = doc.createElement('div');
    panel.id = 'tm-sp-panel';

    const btnCollapse = doc.createElement('button');
    btnCollapse.id = 'tm-collapse-btn';
    btnCollapse.type = 'button';
    btnCollapse.textContent = '×';
    btnCollapse.title = 'Panel ein-/ausklappen';
    panel.appendChild(btnCollapse);

    const btnOverview = doc.createElement('button');
    btnOverview.id = 'tm-overview-btn';
    btnOverview.type = 'button';
    btnOverview.textContent = 'Übersicht';
    btnOverview.title = 'Systempartner-Übersicht (Hintergrundabfrage)';
    panel.appendChild(btnOverview);

    function applyCollapsedState(isCollapsed) {
      panel.classList.toggle('tm-collapsed', !!isCollapsed);
      try { window.localStorage.setItem(COLLAPSE_KEY, isCollapsed ? '1' : '0'); } catch {}
    }

    let collapsed = false;
    try { collapsed = (window.localStorage.getItem(COLLAPSE_KEY) === '1'); } catch {}
    applyCollapsedState(collapsed);

    btnCollapse.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      applyCollapsedState(!panel.classList.contains('tm-collapsed'));
    });

    btnOverview.addEventListener('click', async (e) => {
      e.preventDefault();
      e.stopPropagation();
      try {
        await runOverview(doc);
      } catch (err) {
        alert('Übersicht fehlgeschlagen:\n' + (err?.message || String(err)));
      }
    });

    const title = doc.createElement('h3');
    title.textContent = 'MDE Freigabe PIN (Zustellung)';
    panel.appendChild(title);

    const infoSpan = doc.createElement('div');
    infoSpan.id = 'tm-sp-info';
    infoSpan.textContent = 'Kein Systempartner ausgewählt.';
    panel.appendChild(infoSpan);

    const bubblesContainer = doc.createElement('div');
    bubblesContainer.id = 'tm-sp-bubbles';
    panel.appendChild(bubblesContainer);

    const multiDiv = doc.createElement('div');
    multiDiv.style.margin = '4px 0';

    const multiCheckbox = doc.createElement('input');
    multiCheckbox.type = 'checkbox';
    multiCheckbox.id = 'tm-multi-select';

    const multiLabel = doc.createElement('label');
    multiLabel.htmlFor = 'tm-multi-select';
    multiLabel.textContent = ' Mehrfachauswahl';

    const btnAll = doc.createElement('button');
    btnAll.textContent = 'Alle';
    btnAll.style.marginLeft = '6px';
    btnAll.style.display = 'none';

    const btnNone = doc.createElement('button');
    btnNone.textContent = 'Keine';
    btnNone.style.display = 'none';

    const multiShowBtn = doc.createElement('button');
    multiShowBtn.textContent = 'Anzeigen';
    multiShowBtn.style.marginLeft = '4px';

    multiDiv.appendChild(multiCheckbox);
    multiDiv.appendChild(multiLabel);
    multiDiv.appendChild(btnAll);
    multiDiv.appendChild(btnNone);
    multiDiv.appendChild(multiShowBtn);

    multiDiv.appendChild(doc.createElement('br'));
    const copyRow = doc.createElement('div');
    copyRow.id = 'tm-copy-row';
    copyRow.style.marginTop = '4px';
    multiDiv.appendChild(copyRow);

    addCopyButtons(doc, copyRow);

    panel.appendChild(multiDiv);
    doc.body.appendChild(panel);

    const selectedTours = new Set();

    function clearSelection() {
      selectedTours.clear();
      const bubbles = bubblesContainer.querySelectorAll('.tm-tour-bubble.tm-selected');
      bubbles.forEach(b => b.classList.remove('tm-selected'));
    }

    function selectAllVisibleTours() {
      const bubbles = bubblesContainer.querySelectorAll('.tm-tour-bubble');
      bubbles.forEach(b => {
        const tour = b.dataset.tour;
        if (!tour) return;
        selectedTours.add(tour);
        b.classList.add('tm-selected');
      });
    }

    function getSelectedSystempartnerNames(doc2) {
      const sel = doc2.querySelector('select[name="systempartner"]');
      if (!sel) return [];
      const names = [];
      const selOpts = sel.selectedOptions ? Array.from(sel.selectedOptions) : [];
      if (selOpts.length > 0) {
        selOpts.forEach(o => names.push(o.textContent.trim()));
        return names;
      }
      if (sel.value) {
        const opt = Array.from(sel.options).find(o => o.value === sel.value);
        if (opt) names.push(opt.textContent.trim());
      }
      return names;
    }

    function updateBubbles() {
      const cfg = loadConfig();
      const selectedPartners = getSelectedSystempartnerNames(doc);

      bubblesContainer.innerHTML = '';

      if (!selectedPartners.length) {
        infoSpan.textContent = 'Kein Systempartner ausgewählt.';
        return;
      }

      infoSpan.textContent = 'Systempartner: ' + selectedPartners.join(', ');

      const toursSet = new Set();
      selectedPartners.forEach(name => {
        const norm = normalizeName(name);
        const entry = cfg.find(p => normalizeName(p.name) === norm);
        if (entry && entry.tours) entry.tours.forEach(t => toursSet.add(t));
      });

      const tours = Array.from(toursSet);
      if (!tours.length) {
        const span = doc.createElement('span');
        span.textContent = 'Keine Touren hinterlegt.';
        span.style.color = '#999';
        bubblesContainer.appendChild(span);
        return;
      }

      tours.forEach(tour => {
        const bub = doc.createElement('span');
        bub.className = 'tm-tour-bubble';
        bub.textContent = tour;
        bub.dataset.tour = tour;

        if (selectedTours.has(tour)) bub.classList.add('tm-selected');

        bub.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          if (multiCheckbox.checked) {
            if (selectedTours.has(tour)) {
              selectedTours.delete(tour);
              bub.classList.remove('tm-selected');
            } else {
              selectedTours.add(tour);
              bub.classList.add('tm-selected');
            }
          } else {
            showQrPopup(doc, tour);
          }
        });

        bub.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          e.stopPropagation();
          if (multiCheckbox.checked) {
            const list = selectedTours.size ? Array.from(selectedTours) : [tour];
            showMultiQrPopup(doc, list);
          } else {
            showQrPopup(doc, tour);
          }
        });

        bubblesContainer.appendChild(bub);
      });
    }

    multiCheckbox.addEventListener('change', () => {
      const on = !!multiCheckbox.checked;
      btnAll.style.display = on ? '' : 'none';
      btnNone.style.display = on ? '' : 'none';
      if (!on) clearSelection();
    });

    btnAll.addEventListener('click', () => { selectAllVisibleTours(); });
    btnNone.addEventListener('click', () => { clearSelection(); });

    multiShowBtn.addEventListener('click', () => {
      if (!multiCheckbox.checked) { alert('Bitte zuerst „Mehrfachauswahl“ aktivieren.'); return; }
      const list = selectedTours.size ? Array.from(selectedTours) : [];
      if (!list.length) { alert('Bitte mindestens eine Tour markieren (oder „Alle“ drücken).'); return; }
      showMultiQrPopup(doc, list);
    });

    const sel = doc.querySelector('select[name="systempartner"]');
    if (sel) sel.addEventListener('change', updateBubbles);

    updateBubbles();
  }

  // =========================
  // Starter
  // =========================

  setInterval(() => {
    try {
      const docs = [document];

      if (window.frames && window.frames.length) {
        for (let i = 0; i < window.frames.length; i++) {
          const f = window.frames[i];
          try { if (f.document) docs.push(f.document); } catch {}
        }
      }

      for (const doc of docs) {
        if (!doc || !doc.body) continue;

        if (/Eingangsmengenabgleich/i.test(doc.body.textContent || '')) {
          const sel = doc.querySelector('select[name="systempartner"]');
          if (sel) initInDocument(doc);
        }

        if (/report_inbound_ofd\.cgi/i.test(window.location.pathname)) {
          const sel2 = doc.querySelector('select[name="systempartner"]');
          if (sel2) initInDocument(doc);
        }
      }
    } catch (e) {
      console.error(e);
    }
  }, 800);

})();
