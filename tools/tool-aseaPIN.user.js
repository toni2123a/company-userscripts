// ==UserScript==
// @name         ASEA PIN Freigabe
// @namespace    http://tampermonkey.net/
// @version      6.50
// @description  Eingangsmengenabgleich: Tour-Bubbles + QR-Popup, Mehrfachauswahl + Liste kopieren (WhatsApp-Text) + Kopie (Sammelbild) + Kopie mit Code (Sammelbild inkl. Barcode je Zeile, Spaltenbreite automatisch) + Übersicht (Systempartner -> Anzahl, Zeitfenster aus aktueller Seite + Gesamtsumme) + Einstellungen (Systempartner/Touren, Import und Export).
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/scanmonitor\.cgi.*$/
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @connect      scanserver-d0010195.ssw.dpdit.de
// @connect      scanserver-d0010107.ssw.dpdit.de
// @connect      scanserver-d0010295.ssw.dpdit.de
// @connect      scanserver-d001*
// @connect      barcodeapi.org
// @connect      *
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
  'use strict';

  const DEPOT = (window.location.hostname.match(/scanserver-d(\d{7})\.ssw\.dpdit\.de/i) || [])[1] || '0000000';
  const STORAGE_KEY = 'spTourConfigScanmonitor_v5';
  const COLLAPSE_KEY = STORAGE_KEY + '_collapsed';
  const DEPOTPORTAL_BASE_URL = 'https://depotportal.dpd.com/dp/de_DE/tracking/parcels/';

  // =========================
  // Depotportal / Lebenslauf
  // =========================

  function cleanPsn(value) {
    return String(value || '').replace(/\D/g, '').trim();
  }

  function buildDepotportalUrl(psn) {
    const nr = cleanPsn(psn);
    return nr ? (DEPOTPORTAL_BASE_URL + encodeURIComponent(nr)) : '';
  }

  function openDepotportalTracking(psn) {
    const url = buildDepotportalUrl(psn);
    if (!url) return;
    window.open(url, '_blank', 'noopener');
  }


  // =========================
  // QR / Barcode Helfer
  // =========================

  function gmRequestText(url, headers = {}) {
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest !== 'function') {
        reject(new Error('GM_xmlhttpRequest nicht verfügbar.'));
        return;
      }

      GM_xmlhttpRequest({
        method: 'GET',
        url,
        headers,
        timeout: 30000,
        anonymous: false,
        onload: function (resp) {
          resolve({
            ok: resp.status >= 200 && resp.status < 300,
            status: resp.status,
            statusText: resp.statusText || '',
            text: resp.responseText || ''
          });
        },
        onerror: function () {
          reject(new Error('GM Request fehlgeschlagen.'));
        },
        ontimeout: function () {
          reject(new Error('GM Request Timeout.'));
        }
      });
    });
  }

  async function fetchQrContentForTour(tour) {
    const url = window.location.origin + '/lso/jcrp_ws/scanpocket/qrcode/clearance-granted?tour=' + encodeURIComponent(tour);
    const commonHeaders = {
      'Accept': 'application/json, text/plain, */*',
      'X-Requested-With': 'XMLHttpRequest'
    };

    let lastErr = null;

    try {
      const resp = await fetch(url, {
        method: 'GET',
        credentials: 'include',
        cache: 'no-store',
        headers: commonHeaders
      });

      if (!resp.ok) {
        let extra = '';
        try { extra = await resp.text(); } catch {}
        throw new Error(
          'HTTP-Status ' + resp.status + ' (' + resp.statusText + ')' +
          (extra ? '\nAntwort: ' + extra.slice(0, 300) : '')
        );
      }

      let json;
      try {
        json = await resp.json();
      } catch {
        throw new Error('Antwort ist kein gültiges JSON.');
      }

      if (json && json.qrCode && typeof json.qrCode.code === 'string' && json.qrCode.code.trim() !== '') {
        return json.qrCode.code;
      }
      if (json && typeof json.code === 'string' && json.code.trim() !== '') {
        return json.code;
      }

      console.error('QR-API Antwort für Tour', tour, json);
      if (json && (json.error || json.message)) {
        throw new Error('Server meldet: ' + (json.error || json.message));
      }
      throw new Error('Server liefert keinen QR-Code für diese Tour.');
    } catch (e) {
      lastErr = e;
      console.warn('Normaler fetch für QR fehlgeschlagen, versuche GM_xmlhttpRequest Fallback:', e);
    }

    try {
      const gmResp = await gmRequestText(url, commonHeaders);

      if (!gmResp.ok) {
        throw new Error(
          'HTTP-Status ' + gmResp.status + ' (' + gmResp.statusText + ')' +
          (gmResp.text ? '\nAntwort: ' + gmResp.text.slice(0, 300) : '')
        );
      }

      let json;
      try {
        json = JSON.parse(gmResp.text);
      } catch {
        throw new Error('Antwort ist kein gültiges JSON.');
      }

      if (json && json.qrCode && typeof json.qrCode.code === 'string' && json.qrCode.code.trim() !== '') {
        return json.qrCode.code;
      }
      if (json && typeof json.code === 'string' && json.code.trim() !== '') {
        return json.code;
      }

      console.error('QR-API Antwort für Tour', tour, json);
      if (json && (json.error || json.message)) {
        throw new Error('Server meldet: ' + (json.error || json.message));
      }
      throw new Error('Server liefert keinen QR-Code für diese Tour.');
    } catch (e2) {
      const msg1 = lastErr?.message ? ('Fetch: ' + lastErr.message) : '';
      const msg2 = e2?.message ? ('GM: ' + e2.message) : '';
      throw new Error([msg1, msg2].filter(Boolean).join('\n'));
    }
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

  async function fetchImageBitmap(url) {
    try {
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
    } catch (e1) {
      if (typeof GM_xmlhttpRequest !== 'function') throw e1;

      const blob = await new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
          method: 'GET',
          url,
          responseType: 'blob',
          timeout: 30000,
          anonymous: false,
          onload: function (resp) {
            if (resp.status >= 200 && resp.status < 300 && resp.response) {
              resolve(resp.response);
            } else {
              reject(new Error('GM Bild nicht ladbar: ' + resp.status));
            }
          },
          onerror: function () {
            reject(new Error('GM Bild-Request fehlgeschlagen'));
          },
          ontimeout: function () {
            reject(new Error('GM Bild-Request Timeout'));
          }
        });
      });

      if (window.createImageBitmap) {
        try { return await createImageBitmap(blob); } catch {}
      }

      return await new Promise((resolve, reject) => {
        const img = new Image();
        const objUrl = URL.createObjectURL(blob);
        img.onload = () => {
          URL.revokeObjectURL(objUrl);
          resolve(img);
        };
        img.onerror = () => {
          URL.revokeObjectURL(objUrl);
          reject(new Error('GM Bild decode fehlgeschlagen'));
        };
        img.src = objUrl;
      });
    }
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
  // Tabelle finden / Daten holen
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

  function findBestTableInDoc(doc) {
    let best = null;
    for (const t of Array.from(doc.querySelectorAll('table'))) {
      const s = scoreTable(t);
      if (s <= 0) continue;
      if (!best || s > best.score) best = { table: t, score: s };
    }
    return best;
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

  function extractRowsFromTable(table, options = {}) {
    const {
      visibilityMode = 'visible',
      win = null
    } = options || {};

    if (!table) throw new Error('Tabelle nicht gefunden.');

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
      if (visibilityMode === 'visible' && win && !isRowVisibleInDoc(win, tr)) continue;

      const tds = Array.from(tr.querySelectorAll('td'));
      if (!tds.length) continue;

      const psn = cellText(tds[idx.psn]);
      if (!psn) continue;
      if (!/^\d/.test(psn)) continue;

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

    if (!rows.length) throw new Error('Keine Datenzeilen gefunden.');
    return rows;
  }

  function extractVisibleRows_ANYWHERE() {
    const found = findBestTableAnyDoc();
    if (!found) throw new Error('Tabelle nicht gefunden.');
    const win = found.doc.defaultView || window;
    return extractRowsFromTable(found.table, { visibilityMode: 'visible', win });
  }

  function extractRowsFromHtml(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const best = findBestTableInDoc(doc);
    if (!best) throw new Error('Tabelle in Hintergrundantwort nicht gefunden.');
    return extractRowsFromTable(best.table, { visibilityMode: 'all' });
  }

  // =========================
  // Sammelbild
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

async function buildScreenshotCanvas(rows, { withBarcode } = { withBarcode: false }) {
  const scale = Math.min(2, Math.max(1, (window.devicePixelRatio || 1)));

  const pad = 12;
  const headerH = 34;
  const gap = 10;
  const lineH = 14;
  const rowPadY = 6;

  const measure = document.createElement('canvas');
  const mctx = measure.getContext('2d');
  mctx.font = '12px Arial';

  function ellipsize(ctx, text, maxWidth) {
    let t = String(text || '').trim();
    if (!t) return '';
    if (ctx.measureText(t).width <= maxWidth) return t;
    while (t.length > 1 && ctx.measureText(t + '…').width > maxWidth) {
      t = t.slice(0, -1);
    }
    return t + '…';
  }

  function wrapOneOrTwoLines(ctx, text, maxWidth) {
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
        if (lines.length >= 1) break;
      }
    }
    lines.push(line);

    if (lines.length > 2) lines.length = 2;
    if (lines.length === 2) {
      lines[1] = ellipsize(ctx, lines[1], maxWidth);
    }
    return lines;
  }

  function getAddressLines(r, ctx, width) {
    const name = String(r.name1 || '').trim();
    const street = String(r.str || '').trim();
    const plzOrt = [String(r.plz || '').trim(), String(r.ort || '').trim()].filter(Boolean).join(' ').trim();

    const out = [];
    if (name) out.push(...wrapOneOrTwoLines(ctx, name, width));
    if (street) out.push(...wrapOneOrTwoLines(ctx, street, width));
    if (plzOrt) out.push(...wrapOneOrTwoLines(ctx, plzOrt, width));

    return out.length ? out.slice(0, 6) : [''];
  }

  // Für "Kopie" (ohne Code) bleibt die alte Tabellenart weitgehend erhalten
  if (!withBarcode) {
    const col = computeAutoColumnWidths(mctx, rows);
    const textAreaW =
      col.psn + col.so + col.zc + col.plzort + col.str + col.name + col.umv + (6 * gap);

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
      rowHeights.push(Math.max(textHeight, 26));
    }

    const totalRowsH = rowHeights.reduce((a, b) => a + b, 0);
    const widthCss = pad * 2 + textAreaW;
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
        ctx.fillStyle = '#000';
        ctx.fillText(lines[0] || '', x, baseY);
        if (lines[1]) ctx.fillText(lines[1], x, baseY + lineH);
        x += w + gap;
      };

      const drawSoCell = (so, w) => {
        drawCanvasSoBadge(ctx, so, x, baseY, w);
        x += w + gap;
      };

      drawCell(r.psn, col.psn);
      drawSoCell(r.so, col.so);
      drawCell(r.zc, col.zc);
      drawCell(plzort, col.plzort);
      drawCell(r.str, col.str);
      drawCell(name, col.name);
      drawCell(r.umv, col.umv);

      ctx.strokeStyle = '#eee';
      ctx.beginPath();
      ctx.moveTo(pad, y + rowH);
      ctx.lineTo(widthCss - pad, y + rowH);
      ctx.stroke();

      yCursor += rowH;
    }

    return canvas;
  }

  // Ab hier nur für "Kopie mit Code"
  const addrColW = 240;
  const barcodePadding = 12;

  // Barcode-Bilder vorladen und natürliche Größe verwenden
  const barcodeData = [];
  let barcodeMaxW = 0;
  let barcodeMaxH = 0;

  for (const r of rows) {
    const finalCode = buildFinalBarcodeFromRow(r.plz, r.psn, r.so);
    const url = buildCode128Url(finalCode);

    try {
      const bmp = await fetchImageBitmap(url);
      const iw = bmp?.width || bmp?.naturalWidth || 0;
      const ih = bmp?.height || bmp?.naturalHeight || 0;
      barcodeData.push({ bmp, iw, ih, url, finalCode });
      barcodeMaxW = Math.max(barcodeMaxW, iw);
      barcodeMaxH = Math.max(barcodeMaxH, ih);
    } catch {
      barcodeData.push({ bmp: null, iw: 0, ih: 0, url, finalCode });
    }
  }

  barcodeMaxW = Math.max(barcodeMaxW, 420);
  barcodeMaxH = Math.max(barcodeMaxH, 70);

  const barcodeBoxW = barcodeMaxW + barcodePadding * 2;
  const barcodeBoxH = barcodeMaxH + barcodePadding * 2 + 18; // + Platz für Text unter Barcode

  const textAreaW = addrColW;
  const totalW = textAreaW + gap + barcodeBoxW;

  const rowHeights = rows.map((r) => {
    const addrLines = getAddressLines(r, mctx, addrColW);
    const addrHeight = rowPadY * 2 + (addrLines.length * lineH);
    return Math.max(addrHeight, barcodeBoxH + rowPadY * 2, 92);
  });

  const totalRowsH = rowHeights.reduce((a, b) => a + b, 0);
  const widthCss = pad * 2 + totalW;
  const heightCss = pad * 2 + headerH + totalRowsH;

  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil(widthCss * scale);
  canvas.height = Math.ceil(heightCss * scale);

  const ctx = canvas.getContext('2d');
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.imageSmoothingEnabled = false;

  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, widthCss, heightCss);

  ctx.fillStyle = '#666';
  ctx.font = '10px Arial';
  const ts = new Date().toLocaleString();
  const tw = ctx.measureText(ts).width;
  ctx.fillText(ts, widthCss - pad - tw, pad + 10);

  ctx.fillStyle = '#000';
  ctx.font = 'bold 12px Arial';
  const yHead = pad + 16;
  ctx.fillText('Adresse', pad, yHead);
  ctx.fillText('Barcode', pad + textAreaW + gap, yHead);

  ctx.strokeStyle = '#ddd';
  ctx.beginPath();
  ctx.moveTo(pad, pad + headerH - 8);
  ctx.lineTo(widthCss - pad, pad + headerH - 8);
  ctx.stroke();

  let yCursor = pad + headerH;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowH = rowHeights[i];
    const y = yCursor;

    if (i % 2 === 1) {
      ctx.fillStyle = '#f7f7f7';
      ctx.fillRect(pad, y, widthCss - pad * 2, rowH);
    }

    ctx.fillStyle = '#000';
    ctx.font = '12px Arial';

    // Adresse links
    const addrLines = getAddressLines(r, ctx, addrColW);
    const textX = pad;
    let textY = y + rowPadY + 12;

    for (const line of addrLines) {
      ctx.fillText(line || '', textX, textY);
      textY += lineH;
    }

    // Barcode rechts
    const bd = barcodeData[i];
    const boxX = pad + textAreaW + gap;
    const boxY = y + Math.floor((rowH - barcodeBoxH) / 2);

    ctx.fillStyle = '#fff';
    ctx.fillRect(boxX, boxY, barcodeBoxW, barcodeBoxH);
    ctx.strokeStyle = '#ccc';
    ctx.strokeRect(boxX, boxY, barcodeBoxW, barcodeBoxH);

    if (bd && bd.bmp && bd.iw && bd.ih) {
      const dx = Math.round(boxX + (barcodeBoxW - bd.iw) / 2);
      const dy = Math.round(boxY + barcodePadding);

      ctx.save();
      ctx.imageSmoothingEnabled = false;
      ctx.drawImage(bd.bmp, dx, dy, bd.iw, bd.ih);
      ctx.restore();

ctx.fillStyle = '#000';
ctx.font = '11px Arial';
ctx.fillText(
  String(r.psn || ''),
  boxX + 8,
  boxY + barcodePadding + bd.ih + 14
);
    } else {
      ctx.fillStyle = '#c00';
      ctx.font = '12px Arial';
      ctx.fillText('Barcode nicht ladbar', boxX + 8, boxY + 22);
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

    const btnImg  = containerEl.querySelector('#tm-copy-img');
    const btnCode = containerEl.querySelector('#tm-copy-with-code');

    if (!btnImg && !btnCode) return;

    if (btnImg  && !btnImg.dataset.baseText)  btnImg.dataset.baseText  = (btnImg.textContent  || 'Kopie');
    if (btnCode && !btnCode.dataset.baseText) btnCode.dataset.baseText = (btnCode.textContent || 'Kopie mit Code');

    if (containerEl.dataset.rowCountTimer === '1') return;
    containerEl.dataset.rowCountTimer = '1';

    let last = null;

    const tick = () => {
      const n = getVisibleRowCountSafe();
      if (n === last) return;
      last = n;

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
  // Übersicht: Hintergrundabfrage + Copy + Popup
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

    try {
      const p = new URLSearchParams(window.location.search);
      const from = (p.get('stimestamp_from') || '').trim();
      const till = (p.get('stimestamp_till') || '').trim();
      if (re.test(from) && re.test(till)) return { from, till };
    } catch {}
    return null;
  }

  function getBaseReportUrlAndParamsFromPage() {
    const doc = findReportDoc();
    const baseUrl = window.location.origin + '/cgi-bin/report_inbound_ofd.cgi';

    let form = null;
    const sel = doc.querySelector('select[name="systempartner"]');
    if (sel) form = sel.closest('form');
    if (!form) form = doc.querySelector('form') || null;

    const params = new URLSearchParams();

    if (form) {
      const fd = new FormData(form);
      for (const [k, v] of fd.entries()) {
        if (v === null || v === undefined) continue;
        const s = String(v).trim();
        if (s === '' && !/^systempartner$/i.test(k) && !/^stimestamp_(from|till)$/i.test(k)) continue;
        params.append(k, s);
      }
    }

    if (!params.has('doAction')) params.set('doAction', 'true');
    if (!params.has('reportType')) params.set('reportType', 'simple');
    if (!params.has('reportTypeSelect')) params.set('reportTypeSelect', params.get('reportType') || 'simple');
    if (!params.has('orderby')) params.set('orderby', 'PSN');
    if (!params.has('sortorder')) params.set('sortorder', 'ASC');

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

  function getRouteFilterFromDoc(doc) {
    if (!doc) return null;

    const partnerSelect = doc.querySelector('select[name="systempartner"]');
    const form = partnerSelect?.closest('form') || null;
    const candidates = Array.from((form || doc).querySelectorAll('select'));
    const routeSelect = candidates.find(sel => {
      const key = `${sel.name || ''} ${sel.id || ''}`.toLowerCase();
      return /(^|[^a-z])(route|route_nr|routenr)([^a-z]|$)/i.test(key) || key.includes('route');
    });

    if (!routeSelect || !routeSelect.name) return null;

    const options = Array.from(routeSelect.options || [])
      .map(option => ({
        value: String(option.value ?? ''),
        label: String(option.textContent || option.value || '').trim()
      }))
      .filter(option => option.label);

    if (!options.length) return null;

    return {
      paramName: routeSelect.name,
      selectedValue: String(routeSelect.value ?? options[0].value),
      options
    };
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
      borderRadius: '8px',
      width: 'min(980px, calc(100vw - 32px))',
      minWidth: '0',
      maxWidth: '98vw',
      maxHeight: '82vh',
      overflow: 'auto',
      boxShadow: '0 12px 34px rgba(0,0,0,0.38)',
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
    title.textContent = 'Übersicht – Systempartner (Anzahl angezeigter Zeilen + Direkt-Kopie + QR-Popup)';
    title.style.fontWeight = 'bold';

    const sub = doc.createElement('div');
    sub.style.color = '#666';
    sub.style.marginTop = '2px';
  sub.style.fontSize = '11px';

    const left = doc.createElement('div');
    left.appendChild(title);
    left.appendChild(sub);

    const right = doc.createElement('div');
    right.style.display = 'flex';
    right.style.alignItems = 'center';
    right.style.gap = '6px';

    const btnSettings = doc.createElement('button');
    btnSettings.type = 'button';
    btnSettings.textContent = '⚙';
    btnSettings.title = 'Einstellungen (Systempartner / Touren)';
    btnSettings.style.padding = '3px 8px';
    btnSettings.style.cursor = 'pointer';
    btnSettings.style.fontSize = '15px';
    btnSettings.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (typeof doc._tmOpenSettings === 'function') doc._tmOpenSettings();
      else alert('Einstellungen sind auf dieser Seite nicht verfügbar.');
    };

    const btnRefresh = doc.createElement('button');
    btnRefresh.type = 'button';
    btnRefresh.textContent = 'Aktualisieren';
    btnRefresh.style.padding = '3px 8px';
    btnRefresh.style.cursor = 'pointer';

    const btnClose = doc.createElement('button');
    btnClose.type = 'button';
    btnClose.textContent = 'Schließen';
    btnClose.style.padding = '3px 8px';
    btnClose.style.cursor = 'pointer';

    right.appendChild(btnSettings);
    right.appendChild(btnRefresh);
    right.appendChild(btnClose);

    head.appendChild(left);
    head.appendChild(right);

    const body = doc.createElement('div');
    body.style.padding = '10px 12px 44px 12px';

    const table = doc.createElement('table');
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';

    const thead = doc.createElement('thead');
    const thr = doc.createElement('tr');

    const th0 = doc.createElement('th');
    th0.textContent = '▦';
    th0.title = 'QR/Barcodes';
    th0.style.width = '48px';
    th0.style.borderBottom = '1px solid #ddd';
    th0.style.padding = '6px 6px';
    th0.style.textAlign = 'center';

    const th1 = doc.createElement('th');
    th1.textContent = 'Systempartner';
    th1.style.textAlign = 'left';
    th1.style.borderBottom = '1px solid #ddd';
    th1.style.padding = '6px 6px';

    function makeTh(label, width, align) {
      const th = doc.createElement('th');
      th.textContent = label;
      th.style.textAlign = align || 'center';
      th.style.borderBottom = '1px solid #ddd';
      th.style.padding = '6px 6px';
      if (width) th.style.width = width;
      return th;
    }

    const thExpress = makeTh('Express', '80px', 'center');
    thExpress.style.background = SO_GROUPS.express.bg;
    thExpress.style.color = SO_GROUPS.express.fg;
    thExpress.style.borderRadius = '4px';

    const thGefahrgut = makeTh('Gefahrgut', '90px', 'center');
    thGefahrgut.style.background = SO_GROUPS.gefahrgut.bg;
    thGefahrgut.style.color = SO_GROUPS.gefahrgut.fg;
    thGefahrgut.style.borderRadius = '4px';

    const thPrio = makeTh('Prio', '70px', 'center');
    thPrio.style.background = SO_GROUPS.prio.bg;
    thPrio.style.color = SO_GROUPS.prio.fg;
    thPrio.style.borderRadius = '4px';

    const thTerminHeute = makeTh('Termin für heute', '110px', 'center');
    thTerminHeute.style.background = '#bfe8ff';
    thTerminHeute.style.color = '#000';
    thTerminHeute.style.borderRadius = '4px';

    const th2 = doc.createElement('th');
    th2.textContent = 'Anzahl';
    th2.style.textAlign = 'right';
    th2.style.borderBottom = '1px solid #ddd';
    th2.style.padding = '6px 6px';
    th2.style.width = '120px';

    const th3 = doc.createElement('th');
    th3.textContent = 'Kopieren';
    th3.style.textAlign = 'center';
    th3.style.borderBottom = '1px solid #ddd';
    th3.style.padding = '6px 6px';
    th3.style.width = '110px';

    thr.appendChild(th0);
    thr.appendChild(th1);
    thr.appendChild(thExpress);
    thr.appendChild(thGefahrgut);
    thr.appendChild(thPrio);
    thr.appendChild(thTerminHeute);
    thr.appendChild(th2);
    thr.appendChild(th3);
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
      display: 'grid',
      gridTemplateColumns: '48px 1fr 80px 90px 70px 110px 120px 110px',
      alignItems: 'center',
      columnGap: '0',
      fontWeight: 'bold',
      zIndex: '2'
    });

    const sumQr = doc.createElement('div');

    const sumLabel = doc.createElement('div');
    sumLabel.textContent = 'Summe';

    const sumExpressCell = doc.createElement('div');
    sumExpressCell.style.textAlign = 'center';
    const sumGefahrgutCell = doc.createElement('div');
    sumGefahrgutCell.style.textAlign = 'center';
    const sumPrioCell = doc.createElement('div');
    sumPrioCell.style.textAlign = 'center';
    const sumTerminHeuteCell = doc.createElement('div');
    sumTerminHeuteCell.style.textAlign = 'center';

    const sumCell = doc.createElement('div');
    sumCell.textContent = '0';
    sumCell.style.fontVariantNumeric = 'tabular-nums';
    sumCell.style.textAlign = 'right';

    const btnCopyAll = doc.createElement('button');
    btnCopyAll.type = 'button';
    btnCopyAll.textContent = 'Kopieren';
    btnCopyAll.title = 'Alle geladenen Daten als Bild kopieren';
    btnCopyAll.style.cursor = 'pointer';
    btnCopyAll.style.justifySelf = 'center';

    sticky.appendChild(sumQr);
    sticky.appendChild(sumLabel);
    sticky.appendChild(sumExpressCell);
    sticky.appendChild(sumGefahrgutCell);
    sticky.appendChild(sumPrioCell);
    sticky.appendChild(sumTerminHeuteCell);
    sticky.appendChild(sumCell);
    sticky.appendChild(btnCopyAll);

    box.appendChild(head);
    box.appendChild(body);
    box.appendChild(sticky);
    overlay.appendChild(box);

    let closed = false;

    const close = () => {
      if (closed) return;
      closed = true;
      try { overlay.remove(); } catch {}
      try { doc.removeEventListener('keydown', escClose, true); } catch {}
    };

    const escClose = (e) => {
      if (!overlay.isConnected) return;
      if (e.key !== 'Escape') return;

      const allOverlays = Array.from(doc.querySelectorAll('div')).filter(el =>
        el.style &&
        el.style.position === 'fixed' &&
        (el.style.zIndex === '1000000' || el.style.zIndex === '1000003')
      );
      const topMost = allOverlays[allOverlays.length - 1];
      if (topMost === overlay) close();
    };

    doc.addEventListener('keydown', escClose, true);

    btnClose.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      close();
    };

    overlay.addEventListener('mousedown', (e) => {
      if (e.target === overlay) {
        e.preventDefault();
        e.stopPropagation();
        close();
      }
    }, true);

    box.addEventListener('mousedown', (e) => {
      e.stopPropagation();
    }, true);

    return { overlay, sub, tbody, sumCell, sumExpressCell, sumGefahrgutCell, sumPrioCell, sumTerminHeuteCell, btnCopyAll, btnRefresh, close, sortHeaders: { th1, th2 } };
  }
  function createTinyActionButton(doc, label, title) {
    const btn = doc.createElement('button');
    btn.type = 'button';
    btn.textContent = label;
    btn.title = title;
    Object.assign(btn.style, {
      padding: '1px 6px',
      margin: '0 2px',
      minWidth: '22px',
      height: '22px',
      lineHeight: '18px',
      cursor: 'pointer',
      fontSize: '12px'
    });
    return btn;
  }


  // =========================
  // Sortierung in Script-Tabellen
  // =========================

  function getSortableTextValue(v) {
    return String(v ?? '').replace(/\s+/g, ' ').trim();
  }

  function compareSmartValues(a, b) {
    const av = getSortableTextValue(a);
    const bv = getSortableTextValue(b);

    const an = Number(av.replace(/\./g, '').replace(',', '.'));
    const bn = Number(bv.replace(/\./g, '').replace(',', '.'));

    const aIsNum = av !== '' && !Number.isNaN(an) && /^-?\d+([.,]\d+)?$/.test(av.replace(/\./g, ''));
    const bIsNum = bv !== '' && !Number.isNaN(bn) && /^-?\d+([.,]\d+)?$/.test(bv.replace(/\./g, ''));

    if (aIsNum && bIsNum) return an - bn;

    return av.localeCompare(bv, 'de', {
      numeric: true,
      sensitivity: 'base'
    });
  }

  function applySortableHeader(doc, th, label, key, sortState, renderFn) {
    th.dataset.sortLabel = label;
    th.dataset.sortKey = key;
    th.title = 'Sortieren nach ' + label;
    th.style.cursor = 'pointer';
    th.style.userSelect = 'none';

    if (th._tmSortHandler) {
      try { th.removeEventListener('click', th._tmSortHandler); } catch {}
    }

    th._tmSortHandler = (e) => {
      e.preventDefault();
      e.stopPropagation();

      if (sortState.key === key) {
        sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
      } else {
        sortState.key = key;
        sortState.dir = 'asc';
      }

      renderFn();
    };

    th.addEventListener('click', th._tmSortHandler);
  }

  function updateSortableHeaderLabels(headers, sortState) {
    for (const h of headers) {
      const label = h.dataset.sortLabel || h.textContent || '';
      if (h.dataset.sortKey && sortState.key === h.dataset.sortKey) {
        h.textContent = label + (sortState.dir === 'asc' ? ' ▲' : ' ▼');
      } else {
        h.textContent = label;
      }
    }
  }


  function sortRowsByStreetThenOrt(rows) {
    // Standardsortierung für Systempartner-Details und Übersicht-Kopie:
    // zuerst Ort, danach Straße, danach Paketscheinnummer.
    return (Array.isArray(rows) ? rows.slice() : []).sort((a, b) => {
      const ortCmp = compareSmartValues(a && a.ort, b && b.ort);
      if (ortCmp !== 0) return ortCmp;
      const strCmp = compareSmartValues(a && a.str, b && b.str);
      if (strCmp !== 0) return strCmp;
      return compareSmartValues(a && a.psn, b && b.psn);
    });
  }


  async function fetchReportRowsForSystempartner(baseUrl, baseParams, option, paramOverrides = null) {
    const p = new URLSearchParams(baseParams.toString());
    const sp = option.value || option.label;
    p.set('systempartner', sp);

    if (paramOverrides) {
      for (const [name, value] of Object.entries(paramOverrides)) {
        p.delete(name);
        if (value !== null && value !== undefined) p.append(name, String(value));
      }
    }

    const url = baseUrl + '?' + p.toString();
    const html = await fetchHtml(url);
    let rows;
    try {
      rows = extractRowsFromHtml(html);
    } catch (error) {
      if (/Keine Datenzeilen gefunden/i.test(error?.message || '')) rows = [];
      else throw error;
    }

    return { url, html, rows };
  }

  async function copyRowsAsImage(rows, withBarcode = false) {
    const canvas = await buildScreenshotCanvas(rows, { withBarcode });
    const ok = await copyCanvas(canvas);
    if (!ok) throw new Error('Browser blockiert Clipboard-Bild.');
  }

  const SO_GROUPS = {
    express: { label: 'Express', codes: new Set(['225','228','379','350','530','811','378','155','799']), bg: '#e73545', fg: '#fff' },
    gefahrgut: { label: 'Gefahrgut', codes: new Set(['102','301']), bg: '#fff200', fg: '#000' },
    prio: { label: 'Prio', codes: new Set(['384','387']), bg: '#8e2aa8', fg: '#fff' }
  };

  function cleanSoCode(v) {
    return String(v || '').replace(/\D/g, '').trim();
  }

  function getSoGroupKey(v) {
    const code = cleanSoCode(v);
    if (!code) return '';
    for (const [key, cfg] of Object.entries(SO_GROUPS)) {
      if (cfg.codes.has(code)) return key;
    }
    return '';
  }

  function countSoGroups(rows) {
    const out = { express: 0, gefahrgut: 0, prio: 0 };
    for (const r of (Array.isArray(rows) ? rows : [])) {
      const k = getSoGroupKey(r && r.so);
      if (k && Object.prototype.hasOwnProperty.call(out, k)) out[k]++;
    }
    return out;
  }

  function filterRowsBySoGroup(rows, groupKey) {
    if (!groupKey) return Array.isArray(rows) ? rows.slice() : [];
    return (Array.isArray(rows) ? rows : []).filter(r => getSoGroupKey(r && r.so) === groupKey);
  }

  function getLocalDateParts(d = new Date()) {
    return {
      y: d.getFullYear(),
      m: d.getMonth() + 1,
      day: d.getDate()
    };
  }

  function parseUmverfuegungDate(value) {
    const text = String(value || '').trim();
    if (!text) return null;

    // Unterstützt z. B. 28.09.2026, 28.09.26, 2026-09-28 sowie Datum mit Uhrzeit/Text.
    let m = text.match(/(?:^|\D)(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})(?:\D|$)/);
    if (m) {
      let y = Number(m[3]);
      if (y < 100) y += 2000;
      return { y, m: Number(m[2]), day: Number(m[1]) };
    }

    m = text.match(/(?:^|\D)(\d{4})-(\d{1,2})-(\d{1,2})(?:\D|$)/);
    if (m) return { y: Number(m[1]), m: Number(m[2]), day: Number(m[3]) };

    return null;
  }

  function isUmverfuegungToday(value, now = new Date()) {
    const parsed = parseUmverfuegungDate(value);
    if (!parsed) return false;
    const today = getLocalDateParts(now);
    return parsed.y === today.y && parsed.m === today.m && parsed.day === today.day;
  }

  function getUmverfuegungHighlight(value) {
    const text = String(value || '').replace(/\s+/g, ' ').trim();
    if (!text) return '';

    // ASG bleibt bewusst ohne Markierung.
    if (/\bASG\b/i.test(text)) return '';

    // Jeder erkannte Termin / jedes Datum wird hellblau markiert – unabhängig vom Tag.
    if (parseUmverfuegungDate(text)) return '#bfe8ff';

    // Jede andere Umverfügung (z. B. andere Adresse / Verfügung) wird hellgrün markiert.
    return '#d9f2d9';
  }

  function countTerminHeute(rows) {
    let count = 0;
    for (const r of (Array.isArray(rows) ? rows : [])) {
      if (isUmverfuegungToday(r && r.umv)) count++;
    }
    return count;
  }

  function filterRowsTerminHeute(rows) {
    return (Array.isArray(rows) ? rows : []).filter(r => isUmverfuegungToday(r && r.umv));
  }

  function makeTerminHeuteBadge(doc, count) {
    const badge = doc.createElement('span');
    badge.textContent = String(Number(count) || 0);
    Object.assign(badge.style, {
      display: 'inline-block',
      minWidth: '30px',
      padding: '2px 7px',
      borderRadius: '5px',
      background: '#bfe8ff',
      color: '#000',
      fontWeight: 'bold',
      textAlign: 'center',
      boxSizing: 'border-box'
    });
    return badge;
  }

  function drawCanvasSoBadge(ctx, so, x, baselineY, maxWidth) {
    const code = cleanSoCode(so);
    const groupKey = getSoGroupKey(code);

    if (!groupKey || !SO_GROUPS[groupKey]) {
      ctx.fillStyle = '#000';
      ctx.fillText(String(so || ''), x, baselineY);
      return;
    }

    const cfg = SO_GROUPS[groupKey];
    const text = String(code);
    const padX = 6;
    const h = 18;
    const w = Math.min(Math.max(24, Math.ceil(ctx.measureText(text).width + padX * 2)), Math.max(24, maxWidth || 9999));
    const y = baselineY - 13;

    ctx.save();
    ctx.fillStyle = cfg.bg;

    if (typeof ctx.roundRect === 'function') {
      ctx.beginPath();
      ctx.roundRect(x, y, w, h, 4);
      ctx.fill();
    } else {
      ctx.fillRect(x, y, w, h);
    }

    ctx.fillStyle = cfg.fg;
    ctx.font = 'bold 12px Arial';
    ctx.fillText(text, x + padX, baselineY);
    ctx.restore();
  }

  function makeSoBadge(doc, text, groupKey, title) {
    const span = doc.createElement('span');
    span.textContent = String(text ?? '');
    const cfg = SO_GROUPS[groupKey] || null;
    Object.assign(span.style, {
      display: 'inline-block',
      minWidth: '20px',
      padding: '2px 6px',
      borderRadius: '4px',
      fontWeight: 'bold',
      lineHeight: '16px',
      textAlign: 'center',
      fontVariantNumeric: 'tabular-nums',
      background: cfg ? cfg.bg : 'transparent',
      color: cfg ? cfg.fg : 'inherit',
      boxShadow: cfg ? '0 1px 2px rgba(0,0,0,0.18)' : 'none'
    });
    if (title) span.title = title;
    return span;
  }

  function makeOverviewSoBadge(doc, groupKey, count) {
    const cfg = SO_GROUPS[groupKey];
    const span = makeSoBadge(doc, String(count || 0), groupKey, (cfg ? cfg.label : '') + ' anzeigen');
    Object.assign(span.style, {
      cursor: (count || 0) > 0 ? 'pointer' : 'default',
      opacity: (count || 0) > 0 ? '1' : '0.35'
    });
    return span;
  }

  function normalizeName(name) {
    return String(name || '').trim().toLowerCase();
  }

  function getConfiguredToursForPartner(label) {
    const cfg = loadConfig();
    const norm = normalizeName(label);
    const entry = cfg.find(p => normalizeName(p.name) === norm);
    return entry && Array.isArray(entry.tours) ? entry.tours.map(String).filter(Boolean) : [];
  }

function createBarcodeOverlayForDoc(doc) {
  let overlay = doc.getElementById('tm-barcode-overlay');
  if (overlay) return overlay;

  overlay = doc.createElement('div');
  overlay.id = 'tm-barcode-overlay';
  Object.assign(overlay.style, {
    position: 'fixed',
    inset: '0',
    display: 'none',
    alignItems: 'center',
    justifyContent: 'center',
    background: 'rgba(0,0,0,0.7)',
    zIndex: '1000004'
  });

  const box = doc.createElement('div');
  Object.assign(box.style, {
    background: '#fff',
    padding: '12px',
    borderRadius: '4px',
    textAlign: 'center',
    boxShadow: '0 0 20px #000',
    maxWidth: '95vw',
    maxHeight: '90vh'
  });

  const img = doc.createElement('img');
  img.id = 'tm-barcode-overlay-img';
  Object.assign(img.style, {
    maxWidth: '90vw',
    maxHeight: '70vh',
    marginBottom: '8px',
    display: 'block'
  });

  const info = doc.createElement('div');
  info.id = 'tm-barcode-overlay-info';
  Object.assign(info.style, {
    marginBottom: '8px',
    fontFamily: 'Arial, sans-serif',
    fontSize: '12px',
    color: '#333'
  });

  const btnCopy = doc.createElement('button');
  btnCopy.type = 'button';
  btnCopy.textContent = '▣  Kopieren';
  btnCopy.style.margin = '4px';

  const btnClose = doc.createElement('button');
  btnClose.type = 'button';
  btnClose.textContent = 'Schließen';
  btnClose.style.margin = '4px';

  box.appendChild(img);
  box.appendChild(info);
  box.appendChild(btnCopy);
  box.appendChild(btnClose);
  overlay.appendChild(box);
  doc.body.appendChild(overlay);

  // QR/Barcode-Popup vom darunterliegenden Fenster trennen.
  // Wichtig: Nicht schon bei mousedown schließen, sonst kann der folgende click
  // auf das darunterliegende Systempartner-Fenster fallen und dieses schließen.
  overlay.addEventListener('mousedown', (e) => {
    e.stopPropagation();
  });

  overlay.addEventListener('click', (e) => {
    e.stopPropagation();
    if (e.target === overlay) overlay.style.display = 'none';
  });

  box.addEventListener('mousedown', (e) => {
    e.stopPropagation();
  });

  box.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  btnClose.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    overlay.style.display = 'none';
  });

  doc.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && overlay.style.display !== 'none') {
      overlay.style.display = 'none';
    }
  }, true);

  btnCopy.addEventListener('click', async () => {
    if (!img.src) return;

    const old = btnCopy.textContent;
    btnCopy.disabled = true;
    btnCopy.textContent = 'Kopiere...';

    try {
      const image = await fetchImageBitmap(img.src);

      const canvas = doc.createElement('canvas');
      const iw = image?.width || image?.naturalWidth || 0;
      const ih = image?.height || image?.naturalHeight || 0;
      if (!iw || !ih) throw new Error('Barcode-Bild nicht lesbar.');

      canvas.width = iw;
      canvas.height = ih;

      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, iw, ih);
      ctx.drawImage(image, 0, 0, iw, ih);

      const ok = await copyCanvas(canvas);
      if (!ok) throw new Error('Kopieren vom Browser blockiert.');

      btnCopy.textContent = 'Kopiert ✓';
      setTimeout(() => {
        btnCopy.textContent = old;
        btnCopy.disabled = false;
      }, 1200);
    } catch (e) {
      console.error(e);
      btnCopy.textContent = old;
      btnCopy.disabled = false;
      alert('Kopieren fehlgeschlagen:\n' + (e?.message || String(e)));
    }
  });

  return overlay;
}

function openBarcodeOverlay(doc, row) {
  const overlay = createBarcodeOverlayForDoc(doc);
  const img = overlay.querySelector('#tm-barcode-overlay-img');
  const info = overlay.querySelector('#tm-barcode-overlay-info');

  const plzValue = String(row?.plz || '').trim().replace(/\D/g, '').slice(0, 5);
  const paketNr = String(row?.psn || '').trim();
  const soCode = String(row?.so || '').trim();

  if (!plzValue || !paketNr || !soCode) {
    alert('Barcode kann nicht erzeugt werden: PLZ, PSN oder SO fehlt.');
    return;
  }

  const finalBarcode = buildFinalBarcodeFromRow(plzValue, paketNr, soCode);
  const url = buildCode128Url(finalBarcode);

  img.src = url;
  info.textContent = 'PSN: ' + paketNr + ' | SO: ' + soCode + ' | PLZ: ' + plzValue;
  overlay.style.display = 'flex';
}

function createRowsPreviewOverlay(doc, partnerLabel, rows, options = {}) {
  let sourceRows = sortRowsByStreetThenOrt(Array.isArray(rows) ? rows.slice() : []);
  let displayRows = sourceRows.slice();
  const sortState = { key: '', dir: 'asc' };
  const headerCells = [];
  const selectedRows = new Set();
  const soFilterEnabled = options.soFilter === true;
  let selectedSoCodes = null; // null = alle SO, leere Menge = keine
  const soValue = row => String(row.so ?? '').trim();
  const soFilterBar = doc.createElement('div');
  const soDropdown = doc.createElement('details');
  const soSummary = doc.createElement('summary');
  const soChoices = doc.createElement('div');

  function refreshSoFilter() {
    if (!soFilterEnabled) return;
    const counts = new Map();
    sourceRows.forEach(row => counts.set(soValue(row), (counts.get(soValue(row)) || 0) + 1));
    const codes = Array.from(counts.keys()).sort((a, b) => a.localeCompare(b, 'de', { numeric: true }));
    soSummary.textContent = selectedSoCodes === null ? 'SO: Alle' :
      (selectedSoCodes.size ? 'SO: ' + Array.from(selectedSoCodes).map(code => code || '(ohne SO)').join(', ') : 'SO: Keine');
    soChoices.replaceChildren();
    const actions = doc.createElement('div');
    [['Alle', null], ['Keine', new Set()]].forEach(([label, value]) => {
      const button = doc.createElement('button');
      button.type = 'button';
      button.textContent = label;
      button.style.marginRight = '6px';
      button.addEventListener('click', () => {
        selectedSoCodes = value === null ? null : new Set();
        refreshSoFilter();
        renderRows();
      });
      actions.appendChild(button);
    });
    soChoices.appendChild(actions);
    codes.forEach(code => {
      const label = doc.createElement('label');
      Object.assign(label.style, { display: 'block', padding: '3px 0', cursor: 'pointer' });
      const checkbox = doc.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.checked = selectedSoCodes === null || selectedSoCodes.has(code);
      checkbox.addEventListener('change', () => {
        if (selectedSoCodes === null) selectedSoCodes = new Set(codes);
        if (checkbox.checked) selectedSoCodes.add(code);
        else selectedSoCodes.delete(code);
        if (codes.every(value => selectedSoCodes.has(value))) selectedSoCodes = null;
        renderRows();
        soSummary.textContent = selectedSoCodes === null ? 'SO: Alle' :
          (selectedSoCodes.size ? 'SO: ' + Array.from(selectedSoCodes).map(value => value || '(ohne SO)').join(', ') : 'SO: Keine');
      });
      label.appendChild(checkbox);
      label.appendChild(doc.createTextNode(' ' + (code || '(ohne SO)') + ' (' + counts.get(code) + ')'));
      soChoices.appendChild(label);
    });
  }

  if (soFilterEnabled) {
    Object.assign(soFilterBar.style, { padding: '6px 10px', borderBottom: '1px solid #e5e5e5', flexShrink: '0' });
    Object.assign(soDropdown.style, { width: '280px', maxWidth: '100%', border: '1px solid #aaa', borderRadius: '3px', background: '#fff' });
    Object.assign(soSummary.style, { padding: '4px 6px', cursor: 'pointer' });
    soSummary.title = 'SO-Filter öffnen; mehrere Werte per Checkbox auswählen';
    Object.assign(soChoices.style, { padding: '6px', maxHeight: '180px', overflowY: 'auto', borderTop: '1px solid #ddd' });
    soDropdown.appendChild(soSummary);
    soDropdown.appendChild(soChoices);
    soFilterBar.appendChild(soDropdown);
    refreshSoFilter();
  }

  const overlay = doc.createElement('div');
  Object.assign(overlay.style, {
    position: 'fixed',
    inset: '0',
    background: 'rgba(0,0,0,0.45)',
    zIndex: '1000003',
    display: 'flex',
    alignItems: 'flex-start',
    justifyContent: 'center',
    paddingTop: '50px'
  });

  const box = doc.createElement('div');
  Object.assign(box.style, {
    background: '#fff',
    borderRadius: '6px',
    width: 'min(1450px, calc(100vw - 16px))',
    maxWidth: 'calc(100vw - 8px)',
    maxHeight: '92vh',
    overflow: 'hidden',
    boxShadow: '0 10px 30px rgba(0,0,0,0.35)',
    fontFamily: 'Arial, sans-serif',
    fontSize: '12px',
    position: 'relative',
    display: 'flex',
    flexDirection: 'column'
  });

  const head = doc.createElement('div');
  Object.assign(head.style, {
    padding: '8px 10px',
    borderBottom: '1px solid #e5e5e5',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
    background: '#fff',
    flexWrap: 'nowrap'
  });

  const left = doc.createElement('div');
  Object.assign(left.style, {
    flex: '1 1 auto',
    minWidth: '0',
    overflow: 'hidden'
  });

  const titleRow = doc.createElement('div');
  Object.assign(titleRow.style, {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    minWidth: '0'
  });

  const title = doc.createElement('div');
  title.textContent = 'Details – ' + partnerLabel;
  title.style.fontWeight = 'bold';
  title.style.fontSize = '13px';
  title.style.lineHeight = '16px';
  title.style.whiteSpace = 'nowrap';
  title.style.overflow = 'hidden';
  title.style.textOverflow = 'ellipsis';
  title.style.minWidth = '0';

  const btnRefreshRows = doc.createElement('button');
  btnRefreshRows.type = 'button';
  btnRefreshRows.textContent = 'Aktualisieren';
  btnRefreshRows.title = 'Liste dieses Systempartners neu laden';
  Object.assign(btnRefreshRows.style, {
    padding: '1px 7px',
    height: '21px',
    lineHeight: '17px',
    fontSize: '10px',
    cursor: 'pointer',
    flex: '0 0 auto'
  });

  titleRow.appendChild(title);
  titleRow.appendChild(btnRefreshRows);

  const sub = doc.createElement('div');
  Object.assign(sub.style, {
    color: '#666',
    marginTop: '2px',
    fontSize: '11px',
    display: 'flex',
    alignItems: 'center',
    gap: '4px',
    flexWrap: 'wrap',
    maxHeight: '38px',
    overflow: 'hidden',
    lineHeight: '17px'
  });

  const subText = doc.createElement('span');
  subText.textContent = 'Zeilen: ' + displayRows.length;
  subText.style.whiteSpace = 'nowrap';
  sub.appendChild(subText);

  const routeFilter = options && options.routeFilter;
  let routeSelectForRefresh = null;
  if (routeFilter && Array.isArray(routeFilter.options) && routeFilter.options.length && typeof routeFilter.loadRows === 'function') {
    const routeLabel = doc.createElement('label');
    routeLabel.textContent = 'Route:';
    routeLabel.style.whiteSpace = 'nowrap';
    routeLabel.style.marginLeft = '6px';

    const routeSelect = doc.createElement('select');
    routeSelectForRefresh = routeSelect;
    routeSelect.title = 'Pakete dieses Systempartners nach Route filtern';
    Object.assign(routeSelect.style, {
      height: '19px',
      minWidth: '92px',
      padding: '0 18px 0 4px',
      fontSize: '10px',
      border: '1px solid #aaa',
      borderRadius: '3px',
      background: '#fff',
      color: '#000'
    });

    routeFilter.options.forEach((item) => {
      const option = doc.createElement('option');
      option.value = String(item.value ?? '');
      option.textContent = item.label;
      routeSelect.appendChild(option);
    });
    routeSelect.value = String(routeFilter.selectedValue ?? '');

    let appliedValue = routeSelect.value;
    routeSelect.addEventListener('change', async (e) => {
      e.stopPropagation();
      const requestedValue = routeSelect.value;
      routeSelect.disabled = true;
      subText.textContent = 'Route wird geladen …';

      try {
        const loadedRows = await routeFilter.loadRows(requestedValue);
        sourceRows = sortRowsByStreetThenOrt(Array.isArray(loadedRows) ? loadedRows.slice() : []);
        selectedRows.clear();
        selectedSoCodes = null;
        refreshSoFilter();
        appliedValue = requestedValue;
        renderRows();
      } catch (error) {
        console.error(error);
        routeSelect.value = appliedValue;
        renderRows();
        alert('Route konnte nicht geladen werden:\n' + (error?.message || String(error)));
      } finally {
        routeSelect.disabled = false;
      }
    });

    routeLabel.appendChild(doc.createTextNode(' '));
    routeLabel.appendChild(routeSelect);
    sub.appendChild(routeLabel);
  }

  const configuredToursForDetails = getConfiguredToursForPartner(partnerLabel);
  if (configuredToursForDetails.length) {
    const tourLabel = doc.createElement('span');
    tourLabel.textContent = 'Touren:';
    tourLabel.style.whiteSpace = 'nowrap';
    tourLabel.style.marginLeft = '6px';
    sub.appendChild(tourLabel);

    configuredToursForDetails.forEach((tour) => {
      const tourBtn = doc.createElement('button');
      tourBtn.type = 'button';
      tourBtn.textContent = String(tour);
      tourBtn.title = 'QR-Code für Tour ' + tour + ' anzeigen';
      Object.assign(tourBtn.style, {
        padding: '0 5px',
        height: '17px',
        lineHeight: '15px',
        fontSize: '10px',
        cursor: 'pointer',
        border: '1px solid #bbb',
        borderRadius: '3px',
        background: '#f7f7f7',
        color: '#000'
      });
      tourBtn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        showQrPopup(doc, tour);
      });
      sub.appendChild(tourBtn);
    });
  }

  left.appendChild(titleRow);
  left.appendChild(sub);

  btnRefreshRows.addEventListener('click', async (e) => {
    e.preventDefault();
    e.stopPropagation();

    if (!options || typeof options.refreshRows !== 'function') {
      alert('Für diese Liste ist keine Aktualisierungsfunktion verfügbar.');
      return;
    }

    const oldText = btnRefreshRows.textContent;
    btnRefreshRows.disabled = true;
    btnRefreshRows.textContent = 'Lädt …';
    subText.textContent = 'Liste wird aktualisiert …';

    try {
      const routeValue = routeSelectForRefresh ? routeSelectForRefresh.value : null;
      const loadedRows = await options.refreshRows(routeValue);
      sourceRows = sortRowsByStreetThenOrt(Array.isArray(loadedRows) ? loadedRows.slice() : []);
      selectedRows.clear();
      selectedSoCodes = null;
      refreshSoFilter();
      renderRows();
      btnRefreshRows.textContent = 'Aktualisiert ✓';
      setTimeout(() => {
        if (btnRefreshRows.isConnected) btnRefreshRows.textContent = oldText;
      }, 1000);
    } catch (error) {
      console.error(error);
      renderRows();
      btnRefreshRows.textContent = oldText;
      alert('Liste konnte nicht aktualisiert werden:\n' + (error?.message || String(error)));
    } finally {
      btnRefreshRows.disabled = false;
    }
  });

  const right = doc.createElement('div');
  right.style.display = 'flex';
  right.style.alignItems = 'center';
  right.style.gap = '4px';
  right.style.flexWrap = 'nowrap';
  right.style.marginTop = '0';
  right.style.marginLeft = '8px';

  const btnCopy = doc.createElement('button');
  btnCopy.type = 'button';
  btnCopy.textContent = 'Kopieren';
  btnCopy.title = 'Kopiert die aktuell angezeigte Sortierung als Bild';
  btnCopy.style.padding = '3px 8px';
  btnCopy.style.cursor = 'pointer';
  btnCopy.style.fontSize = '11px';

  const btnCopyCode = doc.createElement('button');
  btnCopyCode.type = 'button';
  btnCopyCode.textContent = 'Kopieren mit Code';
  btnCopyCode.title = 'Kopiert die aktuell angezeigte Sortierung inkl. Barcode je Zeile';
  btnCopyCode.style.padding = '3px 8px';
  btnCopyCode.style.cursor = 'pointer';
  btnCopyCode.style.fontSize = '11px';

  const btnCopySelected = doc.createElement('button');
  btnCopySelected.type = 'button';
  btnCopySelected.textContent = 'Auswahl kopieren';
  btnCopySelected.title = 'Kopiert nur die vorne angehakten Zeilen als Bild';
  btnCopySelected.style.padding = '3px 8px';
  btnCopySelected.style.cursor = 'pointer';
  btnCopySelected.style.fontSize = '11px';

  const btnCopySelectedCode = doc.createElement('button');
  btnCopySelectedCode.type = 'button';
  btnCopySelectedCode.textContent = 'Auswahl kopieren mit Code';
  btnCopySelectedCode.title = 'Kopiert nur die vorne angehakten Zeilen inkl. Barcode je Zeile';
  btnCopySelectedCode.style.padding = '3px 8px';
  btnCopySelectedCode.style.cursor = 'pointer';
  btnCopySelectedCode.style.fontSize = '11px';

  const btnShowSelectedCode = doc.createElement('button');
  btnShowSelectedCode.type = 'button';
  btnShowSelectedCode.textContent = 'Auswahl anzeigen mit Code';
  btnShowSelectedCode.title = 'Zeigt nur die vorne angehakten Zeilen mit Barcode in einem eigenen Fenster';
  btnShowSelectedCode.style.padding = '3px 8px';
  btnShowSelectedCode.style.cursor = 'pointer';
  btnShowSelectedCode.style.fontSize = '11px';

  const btnClose = doc.createElement('button');
  btnClose.type = 'button';
  btnClose.textContent = 'Schließen';
  btnClose.style.padding = '3px 8px';
  btnClose.style.cursor = 'pointer';
  btnClose.style.fontSize = '11px';

  right.appendChild(btnCopy);
  right.appendChild(btnCopyCode);
  right.appendChild(btnCopySelected);
  right.appendChild(btnCopySelectedCode);
  right.appendChild(btnShowSelectedCode);
  right.appendChild(btnClose);

  head.appendChild(left);
  head.appendChild(right);

  const body = doc.createElement('div');
  Object.assign(body.style, {
    overflow: 'auto',
    padding: '0',
    background: '#fff',
    flex: '1 1 auto'
  });

  const table = doc.createElement('table');
  Object.assign(table.style, {
    width: '100%',
    minWidth: '0',
    borderCollapse: 'collapse',
    tableLayout: 'auto',
    fontSize: '11px'
  });

  const thead = doc.createElement('thead');
  const hr = doc.createElement('tr');

  const headers = [
    { label: '', key: '', width: '4%', align: 'center', selectCol: true },
    { label: 'III', key: '', width: '5%', align: 'center' },
    { label: 'Paketscheinnummer', key: 'psn', width: '16%', align: 'left' },
    { label: 'SO', key: 'so', width: '5%', align: 'left' },
    { label: 'Zusatzcodes', key: 'zc', width: '9%', align: 'left' },
    { label: 'PLZ', key: 'plz', width: '6%', align: 'left' },
    { label: 'Ort', key: 'ort', width: '12%', align: 'left' },
    { label: 'Straße', key: 'str', width: '18%', align: 'left' },
    { label: 'Name', key: 'name', width: '19%', align: 'left' },
    { label: 'Umverfügung', key: 'umv', width: '6%', align: 'left' }
  ];

  let selectAllBox = null;

  headers.forEach((h) => {
    const th = doc.createElement('th');
    th.textContent = h.label;
    Object.assign(th.style, {
      position: 'sticky',
      top: '0',
      background: '#f5f5f5',
      borderBottom: '1px solid #dcdcdc',
      padding: h.selectCol ? '4px 6px' : '4px 6px',
      textAlign: h.align,
      fontWeight: 'bold',
      fontSize: '11px',
      width: h.width,
      zIndex: '1'
    });

    if (h.selectCol) {
      selectAllBox = doc.createElement('input');
      selectAllBox.type = 'checkbox';
      selectAllBox.title = 'Alle angezeigten Zeilen auswählen/abwählen';
      Object.assign(selectAllBox.style, {
        cursor: 'pointer',
        width: '13px',
        height: '13px',
        margin: '0',
        verticalAlign: 'middle',
        display: 'inline-block',
        opacity: '1',
        position: 'static',
        appearance: 'auto',
        WebkitAppearance: 'checkbox',
        MozAppearance: 'checkbox'
      });
      function toggleAllRowsFromHeader(e) {
        if (e) {
          e.preventDefault();
          e.stopPropagation();
        }

        const totalCount = displayRows.length;
        const selectedCount = getSelectedDisplayRows().length;
        const shouldSelectAll = totalCount > 0 && selectedCount < totalCount;

        if (shouldSelectAll) {
          displayRows.forEach(r => selectedRows.add(r));
        } else {
          displayRows.forEach(r => selectedRows.delete(r));
        }

        renderRows();
      }

      selectAllBox.addEventListener('click', toggleAllRowsFromHeader);
      th.addEventListener('click', toggleAllRowsFromHeader);
      th.textContent = '';
      th.appendChild(selectAllBox);
    }

    if (h.key) applySortableHeader(doc, th, h.label, h.key, sortState, renderRows);
    headerCells.push(th);
    hr.appendChild(th);
  });

  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = doc.createElement('tbody');

  function rowValue(row, key) {
    if (key === 'name') return (String(row.name1 || '') + (row.name2 ? ' ' + row.name2 : '')).trim();
    return row[key] || '';
  }

  function sortDisplayRows() {
    displayRows = sourceRows.filter(row => !soFilterEnabled || selectedSoCodes === null || selectedSoCodes.has(soValue(row)));
    if (!sortState.key) return;

    displayRows.sort((a, b) => {
      const res = compareSmartValues(rowValue(a, sortState.key), rowValue(b, sortState.key));
      return sortState.dir === 'asc' ? res : -res;
    });
  }

  function getSelectedDisplayRows() {
    return displayRows.filter(r => selectedRows.has(r));
  }

  function updateSelectionUi() {
    const selectedCount = getSelectedDisplayRows().length;
    const totalCount = displayRows.length;

    if (selectAllBox) {
      selectAllBox.checked = totalCount > 0 && selectedCount === totalCount;
      selectAllBox.indeterminate = selectedCount > 0 && selectedCount < totalCount;
    }

    btnCopySelected.textContent = selectedCount ? `Auswahl kopieren (${selectedCount})` : 'Auswahl kopieren';
    btnCopySelectedCode.textContent = selectedCount ? `Auswahl kopieren mit Code (${selectedCount})` : 'Auswahl kopieren mit Code';
    btnShowSelectedCode.textContent = selectedCount ? `Auswahl anzeigen mit Code (${selectedCount})` : 'Auswahl anzeigen mit Code';
  }

  function renderRows() {
    sortDisplayRows();
    updateSortableHeaderLabels(headerCells.filter(h => h.dataset && h.dataset.sortKey), sortState);
    tbody.innerHTML = '';
    subText.textContent = 'Zeilen: ' + displayRows.length + (soFilterEnabled ? ' von ' + sourceRows.length : '') + (sortState.key ? ' | Sortierung: ' + sortState.key + ' ' + (sortState.dir === 'asc' ? 'aufsteigend' : 'absteigend') : '');

    if (!displayRows.length) {
      const emptyTr = doc.createElement('tr');
      const emptyTd = doc.createElement('td');
      emptyTd.colSpan = headers.length;
      emptyTd.textContent = sourceRows.length ? 'Keine Pakete für die gewählten SO. Bitte den SO-Filter ändern.' : 'Keine offenen Pakete vorhanden. Die Touren-QR-Codes oben können trotzdem angeklickt werden.';
      Object.assign(emptyTd.style, {
        padding: '14px 10px',
        color: '#666',
        textAlign: 'left',
        borderBottom: '1px solid #eee'
      });
      emptyTr.appendChild(emptyTd);
      tbody.appendChild(emptyTr);
      updateSelectionUi();
      return;
    }

    displayRows.forEach((r, idx) => {
      const tr = doc.createElement('tr');
      tr.style.background = idx % 2 === 1 ? '#fafafa' : '#fff';

      const tdSelect = doc.createElement('td');
      Object.assign(tdSelect.style, {
        borderBottom: '1px solid #efefef',
        padding: '4px 6px',
        verticalAlign: 'middle',
        textAlign: 'center'
      });

      const rowCheck = doc.createElement('input');
      rowCheck.type = 'checkbox';
      rowCheck.checked = selectedRows.has(r);
      rowCheck.title = 'Diese Zeile für Auswahl-Kopie markieren';
      Object.assign(rowCheck.style, {
        cursor: 'pointer',
        width: '13px',
        height: '13px',
        margin: '0',
        verticalAlign: 'middle',
        display: 'inline-block',
        opacity: '1',
        position: 'static',
        appearance: 'auto',
        WebkitAppearance: 'checkbox',
        MozAppearance: 'checkbox'
      });
      rowCheck.addEventListener('click', (e) => e.stopPropagation());
      rowCheck.addEventListener('change', () => {
        if (rowCheck.checked) selectedRows.add(r);
        else selectedRows.delete(r);
        updateSelectionUi();
      });
      tdSelect.appendChild(rowCheck);
      tr.appendChild(tdSelect);

      const tdBarcode = doc.createElement('td');
      Object.assign(tdBarcode.style, {
        borderBottom: '1px solid #efefef',
        padding: '4px 6px',
        verticalAlign: 'middle',
        textAlign: 'center'
      });

      const barcodeLink = doc.createElement('a');
      barcodeLink.href = '#';
      barcodeLink.textContent = 'III';
      barcodeLink.title = 'Strichcode anzeigen';
      barcodeLink.style.textDecoration = 'underline';
      barcodeLink.style.cursor = 'pointer';

      barcodeLink.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        openBarcodeOverlay(doc, r);
      });

      tdBarcode.appendChild(barcodeLink);
      tr.appendChild(tdBarcode);

      const cells = [
        r.psn || '',
        r.so || '',
        r.zc || '',
        r.plz || '',
        r.ort || '',
        r.str || '',
        (String(r.name1 || '') + (r.name2 ? ' ' + r.name2 : '')).trim(),
        r.umv || ''
      ];

      cells.forEach((val, i) => {
        const td = doc.createElement('td');

        if (i === 0 && cleanPsn(val)) {
          const psnLink = doc.createElement('a');
          psnLink.href = buildDepotportalUrl(val);
          psnLink.target = '_blank';
          psnLink.rel = 'noopener';
          psnLink.textContent = val;
          psnLink.title = 'Paketscheinnummer im Depotportal öffnen';
          psnLink.style.textDecoration = 'underline';
          psnLink.style.cursor = 'pointer';
          psnLink.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            openDepotportalTracking(val);
          });
          td.appendChild(psnLink);
        } else if (i === 1) {
          const g = getSoGroupKey(val);
          if (g) td.appendChild(makeSoBadge(doc, val, g, SO_GROUPS[g].label));
          else td.textContent = val;
        } else {
          td.textContent = val;
        }

        Object.assign(td.style, {
          borderBottom: '1px solid #efefef',
          padding: '4px 6px',
          verticalAlign: 'middle',
          wordBreak: i === 0 ? 'break-all' : 'break-word',
          fontVariantNumeric: i <= 3 ? 'tabular-nums' : 'normal'
        });

        // Umverfügung hervorheben: Termin = hellblau, andere Verfügung/Adresse = hellgrün, ASG = unverändert.
        if (i === 7) {
          const umvBg = getUmverfuegungHighlight(val);
          if (umvBg) {
            td.style.background = umvBg;
            td.style.color = '#000';
            td.style.fontWeight = '600';
          }
        }

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    updateSelectionUi();
  }

  table.appendChild(tbody);
  body.appendChild(table);

  const foot = doc.createElement('div');
  Object.assign(foot.style, {
    padding: '8px 12px',
    borderTop: '1px solid #eee',
    textAlign: 'center',
    background: '#fff'
  });

  const btnCloseBottom = doc.createElement('button');
  btnCloseBottom.type = 'button';
  btnCloseBottom.textContent = 'Schließen';
  Object.assign(btnCloseBottom.style, {
    padding: '4px 24px',
    cursor: 'pointer',
    fontSize: '12px',
    border: '1px solid #cfcfcf',
    borderRadius: '4px',
    background: '#fff'
  });
  foot.appendChild(btnCloseBottom);

  box.appendChild(head);
  if (soFilterEnabled) box.appendChild(soFilterBar);
  box.appendChild(body);
  box.appendChild(foot);
  overlay.appendChild(box);

  const close = () => {
    try { overlay.remove(); } catch {}
    try { doc.removeEventListener('keydown', escClose, true); } catch {}
    try { doc.removeEventListener('click', outsideClickCapture, true); } catch {}
  };

  const escClose = (e) => {
    if (e.key === 'Escape') close();
  };

  const outsideClickCapture = (e) => {
    if (!overlay.isConnected) return;

    const barcodeOverlay = doc.getElementById('tm-barcode-overlay');
    if ((barcodeOverlay && barcodeOverlay.style.display !== 'none') || isAnyQrOverlayOpen(doc)) {
      // Solange ein QR/Barcode offen ist, darf kein Klick das Systempartner-Fenster schließen.
      return;
    }

    if (e.target && typeof e.target.closest === 'function' && (e.target.closest('#tm-barcode-overlay') || e.target.closest('.tm-qr-popup-overlay') || e.target.closest('.tm-barcode-preview-overlay'))) {
      return;
    }

    if (!box.contains(e.target)) close();
  };

  btnClose.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    close();
  });

  btnCloseBottom.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    close();
  });

  overlay.addEventListener('click', (e) => {
    const barcodeOverlay = doc.getElementById('tm-barcode-overlay');
    if ((barcodeOverlay && barcodeOverlay.style.display !== 'none') || isAnyQrOverlayOpen(doc)) {
      e.stopPropagation();
      return;
    }
    if (e.target === overlay) close();
  });

  box.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  doc.addEventListener('keydown', escClose, true);
  doc.addEventListener('click', outsideClickCapture, true);

  async function showSelectedRowsWithBarcodeWindow(rowsToShow) {
    if (!Array.isArray(rowsToShow) || !rowsToShow.length) {
      alert('Keine Zeilen ausgewählt.');
      return;
    }

    const bad = rowsToShow.find(r => !r.plz || !r.psn || !r.so);
    if (bad) {
      alert('Anzeige mit Code nicht möglich:\nFehlende Daten für Barcode.');
      return;
    }

    const previewOverlay = doc.createElement('div');
    previewOverlay.classList.add('tm-barcode-preview-overlay');
    previewOverlay.setAttribute('data-tm-barcode-preview', '1');
    Object.assign(previewOverlay.style, {
      position: 'fixed',
      inset: '0',
      background: 'rgba(0,0,0,0.45)',
      zIndex: '1000005',
      display: 'flex',
      alignItems: 'flex-start',
      justifyContent: 'center',
      paddingTop: '45px'
    });

    const previewBox = doc.createElement('div');
    Object.assign(previewBox.style, {
      background: '#fff',
      borderRadius: '6px',
      width: 'min(1100px, calc(100vw - 32px))',
      maxWidth: '98vw',
      maxHeight: '90vh',
      overflow: 'hidden',
      boxShadow: '0 10px 30px rgba(0,0,0,0.35)',
      fontFamily: 'Arial, sans-serif',
      fontSize: '12px',
      display: 'flex',
      flexDirection: 'column'
    });

    const previewHead = doc.createElement('div');
    Object.assign(previewHead.style, {
      padding: '10px 12px',
      borderBottom: '1px solid #e5e5e5',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      gap: '10px',
      background: '#fff'
    });

    const previewTitle = doc.createElement('div');
    previewTitle.textContent = 'Auswahl mit Barcode – ' + partnerLabel + ' (' + rowsToShow.length + ')';
    previewTitle.style.fontWeight = 'bold';

    const previewActions = doc.createElement('div');
    previewActions.style.display = 'flex';
    previewActions.style.gap = '6px';

    const previewCopy = doc.createElement('button');
    previewCopy.type = 'button';
    previewCopy.textContent = 'Bild kopieren';
    previewCopy.style.padding = '4px 8px';
    previewCopy.style.cursor = 'pointer';

    const previewClose = doc.createElement('button');
    previewClose.type = 'button';
    previewClose.textContent = 'Schließen';
    previewClose.style.padding = '4px 8px';
    previewClose.style.cursor = 'pointer';

    previewActions.appendChild(previewCopy);
    previewActions.appendChild(previewClose);
    previewHead.appendChild(previewTitle);
    previewHead.appendChild(previewActions);

    const previewBody = doc.createElement('div');
    Object.assign(previewBody.style, {
      overflow: 'auto',
      padding: '12px',
      background: '#f7f7f7',
      flex: '1 1 auto',
      textAlign: 'center'
    });

    const loading = doc.createElement('div');
    loading.textContent = 'Barcode-Vorschau wird aufgebaut...';
    loading.style.padding = '20px';
    previewBody.appendChild(loading);

    previewBox.appendChild(previewHead);
    previewBox.appendChild(previewBody);
    previewOverlay.appendChild(previewBox);
    doc.body.appendChild(previewOverlay);

    const closePreview = () => {
      try { previewOverlay.remove(); } catch {}
      try { doc.removeEventListener('keydown', escPreviewClose, true); } catch {}
    };

    const escPreviewClose = (e) => {
      if (e.key === 'Escape') closePreview();
    };

    // Keine Capture-Listener auf dem Vorschaufenster: Sie würden die eigenen Buttons blockieren.
    previewClose.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      closePreview();
      return false;
    };

    previewOverlay.onclick = (e) => {
      if (e.target === previewOverlay) {
        e.preventDefault();
        e.stopPropagation();
        closePreview();
        return false;
      }
    };

    previewBox.onclick = (e) => {
      e.stopPropagation();
    };

    doc.addEventListener('keydown', escPreviewClose, true);

    try {
      const canvas = await buildScreenshotCanvas(rowsToShow, { withBarcode: true });
      const img = doc.createElement('img');
      img.alt = 'Auswahl mit Barcode';
      img.src = canvas.toDataURL('image/png');
      Object.assign(img.style, {
        maxWidth: '100%',
        height: 'auto',
        background: '#fff',
        boxShadow: '0 1px 8px rgba(0,0,0,0.18)'
      });

      previewBody.innerHTML = '';
      previewBody.appendChild(img);

      previewCopy.onclick = async (e) => {
        e.preventDefault();
        e.stopPropagation();

        const old = previewCopy.textContent;
        previewCopy.disabled = true;
        previewCopy.textContent = 'Kopiere...';
        try {
          // Ausschließlich den bereits erzeugten Barcode-Canvas kopieren.
          const ok = await copyCanvas(canvas);
          if (!ok) throw new Error('Browser blockiert Clipboard-Bild.');
          previewCopy.textContent = 'Kopiert ✓';
          setTimeout(() => {
            previewCopy.textContent = old;
            previewCopy.disabled = false;
          }, 1200);
        } catch (e2) {
          console.error(e2);
          previewCopy.textContent = old;
          previewCopy.disabled = false;
          alert('Kopie fehlgeschlagen:\n' + (e2?.message || String(e2)));
        }
        return false;
      };
    } catch (e) {
      console.error(e);
      previewBody.innerHTML = '';
      const err = doc.createElement('div');
      err.textContent = 'Barcode-Vorschau fehlgeschlagen: ' + (e?.message || String(e));
      err.style.color = '#c00';
      err.style.padding = '20px';
      previewBody.appendChild(err);
    }
  }

  async function handleCopyButton(btn, rowsToCopy, withBarcode, emptyMessage, errorTitle) {
    const old = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Kopiere...';
    try {
      if (!Array.isArray(rowsToCopy) || !rowsToCopy.length) {
        throw new Error(emptyMessage || 'Keine Zeilen ausgewählt.');
      }

      if (withBarcode) {
        const bad = rowsToCopy.find(r => !r.plz || !r.psn || !r.so);
        if (bad) throw new Error('Fehlende Daten für Barcode.');
      }

      await copyRowsAsImage(rowsToCopy, withBarcode);

      btn.textContent = 'Kopiert ✓';
      setTimeout(() => {
        btn.textContent = old;
        btn.disabled = false;
        updateSelectionUi();
      }, 1200);
    } catch (e) {
      console.error(e);
      btn.textContent = old;
      btn.disabled = false;
      updateSelectionUi();
      alert(errorTitle + ':\n' + (e?.message || String(e)));
    }
  }

  btnCopy.addEventListener('click', async () => {
    await handleCopyButton(btnCopy, displayRows, false, 'Keine Zeilen vorhanden.', 'Kopie fehlgeschlagen');
  });

  btnCopyCode.addEventListener('click', async () => {
    await handleCopyButton(btnCopyCode, displayRows, true, 'Keine Zeilen vorhanden.', 'Kopie mit Code fehlgeschlagen');
  });

  btnCopySelected.addEventListener('click', async () => {
    await handleCopyButton(btnCopySelected, getSelectedDisplayRows(), false, 'Keine Zeilen ausgewählt.', 'Auswahl-Kopie fehlgeschlagen');
  });

  btnCopySelectedCode.addEventListener('click', async () => {
    await handleCopyButton(btnCopySelectedCode, getSelectedDisplayRows(), true, 'Keine Zeilen ausgewählt.', 'Auswahl-Kopie mit Code fehlgeschlagen');
  });

  btnShowSelectedCode.addEventListener('click', async () => {
    await showSelectedRowsWithBarcodeWindow(getSelectedDisplayRows());
  });

  renderRows();

  return { overlay, close };
}

  async function showRowsPreviewPopup(doc, partnerLabel, rows, options = {}) {
    const safeRows = Array.isArray(rows) ? rows : [];
    const ui = createRowsPreviewOverlay(doc, partnerLabel, safeRows, options);
    doc.body.appendChild(ui.overlay);
  }

  async function runOverview(doc) {
    const { doc: rdoc, baseUrl, params: baseParams } = getBaseReportUrlAndParamsFromPage();
    const options = getSystempartnerOptions(rdoc);
    const routeFilter = getRouteFilterFromDoc(rdoc);

    const ui = createOverviewOverlay(doc);
    doc.body.appendChild(ui.overlay);

    if (!options.length) {
      ui.sub.textContent = '(keine Systempartner gefunden)';
      return;
    }

    const from = (baseParams.get('stimestamp_from') || '').trim();
    const till = (baseParams.get('stimestamp_till') || '').trim();

    async function loadOverviewData() {
      if (!ui.overlay.isConnected) return;

      ui.tbody.innerHTML = '';
      ui.sumCell.textContent = '0';

      const rowMap = new Map();
      let overviewItems = [];
      const overviewSort = { key: 'count', dir: 'desc' };

      function getOverviewSortValue(item, key) {
        if (key === 'partner') return item.label || '';
        if (key === 'count') return item.count || 0;
        return '';
      }

      function renderOverviewRows() {
        if (!ui.overlay.isConnected) return;

        updateSortableHeaderLabels([ui.sortHeaders.th1, ui.sortHeaders.th2], overviewSort);

        const sortedItems = overviewItems
          .slice()
          .sort((a, b) => {
            const res = compareSmartValues(
              getOverviewSortValue(a, overviewSort.key),
              getOverviewSortValue(b, overviewSort.key)
            );
            return overviewSort.dir === 'asc' ? res : -res;
          });

        ui.tbody.innerHTML = '';

        if (!sortedItems.length) {
          const tr = doc.createElement('tr');
          const td = doc.createElement('td');
          td.colSpan = 8;
          td.textContent = 'Keine Systempartner gefunden.';
          td.style.padding = '8px 6px';
          td.style.color = '#666';
          tr.appendChild(td);
          ui.tbody.appendChild(tr);
          ui.sumCell.textContent = '0';
          return;
        }

        let sum = 0;
        let sumExpress = 0;
        let sumGefahrgut = 0;
        let sumPrio = 0;
        let sumTerminHeute = 0;

        for (const it of sortedItems) {
          const rowUi = rowMap.get(it.label);
          if (!rowUi) continue;

          const counts = rowUi.soCounts || { express: 0, gefahrgut: 0, prio: 0 };
          rowUi.tdExpress.innerHTML = '';
          rowUi.tdGefahrgut.innerHTML = '';
          rowUi.tdPrio.innerHTML = '';
          rowUi.tdTerminHeute.innerHTML = '';
          rowUi.tdExpress.appendChild(makeOverviewSoBadge(doc, 'express', counts.express));
          rowUi.tdGefahrgut.appendChild(makeOverviewSoBadge(doc, 'gefahrgut', counts.gefahrgut));
          rowUi.tdPrio.appendChild(makeOverviewSoBadge(doc, 'prio', counts.prio));
          rowUi.tdTerminHeute.appendChild(makeTerminHeuteBadge(doc, rowUi.terminHeuteCount || 0));

          rowUi.td2.textContent = String(it.count);
          ui.tbody.appendChild(rowUi.tr);
          sum += it.count;
          sumExpress += counts.express || 0;
          sumGefahrgut += counts.gefahrgut || 0;
          sumPrio += counts.prio || 0;
          sumTerminHeute += rowUi.terminHeuteCount || 0;
        }

        ui.sumExpressCell.innerHTML = '';
        ui.sumGefahrgutCell.innerHTML = '';
        ui.sumPrioCell.innerHTML = '';
        ui.sumTerminHeuteCell.innerHTML = '';
        ui.sumExpressCell.appendChild(makeOverviewSoBadge(doc, 'express', sumExpress));
        ui.sumGefahrgutCell.appendChild(makeOverviewSoBadge(doc, 'gefahrgut', sumGefahrgut));
        ui.sumPrioCell.appendChild(makeOverviewSoBadge(doc, 'prio', sumPrio));
        ui.sumTerminHeuteCell.appendChild(makeTerminHeuteBadge(doc, sumTerminHeute));
        ui.sumCell.textContent = String(sum);
      }

      applySortableHeader(doc, ui.sortHeaders.th1, 'Systempartner', 'partner', overviewSort, renderOverviewRows);
      applySortableHeader(doc, ui.sortHeaders.th2, 'Anzahl', 'count', overviewSort, renderOverviewRows);

      for (const o of options) {
        const tr = doc.createElement('tr');

        const td0 = doc.createElement('td');
        td0.style.padding = '4px 6px';
        td0.style.borderBottom = '1px solid #f0f0f0';
        td0.style.textAlign = 'center';
        td0.style.whiteSpace = 'nowrap';

        const btnQr = createTinyActionButton(doc, '▦', 'QR/Barcode für die hinterlegten Touren öffnen');
        td0.appendChild(btnQr);

        const td1 = doc.createElement('td');
        td1.textContent = o.label;
        td1.style.padding = '6px 6px';
        td1.style.borderBottom = '1px solid #f0f0f0';
        td1.style.cursor = 'pointer';
        td1.style.textDecoration = 'underline';

        function makeSoCountTd() {
          const td = doc.createElement('td');
          td.style.padding = '4px 6px';
          td.style.borderBottom = '1px solid #f0f0f0';
          td.style.textAlign = 'center';
          td.style.fontVariantNumeric = 'tabular-nums';
          return td;
        }

        const tdExpress = makeSoCountTd();
        const tdGefahrgut = makeSoCountTd();
        const tdPrio = makeSoCountTd();
        const tdTerminHeute = makeSoCountTd();

        const td2 = doc.createElement('td');
        td2.textContent = '…';
        td2.style.padding = '6px 6px';
        td2.style.borderBottom = '1px solid #f0f0f0';
        td2.style.textAlign = 'right';
        td2.style.fontVariantNumeric = 'tabular-nums';
        td2.style.cursor = 'pointer';
        td2.style.textDecoration = 'underline';

        const td3 = doc.createElement('td');
        td3.style.padding = '4px 6px';
        td3.style.borderBottom = '1px solid #f0f0f0';
        td3.style.textAlign = 'center';
        td3.style.whiteSpace = 'nowrap';

        const btnCopyImg = createTinyActionButton(doc, 'Kopieren', 'Tabelle dieses Systempartners im Hintergrund laden und als Bild kopieren');
        td3.appendChild(btnCopyImg);

        tr.appendChild(td0);
        tr.appendChild(td1);
        tr.appendChild(tdExpress);
        tr.appendChild(tdGefahrgut);
        tr.appendChild(tdPrio);
        tr.appendChild(tdTerminHeute);
        tr.appendChild(td2);
        tr.appendChild(td3);

        rowMap.set(o.label, {
          tr, td2, td1, td3, tdExpress, tdGefahrgut, tdPrio, tdTerminHeute, btnQr, btnCopyImg, option: o,
          rows: null, soCounts: { express: 0, gefahrgut: 0, prio: 0 }, terminHeuteCount: 0,
          configuredTours: getConfiguredToursForPartner(o.label)
        });
      }

      let runningSum = 0;
      let done = 0;

      function updateProgress(isFinished = false) {
        const base = (from && till) ? `Zeitraum: ${from} – ${till}` : 'Zeitraum: (nicht gefunden)';
        ui.sub.textContent = isFinished ? `${base}   fertig` : `${base}   ${done}/${options.length}`;
        ui.sumCell.textContent = String(runningSum);
      }

      updateProgress(false);

      async function handleCopy(entry, button) {
        const old = button.textContent;
        button.disabled = true;
        button.textContent = '…';

        try {
          let rows = entry.rows;
          if (!rows) {
            const res = await fetchReportRowsForSystempartner(baseUrl, baseParams, entry.option);
            rows = res.rows;
            entry.rows = rows;
          }

          await copyRowsAsImage(sortRowsByStreetThenOrt(rows), false);

          button.textContent = '✓';
          setTimeout(() => {
            button.textContent = old;
            button.disabled = false;
          }, 1200);
        } catch (e) {
          console.error(e);
          button.textContent = '!';
          setTimeout(() => {
            button.textContent = old;
            button.disabled = false;
          }, 1400);
          alert('Kopieren fehlgeschlagen für "' + entry.option.label + '":\n' + (e?.message || String(e)));
        }
      }

      async function handleOpenQr(entry) {
        try {
          const tours = Array.isArray(entry.configuredTours) ? entry.configuredTours.filter(Boolean) : [];
          if (!tours.length) {
            throw new Error('Für diesen Systempartner sind in den Einstellungen keine Touren hinterlegt.');
          }
          showPartnerTourSelectionPopup(doc, entry.option.label, tours);
        } catch (e) {
          console.error(e);
          alert('QR-Popup fehlgeschlagen für "' + entry.option.label + '":\n' + (e?.message || String(e)));
        }
      }

      function makeRouteDetailOptions(entry, transformRows = null) {
        const detailOptions = {
          refreshRows: async (routeValue = null) => {
            const overrides = (routeFilter && routeValue !== null && routeValue !== undefined)
              ? { [routeFilter.paramName]: routeValue }
              : null;
            const res = await fetchReportRowsForSystempartner(baseUrl, baseParams, entry.option, overrides);
            if (typeof transformRows !== 'function') entry.rows = res.rows;
            return typeof transformRows === 'function' ? transformRows(res.rows) : res.rows;
          }
        };

        if (routeFilter) {
          detailOptions.routeFilter = {
            options: routeFilter.options,
            selectedValue: routeFilter.selectedValue,
            loadRows: async (routeValue) => {
              const res = await fetchReportRowsForSystempartner(
                baseUrl,
                baseParams,
                entry.option,
                { [routeFilter.paramName]: routeValue }
              );
              if (typeof transformRows !== 'function') entry.rows = res.rows;
              return typeof transformRows === 'function' ? transformRows(res.rows) : res.rows;
            }
          };
        }

        return detailOptions;
      }

      async function handleOpenRows(entry) {
        try {
          let rows = entry.rows;
          if (!rows) {
            const res = await fetchReportRowsForSystempartner(baseUrl, baseParams, entry.option);
            rows = res.rows;
            entry.rows = rows;
          }
          await showRowsPreviewPopup(doc, entry.option.label, rows, makeRouteDetailOptions(entry));
        } catch (e) {
          console.error(e);
          alert('Detailanzeige fehlgeschlagen für "' + entry.option.label + '":\n' + (e?.message || String(e)));
        }
      }

      const results = [];
      const CONCURRENCY = 4;
      let idx = 0;

      async function worker() {
        while (idx < options.length && ui.overlay.isConnected) {
          const opt = options[idx++];
          try {
            const { rows } = await fetchReportRowsForSystempartner(baseUrl, baseParams, opt);
            const count = Array.isArray(rows) ? rows.length : 0;
            const soCounts = countSoGroups(rows);
            const terminHeuteCount = countTerminHeute(rows);

            results.push({ label: opt.label, count, rows, soCounts, terminHeuteCount });

            const rowUi = rowMap.get(opt.label);
            if (rowUi) {
              rowUi.rows = rows;
              rowUi.soCounts = soCounts;
              rowUi.terminHeuteCount = terminHeuteCount;
              rowUi.td2.textContent = String(count);
            }

            runningSum += count;
          } catch (e) {
            results.push({ label: opt.label, count: 0, err: (e?.message || String(e)), rows: [], soCounts: { express: 0, gefahrgut: 0, prio: 0 } });

            const rowUi = rowMap.get(opt.label);
            if (rowUi) {
              rowUi.rows = [];
              rowUi.soCounts = { express: 0, gefahrgut: 0, prio: 0 };
              rowUi.terminHeuteCount = 0;
              rowUi.td2.textContent = '0';
            }
          } finally {
            done++;
            updateProgress(false);
          }
        }
      }

      const workers = [];
      for (let i = 0; i < Math.min(CONCURRENCY, options.length); i++) workers.push(worker());
      await Promise.all(workers);

      if (!ui.overlay.isConnected) return;

      overviewItems = results.slice();

      for (const it of overviewItems) {
        const rowUi = rowMap.get(it.label);
        if (!rowUi) continue;

        rowUi.td2.textContent = String(it.count || 0);
        rowUi.soCounts = it.soCounts || countSoGroups(rowUi.rows || []);
        rowUi.terminHeuteCount = (it.terminHeuteCount != null) ? it.terminHeuteCount : countTerminHeute(rowUi.rows || []);

        function bindSoCell(td, groupKey) {
          td.onclick = (e) => {
            e.preventDefault();
            e.stopPropagation();
            const rows = filterRowsBySoGroup(rowUi.rows || [], groupKey);
            if (!rows.length) {
              alert('Keine Treffer für ' + SO_GROUPS[groupKey].label + ' bei "' + rowUi.option.label + '".');
              return;
            }
            showRowsPreviewPopup(
              doc,
              rowUi.option.label + ' – ' + SO_GROUPS[groupKey].label,
              rows,
              { ...makeRouteDetailOptions(rowUi, loadedRows => filterRowsBySoGroup(loadedRows, groupKey)), soFilter: groupKey === 'express' }
            );
          };
        }
        bindSoCell(rowUi.tdExpress, 'express');
        bindSoCell(rowUi.tdGefahrgut, 'gefahrgut');
        bindSoCell(rowUi.tdPrio, 'prio');

        rowUi.tdTerminHeute.onclick = (e) => {
          e.preventDefault();
          e.stopPropagation();
          const rows = filterRowsTerminHeute(rowUi.rows || []);
          if (!rows.length) {
            alert('Keine Umverfügung mit Termin für heute bei "' + rowUi.option.label + '".');
            return;
          }
          showRowsPreviewPopup(
            doc,
            rowUi.option.label + ' – Termin für heute',
            rows,
            makeRouteDetailOptions(rowUi, loadedRows => filterRowsTerminHeute(loadedRows))
          );
        };
        rowUi.tdTerminHeute.style.cursor = 'pointer';

        rowUi.btnQr.onclick = (e) => {
          e.preventDefault();
          e.stopPropagation();
          handleOpenQr(rowUi);
        };

        rowUi.td1.onclick = (e) => {
          e.preventDefault();
          e.stopPropagation();
          handleOpenRows(rowUi);
        };

        rowUi.td2.onclick = (e) => {
          e.preventDefault();
          e.stopPropagation();
          handleOpenRows(rowUi);
        };

        rowUi.btnCopyImg.onclick = (e) => {
          e.preventDefault();
          e.stopPropagation();
          handleCopy(rowUi, rowUi.btnCopyImg);
        };
      }

      ui.sumCell.style.cursor = 'pointer';
      ui.sumCell.style.textDecoration = 'underline';

      ui.sumCell.onclick = async (e) => {
        e.preventDefault();
        e.stopPropagation();

        try {
          const allRows = [];

          for (const entry of rowMap.values()) {
            if (entry && Array.isArray(entry.rows) && entry.rows.length) {
              allRows.push(...entry.rows);
            }
          }

          if (!allRows.length) {
            alert('Keine Daten vorhanden.');
            return;
          }

          await showRowsPreviewPopup(doc, 'Alle Systempartner', allRows);
        } catch (err) {
          console.error(err);
          alert('Fehler beim Öffnen der Gesamtliste:\n' + (err?.message || String(err)));
        }
      };

      function bindTotalSoCell(cell, groupKey) {
        cell.onclick = async (e) => {
          e.preventDefault();
          e.stopPropagation();
          try {
            const allRows = [];
            for (const entry of rowMap.values()) {
              if (entry && Array.isArray(entry.rows) && entry.rows.length) {
                allRows.push(...filterRowsBySoGroup(entry.rows, groupKey));
              }
            }
            if (!allRows.length) {
              alert('Keine Treffer für ' + SO_GROUPS[groupKey].label + '.');
              return;
            }
            await showRowsPreviewPopup(doc, 'Alle Systempartner – ' + SO_GROUPS[groupKey].label, allRows, { soFilter: groupKey === 'express' });
          } catch (err) {
            console.error(err);
            alert('Fehler beim Öffnen der Gesamtliste:\n' + (err?.message || String(err)));
          }
        };
      }
      bindTotalSoCell(ui.sumExpressCell, 'express');
      bindTotalSoCell(ui.sumGefahrgutCell, 'gefahrgut');
      bindTotalSoCell(ui.sumPrioCell, 'prio');

      ui.sumTerminHeuteCell.style.cursor = 'pointer';
      ui.sumTerminHeuteCell.onclick = async (e) => {
        e.preventDefault();
        e.stopPropagation();
        try {
          const allRows = [];
          for (const entry of rowMap.values()) {
            if (entry && Array.isArray(entry.rows) && entry.rows.length) {
              allRows.push(...filterRowsTerminHeute(entry.rows));
            }
          }
          if (!allRows.length) {
            alert('Keine Umverfügung mit Termin für heute vorhanden.');
            return;
          }
          await showRowsPreviewPopup(doc, 'Alle Systempartner – Termin für heute', allRows);
        } catch (err) {
          console.error(err);
          alert('Fehler beim Öffnen der Gesamtliste:\n' + (err?.message || String(err)));
        }
      };

      ui.btnCopyAll.onclick = async (e) => {
        e.preventDefault();
        e.stopPropagation();

        const old = ui.btnCopyAll.textContent;
        ui.btnCopyAll.disabled = true;
        ui.btnCopyAll.textContent = 'Kopiere...';

        try {
          const allRows = [];

          const visibleItems = overviewItems
            .filter(it => (it.count || 0) > 0)
            .slice()
            .sort((a, b) => {
              const res = compareSmartValues(
                getOverviewSortValue(a, overviewSort.key),
                getOverviewSortValue(b, overviewSort.key)
              );
              return overviewSort.dir === 'asc' ? res : -res;
            });

          for (const it of visibleItems) {
            const entry = rowMap.get(it.label);
            if (entry && Array.isArray(entry.rows) && entry.rows.length) {
              allRows.push(...entry.rows);
            }
          }

          if (!allRows.length) {
            alert('Keine Daten vorhanden.');
            return;
          }

          await copyRowsAsImage(sortRowsByStreetThenOrt(allRows), false);
          ui.btnCopyAll.textContent = 'Kopiert ✓';
          setTimeout(() => {
            ui.btnCopyAll.textContent = old;
            ui.btnCopyAll.disabled = false;
          }, 1200);
        } catch (err) {
          console.error(err);
          ui.btnCopyAll.textContent = old;
          ui.btnCopyAll.disabled = false;
          alert('Fehler beim Kopieren:\n' + (err?.message || String(err)));
        }
      };

      renderOverviewRows();
      updateProgress(true);
    }

    ui.btnRefresh.addEventListener('click', async (e) => {
      e.preventDefault();
      e.stopPropagation();
      ui.btnRefresh.disabled = true;
      const old = ui.btnRefresh.textContent;
      ui.btnRefresh.textContent = 'Aktualisiere...';
      try {
        await loadOverviewData();
      } finally {
        ui.btnRefresh.textContent = old;
        ui.btnRefresh.disabled = false;
      }
    });

    await loadOverviewData();
  }

  // =========================
  // Buttons unter Mehrfachauswahl
  // =========================

  function addCopyButtons(doc, containerEl) {
    if (!containerEl) return;

    if (!containerEl.querySelector('#tm-copy-img')) {
      const b = doc.createElement('button');
      b.id = 'tm-copy-img';
      b.type = 'button';
      b.textContent = 'Kopie';
      b.title = 'Kopiert ein Sammelbild wie Screenshot';

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

    if (!containerEl.querySelector('#tm-copy-with-code')) {
      const btn2 = doc.createElement('button');
      btn2.id = 'tm-copy-with-code';
      btn2.type = 'button';
      btn2.textContent = 'Kopie mit Code';
      btn2.title = 'Kopiert ein Sammelbild wie Screenshot inkl. Barcode je Zeile';

btn2.addEventListener('click', () => {
  const old = btn2.dataset.baseText || btn2.textContent;
  btn2.disabled = true;
  btn2.textContent = 'Erzeuge Bild...';

  (async () => {
    try {
      const rows = extractVisibleRows_ANYWHERE();

      const bad = rows.find(r => !r.plz || !r.psn || !r.so);
      if (bad) throw new Error('Fehlende Daten für Barcode.');

      const canvas = await buildScreenshotCanvas(rows, { withBarcode: true });

      // 🔴 WICHTIG: DIREKT HIER, OHNE WEITERE WAITS
      const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
      if (!blob) throw new Error('Blob fehlgeschlagen.');

      await navigator.clipboard.write([
        new ClipboardItem({ 'image/png': blob })
      ]);

      btn2.textContent = 'Kopiert ✓';
      setTimeout(() => {
        btn2.textContent = old;
        btn2.disabled = false;
      }, 1200);

    } catch (e) {
      console.error(e);
      btn2.textContent = old;
      btn2.disabled = false;

      alert(
        'Kopie mit Code fehlgeschlagen.\n\n' +
        'Grund: Browser blockiert Clipboard.\n\n' +
        'Lösung:\n- Seite neu laden\n- Direkt klicken (kein lang warten)\n- Chrome/Edge verwenden'
      );
    }
  })();
});
      containerEl.appendChild(btn2);
    }

    startRowCountAutoUpdate(doc, containerEl);
  }

  // =========================
  // QR Canvas + Popups
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

  function isAnyQrOverlayOpen(doc) {
    return !!(doc && doc.querySelector('.tm-qr-popup-overlay'));
  }

  function markQrPopupOverlay(overlay) {
    if (!overlay) return;
    overlay.classList.add('tm-qr-popup-overlay');
    overlay.setAttribute('data-tm-qr-popup', '1');
  }

  function stopQrPopupEvent(e) {
    if (!e) return;
    e.preventDefault();
    e.stopPropagation();
    if (typeof e.stopImmediatePropagation === 'function') e.stopImmediatePropagation();
  }

  function showQrPopupFromLoadedUrl(doc, tour, imgUrl) {
    const overlay = doc.createElement('div');
    markQrPopupOverlay(overlay);
    Object.assign(overlay.style, {
      position: 'fixed',
      inset: '0',
      background: 'rgba(0,0,0,0.6)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: '1000005'
    });

    const box = doc.createElement('div');
    Object.assign(box.style, {
      background: '#fff',
      padding: '12px',
      borderRadius: '4px',
      textAlign: 'center',
      minWidth: '320px',
      boxShadow: '0 0 15px rgba(0,0,0,0.5)',
      fontFamily: 'Arial, sans-serif',
      fontSize: '12px'
    });

    const title = doc.createElement('div');
    title.textContent = 'MDE Freigabe PIN (Zustellung) – Tour ' + tour;
    title.style.marginBottom = '6px';
    box.appendChild(title);

    const img = doc.createElement('img');
    img.src = imgUrl;
    Object.assign(img.style, {
      maxWidth: '360px',
      maxHeight: '360px',
      marginBottom: '6px'
    });
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
        setTimeout(() => {
          btnCopy.textContent = old;
          btnCopy.disabled = false;
        }, 1200);
        if (!ok) alert('Kopieren nicht möglich.');
      } catch {
        btnCopy.textContent = old;
        btnCopy.disabled = false;
        alert('Kopieren fehlgeschlagen.');
      }
    };
    box.appendChild(btnCopy);

    overlay.appendChild(box);

    overlay.addEventListener('mousedown', (e) => {
      e.stopPropagation();
    });

    overlay.addEventListener('click', (e) => {
      e.stopPropagation();
      if (e.target === overlay) overlay.remove();
    });

    box.addEventListener('mousedown', (e) => {
      e.stopPropagation();
    });

    box.addEventListener('click', (e) => {
      e.stopPropagation();
    });

    doc.body.appendChild(overlay);
  }

  async function showQrPopup(doc, tour) {
    try {
      const content = await fetchQrContentForTour(tour);
      const imgUrl = buildQrImageUrl(content);

      const overlay = doc.createElement('div');
      markQrPopupOverlay(overlay);
      Object.assign(overlay.style, {
        position: 'fixed',
        inset: '0',
        background: 'rgba(0,0,0,0.6)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: '1000005'
      });

      const box = doc.createElement('div');
      Object.assign(box.style, {
        background: '#fff',
        padding: '12px',
        borderRadius: '4px',
        textAlign: 'center',
        minWidth: '320px',
        boxShadow: '0 0 15px rgba(0,0,0,0.5)',
        fontFamily: 'Arial, sans-serif',
        fontSize: '12px'
      });

      const title = doc.createElement('div');
      title.textContent = 'MDE Freigabe PIN (Zustellung) – Tour ' + tour;
      title.style.marginBottom = '6px';
      box.appendChild(title);

      const img = doc.createElement('img');
      img.src = imgUrl;
      Object.assign(img.style, {
        maxWidth: '360px',
        maxHeight: '360px',
        marginBottom: '6px'
      });
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
          setTimeout(() => {
            btnCopy.textContent = old;
            btnCopy.disabled = false;
          }, 1200);
          if (!ok) alert('Kopieren nicht möglich.');
        } catch {
          btnCopy.textContent = old;
          btnCopy.disabled = false;
          alert('Kopieren fehlgeschlagen.');
        }
      };
      box.appendChild(btnCopy);

      overlay.appendChild(box);

      overlay.addEventListener('mousedown', (e) => {
        e.stopPropagation();
      });

      overlay.addEventListener('click', (e) => {
        e.stopPropagation();
        if (e.target === overlay) overlay.remove();
      });

      box.addEventListener('mousedown', (e) => {
        e.stopPropagation();
      });

      box.addEventListener('click', (e) => {
        e.stopPropagation();
      });

      doc.body.appendChild(overlay);
    } catch (e) {
      alert('Fehler beim Laden des QR-Codes für Tour ' + tour + ':\n' + (e?.message || e));
    }
  }

  function showPartnerTourSelectionPopup(doc, partnerLabel, tours) {
    const cleanTours = Array.from(new Set((Array.isArray(tours) ? tours : []).map(t => String(t || '').trim()).filter(Boolean)));
    if (!cleanTours.length) {
      alert('Für diesen Systempartner sind in den Einstellungen keine Touren hinterlegt.');
      return;
    }

    const overlay = doc.createElement('div');
    markQrPopupOverlay(overlay);
    Object.assign(overlay.style, {
      position: 'fixed', inset: '0', background: 'rgba(0,0,0,0.55)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: '1000004'
    });

    const box = doc.createElement('div');
    Object.assign(box.style, {
      background: '#fff', borderRadius: '7px', padding: '12px', minWidth: '320px',
      maxWidth: '90vw', maxHeight: '85vh', overflow: 'auto', boxShadow: '0 10px 28px rgba(0,0,0,.35)',
      fontFamily: 'Arial, sans-serif', fontSize: '12px'
    });

    const title = doc.createElement('div');
    title.textContent = partnerLabel;
    title.style.fontWeight = 'bold';
    title.style.marginBottom = '8px';
    box.appendChild(title);

    const control = doc.createElement('div');
    Object.assign(control.style, { display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '8px' });

    const multi = doc.createElement('input');
    multi.type = 'checkbox';
    multi.id = 'tm-overview-multi-' + Math.random().toString(36).slice(2);
    const label = doc.createElement('label');
    label.htmlFor = multi.id;
    label.textContent = 'Mehrfachauswahl';

    const btnShow = doc.createElement('button');
    btnShow.type = 'button';
    btnShow.textContent = 'Anzeigen';
    btnShow.style.display = 'none';
    btnShow.style.padding = '2px 7px';

    control.appendChild(multi);
    control.appendChild(label);
    control.appendChild(btnShow);
    box.appendChild(control);

    const hint = doc.createElement('div');
    hint.textContent = 'Ohne Mehrfachauswahl öffnet ein Klick die Tour sofort.';
    hint.style.color = '#666';
    hint.style.fontSize = '11px';
    hint.style.marginBottom = '7px';
    box.appendChild(hint);

    const list = doc.createElement('div');
    Object.assign(list.style, { display: 'flex', flexWrap: 'wrap', gap: '6px', marginBottom: '10px' });
    box.appendChild(list);

    const selected = new Set();
    const bubbleByTour = new Map();

    const syncBubble = (tour) => {
      const b = bubbleByTour.get(tour);
      if (!b) return;
      const on = selected.has(tour);
      b.style.background = on ? '#4a4a4a' : '#7d7d7d';
      b.style.borderColor = on ? '#000' : 'transparent';
    };

    for (const tour of cleanTours) {
      const b = doc.createElement('button');
      b.type = 'button';
      b.textContent = tour;
      Object.assign(b.style, {
        borderRadius: '14px', border: '1px solid transparent', background: '#7d7d7d', color: '#fff',
        padding: '4px 10px', cursor: 'pointer', minWidth: '54px'
      });
      bubbleByTour.set(tour, b);
      b.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (!multi.checked) {
          overlay.remove();
          showQrPopup(doc, tour);
          return;
        }
        if (selected.has(tour)) selected.delete(tour); else selected.add(tour);
        syncBubble(tour);
      };
      list.appendChild(b);
    }

    multi.onchange = () => {
      if (!multi.checked) {
        selected.clear();
        cleanTours.forEach(syncBubble);
        btnShow.style.display = 'none';
      } else {
        btnShow.style.display = '';
      }
    };

    btnShow.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();
      const chosen = Array.from(selected);
      if (!chosen.length) {
        alert('Bitte mindestens eine Tour auswählen.');
        return;
      }
      overlay.remove();
      showMultiQrPopup(doc, chosen);
    };

    const footer = doc.createElement('div');
    footer.style.textAlign = 'right';
    const close = doc.createElement('button');
    close.type = 'button';
    close.textContent = 'Schließen';
    close.onclick = () => overlay.remove();
    footer.appendChild(close);
    box.appendChild(footer);

    overlay.appendChild(box);
    overlay.onclick = (e) => { if (e.target === overlay) overlay.remove(); };
    box.onclick = (e) => e.stopPropagation();
    doc.body.appendChild(overlay);
  }

  async function showMultiQrPopup(doc, tours) {
  if (!tours || !tours.length) return;

  const overlay = doc.createElement('div');
  markQrPopupOverlay(overlay);
  Object.assign(overlay.style, {
    position: 'fixed',
    inset: '0',
    background: 'rgba(0,0,0,0.6)',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    zIndex: '1000005'
  });

  const box = doc.createElement('div');
  Object.assign(box.style, {
    background: '#fff',
    padding: '12px',
    borderRadius: '4px',
    textAlign: 'center',
    minWidth: '340px',
    maxWidth: '90vw',
    maxHeight: '90vh',
    overflow: 'auto',
    boxShadow: '0 0 15px rgba(0,0,0,0.5)',
    fontFamily: 'Arial, sans-serif',
    fontSize: '12px'
  });

  const title = doc.createElement('div');
  title.textContent = 'MDE Freigabe PIN – mehrere Touren (' + tours.join(', ') + ')';
  title.style.marginBottom = '6px';
  box.appendChild(title);

  const grid = doc.createElement('div');
  Object.assign(grid.style, {
    display: 'flex',
    flexWrap: 'wrap',
    justifyContent: 'center',
    gap: '10px'
  });
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

  overlay.addEventListener('mousedown', (e) => {
    e.stopPropagation();
  });

  overlay.addEventListener('click', (e) => {
    e.stopPropagation();
    if (e.target === overlay) overlay.remove();
  });

  box.addEventListener('mousedown', (e) => {
    e.stopPropagation();
  });

  box.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  doc.body.appendChild(overlay);

  const imgUrls = {};
  const loadedTours = [];

  for (const oneTour of tours) {
    const card = doc.createElement('div');
    Object.assign(card.style, {
      border: '1px solid #ccc',
      borderRadius: '4px',
      padding: '4px',
      textAlign: 'center',
      minWidth: '150px',
      cursor: 'pointer'
    });
    card.title = 'Einzelnen QR-Code öffnen';

    const h = doc.createElement('div');
    h.textContent = 'Tour ' + oneTour;
    h.style.marginBottom = '4px';
    card.appendChild(h);

    const status = doc.createElement('div');
    status.textContent = 'Lädt...';
    status.style.fontSize = '10px';
    status.style.color = '#666';
    card.appendChild(status);

    grid.appendChild(card);

    try {
      const content = await fetchQrContentForTour(oneTour);
      const url = buildQrImageUrl(content);
      imgUrls[oneTour] = url;
      loadedTours.push(oneTour);

      const img = doc.createElement('img');
      img.src = url;
      Object.assign(img.style, { maxWidth: '160px', maxHeight: '160px' });
      img.title = 'Tour ' + oneTour + ' einzeln öffnen';

      card.onclick = (e) => {
        e.preventDefault();
        e.stopPropagation();
        showQrPopupFromLoadedUrl(doc, oneTour, url);
      };

      card.replaceChild(img, status);
    } catch (e) {
      status.textContent = 'Fehler: ' + (e?.message || e);
      status.style.color = '#c00';
    }
  }

  btnPrint.onclick = () => {
    const w = window.open('', '_blank');
    if (!w) return;

    let html = '<html><head><title>MDE Freigabe PIN – mehrere Touren</title></head><body style="text-align:center;font-family:Arial, sans-serif;">';
    html += '<h3>MDE Freigabe PIN (Zustellung) – mehrere Touren</h3>';

    loadedTours.forEach(oneTour => {
      const url = imgUrls[oneTour];
      if (!url) return;
      html += '<div style="page-break-inside:avoid;margin-bottom:20px;">';
      html += '<div>Tour ' + oneTour + ' | Depot: ' + DEPOT + '</div>';
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

      if (!ok) throw new Error('Bildkopie vom Browser blockiert.');

      btnCopy.textContent = 'Kopiert ✓';
      setTimeout(() => {
        btnCopy.textContent = old;
        btnCopy.disabled = false;
      }, 1200);
    } catch (e) {
      console.error('Multi-QR-Kopieren fehlgeschlagen:', e);
      btnCopy.textContent = old;
      btnCopy.disabled = false;
      alert('Kopieren fehlgeschlagen:\n' + (e?.message || String(e)));
    }
  };
}

  // =========================
  // Config / Einstellungen
  // =========================


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

  function saveConfig(list) {
    try { window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); }
    catch (e) { console.error('Konfig speichern fehlgeschlagen:', e); }
  }

  // =========================
  // Panel UI
  // =========================

  function initInDocument(doc) {
    if (!doc.body || doc.getElementById('tm-sp-panel')) return;

    const css = `
#tm-sp-panel{position:fixed;top:70px;right:10px;width:86px;height:30px;z-index:9999;font-family:Arial,sans-serif;font-size:12px;overflow:visible;}
#tm-overview-btn{width:86px;height:28px;padding:0 8px;margin:0;line-height:26px;text-align:center;font-weight:bold;background:#f5f5f5;border:1px solid #999;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,0.25);cursor:pointer;}
#tm-collapse-btn,#tm-settings-btn,#tm-sp-panel h3,#tm-sp-info,#tm-sp-bubbles,#tm-multi-select,#tm-sp-panel>div:not(#tm-settings-panel){display:none !important;}
#tm-settings-panel{display:none;position:fixed;top:90px;left:50%;transform:translateX(-50%);width:min(760px,calc(100vw - 40px));max-height:78vh;overflow:auto;background:#fff;border:1px solid #bbb;border-radius:8px;padding:12px;box-shadow:0 12px 34px rgba(0,0,0,.4);z-index:1000006;box-sizing:border-box;}
#tm-settings-panel .cap{font-weight:bold;margin:0 0 6px 0;padding-right:80px;}
#tm-settings-panel table{border-collapse:collapse;width:100%;}
#tm-settings-panel th,#tm-settings-panel td{border:1px solid #ddd;padding:2px 3px;font-size:11px;vertical-align:top;}
#tm-settings-panel th{background:#eee;}
#tm-settings-panel input[type="text"],#tm-excel-import{width:100%;box-sizing:border-box;font-size:11px;}
#tm-excel-import{height:110px;resize:vertical;}
`

    const style = doc.createElement('style');
    style.textContent = css;
    doc.head.appendChild(style);

    const panel = doc.createElement('div');
    panel.id = 'tm-sp-panel';

    const btnOverview = doc.createElement('button');
    btnOverview.id = 'tm-overview-btn';
    btnOverview.type = 'button';
    btnOverview.textContent = 'Übersicht';
    btnOverview.title = 'Systempartner-Übersicht inkl. Direkt-Kopie';
    panel.appendChild(btnOverview);

    const btnSettings = doc.createElement('button');
    btnSettings.id = 'tm-settings-btn';
    btnSettings.type = 'button';
    btnSettings.textContent = '⚙';
    btnSettings.title = 'Einstellungen (Systempartner / Touren)';
    panel.appendChild(btnSettings);

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

    const settingsPanel = doc.createElement('div');
    settingsPanel.id = 'tm-settings-panel';

    // Das Einstellungsfenster wird direkt an <body> gehängt und nicht in das
    // kleine Übersicht-Panel. So liegt sein z-index wirklich über der großen
    // Übersicht und wird nicht vom Stacking-Context des Panels begrenzt.
    doc.body.appendChild(panel);
    doc.body.appendChild(settingsPanel);

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

    function renderSettingsPanel() {
      const cfg = loadConfig();
      settingsPanel.innerHTML = '';

      const cap = doc.createElement('div');
      cap.className = 'cap';
      cap.textContent = 'Systempartner / Touren – Einstellungen';
      settingsPanel.appendChild(cap);

      const table = doc.createElement('table');

      const thead = doc.createElement('thead');
      const trh = doc.createElement('tr');
      ['Systempartner', 'Tournummern', 'Aktion'].forEach(txt => {
        const th = doc.createElement('th');
        th.textContent = txt;
        trh.appendChild(th);
      });
      thead.appendChild(trh);
      table.appendChild(thead);

      const tbody = doc.createElement('tbody');

      cfg.forEach((item, index) => {
        const tr = doc.createElement('tr');

        const tdName = doc.createElement('td');
        const inpName = doc.createElement('input');
        inpName.type = 'text';
        inpName.value = item.name;
        inpName.disabled = true;
        tdName.appendChild(inpName);

        const tdTours = doc.createElement('td');
        const inpTours = doc.createElement('input');
        inpTours.type = 'text';
        inpTours.value = item.tours.join(', ');
        inpTours.disabled = true;
        tdTours.appendChild(inpTours);

        const tdAct = doc.createElement('td');
        const btnEdit = doc.createElement('button');
        btnEdit.textContent = 'Bearbeiten';
        const btnDel = doc.createElement('button');
        btnDel.textContent = 'Löschen';

        tdAct.appendChild(btnEdit);
        tdAct.appendChild(doc.createTextNode(' '));
        tdAct.appendChild(btnDel);

        let editing = false;

        btnEdit.addEventListener('click', () => {
          if (!editing) {
            editing = true;
            inpName.disabled = false;
            inpTours.disabled = false;
            btnEdit.textContent = 'Speichern';
          } else {
            const name = inpName.value.trim();
            const toursArr = inpTours.value
              .split(/[,\s]+/)
              .map(t => t.trim())
              .filter(Boolean);

            if (!name || !toursArr.length) {
              alert('Bitte Name und mindestens eine Tournummer angeben.');
              return;
            }

            cfg[index] = { name, tours: toursArr };
            saveConfig(cfg);

            editing = false;
            inpName.disabled = true;
            inpTours.disabled = true;
            btnEdit.textContent = 'Bearbeiten';

            updateBubbles();
          }
        });

        btnDel.addEventListener('click', () => {
          if (!confirm('Diesen Systempartner löschen?')) return;
          cfg.splice(index, 1);
          saveConfig(cfg);
          renderSettingsPanel();
          updateBubbles();
        });

        tr.appendChild(tdName);
        tr.appendChild(tdTours);
        tr.appendChild(tdAct);
        tbody.appendChild(tr);
      });

      const trNew = doc.createElement('tr');

      const tdNewName = doc.createElement('td');
      const inpNewName = doc.createElement('input');
      inpNewName.type = 'text';
      inpNewName.placeholder = 'Systempartner-Name';
      tdNewName.appendChild(inpNewName);

      const tdNewTours = doc.createElement('td');
      const inpNewTours = doc.createElement('input');
      inpNewTours.type = 'text';
      inpNewTours.placeholder = 'Touren, z.B. 627, 623, 638';
      tdNewTours.appendChild(inpNewTours);

      const tdNewAct = doc.createElement('td');
      const btnNewSave = doc.createElement('button');
      btnNewSave.textContent = 'Speichern';
      btnNewSave.disabled = true;
      tdNewAct.appendChild(btnNewSave);

      function checkNewRow() {
        const hasName = inpNewName.value.trim().length > 0;
        const toursArr = inpNewTours.value
          .split(/[,\s]+/)
          .map(t => t.trim())
          .filter(Boolean);
        btnNewSave.disabled = !(hasName && toursArr.length > 0);
      }

      inpNewName.addEventListener('input', checkNewRow);
      inpNewTours.addEventListener('input', checkNewRow);

      btnNewSave.addEventListener('click', () => {
        const name = inpNewName.value.trim();
        const toursArr = inpNewTours.value
          .split(/[,\s]+/)
          .map(t => t.trim())
          .filter(Boolean);

        if (!name || !toursArr.length) {
          alert('Bitte Name und mindestens eine Tournummer angeben.');
          return;
        }

        const cfgNow = loadConfig();
        cfgNow.push({ name, tours: toursArr });
        saveConfig(cfgNow);

        renderSettingsPanel();
        updateBubbles();
      });

      trNew.appendChild(tdNewName);
      trNew.appendChild(tdNewTours);
      trNew.appendChild(tdNewAct);
      tbody.appendChild(trNew);

      table.appendChild(tbody);
      settingsPanel.appendChild(table);

      const importTitle = doc.createElement('div');
      importTitle.textContent = 'Systempartner importieren / exportieren:';
      importTitle.style.marginTop = '8px';
      importTitle.style.fontWeight = 'bold';
      settingsPanel.appendChild(importTitle);

      const importHint = doc.createElement('div');
      importHint.textContent = 'Für andere Rechner am besten als JSON-Datei exportieren und dort über „Datei importieren“ auswählen. Excel-Einfügen (Spalten A:B) bleibt ebenfalls möglich.';
      importHint.style.fontSize = '10px';
      importHint.style.marginBottom = '2px';
      settingsPanel.appendChild(importHint);

      const ta = doc.createElement('textarea');
      ta.id = 'tm-excel-import';
      settingsPanel.appendChild(ta);

      const importActions = doc.createElement('div');
      Object.assign(importActions.style, {
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        gap: '8px',
        marginTop: '4px'
      });
      settingsPanel.appendChild(importActions);

      const importBtn = doc.createElement('button');
      importBtn.textContent = 'Text importieren';
      importActions.appendChild(importBtn);

      const fileImportBtn = doc.createElement('button');
      fileImportBtn.textContent = 'Datei importieren';
      fileImportBtn.title = 'Eine zuvor exportierte JSON-Datei auswählen';
      importActions.appendChild(fileImportBtn);

      const fileInput = doc.createElement('input');
      fileInput.type = 'file';
      fileInput.accept = '.json,application/json,.txt,text/plain';
      fileInput.style.display = 'none';
      settingsPanel.appendChild(fileInput);

      const exportBtn = doc.createElement('button');
      exportBtn.textContent = 'Datei exportieren';
      exportBtn.title = 'Alle Systempartner/Touren als JSON-Datei herunterladen';
      importActions.appendChild(exportBtn);

      function cleanImportedConfig(raw) {
        const list = Array.isArray(raw) ? raw : raw?.systempartners;
        if (!Array.isArray(list)) {
          throw new Error('Die Datei enthält keine gültige Systempartner-Liste.');
        }

        return list.map(item => {
          const name = String(item?.name || '').trim();
          const tours = Array.isArray(item?.tours)
            ? item.tours.map(t => String(t || '').trim()).filter(Boolean)
            : [];
          return { name, tours: Array.from(new Set(tours)) };
        }).filter(item => item.name && item.tours.length);
      }

      function mergeImportedConfig(imported) {
        const cfgNow = loadConfig();
        let addedPartners = 0;
        let addedTours = 0;

        imported.forEach(item => {
          const norm = normalizeName(item.name);
          let entry = cfgNow.find(p => normalizeName(p.name) === norm);
          if (!entry) {
            entry = { name: item.name, tours: [] };
            cfgNow.push(entry);
            addedPartners++;
          }

          item.tours.forEach(tour => {
            if (!entry.tours.includes(tour)) {
              entry.tours.push(tour);
              addedTours++;
            }
          });
        });

        saveConfig(cfgNow);
        renderSettingsPanel();
        updateBubbles();
        alert(
          'Import abgeschlossen.\n' +
          'Neue Systempartner: ' + addedPartners + '\n' +
          'Neue Touren: ' + addedTours +
          ((!addedPartners && !addedTours) ? '\nAlle Daten waren bereits vorhanden.' : '')
        );
      }

      async function saveExportFile(jsonText, suggestedName) {
        const blob = new Blob([jsonText], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = doc.createElement('a');
        a.href = url;
        a.download = suggestedName;
        doc.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }

      exportBtn.addEventListener('click', async () => {
        const cfgNow = loadConfig();

        if (!cfgNow.length) {
          alert('Es sind keine Daten zum Exportieren gespeichert.');
          return;
        }

        const payload = {
          format: 'ASEA-Systempartner',
          version: 1,
          exportedAt: new Date().toISOString(),
          depot: DEPOT,
          systempartners: cleanImportedConfig(cfgNow)
        };
        const date = new Date().toISOString().slice(0, 10);

        try {
          await saveExportFile(JSON.stringify(payload, null, 2), 'ASEA-Systempartner-' + DEPOT + '-' + date + '.json');
          alert('Export heruntergeladen. Die Datei befindet sich im Downloadordner des Browsers.');
        } catch (err) {
          if (err?.name !== 'AbortError') {
            alert('Export fehlgeschlagen:\n' + (err?.message || String(err)));
          }
        }
      });

      fileImportBtn.addEventListener('click', () => fileInput.click());

      fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) return;

        try {
          const text = await file.text();
          const raw = JSON.parse(text.replace(/^\uFEFF/, ''));
          const imported = cleanImportedConfig(raw);
          if (!imported.length) throw new Error('Die Datei enthält keine verwendbaren Einträge.');
          mergeImportedConfig(imported);
        } catch (err) {
          alert('Datei konnte nicht importiert werden:\n' + (err?.message || String(err)));
        } finally {
          fileInput.value = '';
        }
      });

      importBtn.addEventListener('click', () => {
        const text = ta.value;
        if (!text.trim()) {
          alert('Bitte erst Daten aus Excel oder einer zweispaltigen Liste einfügen.');
          return;
        }

        const lines = text.split(/\r?\n/);
        let cfgNow = loadConfig();
        let addedCount = 0;

        lines.forEach(line => {
          const trimmed = line.trim();
          if (!trimmed) return;

          const tabPos = line.indexOf('\t');
          const semiPos = line.indexOf(';');
          let splitPos = -1;

          if (tabPos >= 0 && semiPos >= 0) splitPos = Math.min(tabPos, semiPos);
          else splitPos = Math.max(tabPos, semiPos);

          if (splitPos < 0) return;

          const name = line.slice(0, splitPos).trim();
          const toursText = line.slice(splitPos + 1).trim();

          if (!name || !toursText || name.toLowerCase().startsWith('systempartner')) return;

          const tours = toursText
            .split(/[,\s]+/)
            .map(t => t.replace(/[^\dA-Za-z]/g, '').trim())
            .filter(Boolean);

          if (!tours.length) return;

          const norm = normalizeName(name);
          let entry = cfgNow.find(p => normalizeName(p.name) === norm);
          if (!entry) {
            entry = { name: name, tours: [] };
            cfgNow.push(entry);
          }

          tours.forEach(tour => {
            if (!entry.tours.includes(tour)) {
              entry.tours.push(tour);
              addedCount++;
            }
          });
        });

        if (!addedCount) {
          alert('Es konnten keine neuen Touren importiert werden (evtl. alles schon vorhanden?).');
          return;
        }

        saveConfig(cfgNow);
        alert('Import abgeschlossen. Neue Touren: ' + addedCount);
        ta.value = '';
        renderSettingsPanel();
        updateBubbles();
      });
    }

    function openSettingsPanel() {
      renderSettingsPanel();

      const closeBtn = doc.createElement('button');
      closeBtn.type = 'button';
      closeBtn.textContent = 'Schließen';
      Object.assign(closeBtn.style, {
        position: 'absolute', top: '8px', right: '8px', padding: '2px 8px', cursor: 'pointer'
      });
      closeBtn.onclick = () => {
        settingsPanel.style.display = 'none';
        settingsPanel.innerHTML = '';
      };
      settingsPanel.appendChild(closeBtn);

      // Immer über der Übersicht anzeigen.
      settingsPanel.style.zIndex = '1000010';
      settingsPanel.style.display = 'block';
    }

    doc._tmOpenSettings = openSettingsPanel;

    btnSettings.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (settingsPanel.style.display === 'block') settingsPanel.style.display = 'none';
      else openSettingsPanel();
    });

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
