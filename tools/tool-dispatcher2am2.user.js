// ==UserScript==
// @name         DPD Dispatcher – Prio / Express
// @namespace    bodo.dpd.custom
// @version      9.12.4
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @description  PRIO / EXPRESS auf einen Blick – inklusive fehlendem Eingang aus 2Use (konfiguriertes Depot und anpassbarer Zeitraum).
// @match        https://dispatcher2-de.geopost.com/*
// @match        https://2use-prod.dpdit.de/Report/*
// @match        https://2use-render-prod.dpdit.de/*
// @match        https://bipvmssrs1.dpdit.de/*
// @match        https://dpd360.dpd.de/ops/express_ticker.aspx*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_openInTab
// @connect      2use-prod.dpdit.de
// @connect      2use-render-prod.dpdit.de
// @connect      bipvmssrs1.dpdit.de
// @require      https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';

  const pageWindow = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;

  if (location.hostname === 'dpd360.dpd.de' && /\/ops\/express_ticker\.aspx/i.test(location.pathname)) {
    const clean = value => String(value || '').replace(/\s+/g, ' ').trim();
    let payload = null;
    try { payload = JSON.parse(String(GM_getValue('pmExpressTickerPackages', '') || 'null')); } catch {}
    const packageNumbers = Array.from(new Set((Array.isArray(payload?.packages) ? payload.packages : [])
      .map(value => String(value || '').replace(/\D+/g, ''))
      .filter(Boolean)));
    if (packageNumbers.length) {
      let attempts = 0;
      const timer = setInterval(() => {
        attempts++;
        const textareas = Array.from(document.querySelectorAll('textarea'));
        const label = Array.from(document.querySelectorAll('label, div, span')).find(element =>
          element.children.length === 0 && /Pakete\s+hinzuf(?:ü|u)gen/i.test(clean(element.textContent))
        );
        const area = (label && textareas.find(textarea => label.parentElement?.contains(textarea))) || textareas[0];
        if (area) {
          const value = packageNumbers.join('\n');
          const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set;
          if (setter) setter.call(area, value); else area.value = value;
          area.dispatchEvent(new Event('input', { bubbles: true }));
          area.dispatchEvent(new Event('change', { bubbles: true }));
          area.focus();
          clearInterval(timer);
          try { GM_setValue('pmExpressTickerPackages', ''); } catch {}
        } else if (attempts >= 240) {
          clearInterval(timer);
        }
      }, 250);
    }
    return;
  }

  if (location.hostname === '2use-render-prod.dpdit.de' || location.hostname === 'bipvmssrs1.dpdit.de') {
    GM_setValue('pmTwoUseRenderHeartbeat', Date.now());
    const renderRunId = Number(GM_getValue('pmTwoUseActiveRun', 0) || 0);
    const clean = value => String(value || '').replace(/\s+/g, ' ').trim();
    const activateControl = control => {
      if (!control) return false;
      const view = control.ownerDocument?.defaultView || window;
      const mouseOptions = { bubbles: true, cancelable: true, view };
      try {
        if (view.PointerEvent) control.dispatchEvent(new view.PointerEvent('pointerdown', mouseOptions));
        control.dispatchEvent(new view.MouseEvent('mousedown', mouseOptions));
        if (view.PointerEvent) control.dispatchEvent(new view.PointerEvent('pointerup', mouseOptions));
        control.dispatchEvent(new view.MouseEvent('mouseup', mouseOptions));
        control.click();
        return true;
      } catch {
        try { control.click(); return true; } catch { return false; }
      }
    };
    const leafTexts = root => Array.from(root.querySelectorAll('*'))
      .filter(el => el.children.length === 0)
      .map(el => clean(el.textContent))
      .filter(Boolean);

    const parseDelimited = (text, delimiter) => {
      const rows = [];
      let row = [], value = '', quoted = false;
      const source = String(text || '').replace(/^\uFEFF/, '');
      for (let i = 0; i < source.length; i++) {
        const ch = source[i];
        if (quoted) {
          if (ch === '"' && source[i + 1] === '"') { value += '"'; i++; }
          else if (ch === '"') quoted = false;
          else value += ch;
        } else if (ch === '"') quoted = true;
        else if (ch === delimiter) { row.push(value); value = ''; }
        else if (ch === '\n') { row.push(value.replace(/\r$/, '')); rows.push(row); row = []; value = ''; }
        else value += ch;
      }
      if (value || row.length) { row.push(value.replace(/\r$/, '')); rows.push(row); }
      return rows;
    };

    const parseExportRows = csvText => {
      let matrix = [];
      for (const delimiter of [';', ',', '\t']) {
        const candidate = parseDelimited(csvText, delimiter);
        const headerAt = candidate.findIndex(row => {
          const joined = row.map(clean).join(' | ');
          return /Produkt/i.test(joined) && /PKNR/i.test(joined) && /Scan/i.test(joined);
        });
        if (headerAt >= 0) { matrix = candidate.slice(headerAt); break; }
      }
      if (!matrix.length) return [];
      const headers = matrix[0].map(clean);
      const headerKey = value => clean(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/gi, '').toLowerCase();
      const headerKeys = headers.map(headerKey);
      const find = (...patterns) => headerKeys.findIndex(value => patterns.some(pattern => pattern.test(value)));
      const idx = {
        status: find(/^status(?:\d+)?$/i),
        lastScan: find(/letzter.*scan/i, /last.*scan/i, /scanart.*vorhandenen/i),
        scanTime: find(/scanzeit/i, /scantime/i, /datum.*vorhandenen.*scan/i, /timescan/i),
        product: find(/^produkt(?:\d+)?$/i, /^product(?:\d+)?$/i),
        pknr: find(/^pknr(?:\d+)?$/i, /paketnummer/i, /parcelnumber/i),
        vd: find(/^vd(?:\d+)?$/i, /versanddepot/i),
        route: find(/^route(?:\d+)?$/i, /^tour(?:\d+)?$/i),
        targetZip: find(/ziel.*plz/i, /target.*zip/i, /empfanger.*plz/i),
        ticker: find(/^ticker(?:\d+)?$/i, /^ticket(?:\d+)?$/i)
      };
      const get = (cells, index) => index >= 0 ? clean(cells[index]) : '';
      return matrix.slice(1).map(cells => {
        const values = cells.map(clean);
        const detectedProduct = values.find(value => /^DPD\s*(?:Priority|Prio|Express|\d{1,2}:\d{2}|Food)/i.test(value)) || '';
        const detectedPknr = values.find(value => /^\d{10,16}$/.test(value)) || '';
        const detectedScanTime = values.find(value => /^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(value)) || '';
        const productIndex = values.indexOf(detectedProduct);
        const pknrIndex = values.indexOf(detectedPknr);
        const scanIndex = values.indexOf(detectedScanTime);
        const afterPknr = pknrIndex >= 0 ? values.slice(pknrIndex + 1) : [];
        const shortNumbers = afterPknr.filter(value => /^\d{3,4}$/.test(value));
        const detectedVd = shortNumbers.find(value => /^\d{4}$/.test(value)) || '';
        const detectedRoute = shortNumbers.find(value => value !== detectedVd) || '';
        const detectedZip = afterPknr.find(value => /^\d{5}$/.test(value)) || '';
        const detectedTicker = values.find(value => /tick(?:er|et)\s*erstellen|DPD\s*360/i.test(value)) || '';
        const detectedLastScan = scanIndex > 0
          ? values.slice(0, scanIndex).reverse().find(value => value && value !== detectedProduct && !/^\d+$/.test(value)) || ''
          : '';
        return {
          __twoUse: true,
          status: get(values, idx.status), lastScan: get(values, idx.lastScan) || detectedLastScan,
          scanTime: detectedScanTime || get(values, idx.scanTime),
          product: detectedProduct || get(values, idx.product),
          pknr: detectedPknr || get(values, idx.pknr), vd: get(values, idx.vd) || detectedVd,
          route: get(values, idx.route) || detectedRoute, targetZip: get(values, idx.targetZip) || detectedZip,
          ticker: get(values, idx.ticker) || detectedTicker, pknrUrl: '', tickerUrl: ''
        };
      }).filter(row => row.product && row.pknr);
    };

    let exportMenuOpenedAt = 0;
    const revealSsrsExportMenu = docs => {
      if (exportMenuOpenedAt && Date.now() - exportMenuOpenedAt < 5000) return;
      const controls = docs.flatMap(doc => Array.from(doc.querySelectorAll(
        'button, a, input, [role="button"], [title], [aria-label]'
      )));
      const exportButton = controls.find(control => {
        const label = [control.title, control.getAttribute?.('aria-label'), control.getAttribute?.('alt'), control.value, control.textContent]
          .filter(Boolean).join(' ');
        return /export|speichern|save/i.test(label) && !/excel|csv|pdf|word/i.test(label) && !control.disabled;
      });
      if (exportButton) {
        exportMenuOpenedAt = Date.now();
        activateControl(exportButton);
      }
    };

    const directSsrsExcelUrl = docs => {
      let range = {};
      try { range = JSON.parse(String(GM_getValue('pmTwoUseLastDateRange', '') || '{}')); } catch {}
      const compact = value => String(value || '').replace(/\D+/g, '');
      const from = compact(range.from);
      const to = compact(range.to);
      if (from.length !== 8 || to.length !== 8) return '';
      const allSource = docs.map(doc => String(doc.documentElement?.outerHTML || '')).join('\n');
      const username = allSource.match(/[\w.+-]+@depot\d+\.dpd\.de/i)?.[0]
        || String(GM_getValue('pmTwoUseUsername', '') || '');
      const params = [];
      const add = (name, value) => params.push(`${encodeURIComponent(name)}=${encodeURIComponent(value)}`);
      add('rs:Command', 'Render'); add('rs:Format', 'EXCELOPENXML');
      add('paraPeriodFrom', from); add('paraPeriodTo', to);
      add('paraPeriodFromLabel', `${from.slice(6, 8)}.${from.slice(4, 6)}.${from.slice(0, 4)}`);
      add('paraPeriodToLabel', `${to.slice(6, 8)}.${to.slice(4, 6)}.${to.slice(0, 4)}`);
      add('paraInterval', 'day'); add('paraDifferenzabgleich', '1');
      ['107', '195', '295'].forEach(value => add('paraLocation', value));
      [
        'DPD 8:30', 'DPD 10:00', 'DPD 12:00', 'DPD 18:00', 'DPD Express',
        'DPD 10:00 (Mo-Sa)', 'DPD 12:00 (Mo-Sa)', 'DPD Food 12:00 (Mo-Sa)',
        'DPD Food 18:00 (Mo-Sa)', 'DPD Priority'
      ].forEach(value => add('paraExpressProdukt', value));
      add('paraSendingCustomer', ' '); add('paraMissing', 'Ja');
      if (username) add('paraUsername', username);
      return `https://bipvmssrs1.dpdit.de/ReportServer?/Reports/Operations_ExpressOnline_Differenzabgleich&${params.join('&')}`;
    };

    const discoverSsrsExportUrls = docs => {
      const found = [];
      for (const doc of docs) {
        const source = String(doc.documentElement?.outerHTML || '')
          .replace(/&amp;/gi, '&').replace(/\\u0026/gi, '&').replace(/\\x26/gi, '&');
        const hits = source.match(/(?:https?:\/\/[^"'<>\s]+)?(?:\/[^"'<>\s]*)?Reserved\.ReportViewerWebControl\.axd\?[^"'<>\s]+/gi) || [];
        for (let raw of hits) {
          raw = raw.replace(/&quot;.*$/i, '').replace(/[),;]+$/, '');
          try {
            const url = new URL(raw, doc.location?.href || location.href);
            if (!url.searchParams.get('ReportSession') || !url.searchParams.get('ControlID')) continue;
            url.searchParams.set('OpType', 'Export');
            url.searchParams.set('Format', 'EXCELOPENXML');
            url.searchParams.set('ContentDisposition', 'AlwaysAttachment');
            found.push(url.href);
          } catch {}
        }
      }
      const direct = directSsrsExcelUrl(docs);
      // Prefer a fresh, parameterized full export. Viewer-session URLs can
      // point at an old/paged render and occasionally never answer.
      if (direct) found.unshift(direct);
      return Array.from(new Set(found));
    };

    let exportInFlight = false;
    let exportSucceeded = false;
    let exportExhausted = false;
    const exportFailedUrls = new Set();
    const requestExportFile = url => new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: 'GET', url, anonymous: false, withCredentials: true,
        timeout: 120000, responseType: 'arraybuffer',
        headers: { Accept: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*' },
        onload: response => {
          if (response.status < 200 || response.status >= 300) {
            reject(new Error(`HTTP ${response.status}`));
            return;
          }
          resolve({
            data: response.response,
            contentType: String(response.responseHeaders || '').match(/content-type:\s*([^\r\n]+)/i)?.[1] || ''
          });
        },
        onerror: () => reject(new Error('Netzwerkfehler beim SSRS-Export')),
        ontimeout: () => reject(new Error('Zeitüberschreitung beim SSRS-Export'))
      });
    });
    const parseExcelExportRows = arrayBuffer => {
      if (typeof XLSX === 'undefined') throw new Error('Excel-Lesemodul wurde nicht geladen');
      if (!(arrayBuffer instanceof ArrayBuffer) || !arrayBuffer.byteLength) throw new Error('Excel-Export war leer');
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: false, raw: false });
      let best = [];
      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        if (!sheet) continue;
        const csvText = XLSX.utils.sheet_to_csv(sheet, { FS: ';', RS: '\n', blankrows: false });
        const rows = parseExportRows(csvText);
        if (rows.length > best.length) best = rows;
      }
      return best;
    };
    const tryBackgroundExcelExport = async (docs, expectedCount) => {
      const exportUrl = discoverSsrsExportUrls(docs).find(url => !exportFailedUrls.has(url));
      if (!exportUrl) { exportExhausted = true; return false; }
      if (exportInFlight) return false;
      exportInFlight = true;
      try {
        GM_setValue('pmTwoUseProgress', JSON.stringify({
          runId: renderRunId, found: 0, expected: expectedCount, at: Date.now(),
          source: 'Excel', action: 'download'
        }));
        const response = await requestExportFile(exportUrl);
        try {
          GM_setValue('pmTwoUseExportDebug', JSON.stringify({
            at: Date.now(), url: exportUrl, contentType: response.contentType,
            bytes: Number(response.data?.byteLength || 0)
          }));
        } catch {}
        const rows = parseExcelExportRows(response.data);
        if (!rows.length) throw new Error('Excel enthielt keine Paketzeilen');
        const isSessionExport = /Reserved\.ReportViewerWebControl\.axd/i.test(exportUrl);
        if (isSessionExport && expectedCount != null && rows.length !== expectedCount) {
          throw new Error(`Excel unvollst\u00e4ndig: ${rows.length} von ${expectedCount}`);
        }
        const authoritativeCount = rows.length;
        GM_setValue('pmTwoUseRows', JSON.stringify(rows));
        GM_setValue('pmTwoUseRowsAt', Date.now());
        GM_setValue('pmTwoUseRowsRun', renderRunId);
        GM_setValue('pmTwoUseExpectedCount', authoritativeCount);
        GM_setValue('pmTwoUseImportSource', 'Excel-Hintergrundexport');
        GM_setValue('pmTwoUseRenderError', '');
        GM_setValue('pmTwoUseProgress', JSON.stringify({ runId: renderRunId, found: rows.length, expected: authoritativeCount, at: Date.now(), source: 'Excel' }));
        exportSucceeded = true;
        return true;
      } catch (error) {
        exportFailedUrls.add(exportUrl);
        console.warn('2Use-Hintergrundexport nicht verf\u00fcgbar; Seitenerfassung bleibt aktiv:', error);
        return false;
      } finally {
        exportInFlight = false;
      }
    };

    const parseVisualRows = doc => {
      const products = Array.from(doc.querySelectorAll('*')).filter(el =>
        el.children.length === 0 && /^DPD\s+(?:Priority|Express|\d{1,2}:\d{2}|Food)/i.test(clean(el.textContent))
      );
      const found = new Map();

      for (const productEl of products) {
        let container = productEl;
        let tokens = [];
        for (let depth = 0; depth < 10 && container; depth++, container = container.parentElement) {
          const candidate = leafTexts(container);
          const hasTime = candidate.some(value => /^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(value));
          const hasPknr = candidate.some(value => /^\d{10,16}$/.test(value));
          if (hasTime && hasPknr && candidate.length >= 5 && candidate.length <= 30) {
            tokens = candidate;
            break;
          }
        }
        if (!tokens.length) continue;

        const productIndex = tokens.findIndex(value => /^DPD\s+(?:Priority|Express|\d{1,2}:\d{2}|Food)/i.test(value));
        const scanIndex = tokens.findIndex(value => /^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(value));
        const pknrIndex = tokens.findIndex((value, index) => index > productIndex && /^\d{10,16}$/.test(value));
        const zipIndex = (() => {
          for (let i = tokens.length - 1; i > pknrIndex; i--) if (/^\d{5}$/.test(tokens[i])) return i;
          return -1;
        })();
        const routeIndex = (() => {
          for (let i = zipIndex - 1; i > pknrIndex; i--) if (/^\d{4}$/.test(tokens[i])) return i;
          return -1;
        })();
        const vdIndex = (() => {
          for (let i = routeIndex - 1; i > pknrIndex; i--) if (/^\d{4}$/.test(tokens[i])) return i;
          return -1;
        })();
        if (productIndex < 0 || scanIndex < 0 || pknrIndex < 0) continue;

        const row = {
          __twoUse: true,
          status: '',
          lastScan: scanIndex > 0 ? tokens[scanIndex - 1] : '',
          scanTime: tokens[scanIndex],
          product: tokens[productIndex],
          pknr: tokens[pknrIndex],
          vd: vdIndex >= 0 ? tokens[vdIndex] : '',
          route: routeIndex >= 0 ? tokens[routeIndex] : '',
          targetZip: zipIndex >= 0 ? tokens[zipIndex] : '',
          ticker: zipIndex >= 0 ? tokens.slice(zipIndex + 1).join(' ') : '',
          pknrUrl: '',
          tickerUrl: ''
        };
        found.set([row.pknr, row.scanTime, row.product].join('|'), row);
      }
      return Array.from(found.values());
    };

    const parseRowsByScreenLine = doc => {
      const leaves = Array.from(doc.querySelectorAll('*')).map(element => {
        if (element.children.length) return null;
        const text = clean(element.textContent);
        if (!text) return null;
        const rect = element.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) return null;
        return { element, text, rect, centerY: (rect.top + rect.bottom) / 2 };
      }).filter(Boolean);
      const productLeaves = leaves.filter(item => /^DPD\s*(?:Priority|Prio|Express|\d{1,2}:\d{2}|Food)/i.test(item.text));
      const found = new Map();
      productLeaves.forEach(productItem => {
        const line = leaves.filter(item => Math.abs(item.centerY - productItem.centerY) <= 6)
          .sort((a, b) => a.rect.left - b.rect.left);
        const tokens = line.map(item => item.text);
        const productIndex = tokens.findIndex(value => /^DPD\s*(?:Priority|Prio|Express|\d{1,2}:\d{2}|Food)/i.test(value));
        const pknrIndex = tokens.findIndex(value => /^\d{10,16}$/.test(value));
        const scanIndex = tokens.findIndex(value => /^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(value));
        if (productIndex < 0 || pknrIndex < 0 || scanIndex < 0) return;
        const afterPknr = tokens.slice(pknrIndex + 1);
        const shortNumbers = afterPknr.filter(value => /^\d{3,4}$/.test(value));
        const vd = shortNumbers.find(value => /^\d{4}$/.test(value)) || '';
        const route = shortNumbers.find(value => value !== vd) || '';
        const targetZip = afterPknr.find(value => /^\d{5}$/.test(value)) || '';
        const ticker = tokens.find(value => /tick(?:er|et)\s*erstellen|DPD\s*360/i.test(value)) || '';
        const lastScan = tokens.slice(0, scanIndex).reverse().find(value =>
          value && value !== tokens[productIndex] && !/^\d+$/.test(value)
        ) || '';
        const row = {
          __twoUse: true, status: '', lastScan, scanTime: tokens[scanIndex],
          product: tokens[productIndex], pknr: tokens[pknrIndex], vd, route,
          targetZip, ticker, pknrUrl: '', tickerUrl: ''
        };
        found.set([row.pknr, row.scanTime, row.product].join('|'), row);
      });
      return Array.from(found.values());
    };

    const parseRenderedReport = () => {
      const docs = [document];
      for (const frame of Array.from(document.querySelectorAll('iframe'))) {
        try { if (frame.contentDocument) docs.push(frame.contentDocument); } catch {}
      }
      const tables = docs.flatMap(doc => Array.from(doc.querySelectorAll('table')));
      for (const table of tables) {
        const trs = Array.from(table.querySelectorAll('tr'));
        const headerAt = trs.findIndex(tr => /Produkt/i.test(clean(tr.textContent)) && /PKNR/i.test(clean(tr.textContent)) && /Scan/i.test(clean(tr.textContent)));
        if (headerAt < 0) continue;
        const headers = Array.from(trs[headerAt].children).map(cell => clean(cell.textContent));
        const find = re => headers.findIndex(value => re.test(value));
        const idx = {
          status: find(/^Status$/i), lastScan: find(/letzter\s*Scan/i), scanTime: find(/Scanzeitpunkt/i),
          product: find(/^Produkt$/i), pknr: find(/^PKNR$/i), vd: find(/^VD$/i), route: find(/^Route$/i),
          targetZip: find(/Ziel\s*PLZ/i), ticker: find(/^Ticker$/i)
        };
        const get = (cells, index) => index >= 0 ? (cells[index] || '') : '';
        const tableRows = trs.slice(headerAt + 1).map(tr => {
          const cellElements = Array.from(tr.children);
          const cells = cellElements.map(cell => clean(cell.textContent));
          const detectedProduct = cells.find(value => /^DPD\s*(?:Priority|Prio|Express|\d{1,2}:\d{2}|Food)/i.test(value)) || '';
          const detectedPknr = cells.find(value => /^\d{10,16}$/.test(value)) || '';
          const detectedScanTime = cells.find(value => /^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(value)) || '';
          const pknrIndex = cells.indexOf(detectedPknr);
          const scanIndex = cells.indexOf(detectedScanTime);
          const afterPknr = pknrIndex >= 0 ? cells.slice(pknrIndex + 1) : [];
          const shortNumbers = afterPknr.filter(value => /^\d{3,4}$/.test(value));
          const detectedVd = shortNumbers.find(value => /^\d{4}$/.test(value)) || '';
          const detectedRoute = shortNumbers.find(value => value !== detectedVd) || '';
          const detectedZip = afterPknr.find(value => /^\d{5}$/.test(value)) || '';
          const detectedTicker = cells.find(value => /tick(?:er|et)\s*erstellen|DPD\s*360/i.test(value)) || '';
          const detectedLastScan = scanIndex > 0
            ? cells.slice(0, scanIndex).reverse().find(value => value && value !== detectedProduct && !/^\d+$/.test(value)) || ''
            : '';
          return {
            __twoUse: true,
            status: get(cells, idx.status), lastScan: get(cells, idx.lastScan) || detectedLastScan, scanTime: detectedScanTime || get(cells, idx.scanTime),
            product: detectedProduct || get(cells, idx.product), pknr: detectedPknr || get(cells, idx.pknr), vd: get(cells, idx.vd) || detectedVd,
            route: get(cells, idx.route) || detectedRoute, targetZip: get(cells, idx.targetZip) || detectedZip,
            ticker: get(cells, idx.ticker) || detectedTicker,
            pknrUrl: '', tickerUrl: ''
          };
        }).filter(row => row.pknr && row.product);
        if (tableRows.length) return tableRows;
      }
      // SSRS does not always repeat the column header on later report pages.
      // Parse those rows by their unmistakable values instead of requiring a
      // header containing Produkt / PKNR / Scanzeitpunkt.
      for (const table of tables) {
        const headerlessRows = Array.from(table.querySelectorAll('tr')).map(tr => {
          const cells = Array.from(tr.children).map(cell => clean(cell.textContent));
          const productIndex = cells.findIndex(value => /^DPD\s*(?:Priority|Prio|Express|\d{1,2}:\d{2}|Food)/i.test(value));
          const pknrIndex = cells.findIndex(value => /^\d{10,16}$/.test(value));
          const scanIndex = cells.findIndex(value => /^\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2}:\d{2}$/.test(value));
          if (productIndex < 0 || pknrIndex < 0 || scanIndex < 0) return null;
          const afterPknr = cells.slice(pknrIndex + 1);
          const shortNumbers = afterPknr.filter(value => /^\d{3,4}$/.test(value));
          const vd = shortNumbers.find(value => /^\d{4}$/.test(value)) || '';
          const route = shortNumbers.find(value => value !== vd) || '';
          const targetZip = afterPknr.find(value => /^\d{5}$/.test(value)) || '';
          const ticker = cells.find(value => /tick(?:er|et)\s*erstellen|DPD\s*360/i.test(value)) || '';
          const lastScan = cells.slice(0, scanIndex).reverse().find(value =>
            value && value !== cells[productIndex] && !/^\d+$/.test(value)
          ) || '';
          return {
            __twoUse: true, status: '', lastScan, scanTime: cells[scanIndex],
            product: cells[productIndex], pknr: cells[pknrIndex], vd, route,
            targetZip, ticker, pknrUrl: '', tickerUrl: ''
          };
        }).filter(Boolean);
        if (headerlessRows.length) return headerlessRows;
      }
      for (const doc of docs) {
        const screenRows = parseRowsByScreenLine(doc);
        if (screenRows.length) return screenRows;
        const visualRows = parseVisualRows(doc);
        if (visualRows.length) return visualRows;
      }
      return null;
    };

    const collected = new Map();
    let persistedPages = { pages: {} };
    try {
      const savedPages = JSON.parse(String(GM_getValue('pmTwoUsePageRows', '') || '{}'));
      if (Number(savedPages?.runAt || 0) === renderRunId && savedPages.pages && typeof savedPages.pages === 'object') persistedPages = savedPages;
    } catch {}
    const rebuildCollectedFromPages = () => {
      collected.clear();
      Object.keys(persistedPages.pages || {}).sort((a, b) => Number(a) - Number(b)).forEach(page => {
        const pageRows = Array.isArray(persistedPages.pages[page]) ? persistedPages.pages[page] : [];
        pageRows.forEach((row, index) => collected.set(`page-${page}|row-${index}`, row));
      });
    };
    rebuildCollectedFromPages();
    let attempts = 0;
    let lastClickedPage = 0;
    let lastPageClickAt = 0;
    let lastInputPage = 0;
    let stableWithoutPaging = 0;
    let lastVirtualScrollAt = 0;
    let virtualScrollPass = 0;
    let observedPage = 0;
    let observedPageSince = 0;
    let observedRowsSignature = '';
    let stablePageSamples = 0;
    const timer = setInterval(() => {
      attempts++;
      const rows = parseRenderedReport();
      const reportDocs = [document];
      let pageText = clean(document.body?.textContent);
      for (const frame of Array.from(document.querySelectorAll('iframe'))) {
        try {
          if (frame.contentDocument) reportDocs.push(frame.contentDocument);
          pageText += ' ' + clean(frame.contentDocument?.body?.textContent);
        } catch {}
      }
      const zeroResult = /Anzahl\s+Pakete\s*:\s*0(?:\D|$)/i.test(pageText);
      let expectedCount = null;
      const expectedPattern = /Anzahl\s+Pakete\s*:?\s*(\d[\d.]*)/gi;
      let expectedMatch;
      while ((expectedMatch = expectedPattern.exec(pageText)) !== null) {
        const value = Number(String(expectedMatch[1] || '').replace(/\D/g, ''));
        if (Number.isFinite(value) && (expectedCount == null || value > expectedCount)) expectedCount = value;
      }

      // Primary path: one complete in-memory Excel export, regardless of report page count.
      // The old page-by-page reader remains only as a fallback when the server
      // does not expose an export URL.
      if (!zeroResult && expectedCount != null && !exportSucceeded && !exportExhausted) {
        void tryBackgroundExcelExport(reportDocs, expectedCount).then(success => {
          if (!success) return;
          clearInterval(timer);
          try { window.top.postMessage({ type: 'pm-two-use-complete', runId: renderRunId, rows: expectedCount, expected: expectedCount }, '*'); } catch {}
          if (window.name === 'pmTwoUseBridge') setTimeout(() => window.close(), 400);
        });
        return;
      }
      let currentPage = 0;
      let totalPages = 0;
      const pagePattern = /\b(\d+)\s+(?:von|of)\s+(\d+)\b/gi;
      let pageMatch;
      while ((pageMatch = pagePattern.exec(pageText)) !== null) {
        const current = Number(pageMatch[1] || 0);
        const total = Number(pageMatch[2] || 0);
        if (current > 0 && total >= current && total > totalPages) {
          currentPage = current;
          totalPages = total;
        }
      }
      const pageInputForNav = reportDocs.flatMap(doc => Array.from(doc.querySelectorAll('input'))).find(input => {
        const type = String(input.type || 'text').toLowerCase();
        const rect = input.getBoundingClientRect();
        return /^(text|number)$/.test(type) && /^\d+$/.test(String(input.value || '')) && rect.width > 20 && rect.height > 10;
      });
      if (!currentPage && pageInputForNav) {
        const inputPage = Number(pageInputForNav.value || 0);
        const inputDocText = clean(pageInputForNav.ownerDocument?.body?.textContent);
        // input.value is not part of body.textContent. Read the current page
        // from the field and only the total page count from the visible text.
        const inputPages = inputDocText.match(/\b(?:von|of)\s*(\d+)\b/i);
        if (inputPage > 0 && Number(inputPages?.[1] || 0) >= inputPage) {
          currentPage = inputPage;
          totalPages = Number(inputPages[1]);
        }
      }

      // The SSRS toolbar changes its page number before the report body has
      // necessarily finished rendering.  Never save/click on that transient
      // state: require both a content change from the preceding page and a
      // stable row set for several polling cycles.
      const rowsSignature = Array.isArray(rows) && rows.length
        ? rows.map(row => [row.pknr, row.scanTime, row.product, row.lastScan].map(value => String(value || '')).join('|')).join('\n')
        : '';
      let pageRowsReady = !currentPage;
      if (currentPage && rowsSignature) {
        if (observedPage !== currentPage || observedRowsSignature !== rowsSignature) {
          observedPage = currentPage;
          observedPageSince = Date.now();
          observedRowsSignature = rowsSignature;
          stablePageSamples = 1;
        } else {
          stablePageSamples++;
        }
        const previousSignature = String(persistedPages.pageSignatures?.[String(currentPage - 1)] || '');
        const contentChangedFromPreviousPage = currentPage <= 1 || !previousSignature || rowsSignature !== previousSignature;
        pageRowsReady = contentChangedFromPreviousPage && stablePageSamples >= 8 && Date.now() - observedPageSince >= 1750;
      }

      if (Array.isArray(rows) && rows.length > 0 && pageRowsReady) {
        if (currentPage) {
          // SSRS reloads the embedded viewer on every page change. Persist the
          // current page before clicking next so earlier rows are not lost.
          if (expectedCount != null && persistedPages.expected != null && Number(persistedPages.expected) !== Number(expectedCount)) {
            persistedPages.pages = {};
            persistedPages.pageSignatures = {};
          }
          persistedPages.pages[String(currentPage)] = rows;
          persistedPages.pageSignatures = persistedPages.pageSignatures && typeof persistedPages.pageSignatures === 'object'
            ? persistedPages.pageSignatures : {};
          persistedPages.pageSignatures[String(currentPage)] = rowsSignature;
          persistedPages.expected = expectedCount;
          persistedPages.totalPages = totalPages;
          persistedPages.at = Date.now();
          try { GM_setValue('pmTwoUsePageRows', JSON.stringify(persistedPages)); } catch {}
          rebuildCollectedFromPages();
          if (expectedCount != null && collected.size > expectedCount) {
            persistedPages.pages = { [String(currentPage)]: rows };
            persistedPages.pageSignatures = { [String(currentPage)]: rowsSignature };
            try { GM_setValue('pmTwoUsePageRows', JSON.stringify(persistedPages)); } catch {}
            rebuildCollectedFromPages();
          }
        }
        const occurrences = new Map();
        if (!currentPage) rows.forEach((row, index) => {
          // Eine Sendungsnummer kann im Bericht mehrfach mit unterschiedlichen
          // Scans/Produkten vorkommen. Nur nach PKNR zu deduplizieren ließ z. B.
          // 34 sichtbare Zeilen fälschlich als 33 Pakete erscheinen.
          const fingerprint = [
            row.pknr, row.scanTime, row.product, row.lastScan,
            row.vd, row.route, row.targetZip, row.ticker
          ].map(value => String(value || '')).join('|') || `row-${index}`;
          const occurrence = (occurrences.get(fingerprint) || 0) + 1;
          occurrences.set(fingerprint, occurrence);
          const key = currentPage ? `page-${currentPage}|row-${index}` : `${fingerprint}|#${occurrence}`;
          collected.set(key, row);
        });
        try {
          GM_setValue('pmTwoUseProgress', JSON.stringify({ runId: renderRunId, found: expectedCount == null ? collected.size : Math.min(collected.size, expectedCount), expected: expectedCount, at: Date.now() }));
        } catch {}

        if (currentPage && totalPages && currentPage < totalPages) {
          stableWithoutPaging = 0;
          const exactNextCandidates = reportDocs.flatMap(doc => [
            ...Array.from(doc.querySelectorAll('[title], [aria-label]')),
            ...Array.from(doc.querySelectorAll('svg title, title')).map(title => title.parentElement).filter(Boolean)
          ]);
          const exactNext = exactNextCandidates.find(control => {
            const label = [control.getAttribute('title'), control.getAttribute('aria-label')]
              .filter(Boolean).join(' ') || clean(control.textContent);
            const normalizedLabel = label.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
            return normalizedLabel === 'nachste seite' && !control.disabled;
          });
          if (exactNext && (lastClickedPage !== currentPage || Date.now() - lastPageClickAt > 5000)) {
            lastClickedPage = currentPage;
            lastPageClickAt = Date.now();
            const clickable = exactNext.closest?.('button, a, input, [role="button"], [onclick], [tabindex]') || exactNext;
            activateControl(clickable);
            try {
              GM_setValue('pmTwoUseProgress', JSON.stringify({
                runId: renderRunId, found: expectedCount == null ? collected.size : Math.min(collected.size, expectedCount), expected: expectedCount, page: currentPage,
                pages: totalPages, action: 'click-title-naechste-seite', at: Date.now()
              }));
            } catch {}
            return;
          }
          if (!exactNext && pageInputForNav && (lastClickedPage !== currentPage || Date.now() - lastPageClickAt > 5000)) {
            const rect = pageInputForNav.getBoundingClientRect();
            const doc = pageInputForNav.ownerDocument;
            // Im SSRS-Toolbar liegt "Naechste Seite" konstant rechts neben
            // dem Seitenfeld und dem Text "von N".
            const hit = doc.elementFromPoint(rect.right + 82, (rect.top + rect.bottom) / 2);
            const clickable = hit?.closest?.('button, a, input, [role="button"], [onclick], [tabindex]') || hit;
            if (clickable) {
              lastClickedPage = currentPage;
              lastPageClickAt = Date.now();
              activateControl(clickable);
              try {
                GM_setValue('pmTwoUseProgress', JSON.stringify({
                  runId: renderRunId, found: expectedCount == null ? collected.size : Math.min(collected.size, expectedCount), expected: expectedCount, page: currentPage,
                  pages: totalPages, action: 'click-position-next', at: Date.now()
                }));
              } catch {}
              return;
            }
          }
          if (lastClickedPage !== currentPage || Date.now() - lastPageClickAt > 1500) {
            const rawControls = reportDocs.flatMap(doc => Array.from(doc.querySelectorAll(
              'button, a, input[type="image"], input[type="button"], [role="button"], [onclick], [tabindex], img, svg'
            )));
            const controls = Array.from(new Set(rawControls.map(control => {
              if (typeof control.closest !== 'function') return control;
              return control.closest('button, a, input, [role="button"], [onclick], [tabindex]') || control;
            })));
            let next = controls.find(control => {
              const label = [
                control.getAttribute('title'), control.getAttribute('aria-label'),
                control.getAttribute('name'), control.id, control.textContent
              ].filter(Boolean).join(' ');
              return /next\s*page|nextpage|nächste\s*seite|seiten?vor/i.test(label) && !control.disabled;
            });

            if (!next) {
              next = controls.find(control => {
                const label = [
                  control.getAttribute?.('title'), control.getAttribute?.('aria-label'),
                  control.getAttribute?.('alt'), control.getAttribute?.('data-original-title'),
                  control.getAttribute?.('name'), control.id, control.textContent
                ].filter(Boolean).join(' ').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
                return /next\s*page|nextpage|nachste\s*seite|seite\s*(?:vor|weiter)|nextbutton/.test(label) && !control.disabled;
              });
            }

            const directPageInput = reportDocs.flatMap(doc => Array.from(doc.querySelectorAll('input'))).find(input => {
              const type = String(input.type || 'text').toLowerCase();
              const rect = input.getBoundingClientRect();
              return String(input.value) === String(currentPage) && /^(text|number)$/.test(type) && rect.width > 20 && rect.height > 10;
            });
            // Nicht zuerst das Seitenfeld manipulieren: SSRS setzt dessen Wert
            // oft zurueck, ohne die Seite zu wechseln. Der echte Pfeil folgt unten.
            if (false && directPageInput && lastInputPage !== currentPage) {
              lastInputPage = currentPage;
              lastClickedPage = currentPage;
              lastPageClickAt = Date.now();
              directPageInput.focus();
              const inputView = directPageInput.ownerDocument?.defaultView || window;
              const inputCtor = inputView.HTMLInputElement || HTMLInputElement;
              const setter = Object.getOwnPropertyDescriptor(inputCtor.prototype, 'value')?.set;
              if (setter) setter.call(directPageInput, String(currentPage + 1));
              else directPageInput.value = String(currentPage + 1);
              if (typeof directPageInput.setAttribute === 'function') directPageInput.setAttribute('value', String(currentPage + 1));
              directPageInput.dispatchEvent(new inputView.Event('input', { bubbles: true }));
              directPageInput.dispatchEvent(new inputView.Event('change', { bubbles: true }));
              ['keydown', 'keypress', 'keyup'].forEach(type => directPageInput.dispatchEvent(new inputView.KeyboardEvent(type, {
                key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
              })));
              try { directPageInput.blur(); } catch {}
              return;
            }

            if (!next) {
              const pageInput = reportDocs.flatMap(doc => Array.from(doc.querySelectorAll('input'))).find(input => {
                const type = String(input.type || 'text').toLowerCase();
                const rect = input.getBoundingClientRect();
                return String(input.value) === String(currentPage) && /^(text|number)$/.test(type) && rect.width > 20 && rect.height > 10;
              });
              if (pageInput) {
                const inputRect = pageInput.getBoundingClientRect();
                next = controls
                  .filter(control => {
                    if (control.disabled || control.ownerDocument !== pageInput.ownerDocument) return false;
                    const rect = control.getBoundingClientRect();
                    const sameToolbarLine = Math.abs((rect.top + rect.bottom) / 2 - (inputRect.top + inputRect.bottom) / 2) < 24;
                    return rect.width > 8 && rect.height > 8 && sameToolbarLine &&
                      rect.left > inputRect.right + 20 && rect.left < inputRect.right + 260;
                  })
                  .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)[0];

                // Nur wenn der echte Weiter-Pfeil nicht auffindbar ist, das
                // Seitenfeld als Rückfall direkt auf die nächste Seite setzen.
                if (!next && lastInputPage !== currentPage) {
                  lastInputPage = currentPage;
                  lastClickedPage = currentPage;
                  lastPageClickAt = Date.now();
                  pageInput.focus();
                  const inputCtor = pageInput.ownerDocument?.defaultView?.HTMLInputElement || HTMLInputElement;
                  const setter = Object.getOwnPropertyDescriptor(inputCtor.prototype, 'value')?.set;
                  if (setter) setter.call(pageInput, String(currentPage + 1));
                  else pageInput.value = String(currentPage + 1);
                  pageInput.dispatchEvent(new Event('input', { bubbles: true }));
                  pageInput.dispatchEvent(new Event('change', { bubbles: true }));
                  pageInput.dispatchEvent(new KeyboardEvent('keydown', {
                    key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
                  }));
                  pageInput.dispatchEvent(new KeyboardEvent('keyup', {
                    key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
                  }));
                  return;
                }
              }
            }

            if (next) {
              if (typeof next.closest === 'function') {
                const clickable = next.closest('button, a, input, [role="button"], [onclick], [tabindex]');
                if (clickable) next = clickable;
              }
              lastClickedPage = currentPage;
              lastPageClickAt = Date.now();
              activateControl(next);
              lastInputPage = 0;
              try {
                GM_setValue('pmTwoUseProgress', JSON.stringify({
                  runId: renderRunId, found: expectedCount == null ? collected.size : Math.min(collected.size, expectedCount), expected: expectedCount, page: currentPage,
                  pages: totalPages, action: 'next', at: Date.now()
                }));
              } catch {}
            } else {
              // Wenn kein Schalter erkannt wurde, die Seiteneingabe beim
              // nÃ¤chsten Durchlauf erneut versuchen.
              lastInputPage = 0;
            }
          }
          return;
        }

        if (!currentPage || !totalPages) {
          if (expectedCount != null && collected.size < expectedCount) {
            const scrollables = reportDocs.flatMap(doc => {
              const candidates = [doc.scrollingElement, ...Array.from(doc.querySelectorAll('*'))];
              return candidates.filter(element => element && element.clientHeight > 80 && element.scrollHeight > element.clientHeight + 20);
            });
            const uniqueScrollables = Array.from(new Set(scrollables));
            const reportTables = reportDocs.flatMap(doc => Array.from(doc.querySelectorAll('table')).filter(table => {
              const text = clean(table.textContent);
              return /Produkt/i.test(text) && /PKNR/i.test(text) && /Scan/i.test(text);
            }));
            // Zuerst den Scroll-Container bewegen, in dem die eigentliche
            // Berichtstabelle liegt. Das Seiten-Scrolling lädt keine weiteren
            // virtualisierten Tabellenzeilen nach.
            const tableScrollables = uniqueScrollables.filter(element =>
              reportTables.some(table => element.contains?.(table))
            );
            const activeScrollables = tableScrollables.length ? tableScrollables : uniqueScrollables;
            const tableDistance = element => {
              let best = Number.MAX_SAFE_INTEGER;
              reportTables.forEach(table => {
                let node = table;
                let distance = 0;
                while (node && distance < best) {
                  if (node === element) { best = distance; break; }
                  node = node.parentElement;
                  distance++;
                }
              });
              return best;
            };
            if (Date.now() - lastVirtualScrollAt < 700) return;
            const target = activeScrollables.filter(element => element.scrollHeight - element.clientHeight - element.scrollTop > 3).sort((a, b) => {
              const distanceDifference = tableDistance(a) - tableDistance(b);
              if (distanceDifference) return distanceDifference;
              const remainingA = a.scrollHeight - a.clientHeight - a.scrollTop;
              const remainingB = b.scrollHeight - b.clientHeight - b.scrollTop;
              return remainingB - remainingA;
            })[0];
            if (target) {
              const before = target.scrollTop;
              const step = Math.max(180, Math.floor(target.clientHeight * 0.42));
              const nextTop = Math.min(target.scrollHeight - target.clientHeight, before + step);
              const view = target.ownerDocument?.defaultView;
              if (target === target.ownerDocument?.scrollingElement && view) view.scrollTo(0, nextTop);
              else target.scrollTop = nextTop;
              target.dispatchEvent(new Event('scroll', { bubbles: true }));
              try { view?.dispatchEvent(new Event('scroll')); } catch {}
              lastVirtualScrollAt = Date.now();
              if (target.scrollTop !== before) {
                stableWithoutPaging = 0;
                return;
              }
            }
            if (!target && activeScrollables.length && virtualScrollPass < 2) {
              virtualScrollPass++;
              activeScrollables.forEach(element => {
                const view = element.ownerDocument?.defaultView;
                if (element === element.ownerDocument?.scrollingElement && view) view.scrollTo(0, 0);
                else element.scrollTop = 0;
                element.dispatchEvent(new Event('scroll', { bubbles: true }));
              });
              lastVirtualScrollAt = Date.now();
              stableWithoutPaging = 0;
              return;
            }
          }
          stableWithoutPaging++;
          if (stableWithoutPaging < 3) return;
        }
      }

      // A page number without a stable, newly rendered body is not a loaded
      // report page yet. Keep polling instead of storing stale rows or closing.
      if (currentPage && totalPages && !pageRowsReady) return;

      let finishedRows = Array.from(collected.values());
      // A one-page report is authoritative on its own and must not be blocked
      // by remnants in the cross-reload page cache.
      if (currentPage === 1 && totalPages === 1 && expectedCount != null && Array.isArray(rows) && rows.length === expectedCount) {
        finishedRows = rows.slice();
        persistedPages.pages = { '1': rows.slice() };
        persistedPages.expected = expectedCount;
        try { GM_setValue('pmTwoUsePageRows', JSON.stringify(persistedPages)); } catch {}
      }
      if (!zeroResult && expectedCount == null) {
        // Ohne die aktuelle 2Use-Gesamtzahl darf eine sichtbare Einzelseite
        // niemals als vollstÃ¤ndiges Live-Ergebnis gespeichert werden.
        if (attempts >= 720) {
          clearInterval(timer);
          GM_setValue('pmTwoUseRenderError', `2Use-Gesamtzahl nicht erkannt; ${finishedRows.length} sichtbare Zeilen wurden nicht Ã¼bernommen.`);
        }
        return;
      }
      if (expectedCount != null && expectedCount > 0 && finishedRows.length !== expectedCount) {
        // Keine Teilübernahme: Der bisherige Tabellenstand bleibt erhalten,
        // bis der neue Bericht vollständig eingelesen wurde.
        if (attempts >= 720) {
          clearInterval(timer);
          GM_setValue('pmTwoUseRenderError', `2Use unvollständig: ${finishedRows.length} von ${expectedCount} Paketen eingelesen.`);
          return;
        }
        return;
      }
      if (finishedRows.length > 0 || zeroResult) {
        clearInterval(timer);
        GM_setValue('pmTwoUseRows', JSON.stringify(finishedRows));
        GM_setValue('pmTwoUseRowsAt', Date.now());
        GM_setValue('pmTwoUseRowsRun', renderRunId);
        GM_setValue('pmTwoUseExpectedCount', expectedCount == null ? finishedRows.length : expectedCount);
        GM_setValue('pmTwoUseImportSource', 'SSRS-Seitenansicht');
        try { window.top.postMessage({ type: 'pm-two-use-complete', runId: renderRunId, rows: finishedRows.length, expected: expectedCount }, '*'); } catch {}
        if (window.name === 'pmTwoUseBridge') setTimeout(() => window.close(), 400);
      } else if (attempts >= 720) {
        clearInterval(timer);
        GM_setValue('pmTwoUseRenderError', 'Die gerenderte 2Use-Tabelle wurde nicht gefunden.');
      }
    }, 250);
    return;
  }

  if (location.hostname === '2use-prod.dpdit.de') {
    const portalCommandStartedAt = Number(new URLSearchParams(location.search).get('pmCommandId') || 0);
    if (portalCommandStartedAt) {
      const closeCompletedPortal = () => {
        let rowsAt = 0, rowsRun = 0;
        try {
          rowsAt = Number(GM_getValue('pmTwoUseRowsAt', 0) || 0);
          rowsRun = Number(GM_getValue('pmTwoUseRowsRun', 0) || 0);
        } catch {}
        if (rowsRun === portalCommandStartedAt && rowsAt >= portalCommandStartedAt) {
          try { window.close(); } catch {}
          return true;
        }
        return false;
      };
      window.addEventListener('message', event => {
        let sourceHost = '';
        try { sourceHost = new URL(event.origin).hostname; } catch { return; }
        if (!/\.dpdit\.de$/i.test(sourceHost)) return;
        if (event.data?.type === 'pm-two-use-complete' && Number(event.data.runId || 0) === portalCommandStartedAt) setTimeout(() => {
          closeCompletedPortal();
          try { window.close(); } catch {}
        }, 150);
      });
      const closeTimer = setInterval(() => {
        if (closeCompletedPortal()) clearInterval(closeTimer);
      }, 250);
      setTimeout(() => clearInterval(closeTimer), 240000);
    }
    const captureReportUrl = () => {
      let reportUrl = '';
      try { reportUrl = String(pageWindow.REPORTFORMURL || ''); } catch {}
      if (!reportUrl) {
        try {
          reportUrl = String(pageWindow.eval('typeof REPORTFORMURL !== "undefined" ? REPORTFORMURL : ""') || '');
        } catch {}
      }
      if (!reportUrl) {
        const frame = document.getElementById('idIframeReport');
        const src = frame?.getAttribute('src') || frame?.src || '';
        if (/2use-render-prod\.dpdit\.de/i.test(src)) reportUrl = src;
      }
      if (!reportUrl) return false;
      try { reportUrl = new URL(reportUrl, location.href).href; } catch { return false; }
      GM_setValue('pmTwoUseReportUrl', reportUrl);
      return true;
    };

    if (!captureReportUrl()) {
      let attempts = 0;
      const timer = setInterval(() => {
        attempts++;
        if (captureReportUrl() || attempts >= 240) clearInterval(timer);
      }, 250);
    }

    const autoQuery = new URLSearchParams(location.search);
    if (autoQuery.get('pmAutoRun') === '1' && sessionStorage.getItem('pmTwoUseAutoCommand') !== autoQuery.get('pmCommandId')) {
      const query = autoQuery;
      const event = type => new pageWindow.Event(type, { bubbles: true });
      const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
      const waitFor = async (test, timeout = 30000, interval = 150) => {
        const until = Date.now() + timeout;
        while (Date.now() < until) {
          try {
            const result = test();
            if (result) return result;
          } catch {}
          await sleep(interval);
        }
        throw new Error('2Use hat das Parameterformular nicht rechtzeitig aktualisiert.');
      };
      const waitStable = async (test, stableMs = 2200, timeout = 12000) => {
        const until = Date.now() + timeout;
        let stableSince = 0;
        while (Date.now() < until) {
          let okay = false;
          try { okay = !!test(); } catch {}
          if (okay) {
            if (!stableSince) stableSince = Date.now();
            if (Date.now() - stableSince >= stableMs) return true;
          } else {
            stableSince = 0;
          }
          await sleep(150);
        }
        return false;
      };
      const form = () => document.getElementById('frmSelectReportParameter2187');
      const setStatus = (message, error = false) => {
        let box = document.getElementById('pm-two-use-auto-status');
        if (!box) {
          box = document.createElement('div');
          box.id = 'pm-two-use-auto-status';
          box.style.cssText = 'position:fixed;right:16px;top:16px;z-index:2147483647;max-width:460px;padding:11px 14px;border-radius:8px;box-shadow:0 5px 22px rgba(0,0,0,.25);font:600 13px system-ui';
          document.body.appendChild(box);
        }
        box.style.background = error ? '#fee2e2' : '#eff6ff';
        box.style.color = error ? '#991b1b' : '#1e3a8a';
        box.textContent = message;
      };
      const selectedProducts = () => Array.from(document.querySelectorAll('#paraExpressProdukt option:checked'))
        .map(option => String(option.value || '')).filter(value => value && value.toLowerCase() !== 'all');
      const currentLocations = () => Array.from(document.querySelectorAll('[name="paraLocation"]'))
        .map(field => String(field.value || '').replace(/^0+/, '')).filter(Boolean);
      const setLocations = async depot => {
        const targetLocations = depot === '195' ? ['107', '195', '295'] : [depot];
        if (targetLocations.every(value => currentLocations().includes(value))) return;

        const currentForm = await waitFor(() => form());
        currentForm.querySelectorAll('[name="paraLocation"], [name="paraLocationLabel"]').forEach(field => field.remove());
        const labels = depot === '195'
          ? ['0107 - belongs to 0195', ' Leupoldsgrün', '0195 - Leupoldsgrün', '0295 - belongs to 0195', ' Leupoldsgrün']
          : [`${String(depot).padStart(4, '0')} - Depot ${String(depot).padStart(3, '0')}`];
        const addHidden = (name, value, index) => {
          const input = document.createElement('input');
          input.type = 'hidden';
          input.name = name;
          input.id = `${name}${index + 1}`;
          input.value = value;
          currentForm.appendChild(input);
        };
        targetLocations.forEach((value, index) => addHidden('paraLocation', value, index));
        labels.forEach((value, index) => addHidden('paraLocationLabel', value, index));
        if (!targetLocations.every(value => currentLocations().includes(value))) {
          throw new Error(`Depot ${depot} konnte nicht in das 2Use-Formular eingetragen werden.`);
        }
      };
      const setSelect = async (id, value, label, forceChange = false) => {
        for (let attempt = 1; attempt <= 4; attempt++) {
          const select = await waitFor(() => {
            const current = document.getElementById(id);
            return current && Array.from(current.options).some(item => String(item.value) === String(value)) ? current : null;
          }, 30000);
          if (String(select.value) !== String(value) || forceChange) {
            select.value = String(value);
            select.dispatchEvent(event('input'));
            select.dispatchEvent(event('change'));
            await sleep(700);
          }
          if (await waitStable(() => String(document.getElementById(id)?.value || '') === String(value), 900, 6000)) return;
        }
        throw new Error(`${label} wurde von 2Use nach dem Setzen wieder verworfen.`);
      };
      const setAllProducts = async () => {
        for (let attempt = 1; attempt <= 4; attempt++) {
          const select = await waitFor(() => {
            const current = document.getElementById('paraExpressProdukt');
            return current && Array.from(current.options).some(option => option.value && option.value.toLowerCase() !== 'all')
              ? current
              : null;
          }, 30000);
          const products = Array.from(select.options).filter(option => option.value && option.value.toLowerCase() !== 'all');
          const allSelected = products.every(option => option.selected);
          const invalidPlaceholderSelected = Array.from(select.options)
            .some(option => !String(option.value || '').trim() && option.selected);
          if (!allSelected || invalidPlaceholderSelected) {
            Array.from(select.options).forEach(option => {
              const value = String(option.value || '').trim();
              option.selected = Boolean(value) && value.toLowerCase() !== 'all';
            });
            select.dispatchEvent(event('input'));
            select.dispatchEvent(event('change'));
            await sleep(700);
          }
          if (await waitStable(() => {
            const current = document.getElementById('paraExpressProdukt');
            const available = current ? Array.from(current.options).filter(option => option.value && option.value.toLowerCase() !== 'all').length : 0;
            const hasInvalidPlaceholder = current
              ? Array.from(current.options).some(option => !String(option.value || '').trim() && option.selected)
              : true;
            return available > 0 && !hasInvalidPlaceholder && selectedProducts().length === available;
          }, 900, 6000)) return;
        }
        throw new Error('Alle Produkte wurden von 2Use nach dem Setzen wieder verworfen.');
      };
      const setAtomicParameters = async (depot, fromValue, toValue) => {
        const controls = await waitFor(() => {
          const from = document.getElementById('paraPeriodFrom');
          const to = document.getElementById('paraPeriodTo');
          const comparison = document.getElementById('paraDifferenzabgleich');
          const products = document.getElementById('paraExpressProdukt');
          const hasFrom = from && Array.from(from.options).some(option => String(option.value) === fromValue);
          const hasTo = to && Array.from(to.options).some(option => String(option.value) === toValue);
          const hasProducts = products && Array.from(products.options).some(option => {
            const value = String(option.value || '').trim();
            return value && value.toLowerCase() !== 'all';
          });
          return hasFrom && hasTo && comparison && hasProducts ? { from, to, comparison, products } : null;
        }, 30000);

        await setLocations(depot);
        const selectSingle = (select, value) => {
          Array.from(select.options).forEach(option => {
            const selected = String(option.value) === String(value);
            option.selected = selected;
            option.defaultSelected = selected;
            option.toggleAttribute('selected', selected);
          });
          select.value = String(value);
        };
        selectSingle(controls.comparison, '1');
        selectSingle(controls.from, fromValue);
        selectSingle(controls.to, toValue);
        Array.from(controls.products.options).forEach(option => {
          const value = String(option.value || '').trim();
          const selected = Boolean(value) && value.toLowerCase() !== 'all';
          option.selected = selected;
          option.defaultSelected = selected;
          option.toggleAttribute('selected', selected);
        });

        const setNamedValue = (name, value) => {
          form().querySelectorAll(`[name="${name}"]`).forEach(field => { field.value = value; });
        };
        const optionLabel = select => String(select.options[select.selectedIndex]?.textContent || '')
          .replace(/\s+/g, ' ').trim();
        setNamedValue('paraPeriodFromLabel', optionLabel(controls.from));
        setNamedValue('paraPeriodToLabel', optionLabel(controls.to));
        setNamedValue('paraSendingCustomer', ' ');
        const yes = form().querySelector('input[name="paraMissing"][value="Ja"]');
        const no = form().querySelector('input[name="paraMissing"][value="Nein"]');
        if (yes) yes.checked = true;
        if (no) no.checked = false;
        controls.products.classList.remove('missingValBorder');
        controls.products.removeAttribute('title');

        const button = await waitFor(() => document.getElementById('btnExecReport'));
        button.disabled = false;
        button.removeAttribute('disabled');
        button.setAttribute('aria-disabled', 'false');
        return button;
      };
      const keepVisibleDateRange = (fromValue, toValue) => {
        const format = value => `${String(value).slice(6, 8)}.${String(value).slice(4, 6)}.${String(value).slice(0, 4)}`;
        const show = (id, value) => {
          const select = document.getElementById(id);
          if (!select) return;
          let option = Array.from(select.options).find(item => String(item.value) === String(value));
          if (!option) {
            option = document.createElement('option');
            option.value = String(value);
            option.textContent = format(value);
            select.appendChild(option);
          }
          Array.from(select.options).forEach(item => { item.selected = item === option; });
          select.value = String(value);
          // vanillaSelectBox versteckt das native Select und zeigt einen eigenen
          // Button. Dessen Beschriftung muss separat synchronisiert werden.
          const container = select.closest('.col') || select.parentElement;
          container?.querySelectorAll('button').forEach(button => {
            if (/^\s*\d{2}\.\d{2}\.\d{4}\s*$/.test(String(button.textContent || ''))) {
              button.textContent = format(value);
            }
          });
        };
        const timer = setInterval(() => {
          show('paraPeriodFrom', fromValue);
          show('paraPeriodTo', toValue);
        }, 250);
        setTimeout(() => clearInterval(timer), 120000);
      };

      (async () => {
        try {
          setStatus('2Use wird vorbereitet …');
          await waitFor(() => form() && document.getElementById('paraPeriodFrom'));
          const username = String(form()?.querySelector('[name="paraUsername"]')?.value || '').trim();
          if (username) GM_setValue('pmTwoUseUsername', username);
          const depot = String(query.get('pmDepot') || '').replace(/\D+/g, '').replace(/^0+/, '');
          if (!depot) throw new Error('In den Dispatcher-Einstellungen fehlt die Depotnummer.');

          // Erst die Abfrage wählen: 2Use stellt die historischen Datumswerte
          // und die Produktliste erst abhängig von dieser Auswahl bereit.
          setStatus('2Use: Eingangsabgleich wird ausgewählt …');
          await setSelect('paraDifferenzabgleich', '1', 'Abfrage Einrollung VD – Eingang ED', true);

          // Keine weiteren Change-Events: Jede einzelne Änderung würde das ganze
          // 2Use-Formular erneut laden und dabei zuvor gesetzte Werte zurücksetzen.
          setStatus('2Use: Parameter werden gemeinsam gesetzt …');
          let executeButton = await setAtomicParameters(depot, query.get('pmFrom'), query.get('pmTo'));
          const displayDate = value => `${String(value).slice(6, 8)}.${String(value).slice(4, 6)}.${String(value).slice(0, 4)}`;
          setStatus(`2Use: ${displayDate(query.get('pmFrom'))} bis ${displayDate(query.get('pmTo'))} · alle Produkte`);
          await sleep(1200);
          // Unmittelbar vor dem Klick nochmals setzen, falls ein UI-Plugin nur
          // anhand der selected-Attribute neu aufgebaut wurde.
          executeButton = await setAtomicParameters(depot, query.get('pmFrom'), query.get('pmTo'));

          GM_setValue('pmTwoUseBridgeSubmittedAt', Date.now());
          sessionStorage.setItem('pmTwoUseAutoCommand', query.get('pmCommandId') || String(Date.now()));
          setStatus('2Use: Bericht wird gestartet …');
          keepVisibleDateRange(query.get('pmFrom'), query.get('pmTo'));
          executeButton.click();
          // 2Use blendet trotz erfolgreich erzeugtem Bericht gelegentlich noch
          // die veraltete Meldung "Bitte alle Parameter setzen!" darüber ein.
          const warningTimer = setInterval(() => {
            Array.from(document.querySelectorAll('div, section, aside')).forEach(element => {
              if (/^Bitte\s+alle\s+Parameter\s+setzen!?$/i.test(String(element.textContent || '').replace(/\s+/g, ' ').trim())) {
                element.style.setProperty('display', 'none', 'important');
              }
            });
          }, 500);
          setTimeout(() => clearInterval(warningTimer), 120000);
        } catch (error) {
          const message = String(error?.message || error || 'Unbekannter Fehler');
          GM_setValue('pmTwoUseRenderError', message);
          setStatus(`2Use-Automatik: ${message}`, true);
        }
      })();
    }
    return;
  }

  const moduleDef = {
    id: 'prio-express-monitor',
    label: 'Prio / Express',
    run: () => startModuleOnce()
  };

  try {
    if (pageWindow.TM && typeof pageWindow.TM.register === 'function') {
      pageWindow.TM.register(moduleDef);
    } else {
      pageWindow.__tmQueue = pageWindow.__tmQueue || [];
      pageWindow.__tmQueue.push(moduleDef);
    }
  } catch (error) {
    console.warn('Prio / Express: Registrierung im Tool-Menü fehlgeschlagen.', error);
  }

  let started = false;
  function startModuleOnce() {
    if (started) {
      togglePanel(true);
      connectTwoUse();
      initialOpenRefresh().catch(console.error);
      return;
    }
    started = true;
    boot();
    togglePanel(true);
    connectTwoUse();
    initialOpenRefresh().catch(console.error);
  }

  const NS  = 'pm-';
  const esc = s => String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => String(s || '').replace(/\s+/g, ' ').trim();
  const tourKey = t => String(t || '').replace(/[^\dA-Za-z]/g, '').trim();
  const collator = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });

  const state = {
    events: [],
    nextId: 1,
    _bootShown: false,
    _prioAllList: [],
    _prioOpenList: [],
    _expAllList: [],
    _expOpenList: [],
    _expLate11List: [],
    _noRolloutList: [],
    _twoUseError: '',
    _twoUseLoading: false,
    _twoUseRangeLabel: '',
    _twoUseProgressLabel: '',
    _dispatcherLoading: false,
    _footerMessage: '',
    _pendingTwoUseKind: '',
    _modal: { rows: [], opts: {}, title: '', selected: new Set() },
    lastRefreshAt: 0
  };

  let lastOkRequest = null;
  let isBusy = false;
  let isLoading = false;
  let sourceRefreshTimer = null;
  let suppressSourceRefreshUntil = 0;
  let twoUseConnectionActive = false;
  let lastTwoUseStartedAt = 0;
  let twoUseAutoTimer = null;
  let twoUsePollTimer = null;
  let twoUsePopup = null;

  const TWO_USE_PAGE = 'https://2use-prod.dpdit.de/Report/Index/2187?portalId=79';
  const TWO_USE_PRODUCTS = [
    'DPD 8:30', 'DPD 10:00', 'DPD 12:00', 'DPD 18:00', 'DPD Express',
    'DPD 10:00 (Mo-Sa)', 'DPD 12:00 (Mo-Sa)',
    'DPD Food 12:00 (Mo-Sa)', 'DPD Food 18:00 (Mo-Sa)', 'DPD Priority'
  ];

  function twoUseRequest({ method = 'GET', url, data = null, headers = {}, timeout = 45000 }) {
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest !== 'function') {
        reject(new Error('GM_xmlhttpRequest ist nicht verfügbar. Userscript-Berechtigungen prüfen.'));
        return;
      }
      GM_xmlhttpRequest({
        method,
        url,
        data,
        headers,
        timeout,
        anonymous: false,
        withCredentials: true,
        onload: r => {
          if (r.status >= 200 && r.status < 400) resolve(r);
          else reject(new Error(`2Use HTTP ${r.status} bei ${url}`));
        },
        ontimeout: () => reject(new Error('2Use-Abfrage hat das Zeitlimit überschritten.')),
        onerror: () => reject(new Error('2Use ist nicht erreichbar oder die Anmeldung fehlt.'))
      });
    });
  }

  function localDateParts(date) {
    const pad = n => String(n).padStart(2, '0');
    const y = date.getFullYear();
    const m = pad(date.getMonth() + 1);
    const d = pad(date.getDate());
    return { value: `${y}${m}${d}`, label: `${d}.${m}.${y}` };
  }

  function localIsoDate(date) {
    const pad = n => String(n).padStart(2, '0');
    return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
  }

  function parseLocalIsoDate(value) {
    const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12) : null;
  }

  const TWO_USE_SESSION_DATE_KEY = 'pmTwoUseDateRangeForCurrentPage';
  // sessionStorage hält den Wert auch bei internen SPA-Aktualisierungen. Beim
  // echten Seitenladen ist das Window-Flag wieder weg und der Wert wird gelöscht.
  if (!pageWindow.__pmTwoUseDateSessionInitialized) {
    try { sessionStorage.removeItem(TWO_USE_SESSION_DATE_KEY); } catch {}
    pageWindow.__pmTwoUseDateSessionInitialized = true;
  }

  function getManualTwoUseDateRange() {
    try {
      const saved = JSON.parse(String(sessionStorage.getItem(TWO_USE_SESSION_DATE_KEY) || 'null'));
      const from = parseLocalIsoDate(saved?.from);
      const to = parseLocalIsoDate(saved?.to);
      return from && to && from <= to ? { from, to } : null;
    } catch {
      return null;
    }
  }

  function setManualTwoUseDateRange(range) {
    try {
      if (!range) sessionStorage.removeItem(TWO_USE_SESSION_DATE_KEY);
      else sessionStorage.setItem(TWO_USE_SESSION_DATE_KEY, JSON.stringify({
        from: localIsoDate(range.from), to: localIsoDate(range.to)
      }));
    } catch {}
  }

  function automaticTwoUseDateRange() {
    const to = new Date();
    // So -> Fr, Mo -> Fr, Di-Sa -> jeweiliger Vortag.
    const daysBackByWeekday = [2, 3, 1, 1, 1, 1, 1];
    const daysBack = daysBackByWeekday[to.getDay()];
    const from = new Date(to.getFullYear(), to.getMonth(), to.getDate() - daysBack, 12);
    return { from, to };
  }

  function selectedTwoUseDateRange() {
    const manual = getManualTwoUseDateRange();
    if (manual) {
      return { from: manual.from, to: manual.to, manual: true };
    }
    return { ...automaticTwoUseDateRange(), manual: false };
  }

  function openTwoUseDateSettings() {
    const current = selectedTwoUseDateRange();
    document.getElementById(NS + 'two-use-date-dialog')?.remove();
    const overlay = document.createElement('div');
    overlay.id = NS + 'two-use-date-dialog';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:100010;background:rgba(15,23,42,.38);display:flex;align-items:center;justify-content:center;font-family:system-ui';
    overlay.innerHTML = `
      <form style="width:min(440px,calc(100vw - 32px));background:#fff;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 18px 50px rgba(0,0,0,.25);padding:18px">
        <div style="font-size:16px;font-weight:800;margin-bottom:5px">2Use-Datum anpassen</div>
        <div style="font-size:12px;color:#475569;margin-bottom:16px">Ein Datum auswählen – die Abfrage startet sofort. Beim Neuladen der Seite gilt wieder automatisch die Wochenregel.</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
          <label style="font-size:12px;font-weight:700">Datum von<input name="from" type="date" value="${localIsoDate(current.from)}" required style="display:block;width:100%;box-sizing:border-box;margin-top:5px;padding:8px;border:1px solid #cbd5e1;border-radius:7px"></label>
          <label style="font-size:12px;font-weight:700">Datum bis<input name="to" type="date" value="${localIsoDate(current.to)}" required style="display:block;width:100%;box-sizing:border-box;margin-top:5px;padding:8px;border:1px solid #cbd5e1;border-radius:7px"></label>
        </div>
        <div data-error style="min-height:18px;margin-top:8px;color:#b91c1c;font-size:12px"></div>
        <div style="display:flex;justify-content:space-between;gap:8px;margin-top:8px">
          <button type="button" data-auto style="padding:7px 11px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;cursor:pointer">Automatische Wochenregel</button>
          <button type="button" data-cancel style="padding:7px 13px;border:1px solid #cbd5e1;border-radius:7px;background:#fff;cursor:pointer">Abbrechen</button>
        </div>
      </form>`;
    const close = () => overlay.remove();
    overlay.querySelector('[data-cancel]').addEventListener('click', close);
    overlay.querySelector('[data-auto]').addEventListener('click', () => {
      setManualTwoUseDateRange(null);
      close();
      connectTwoUse(false, true);
    });
    overlay.addEventListener('mousedown', event => { if (event.target === overlay) close(); });
    overlay.querySelector('form').addEventListener('submit', event => event.preventDefault());
    const applySelectedDate = () => {
      const formElement = overlay.querySelector('form');
      const from = parseLocalIsoDate(formElement.elements.from.value);
      const to = parseLocalIsoDate(formElement.elements.to.value);
      if (!from || !to || from > to) {
        overlay.querySelector('[data-error]').textContent = '„Von“ darf nicht nach „Bis“ liegen.';
        return;
      }
      setManualTwoUseDateRange({ from, to });
      close();
      connectTwoUse(false, true);
    };
    overlay.querySelectorAll('input[type="date"]').forEach(input => input.addEventListener('change', applySelectedDate));
    document.body.appendChild(overlay);
  }

  function discoverTwoUseReportUrl(html) {
    const patterns = [
      /(?:const|let|var)\s+REPORTFORMURL\s*=\s*["'`]([^"'`]+)["'`]/i,
      /REPORTFORMURL["']?\s*:\s*["'`]([^"'`]+)["'`]/i,
      /data-reportformurl=["']([^"']+)["']/i
    ];
    for (const re of patterns) {
      const hit = String(html || '').match(re)?.[1];
      if (hit) return new URL(hit.replace(/&amp;/g, '&'), TWO_USE_PAGE).href;
    }
    return '';
  }

  function buildTwoUseParameters(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const form = doc.querySelector('#frmSelectReportParameter2187');
    if (!form) throw new Error('2Use-Anmeldung fehlt oder das Parameterformular wurde nicht gefunden.');

    const params = new URLSearchParams();
    form.querySelectorAll('input[type="hidden"][name]').forEach(input => {
      if (input.name && input.value != null) params.append(input.name, input.value);
    });

    const dateRange = selectedTwoUseDateRange();
    const from = localDateParts(dateRange.from);
    const to = localDateParts(dateRange.to);
    try {
      GM_setValue('pmTwoUseLastDateRange', JSON.stringify({ from: localIsoDate(dateRange.from), to: localIsoDate(dateRange.to) }));
    } catch {}
    const set = (key, value) => { params.delete(key); params.set(key, value); };

    set('paraPeriodFrom', from.value);
    set('paraPeriodFromLabel', from.label);
    set('paraPeriodTo', to.value);
    set('paraPeriodToLabel', to.label);
    set('paraInterval', 'day');
    set('paraDifferenzabgleich', '1');
    const depot = String(getSetting('depotSuffix') || '').replace(/\D+/g, '').slice(-3);
    const depotValue = depot.replace(/^0+/, '');
    const currentLocations = Array.from(form.querySelectorAll('[name="paraLocation"]'))
      .map(field => String(field.value || '').replace(/^0+/, '')).filter(Boolean);
    const currentLocationLabels = Array.from(form.querySelectorAll('[name="paraLocationLabel"]'))
      .map(field => String(field.value || '')).filter(Boolean);
    const isDepot195 = depotValue === '195';
    const useCurrentTree = depotValue && currentLocations.includes(depotValue);
    const locations = isDepot195 ? ['107', '195', '295'] : useCurrentTree ? currentLocations : [depotValue].filter(Boolean);
    const locationLabels = isDepot195
      ? ['0107 - belongs to 0195', ' Leupoldsgrün', '0195 - Leupoldsgrün', '0295 - belongs to 0195', ' Leupoldsgrün']
      : useCurrentTree && currentLocationLabels.length
        ? currentLocationLabels
        : locations.map(value => `${String(value).padStart(4, '0')} - Depot ${String(value).padStart(3, '0')}`);
    if (!locations.length) throw new Error('In den Einstellungen ist keine Depotkennung hinterlegt.');
    params.delete('paraLocation');
    locations.forEach(value => params.append('paraLocation', value));
    params.delete('paraLocationLabel');
    locationLabels.forEach(value => params.append('paraLocationLabel', value));
    set('paraSendingCustomer', ' ');
    set('paraMissing', 'Ja');
    set('forAboFramework', 'false');

    params.delete('paraExpressProdukt');
    const currentProducts = Array.from(doc.querySelectorAll('select[name="paraExpressProdukt"] option'))
      .map(option => String(option.value || '').trim())
      .filter(value => value && value.toLowerCase() !== 'all');
    const products = currentProducts.length ? Array.from(new Set(currentProducts)) : TWO_USE_PRODUCTS;
    products.forEach(product => params.append('paraExpressProdukt', product));
    return params;
  }

  function findHeaderIndex(headers, pattern) {
    return headers.findIndex(h => pattern.test(norm(h)));
  }

  function parseTwoUseReport(html) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const tables = Array.from(doc.querySelectorAll('table'));
    let result = [];

    for (const table of tables) {
      const rows = Array.from(table.querySelectorAll('tr'));
      const headerAt = rows.findIndex(tr => {
        const text = norm(tr.textContent);
        return /Produkt/i.test(text) && /PKNR/i.test(text) && /Scan/i.test(text);
      });
      if (headerAt < 0) continue;

      const headers = Array.from(rows[headerAt].querySelectorAll('th,td')).map(c => norm(c.textContent));
      const idx = {
        status: findHeaderIndex(headers, /^Status$/i),
        lastScan: findHeaderIndex(headers, /letzter\s*Scan/i),
        scanTime: findHeaderIndex(headers, /Scanzeitpunkt/i),
        product: findHeaderIndex(headers, /^Produkt$/i),
        pknr: findHeaderIndex(headers, /^PKNR$/i),
        vd: findHeaderIndex(headers, /^VD$/i),
        route: findHeaderIndex(headers, /^Route$/i),
        zip: findHeaderIndex(headers, /Ziel\s*PLZ/i),
        ticker: findHeaderIndex(headers, /^Ticker$/i)
      };
      if (idx.product < 0) continue;

      result = rows.slice(headerAt + 1).map((tr, rowIndex) => {
        const cellElements = Array.from(tr.querySelectorAll(':scope > th, :scope > td'));
        const cells = cellElements.map(c => norm(c.textContent));
        const get = i => i >= 0 ? (cells[i] || '') : '';
        return {
          __twoUse: true,
          __twoUseId: rowIndex,
          status: get(idx.status),
          lastScan: get(idx.lastScan),
          scanTime: get(idx.scanTime),
          product: get(idx.product),
          pknr: get(idx.pknr),
          vd: get(idx.vd),
          route: get(idx.route),
          targetZip: get(idx.zip),
          ticker: get(idx.ticker),
          pknrUrl: '',
          tickerUrl: ''
        };
      }).filter(r => r.product && /DPD/i.test(r.product));
      return result;
    }
    return result;
  }

  async function refreshTwoUseMissing() {
    let raw = '', loadedAt = 0;
    try {
      raw = String(GM_getValue('pmTwoUseRows', '') || '');
      loadedAt = Number(GM_getValue('pmTwoUseRowsAt', 0) || 0);
    } catch {}
    const isToday = loadedAt && new Date(loadedAt).toDateString() === new Date().toDateString();
    if (!raw || !isToday) {
      throw new Error('Die 2Use-Daten für heute müssen einmal geladen werden.');
    }
    const rows = JSON.parse(raw);
    state._noRolloutList = Array.isArray(rows) ? rows : [];
    state._twoUseError = '';
    updateKpisOverview();
  }

  function connectTwoUse(silent = false, forceRestart = false) {
    if (twoUseConnectionActive && !forceRestart) return;
    if (forceRestart) {
      if (twoUsePollTimer) clearInterval(twoUsePollTimer);
      twoUsePollTimer = null;
      try { if (twoUsePopup && !twoUsePopup.closed) twoUsePopup.close(); } catch {}
      twoUsePopup = null;
      twoUseConnectionActive = false;
    }
    const startedAt = Date.now();
    const dateRange = selectedTwoUseDateRange();
    const from = localDateParts(dateRange.from);
    const to = localDateParts(dateRange.to);
    const depot = String(getSetting('depotSuffix') || '').replace(/\D+/g, '').slice(-3);
    if (!depot) {
      if (!silent) alert('Bitte zuerst in den Einstellungen die dreistellige Depotnummer eintragen.');
      return;
    }
    state._twoUseLoading = true;
    state._twoUseRangeLabel = `${from.label} bis ${to.label}`;
    state._twoUseProgressLabel = 'Abfrage wird vorbereitet';
    state._footerMessage = '';
    updateFooterStatus();
    try {
      GM_setValue('pmTwoUseRenderError', '');
      GM_setValue('pmTwoUseRenderHeartbeat', 0);
      GM_setValue('pmTwoUseProgress', '');
      GM_setValue('pmTwoUseActiveRun', startedAt);
      GM_setValue('pmTwoUseRowsRun', 0);
      GM_setValue('pmTwoUsePageRows', JSON.stringify({ runAt: startedAt, pages: {} }));
      GM_setValue('pmTwoUseReportUrl', '');
      GM_setValue('pmTwoUseBridgeSubmittedAt', 0);
      GM_setValue('pmTwoUseLastDateRange', JSON.stringify({ from: localIsoDate(dateRange.from), to: localIsoDate(dateRange.to) }));
    } catch {}
    const bridgeParams = new URLSearchParams({
      pmPrioBridge: '1',
      pmAutoRun: '1',
      pmDepot: depot,
      pmFrom: from.value,
      pmTo: to.value,
      pmCommandId: String(startedAt)
    });
    const url = TWO_USE_PAGE + '&' + bridgeParams.toString();
    let popup = null;
    if (typeof GM_openInTab === 'function') {
      try {
        popup = GM_openInTab(url, { active: false, insert: true, setParent: true });
      } catch {}
    }
    // Fallback without a visible window. The 2Use userscript also runs in
    // this frame and communicates through the shared GM storage.
    if (!popup) {
      const frame = document.createElement('iframe');
      frame.name = 'pmTwoUseBridge';
      frame.src = url;
      frame.setAttribute('aria-hidden', 'true');
      frame.style.cssText = 'position:fixed!important;left:-10000px!important;top:-10000px!important;width:1px!important;height:1px!important;opacity:0!important;pointer-events:none!important;border:0!important;';
      (document.body || document.documentElement).appendChild(frame);
      popup = {
        get closed() { return !frame.isConnected; },
        close() { frame.remove(); }
      };
    }
    twoUseConnectionActive = true;
    twoUsePopup = popup;
    lastTwoUseStartedAt = startedAt;

    let attempts = 0;
    const timer = setInterval(() => {
      if (lastTwoUseStartedAt !== startedAt) {
        clearInterval(timer);
        if (twoUsePollTimer === timer) twoUsePollTimer = null;
        return;
      }
      attempts++;
      let rawRows = '', rowsAt = 0, rowsRun = 0, renderError = '', renderHeartbeat = 0, progressRaw = '', submittedAt = 0;
      try {
        rawRows = String(GM_getValue('pmTwoUseRows', '') || '');
        rowsAt = Number(GM_getValue('pmTwoUseRowsAt', 0) || 0);
        rowsRun = Number(GM_getValue('pmTwoUseRowsRun', 0) || 0);
        renderError = String(GM_getValue('pmTwoUseRenderError', '') || '');
        renderHeartbeat = Number(GM_getValue('pmTwoUseRenderHeartbeat', 0) || 0);
        progressRaw = String(GM_getValue('pmTwoUseProgress', '') || '');
        submittedAt = Number(GM_getValue('pmTwoUseBridgeSubmittedAt', 0) || 0);
      } catch {}

      try {
        const progress = JSON.parse(progressRaw || '{}');
        if (Number(progress.runId || 0) === startedAt && Number(progress.at || 0) >= startedAt && Number(progress.expected || 0) > 0) {
          state._twoUseProgressLabel = progress.source === 'Excel' && progress.action === 'download'
            ? `vollständiger Excel-Bericht mit ${Number(progress.expected)} Paketen wird im Hintergrund geladen`
            : `${Number(progress.found || 0)} von ${Number(progress.expected)} Paketen werden eingelesen`;
          updateFooterStatus();
        }
      } catch {}

      if (rawRows && rowsRun === startedAt && rowsAt >= startedAt) {
        clearInterval(timer);
        if (twoUsePollTimer === timer) twoUsePollTimer = null;
        twoUseConnectionActive = false;
        twoUsePopup = null;
        try { popup.close(); } catch {}
        const rows = JSON.parse(rawRows);
        state._noRolloutList = Array.isArray(rows) ? rows : [];
        state._twoUseError = '';
        state._twoUseLoading = false;
        state._twoUseProgressLabel = '';
        state._footerMessage = `2Use-Daten vom ${state._twoUseRangeLabel} aktualisiert: ${state._noRolloutList.length} Pakete · ${formatDateTime(Date.now())}`;
        updateKpisOverview();
        setLastRefreshNow();
        if (state._pendingTwoUseKind) {
          const kind = state._pendingTwoUseKind;
          state._pendingTwoUseKind = '';
          showMetricList(kind, 'norollout');
        }
      } else if (renderError || attempts >= 960 || (popup.closed && rowsAt < startedAt)) {
        clearInterval(timer);
        if (twoUsePollTimer === timer) twoUsePollTimer = null;
        twoUseConnectionActive = false;
        twoUsePopup = null;
        state._twoUseError = renderError || (
          !submittedAt
            ? 'Die 2Use-Automatik wurde nicht gestartet oder das Parameterformular war nicht erreichbar.'
            : !renderHeartbeat
              ? 'Tampermonkey ist auf 2use-render-prod.dpdit.de beziehungsweise bipvmssrs1.dpdit.de nicht aktiv.'
              : 'Die 2Use-Auswertung konnte nicht vollständig geladen werden.'
        );
        state._twoUseLoading = false;
        state._twoUseProgressLabel = '';
        state._footerMessage = `2Use-Aktualisierung vom ${state._twoUseRangeLabel} fehlgeschlagen: ${state._twoUseError}`;
        updateFooterStatus();
        try { popup.close(); } catch {}
      }
    }, 250);
    twoUsePollTimer = timer;
  }

  const TP_IDB_NAME = 'fvpr_db';
  const TP_STORE = 'tourMap';
  const TP_PARTNERS_STORE = 'partners';

  let tourPartnerMap = new Map();

  function tpIdbOpen() {
    return new Promise((res, rej) => {
      const req = indexedDB.open(TP_IDB_NAME);
      req.onsuccess = () => res(req.result);
      req.onerror = () => rej(req.error);
    });
  }

  async function tpIdbAll(store) {
    try {
      const db = await tpIdbOpen();
      if (!db.objectStoreNames.contains(store)) return [];
      return new Promise((res) => {
        const r = db.transaction(store, 'readonly').objectStore(store).getAll();
        r.onsuccess = () => res(r.result || []);
        r.onerror = () => res([]);
      });
    } catch {
      return [];
    }
  }

  async function tpIdbGet(store, key) {
    try {
      const db = await tpIdbOpen();
      if (!db.objectStoreNames.contains(store)) return null;
      return new Promise((res) => {
        const r = db.transaction(store, 'readonly').objectStore(store).get(key);
        r.onsuccess = () => res(r.result || null);
        r.onerror = () => res(null);
      });
    } catch {
      return null;
    }
  }

  async function loadTourPartnerMap() {
    try {
      const rows = await tpIdbAll(TP_STORE);
      const m = new Map();
      for (const r of rows) {
        const t = tourKey(r.tour || '');
        const p = norm(r.partner || '');
        if (t && p) m.set(t, p);
      }
      tourPartnerMap = m;
    } catch {
      tourPartnerMap = new Map();
    }
  }

  async function getPartnerMailRecord(partnerName) {
    const key = norm(partnerName || '');
    if (!key) return null;
    return await tpIdbGet(TP_PARTNERS_STORE, key);
  }

  function splitEmails(raw) { return (raw || '').split(/[,;\s]+/).map(s => s.trim()).filter(Boolean); }
  function isEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); }
  function normalizeEmailList(raw) {
    const arr = splitEmails(raw);
    const valid = [], invalid = [], seen = new Set();
    for (const a of arr) {
      const low = a.toLowerCase();
      if (seen.has(low)) continue;
      seen.add(low);
      (isEmail(a) ? valid : invalid).push(a);
    }
    return { valid, invalid };
  }

  function openMailto(subject, to, cc) {
    const params = new URLSearchParams();
    if (subject) params.set('subject', subject);
    if (cc) params.set('cc', cc);
    window.location.href = `mailto:${encodeURIComponent(to || '')}?${params.toString()}`;
  }

  async function copyHtmlToClipboard(html) {
    try {
      if (!navigator.clipboard || !window.ClipboardItem) throw new Error('ClipboardItem nicht verfügbar');
      const blobHtml = new Blob([html], { type: 'text/html' });
      const blobTxt  = new Blob([html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim()], { type: 'text/plain' });
      const item = new ClipboardItem({ 'text/html': blobHtml, 'text/plain': blobTxt });
      await navigator.clipboard.write([item]);
      return true;
    } catch {
      try {
        await navigator.clipboard.writeText(html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim());
        return true;
      } catch {
        return false;
      }
    }
  }

  function ensureStyles() {
    if (document.getElementById(NS + 'style')) return;
    const style = document.createElement('style');
    style.id = NS + 'style';
    style.textContent = `
      .${NS}panel{position:fixed;top:72px;left:50%;transform:translateX(-50%);width:min(820px,95vw);max-height:78vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100000;display:none}
      .${NS}header{display:grid;grid-template-columns:1fr;gap:10px;align-items:start;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
      .${NS}toolbar{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:8px}
      .${NS}title{font:800 13px system-ui;color:#1e293b;letter-spacing:.01em}
      .${NS}group{display:flex;flex-wrap:wrap;align-items:center;gap:6px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:5px 6px}
      .${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer}
      .${NS}btn-sm.active{background:#c1121f;color:#fff;border-color:#9b0d18}
      .${NS}btn-sm.dim{opacity:.6;pointer-events:none}
      .${NS}icon-btn{display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;padding:0;font-size:17px;line-height:1;background:#fff;color:#475569}
      .${NS}icon-btn:hover{background:#eef2f7;color:#0f172a}
      .${NS}warning-btn{border-color:#f59e0b;background:linear-gradient(180deg,#fff7d6,#ffedaa);color:#8a4b00;border-radius:999px;padding:6px 12px;box-shadow:0 2px 6px rgba(217,119,6,.16)}
      .${NS}warning-btn:hover{background:linear-gradient(180deg,#ffefb4,#ffe083);border-color:#d97706;box-shadow:0 3px 8px rgba(217,119,6,.25)}
      .${NS}kpis{display:grid;grid-template-columns:112px repeat(4,minmax(86px,1fr));gap:6px;align-items:center;margin-top:4px}
      .${NS}kpi-status-head{height:28px;display:flex;align-items:center;justify-content:flex-end;padding:0 10px;color:#64748b;font:800 11px system-ui;text-transform:uppercase;letter-spacing:.06em}
      .${NS}kpi-head{padding:6px 8px;border-radius:5px;color:#fff;text-align:center;font:700 12px system-ui;white-space:nowrap}
      .${NS}kpi-head-prio{background:#8e24aa}
      .${NS}kpi-head-express{background:#ef3340}
      .${NS}kpi-head-express12{background:#d97706}
      .${NS}kpi-head-express18{background:#2563eb}
      .${NS}kpi-label{display:flex;align-items:center;justify-content:flex-start;gap:7px;min-height:28px;padding:5px 8px;border:1px solid #dbe3ec;border-left:4px solid #64748b;border-radius:5px;background:#f1f5f9;font:800 12px system-ui;color:#1e293b;white-space:nowrap;box-sizing:border-box}
      .${NS}kpi-total{display:inline-flex;align-items:center;justify-content:center;min-width:22px;height:20px;padding:0 5px;border:0;border-radius:999px;background:#64748b;color:#fff;font:800 10px system-ui;box-sizing:border-box;cursor:pointer}
      .${NS}kpi-total:hover{background:#334155;transform:scale(1.04)}
      .${NS}kpi{appearance:none;border:1px solid rgba(0,0,0,.08);background:#f1f5f9;padding:5px 8px;border-radius:5px;font:700 12px system-ui;cursor:pointer;min-height:28px}
      .${NS}kpi:hover{filter:brightness(.96);box-shadow:0 1px 3px rgba(0,0,0,.14)}
      .${NS}kpi[data-kind="prio"]{background:#ead2f0;color:#6b167c}
      .${NS}kpi[data-kind="express"]{background:#fecdd3;color:#9f1239}
      .${NS}kpi[data-kind="express12"]{background:#fde7bd;color:#92400e}
      .${NS}kpi[data-kind="express18"]{background:#dbeafe;color:#1d4ed8}
      .${NS}list{list-style:none;margin:0;padding:0}
      .${NS}empty{padding:14px 12px;opacity:.75;text-align:center;font:500 12px system-ui}
      .${NS}loading{display:none;padding:8px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:600 12px system-ui;background:#fffbe6}
      .${NS}loading.on{display:block}
      .${NS}foot{padding:8px 12px;border-top:1px solid rgba(0,0,0,.08);font:600 12px system-ui;color:#475569;background:#fafafa}
      .${NS}modal{position:fixed;inset:0;display:none;align-items:flex-start;justify-content:center;background:rgba(0,0,0,.35);z-index:100001}
      .${NS}modal-inner{background:#fff;width:min(1600px,96vw);height:min(88vh,1000px);overflow:auto;border-radius:12px;box-shadow:0 12px 28px rgba(0,0,0,.2);border:1px solid rgba(0,0,0,.12)}
      .${NS}modal-head{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui;position:sticky;top:0;background:#fff;z-index:2}
      .${NS}modal-body{padding:8px 12px;max-height:calc(100% - 46px);overflow:auto}
      .${NS}tbl{width:100%;table-layout:fixed;border-collapse:collapse;font:12px system-ui}
      .${NS}tbl th,.${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .${NS}tbl th{text-align:left;background:#fafafa;position:sticky;top:0;cursor:pointer;user-select:none;z-index:1}
      .${NS}sort-asc::after{content:" ▲";font-size:11px}
      .${NS}sort-desc::after{content:" ▼";font-size:11px}
      .${NS}eye{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border:1px solid rgba(0,0,0,.2);border-radius:6px;background:#fff;margin-right:6px;cursor:pointer;font-size:12px;line-height:1}
      .${NS}eye:hover{background:#f3f4f6}
      .${NS}badge{display:inline-block;padding:2px 6px;border-radius:999px;font-size:11px;border:1px solid rgba(0,0,0,.15);background:#f3f4f6}
      .${NS}badge-status-ok{background:#16a34a;color:#fff;border-color:#15803d}
      .${NS}badge-status-problem{background:#dc2626;color:#fff;border-color:#b91c1c}
      .${NS}badge-status-run{background:#eab308;color:#111827;border-color:#ca8a04}
      .${NS}detail-row > td{background:#f9fafb;padding:6px 8px}
      .${NS}detail-inner{border-top:1px solid rgba(0,0,0,.08);margin-top:4px;padding-top:4px}
      .${NS}detail-inner table{width:100%;border-collapse:collapse;font-size:11px}
      .${NS}detail-inner th,.${NS}detail-inner td{border-bottom:1px solid rgba(0,0,0,.06);padding:3px 4px;white-space:nowrap}
      .${NS}row-express{background:#dcfce7}
    `;
    document.head.appendChild(style);
  }

  function formatDateTime(ts) {
    if (!ts) return '—';
    const d = new Date(ts);
    return d.toLocaleString('de-DE', { hour12: false });
  }

  function setLastRefreshNow() {
    state.lastRefreshAt = Date.now();
    updateFooterStatus();
  }

  function updateFooterStatus() {
    const el = document.getElementById(NS + 'last-refresh');
    if (el) {
      const activities = [];
      if (state._dispatcherLoading) activities.push('Dispatcher-Daten werden aktualisiert');
      if (state._twoUseLoading) {
        let text = `2Use-Daten${state._twoUseRangeLabel ? ` vom ${state._twoUseRangeLabel}` : ''} werden aktualisiert`;
        if (state._twoUseProgressLabel) text += ` · ${state._twoUseProgressLabel}`;
        activities.push(text);
      }
      if (activities.length) {
        el.textContent = activities.join(' · ') + ' …';
        return;
      }
      if (state._footerMessage) {
        el.textContent = state._footerMessage;
        return;
      }
      let expected = null;
      try {
        const stored = GM_getValue('pmTwoUseExpectedCount', null);
        if (stored != null && stored !== '') expected = Number(stored);
      } catch {}
      const imported = state._noRolloutList.length;
      const twoUse = expected != null && Number.isFinite(expected)
        ? ` · 2Use: ${imported} von ${expected} eingelesen`
        : '';
      el.textContent = `Aktualisiert: ${formatDateTime(state.lastRefreshAt)}${twoUse}`;
    }
  }

  function mountUI() {
    ensureStyles();
    if (document.getElementById(NS + 'panel')) return;

    const panel = document.createElement('div');
    panel.id = NS + 'panel';
    panel.className = NS + 'panel';
    panel.innerHTML = `
      <div class="${NS}header">
        <div class="${NS}toolbar">
          <div class="${NS}title">PRIO / EXPRESS – Übersicht</div>
          <div class="${NS}group">
            <button class="${NS}btn-sm ${NS}icon-btn" data-action="openSettings" title="Einstellungen" aria-label="Einstellungen">⚙</button>
            <button class="${NS}btn-sm ${NS}icon-btn" data-action="twoUseDate" title="2Use-Datum anpassen" aria-label="2Use-Datum anpassen">📅</button>
            <button class="${NS}btn-sm" data-action="refreshApi">Aktualisieren</button>
            <button class="${NS}btn-sm ${NS}warning-btn" data-action="showExpLate11">Express 12: zu knapp / falsch einsortiert</button>
          </div>
        </div>
        <div class="${NS}kpis">
          <div class="${NS}kpi-status-head">Status</div>
          <div class="${NS}kpi-head ${NS}kpi-head-prio">Prio</div>
          <div class="${NS}kpi-head ${NS}kpi-head-express">Express</div>
          <div class="${NS}kpi-head ${NS}kpi-head-express12">Express 12</div>
          <div class="${NS}kpi-head ${NS}kpi-head-express18">Express 18</div>

          <div class="${NS}kpi-label"><button type="button" class="${NS}kpi-total" id="${NS}total-norollout" data-total-status="norollout" title="Gesamtliste öffnen">0</button><span>Kein Eingang</span></div>
          <button class="${NS}kpi" id="${NS}kpi-norollout-prio" data-kind="prio" data-status="norollout">0</button>
          <button class="${NS}kpi" id="${NS}kpi-norollout-express" data-kind="express" data-status="norollout">0</button>
          <button class="${NS}kpi" id="${NS}kpi-norollout-express12" data-kind="express12" data-status="norollout">0</button>
          <button class="${NS}kpi" id="${NS}kpi-norollout-express18" data-kind="express18" data-status="norollout">0</button>

          <div class="${NS}kpi-label"><button type="button" class="${NS}kpi-total" id="${NS}total-all" data-total-status="all" title="Gesamtliste öffnen">0</button><span>Ausrollung</span></div>
          <button class="${NS}kpi" id="${NS}kpi-all-prio" data-kind="prio" data-status="all">0</button>
          <button class="${NS}kpi" id="${NS}kpi-all-express" data-kind="express" data-status="all">0</button>
          <button class="${NS}kpi" id="${NS}kpi-all-express12" data-kind="express12" data-status="all">0</button>
          <button class="${NS}kpi" id="${NS}kpi-all-express18" data-kind="express18" data-status="all">0</button>

          <div class="${NS}kpi-label"><button type="button" class="${NS}kpi-total" id="${NS}total-delivered" data-total-status="delivered" title="Gesamtliste öffnen">0</button><span>Zugestellt</span></div>
          <button class="${NS}kpi" id="${NS}kpi-delivered-prio" data-kind="prio" data-status="delivered">0</button>
          <button class="${NS}kpi" id="${NS}kpi-delivered-express" data-kind="express" data-status="delivered">0</button>
          <button class="${NS}kpi" id="${NS}kpi-delivered-express12" data-kind="express12" data-status="delivered">0</button>
          <button class="${NS}kpi" id="${NS}kpi-delivered-express18" data-kind="express18" data-status="delivered">0</button>

          <div class="${NS}kpi-label"><button type="button" class="${NS}kpi-total" id="${NS}total-open" data-total-status="open" title="Gesamtliste öffnen">0</button><span>Offen</span></div>
          <button class="${NS}kpi" id="${NS}kpi-open-prio" data-kind="prio" data-status="open">0</button>
          <button class="${NS}kpi" id="${NS}kpi-open-express" data-kind="express" data-status="open">0</button>
          <button class="${NS}kpi" id="${NS}kpi-open-express12" data-kind="express12" data-status="open">0</button>
          <button class="${NS}kpi" id="${NS}kpi-open-express18" data-kind="express18" data-status="open">0</button>
        </div>
      </div>
      <div id="${NS}loading" class="${NS}loading">Lade Daten …</div>
      <ul id="${NS}list" class="${NS}list"></ul>
      <div class="${NS}foot" id="${NS}last-refresh">Aktualisiert: —</div>
    `;
    document.body.appendChild(panel);

    const modal = document.createElement('div');
    modal.id = NS + 'modal';
    modal.className = NS + 'modal';
    modal.innerHTML = `
      <div class="${NS}modal-inner">
        <div class="${NS}modal-head">
          <div id="${NS}modal-title">Liste</div>
          <div style="display:flex;gap:8px;align-items:center">
            <button class="${NS}btn-sm" data-action="mailSelected" style="display:none" id="${NS}mail-selected">Mail an Systempartner</button>
            <button class="${NS}btn-sm" data-action="closeModal">Schließen</button>
          </div>
        </div>
        <div class="${NS}modal-body" id="${NS}modal-body"></div>
      </div>
    `;
    document.body.appendChild(modal);

    panel.addEventListener('click', async (e) => {
      const totalBadge = e.target.closest('button[data-total-status]');
      if (totalBadge) {
        showStatusTotal(String(totalBadge.dataset.totalStatus || ''));
        return;
      }
      const chip = e.target.closest('.' + NS + 'kpi');
      if (chip) {
        showMetricList(chip.dataset.kind, chip.dataset.status);
        return;
      }

      const b = e.target.closest('.' + NS + 'btn-sm');
      if (!b) return;
      const a = b.dataset.action;
      if (a === 'openSettings') { openSettingsModal(); return; }
      if (a === 'twoUseDate') { openTwoUseDateSettings(); return; }
      if (a === 'refreshApi') {
        connectTwoUse();
        await fullRefresh(true).catch(console.error);
        return;
      }
      if (a === 'showExpLate11') { showExpLate11(); return; }
    });

    modal.addEventListener('click', e => {
      if (e.target.dataset.action === 'closeModal' || e.target === modal) { hideModal(); return; }
      const ticketAll = e.target.closest('input[data-two-use-ticket-all]');
      if (ticketAll) {
        const table = ticketAll.closest('table');
        table?.querySelectorAll('tbody input[data-two-use-ticket]').forEach(checkbox => {
          checkbox.checked = ticketAll.checked;
        });
        ticketAll.indeterminate = false;
        return;
      }
      const ticketChoice = e.target.closest('input[data-two-use-ticket]');
      if (ticketChoice) {
        const table = ticketChoice.closest('table');
        const choices = Array.from(table?.querySelectorAll('tbody input[data-two-use-ticket]') || []);
        const selected = choices.filter(checkbox => checkbox.checked).length;
        const all = table?.querySelector('thead input[data-two-use-ticket-all]');
        if (all) {
          all.checked = choices.length > 0 && selected === choices.length;
          all.indeterminate = selected > 0 && selected < choices.length;
        }
        return;
      }
      const sortButton = e.target.closest('button[data-two-use-sort]');
      if (sortButton) {
        const table = sortButton.closest('table');
        const tbody = table?.querySelector('tbody');
        if (!table || !tbody) return;
        const index = Number(sortButton.dataset.twoUseSort);
        const previousIndex = Number(table.dataset.sortIndex ?? -1);
        const direction = previousIndex === index && table.dataset.sortDirection === 'asc' ? 'desc' : 'asc';
        const collator = new Intl.Collator('de', { numeric: true, sensitivity: 'base' });
        const sortValue = row => {
          const cell = row.children[index];
          return String(cell?.dataset.sortValue ?? cell?.textContent ?? '').replace(/\s+/g, ' ').trim();
        };
        const dateValue = value => {
          const match = value.match(/^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
          return match ? Date.UTC(+match[3], +match[2] - 1, +match[1], +(match[4] || 0), +(match[5] || 0), +(match[6] || 0)) : null;
        };
        const rows = Array.from(tbody.rows);
        rows.sort((left, right) => {
          const a = sortValue(left);
          const b = sortValue(right);
          const ad = dateValue(a);
          const bd = dateValue(b);
          const compared = ad != null && bd != null ? ad - bd : collator.compare(a, b);
          return direction === 'asc' ? compared : -compared;
        });
        rows.forEach(row => tbody.appendChild(row));
        table.dataset.sortIndex = String(index);
        table.dataset.sortDirection = direction;
        table.querySelectorAll('[data-sort-indicator]').forEach(indicator => { indicator.textContent = '↕'; });
        const indicator = sortButton.querySelector('[data-sort-indicator]');
        if (indicator) indicator.textContent = direction === 'asc' ? '↑' : '↓';
        return;
      }
      const popupEye = e.target.closest('button[data-scan-popup]');
      if (popupEye) { openScanserverPopup(String(popupEye.dataset.scanPopup || '')); return; }
      const external = e.target.closest('button[data-pm-external]');
      if (external) {
        e.preventDefault();
        e.stopPropagation();
        const url = String(external.dataset.pmExternal || '');
        const currentPackage = String(external.dataset.expressTicket || '').replace(/\D+/g, '');
        if (currentPackage && /\/ops\/express_ticker\.aspx/i.test(url)) {
          const table = external.closest('table');
          const selectedPackages = Array.from(table?.querySelectorAll('tbody input[data-two-use-ticket]:checked') || [])
            .map(checkbox => String(checkbox.dataset.pknr || '').replace(/\D+/g, ''))
            .filter(Boolean);
          const packages = Array.from(new Set(selectedPackages.length ? selectedPackages : [currentPackage]));
          try {
            GM_setValue('pmExpressTickerPackages', JSON.stringify({ packages, at: Date.now() }));
          } catch {}
        }
        if (/^https:\/\//i.test(url)) window.open(url, '_blank', 'noopener,noreferrer');
        return;
      }
      const eye = e.target.closest('button.' + NS + 'eye[data-psn]');
      if (eye) { openScanserverPopup(String(eye.dataset.psn || '')); return; }
      const btn = e.target.closest('button[data-action]');
      if (!btn) return;
      const a = btn.dataset.action;
      if (a === 'guessDepot') { guessDepotFromVehicles(); return; }
      if (a === 'saveSettings') { saveSettingsFromModal(); return; }
      if (a === 'mailSelected') { mailSelectedLate11().catch(console.error); return; }
      if (a === 'connectTwoUse') { connectTwoUse(); return; }
    });

    if (!state._bootShown) {
      addEvent({
        title: 'Bereit',
        meta: 'Quelle: /dispatcher/api/pickup-delivery · Fahrer aus Fahrzeugübersicht · Systempartner aus lokaler TourMap',
        sev: 'info',
        read: true
      });
      state._bootShown = true;
    }

    render();

    // Klick außerhalb schließt Modal bzw. Hauptfenster
    document.addEventListener('mousedown', function(e) {

      const twoUseDateDialog = document.getElementById(NS + 'two-use-date-dialog');
      if (twoUseDateDialog && twoUseDateDialog.contains(e.target)) return;

      const scanPopup = document.getElementById(NS + 'scan-popup');
      if (scanPopup && scanPopup.contains(e.target)) return;

      const modal = document.getElementById(NS + 'modal');
      const modalInner = modal?.querySelector('.' + NS + 'modal-inner');

      if (modal && modal.style.display === 'flex') {
        if (modalInner && !modalInner.contains(e.target)) {
          hideModal();
        }
        return;
      }

      const panel = document.getElementById(NS + 'panel');

      if (
        panel &&
        getComputedStyle(panel).display !== 'none' &&
        !panel.contains(e.target)
      ) {
        panel.style.setProperty('display', 'none', 'important');
      }

    }, true);

  }

  function togglePanel(force) {
    const panel = document.getElementById(NS + 'panel');
    if (!panel) { mountUI(); return; }
    const isHidden = getComputedStyle(panel).display === 'none';
    const show = force != null ? !!force : isHidden;
    panel.style.setProperty('display', show ? 'block' : 'none', 'important');
  }

  const LSKEY = 'pmSettings';
  function loadSettings() {
    try { return Object.assign({ scanserverPass: '', depotSuffix: '' }, JSON.parse(localStorage.getItem(LSKEY) || '{}')); }
    catch { return { scanserverPass: '', depotSuffix: '' }; }
  }
  function saveSettingsObj(s) { try { localStorage.setItem(LSKEY, JSON.stringify(s)); } catch {} }
  function setSetting(k, v) { const s = loadSettings(); s[k] = v; saveSettingsObj(s); }
  function getSetting(k) { return loadSettings()[k]; }

  function openSettingsModal() {
    const s = loadSettings();
    const html = `
      <div style="display:grid;gap:10px;max-width:520px">
        <label style="display:grid;gap:6px;font:600 12px system-ui">
          Scanserver-Passwort
          <input id="${NS}inp-pass" type="password" placeholder="••••••••" value="${esc(s.scanserverPass || '')}"
                 style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px"/>
        </label>
        <label style="display:grid;gap:6px;font:600 12px system-ui">
          Depotkennung (3-stellig, z. B. 157)
          <div style="display:flex;gap:8px;align-items:center">
            <input id="${NS}inp-depot" type="text" pattern="\\d{3}" maxlength="3" placeholder="157" value="${esc(String(s.depotSuffix || '').slice(-3))}"
                   style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px;width:100px;text-align:center;font-weight:700;letter-spacing:.5px"/>
            <button class="${NS}btn-sm" data-action="guessDepot">Auto erkennen</button>
          </div>
          <div style="opacity:.7;font:12px system-ui">Host: <code>scanserver-d0010<strong>${esc(String(s.depotSuffix || '').slice(-3) || '157')}</strong>.ssw.dpdit.de</code></div>
        </label>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button class="${NS}btn-sm" data-action="saveSettings">Speichern</button>
        </div>
      </div>`;
    openModal('Einstellungen', html);
  }

  function findVehicleGridContainer() {
    const cands = Array.from(document.querySelectorAll(
      '.MuiDataGrid-virtualScroller, .MuiDataGrid-main, [data-rttable="true"], [role="grid"], table'
    )).filter(el => el.offsetParent !== null);
    if (!cands.length) return null;
    const scrollable = cands.filter(el => {
      const cs = getComputedStyle(el);
      const ov = cs.overflowY || cs.overflow;
      return /auto|scroll/i.test(ov || '') && el.scrollHeight > el.clientHeight;
    });
    const pool = scrollable.length ? scrollable : cands;
    return pool.reduce((best, el) => ((el.scrollHeight || 0) > ((best?.scrollHeight) || 0) ? el : best), null);
  }

  function guessDepotSuffixFromVehicleTable(root) {
    const grid = root || findVehicleGridContainer();
    if (!grid) return '';
    const counts = new Map();

    const pick = (t) => {
      t = String(t || '');
      let m = t.match(/d0010(\d{3})/i) || t.match(/0010(\d{3})/) || t.match(/010(\d{3})/);
      if (m) return m[1];
      m = t.match(/\b(\d{3})\b/);
      if (m) return m[1];
      m = t.match(/0{0,2}(\d{3})\b/);
      return m ? m[1] : '';
    };

    const ths = Array.from(grid.querySelectorAll('thead th,[role="columnheader"]'));
    const iDepot = ths.findIndex(th => /^Depot$/i.test((th.textContent || th.title || '').trim()));
    if (iDepot < 0) return '';

    const rows = Array.from(grid.querySelectorAll('tbody tr,[role="row"]'))
      .filter(r => r.querySelector('td,[role="gridcell"]'))
      .slice(0, 800);

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const raw = (tds[iDepot]?.getAttribute?.('aria-label') ||
                   tds[iDepot]?.getAttribute?.('data-title') ||
                   tds[iDepot]?.querySelector?.('[title]')?.getAttribute('title') ||
                   tds[iDepot]?.innerText || tds[iDepot]?.textContent || '').trim();
      if (!raw) continue;
      const suf = pick(raw);
      if (suf && suf !== '000') counts.set(suf, (counts.get(suf) || 0) + 1);
    }

    if (!counts.size) return '';
    let best = '', bestN = -1;
    for (const [k, n] of counts) {
      if (n > bestN) { best = k; bestN = n; }
    }
    return best;
  }

  function guessDepotFromVehicles() {
    const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
    if (g) {
      const inp = document.getElementById(NS + 'inp-depot');
      if (inp) inp.value = g;
      addEvent({ title: 'Einstellungen', meta: `Depotkennung erkannt: ${g}`, sev: 'info', read: true });
    } else {
      addEvent({ title: 'Einstellungen', meta: 'Depotkennung konnte nicht ermittelt werden.', sev: 'warn', read: true });
    }
  }

  function saveSettingsFromModal() {
    const pass = (document.getElementById(NS + 'inp-pass')?.value || '');
    const dep  = (document.getElementById(NS + 'inp-depot')?.value || '').replace(/\D+/g, '').slice(-3);
    saveSettingsObj({ ...loadSettings(), scanserverPass: pass, depotSuffix: dep });
    addEvent({ title: 'Einstellungen', meta: 'Gespeichert.', sev: 'info', read: true });
    hideModal();
  }

  function getScanserverBase() {
    let suf = String(getSetting('depotSuffix') || '').replace(/\D+/g, '').slice(-3);
    if (!suf) {
      const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
      if (g) { suf = g; setSetting('depotSuffix', g); }
    }
    if (!suf) suf = '157';
    return `https://scanserver-d0010${suf}.ssw.dpdit.de/cgi-bin/pa.cgi`;
  }

  function buildScanserverUrl(psnRaw) {
    const pass = getSetting('scanserverPass') || '';
    if (!pass) {
      addEvent({ title: 'Scanserver', meta: 'Kein Passwort hinterlegt. Bitte in den Einstellungen setzen.', sev: 'warn', read: true });
      openSettingsModal();
      return '';
    }
    let psn = String(psnRaw || '').replace(/\D+/g, '');
    if (psn.length === 13) psn = '0' + psn;

    const base = getScanserverBase();
    const params = new URLSearchParams();
    params.set('_url', 'file');
    params.set('_passwd', pass);
    params.set('_disp', '3');
    params.set('_pivotxx', '0');
    params.set('_rastert', '4');
    params.set('_rasteryt', '0');
    params.set('_rasterx', '0');
    params.set('_rastery', '0');
    params.set('_pivot', '0');
    params.set('_pivotbp', '0');
    params.set('_sortby', 'date|time');
    params.set('_dca', '0');
    params.set('_tabledef', 'psn|date|time|sa|tour|zc|sc|adr1|str|hno|plz1|city|dc|etafrom|etato');
    params.set('_arg59', 'dpd');
    params.set('_arg0a', psn);
    params.set('_arg0b', psn);
    params.set('_arg0', psn + ',' + psn);
    params.set('_csv', '0');
    return `${base}?${params.toString()}`;
  }

  function openScanserverPopup(psn) {
    const url = buildScanserverUrl(psn);
    if (!url) return;
    document.getElementById(NS + 'scan-popup')?.remove();
    const overlay = document.createElement('div');
    overlay.id = NS + 'scan-popup';
    overlay.style.cssText = 'position:fixed;inset:0;z-index:100020;background:rgba(15,23,42,.42);display:flex;align-items:center;justify-content:center;padding:20px;box-sizing:border-box';
    overlay.innerHTML = `
      <div data-scan-popup-inner style="width:min(1050px,94vw);height:min(430px,78vh);background:#fff;border:1px solid #94a3b8;border-radius:10px;box-shadow:0 20px 60px rgba(0,0,0,.3);overflow:hidden;display:flex;flex-direction:column;resize:both">
        <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;padding:8px 12px;border-bottom:1px solid #dbe2ea;background:#f8fafc;font:700 13px system-ui">
          <span>Lokale Paketauskunft · ${esc(String(psn || ''))}</span>
          <button type="button" data-scan-close style="padding:5px 9px;border:1px solid #cbd5e1;border-radius:6px;background:#fff;cursor:pointer">Schließen</button>
        </div>
        <iframe src="${esc(url)}" title="Lokale Paketauskunft" style="width:100%;height:100%;border:0;background:#fff" referrerpolicy="no-referrer"></iframe>
      </div>`;
    overlay.addEventListener('mousedown', event => {
      if (event.target === overlay) overlay.remove();
    });
    overlay.querySelector('[data-scan-close]').addEventListener('click', () => overlay.remove());
    document.body.appendChild(overlay);
    const frame = overlay.querySelector('iframe');
    const inner = overlay.querySelector('[data-scan-popup-inner]');
    frame.addEventListener('load', () => {
      try {
        const doc = frame.contentDocument;
        const width = Math.min(1200, Math.max(700, (doc.documentElement.scrollWidth || doc.body.scrollWidth) + 28));
        const height = Math.min(700, Math.max(220, (doc.documentElement.scrollHeight || doc.body.scrollHeight) + 48));
        inner.style.width = `min(${width}px,94vw)`;
        inner.style.height = `min(${height}px,78vh)`;
      } catch {}
    });
  }

  function storeCapturedPickupRequest(urlString, headersMaybe) {
    try {
      const u = new URL(urlString, location.origin);
      if (!u.href.includes('/dispatcher/api/pickup-delivery')) return;

      const q = u.searchParams;
      if (q.get('parcelNumber')) return;

      const h = {};
      const src = headersMaybe || {};

      if (src instanceof Headers) {
        src.forEach((v, k) => h[String(k).toLowerCase()] = String(v));
      } else if (Array.isArray(src)) {
        src.forEach(([k, v]) => h[String(k).toLowerCase()] = String(v));
      } else if (src && typeof src === 'object') {
        Object.entries(src).forEach(([k, v]) => h[String(k).toLowerCase()] = String(v));
      }

      if (!h['authorization']) {
        const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
        if (m) h['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
      }

      lastOkRequest = { url: u, headers: h };

      if (Date.now() > suppressSourceRefreshUntil && !isBusy) {
        queueRefreshFromSource();
      }
    } catch {}
  }

  function queueRefreshFromSource() {
    if (sourceRefreshTimer) clearTimeout(sourceRefreshTimer);
    sourceRefreshTimer = setTimeout(() => {
      sourceRefreshTimer = null;
      if (!document.hidden) {
        fullRefresh(false).catch(console.error);
        if (Date.now() - lastTwoUseStartedAt > 30000) connectTwoUse(true);
      }
    }, 700);
  }

  (function hookNetwork() {
    if (!window.__pm_fetch_hooked && window.fetch) {
      const orig = window.fetch;
      window.fetch = async function (input, init = {}) {
        const res = await orig.apply(this, arguments);
        try {
          const uStr = typeof input === 'string' ? input : (input && input.url) || '';
          if (res.ok) storeCapturedPickupRequest(uStr, init?.headers || input?.headers);
        } catch {}
        return res;
      };
      window.__pm_fetch_hooked = true;
    }

    if (!window.__pm_xhr_hooked && window.XMLHttpRequest) {
      const X = window.XMLHttpRequest;
      const open = X.prototype.open, send = X.prototype.send, setH = X.prototype.setRequestHeader;

      X.prototype.open = function (m, u) {
        this.__pm_url = typeof u === 'string' ? new URL(u, location.origin) : null;
        this.__pm_headers = {};
        return open.apply(this, arguments);
      };

      X.prototype.setRequestHeader = function (k, v) {
        try { this.__pm_headers[String(k).toLowerCase()] = String(v); } catch {}
        return setH.apply(this, arguments);
      };

      X.prototype.send = function () {
        const onload = () => {
          try {
            if (this.__pm_url && this.status >= 200 && this.status < 300) {
              storeCapturedPickupRequest(this.__pm_url.href, this.__pm_headers);
            }
          } catch {}
          this.removeEventListener('load', onload);
        };
        this.addEventListener('load', onload);
        return send.apply(this, arguments);
      };

      window.__pm_xhr_hooked = true;
    }
  })();

  const gridIndex = { tour2driver: new Map() };
  const deliveryDetailsCache = new Map();

  function detectTourDriverCols(tbl) {
    const ths = Array.from(tbl.querySelectorAll('thead th,[role="columnheader"]'))
      .map(el => ({ el, txt: norm(el.textContent || el.title || '') }));
    let iTour = -1, iDrv = -1;
    ths.forEach((h, i) => {
      if (iTour < 0 && /\bTour(\s*nr|nummer)?\b/i.test(h.txt)) iTour = i;
      if (iDrv < 0 && /(Zusteller(\s*name)?|Fahrer)/i.test(h.txt)) iDrv = i;
    });
    return { iTour, iDrv };
  }

  function collectTourDriverFromTable(tbl, map) {
    const { iTour, iDrv } = detectTourDriverCols(tbl);
    if (iTour < 0 || iDrv < 0) return 0;
    const rows = Array.from(tbl.querySelectorAll('tbody tr,[role="row"]'))
      .filter(r => r.querySelector('td,[role="gridcell"]'));
    let added = 0;

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const get = i => norm(
        tds[i]?.getAttribute?.('aria-label') ||
        tds[i]?.getAttribute?.('data-title') ||
        tds[i]?.querySelector?.('[title]')?.getAttribute('title') ||
        tds[i]?.innerText || tds[i]?.textContent || ''
      );
      const tour = tourKey(get(iTour));
      const drv  = get(iDrv);
      if (tour && drv && !map.has(tour)) { map.set(tour, drv); added++; }
    }
    return added;
  }

  async function buildTourDriverMap() {
    try {
      const map = new Map();
      Array.from(document.querySelectorAll('table,[role="grid"]'))
        .filter(el => el.offsetParent !== null)
        .forEach(tbl => collectTourDriverFromTable(tbl, map));
      if (map.size) {
        gridIndex.tour2driver = map;
        window.__pmTour2Driver = map;
      }
    } catch {}
  }

  function getJwtAuthHeader() {
    const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
    return m ? 'Bearer ' + decodeURIComponent(m[1]) : '';
  }

  function buildHeaders(h) {
    const H = new Headers();
    try {
      if (h) {
        Object.entries(h).forEach(([k, v]) => {
          const key = k.toLowerCase();
          if (['authorization', 'accept', 'x-xsrf-token', 'x-csrf-token'].includes(key)) {
            H.set(key === 'accept' ? 'Accept' : key.replace(/(^.|-.)/g, s => s.toUpperCase()), v);
          }
        });
      }
      if (!H.has('Authorization')) {
        const auth = getJwtAuthHeader();
        if (auth) H.set('Authorization', auth);
      }
      if (!H.has('Accept')) H.set('Accept', 'application/json, text/plain, */*');
    } catch {}
    return H;
  }

  function buildBasePickupUrlFromToday() {
    const d = new Date();
    const ds = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const u = new URL(location.origin + '/dispatcher/api/pickup-delivery');
    u.searchParams.set('page', '1');
    u.searchParams.set('pageSize', '500');
    u.searchParams.set('dateFrom', ds);
    u.searchParams.set('dateTo', ds);
    u.searchParams.set('active', 'true');
    return u;
  }

  function buildUrlPrio(base, page) {
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize', '500');
    q.set('priority', 'prio');
    q.delete('elements');
    q.delete('parcelNumber');
    return u;
  }

  function buildUrlElements(base, page, el) {
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize', '500');
    q.set('elements', String(el));
    q.delete('priority');
    q.delete('parcelNumber');
    return u;
  }

  function pickArray(payload) {
    if (Array.isArray(payload)) return payload;
    if (payload && Array.isArray(payload.items)) return payload.items;
    if (payload && Array.isArray(payload.content)) return payload.content;
    if (payload && payload.data) {
      if (Array.isArray(payload.data)) return payload.data;
      if (Array.isArray(payload.data.items)) return payload.data.items;
      if (Array.isArray(payload.data.content)) return payload.data.content;
    }
    if (payload && Array.isArray(payload.results)) return payload.results;
    if (payload && payload._embedded) {
      const v = Object.values(payload._embedded).find(Array.isArray);
      if (Array.isArray(v)) return v;
    }
    return [];
  }

  function pickTotals(payload, sizeFallback) {
    const p = payload || {};
    const pg = p.page || {};
    const totalElements = Number(p.totalElements ?? p.total ?? p.count ?? pg.totalElements ?? pg.total ?? 0);
    const totalPages = Number(p.totalPages ?? pg.totalPages ?? (totalElements ? Math.ceil(totalElements / (sizeFallback || 500)) : 0));
    return {
      totalElements: Number.isFinite(totalElements) ? totalElements : 0,
      totalPages: Number.isFinite(totalPages) ? totalPages : 0
    };
  }

  async function fetchPagedFast(builder, { concurrency = 6, size = 500, hardMaxPages = 200 } = {}) {
    const baseUrl = lastOkRequest?.url ? new URL(lastOkRequest.url.href) : buildBasePickupUrlFromToday();
    const headers = buildHeaders(lastOkRequest?.headers || {});
    suppressSourceRefreshUntil = Date.now() + 5000;

    const u1 = builder(baseUrl, 1);
    const r1 = await fetch(u1.toString(), { credentials: 'include', headers });
    if (!r1.ok) throw new Error(`HTTP ${r1.status}`);
    const j1 = await r1.json();
    const chunk1 = pickArray(j1);

    const { totalPages: tpRaw } = pickTotals(j1, size);
    let totalPages = tpRaw || 0;

    if (totalPages > 1) {
      totalPages = Math.min(totalPages, hardMaxPages);
      const pages = [];
      for (let p = 2; p <= totalPages; p++) pages.push(p);

      const out = chunk1.slice();
      let idx = 0;

      async function worker() {
        while (idx < pages.length) {
          const p = pages[idx++];
          const u = builder(baseUrl, p);
          const res = await fetch(u.toString(), { credentials: 'include', headers });
          if (!res.ok) continue;
          const j = await res.json();
          const arr = pickArray(j);
          if (arr.length) out.push(...arr);
          if (arr.length < size) { idx = pages.length; break; }
        }
      }

      const workers = Array.from({ length: Math.max(1, Math.min(concurrency, pages.length)) }, worker);
      await Promise.all(workers);
      return out;
    }

    if (chunk1.length < size) return chunk1;

    const out = chunk1.slice();
    let page = 2;
    while (page <= hardMaxPages) {
      const u = builder(baseUrl, page);
      const r = await fetch(u.toString(), { credentials: 'include', headers });
      if (!r.ok) break;
      const j = await r.json();
      const arr = pickArray(j);
      if (!arr.length) break;
      out.push(...arr);
      if (arr.length < size) break;
      page++;
    }
    return out;
  }

  const parcelId = r => r.__pidOverride || r.parcelNumber || (Array.isArray(r.parcelNumbers) && r.parcelNumbers[0]) || r.id || '';
  const addrOf = r => [r.street, r.houseno].filter(Boolean).join(' ');
  const placeOf = r => [r.postalCode, r.city].filter(Boolean).join(' ');
  const addCodes = r => Array.isArray(r.additionalCodes) ? r.additionalCodes.map(String) : [];
  const isDelivery = r => String(r?.orderType || '').toUpperCase() === 'DELIVERY';
  const isPRIO = r => String(r?.priority || r?.prio || '').toUpperCase() === 'PRIO';

  const composeDateTime = (dateStr, timeStr) => {
    if (!dateStr || !timeStr) return null;
    const s = `${dateStr}T${String(timeStr).slice(0, 8)}`;
    const d = new Date(s);
    return isNaN(d) ? null : d;
  };

  const fromTime = r => r.from2 ? composeDateTime(r.date, r.from2) : null;
  const toTime = r => r.to2 ? composeDateTime(r.date, r.to2) : null;
  const deliveredTime = r => r.deliveredTime ? new Date(r.deliveredTime) : null;

  function apiStatus(r) {
    const v =
      r?.statusDisplay || r?.statusLabel || r?.statusDescription ||
      r?.statusName || r?.statusText || r?.stateText ||
      r?.deliveryStatus || r?.parcelStatus || '';
    return String(v || '').trim();
  }

  const statusOf = r => apiStatus(r);

  const isDeliveryProblemForceOpen = r => {
    const s = (statusOf(r) || '').toUpperCase();
    if (!/DELIVERY_PROBLEM/.test(s)) return false;
    const codes = addCodes(r);
    return codes.includes('041') || codes.includes('061');
  };

  const delivered = r => {
    if (isDeliveryProblemForceOpen(r)) return false;
    if (r.deliveredTime) return true;
    const s = (statusOf(r) || '').toUpperCase();
    return /ZUGESTELLT|DELIVERED/.test(s);
  };

  const tourOf = r => r.tour ? String(r.tour) : '';

  function driverOf(r) {
    const direct = r.driverName || r.driver || r.courierName || r.riderName || r.tourDriver || '';
    if (direct && direct.trim()) return direct.trim();
    const key = tourKey(tourOf(r) || '');
    const viaGrid =
      (gridIndex.tour2driver && gridIndex.tour2driver.get(key)) ||
      (window.__pmTour2Driver instanceof Map ? window.__pmTour2Driver.get(key) : '');
    return (viaGrid || '—').trim() || '—';
  }

  function partnerOfTour(tour) {
    const k = tourKey(tour || '');
    return tourPartnerMap.get(k) || '—';
  }

  function formatHHMM(d) {
    return d ? d.toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }) : '';
  }

  function buildPredictMeta(r) {
    const pf = fromTime(r);
    const pt = toTime(r);
    const pfTs = pf ? +pf : 0;
    const ptTs = pt ? +pt : 0;
    let range = '—';
    if (pf && pt) range = `${formatHHMM(pf)} - ${formatHHMM(pt)}`;
    else if (pf) range = formatHHMM(pf);
    else if (pt) range = formatHHMM(pt);
    return { pfTs, ptTs, range };
  }

  function serviceCodesOf(r) {
    if (!r) return [];
    const set = new Set();

    const addFromVal = v => {
      if (v == null) return;
      String(v)
        .split(/[^\dA-Za-z]+/)
        .map(s => s.trim())
        .filter(Boolean)
        .forEach(code => set.add(code));
    };

    const addFromArr = arr => {
      if (!Array.isArray(arr)) return;
      arr.forEach(addFromVal);
    };

    addFromVal(r.serviceCode);
    addFromVal(r.servicecode);
    addFromVal(r.service_code);
    addFromArr(r.serviceCodes);

    if (r.service && typeof r.service === 'object') {
      addFromVal(r.service.code);
      addFromVal(r.service.serviceCode);
      addFromVal(r.service.id);
      addFromArr(r.service.serviceCodes);
    }

    if (r.product && typeof r.product === 'object') {
      addFromVal(r.product.serviceCode);
      addFromVal(r.product.code);
      addFromVal(r.product.id);
      addFromArr(r.product.serviceCodes);
    }

    const arr = Array.from(set);
    arr.sort((a, b) => collator.compare(a, b));
    return arr;
  }

  const EXPRESS12_CODES = new Set([
    '104','107','135','196','210','225','226','227','228','231','232','234','237','238','239','240',
    '243','245','247','249','255','261','262','267','269','286','310','311','323','379','412','414',
    '452','453','458','459','488','490','503','505','530','531','537','538','542','547','567',
    '786','797','811'
  ]);

  const EXPRESS18_CODES = new Set([
    '155','157','158','161','163','164','166','168','171','174',
    '219','224','230','236','265','318','324','378','419','422','423',
    '483','499','500','511','512','535','562','564','662','664',
    '787','799','812'
  ]);

  const EXPRESS_SERVICE_WHITELIST = new Set([...EXPRESS12_CODES, ...EXPRESS18_CODES]);

  function rowHasExpress12BySvc(r) {
    const arr = r.__serviceCodes || serviceCodesOf(r);
    return arr.some(c => EXPRESS12_CODES.has(String(c)));
  }

  function rowHasExpress18BySvc(r) {
    const arr = r.__serviceCodes || serviceCodesOf(r);
    return arr.some(c => EXPRESS18_CODES.has(String(c)));
  }

  function statusClass(text) {
    const t = String(text || '').toUpperCase();
    if (/PROBLEM|FAIL|NICHT/.test(t)) return NS + 'badge-status-problem';
    if (/ZUGESTELLT|DELIVERED/.test(t)) return NS + 'badge-status-ok';
    if (/ZUSTELLUNG|OUT_FOR_DELIVERY|IN_DELIVERY/.test(t)) return NS + 'badge-status-run';
    return '';
  }

  function normRow(r) {
    const pid = parcelId(r) || '';
    const { pfTs, ptTs, range } = buildPredictMeta(r);
    const svcArr = serviceCodesOf(r);
    const isExpSvc = svcArr.some(c => EXPRESS_SERVICE_WHITELIST.has(String(c)));
    const isExp12 = svcArr.some(c => EXPRESS12_CODES.has(String(c)));
    const isExp18 = !isExp12 && svcArr.some(c => EXPRESS18_CODES.has(String(c)));
    const expType = isExp12 ? '12' : (isExp18 ? '18' : '');
    const isDel = delivered(r);

    let highlightLate12 = false;
    if (isExp12 && !isDel && (pfTs || ptTs) && r.date) {
      const cut = new Date(`${r.date}T12:00:00`);
      if (!isNaN(cut)) {
        const cutTs = +cut;
        if ((pfTs && pfTs > cutTs) || (ptTs && ptTs > cutTs)) highlightLate12 = true;
      }
    }

    const t = tourOf(r) || '';
    const sysPartner = partnerOfTour(t);

    return {
      ...r,
      __pid: pid,
      __addr: [addrOf(r), placeOf(r)].filter(Boolean).join(' · ') || '—',
      __driver: driverOf(r),
      __tourNum: Number(tourOf(r) || 0),
      __systempartner: sysPartner,
      __status: statusOf(r) || '',
      __delivTs: deliveredTime(r) ? deliveredTime(r).getTime() : 0,
      __predFromTs: pfTs,
      __predToTs: ptTs,
      __predRangeStr: range,
      __codesStr: (addCodes(r) || []).join(', ') || '—',
      __expType: expType,
      __serviceCode: svcArr[0] || '',
      __serviceCodes: svcArr,
      __highlightLatePredict12: highlightLate12,
      __isExpressSvc: isExpSvc
    };
  }

  function expandAndNorm(rows) {
    const out = [];
    for (const r of rows) {
      const list = Array.isArray(r.parcelNumbers) && r.parcelNumbers.length ? r.parcelNumbers : [parcelId(r)];
      const seen = new Set();
      for (const raw of list) {
        const psn = String(raw || '').replace(/\D+/g, '');
        if (!psn || seen.has(psn)) continue;
        seen.add(psn);
        const rr = { ...r, __pidOverride: psn.length === 13 ? '0' + psn : psn };
        out.push(normRow(rr));
      }
    }
    return out;
  }

  function buildHeaderHtml(selectable = false) {
    const base = ['Paketscheinnummer','Adresse','Fahrer','Tour','Systempartner','Status','Zustellzeit','Zusatzcode','Servicecode','Predict'];
    const ths = selectable ? ['✓', ...base] : base;
    return `<tr>${ths.map((h, i) => `<th data-col="${i}">${h}</th>`).join('')}</tr>`;
  }

  function buildColgroupHtml(selectable = false) {
    // Feste, prozentuale Spaltenbreiten zusammen mit table-layout:fixed
    // verhindern, dass die Modal-Tabelle auf FHD-Monitoren über die
    // Fensterbreite hinauswächst und einen horizontalen Scrollbalken erzwingt.
    const widths = selectable
      ? [3, 12, 20, 12, 7, 9, 8, 8, 7, 7, 7]
      : [13, 20, 12, 7, 9, 8, 8, 8, 8, 7];
    return `<colgroup>${widths.map(w => `<col style="width:${w}%">`).join('')}</colgroup>`;
  }

  function buildTableShell(selectable = false) {
    return `
      <div id="${NS}vt-wrap" style="position:relative;height:min(70vh,720px);overflow:auto">
        <table class="${NS}tbl">
          ${buildColgroupHtml(selectable)}
          <thead>${buildHeaderHtml(selectable)}</thead>
          <tbody id="${NS}vt-body"></tbody>
        </table>
      </div>`;
  }

  function rowHtml(r, selectable = false) {
    const selKey = String(r.stopId ?? r.id ?? r.__pid ?? '');
    const checked = selectable && state._modal.selected?.has(selKey) ? 'checked' : '';
    const selCell = selectable ? `<input type="checkbox" data-sel="1" data-key="${esc(selKey)}" ${checked} />` : null;
    const pkgCount = Number(r.__pkgCount || 1);
    const psnLabel = (r.__pid && pkgCount > 1) ? `${r.__pid} (+${pkgCount - 1})` : (r.__pid || '—');

    const pLink = r.__pid
      ? `<a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${r.__pid}" target="_blank" rel="noopener">${esc(psnLabel)}</a>`
      : '—';

    const eye = r.__pid ? `<button class="${NS}eye" title="Lokale Paketauskunft als Popup öffnen" data-psn="${esc(r.__pid)}">👁</button>` : '';
    const dtime = r.__delivTs ? formatHHMM(new Date(r.__delivTs)) : '—';

    const statusText = r.__status || '';
    const statusCls = statusClass(statusText);
    const statusCell = statusText ? `<span class="${NS}badge ${statusCls}">${esc(statusText)}</span>` : '';

    const serviceBadges = (r.__serviceCodes && r.__serviceCodes.length)
      ? r.__serviceCodes.map(c => `<span class="${NS}badge">${esc(c)}</span>`).join(' ')
      : '';

    const pred = r.__predRangeStr || '—';

    const cells = [
      `${eye}${pLink}`,
      esc(r.__addr),
      esc(r.__driver),
      esc(String(r.__tourNum || '—')),
      esc(String(r.__systempartner || '—')),
      statusCell,
      esc(dtime),
      esc(r.__codesStr),
      serviceBadges,
      esc(pred)
    ];

    // Titles fürs Mouseover: bei gestauchten Spalten (table-layout:fixed +
    // Ellipsis) soll der volle Text per Tooltip sichtbar bleiben.
    const cellTitles = [
      r.__pid || '',
      r.__addr || '',
      r.__driver || '',
      String(r.__tourNum || ''),
      String(r.__systempartner || ''),
      statusText,
      dtime,
      r.__codesStr || '',
      (r.__serviceCodes || []).join(' '),
      pred
    ];

    const finalCells = selectable ? [selCell, ...cells] : cells;
    const finalTitles = selectable ? ['', ...cellTitles] : cellTitles;
    const dataAttrs = [];
    if (r.id != null) dataAttrs.push(`data-delivery="${esc(String(r.id))}"`);
    if (r.stopId != null) dataAttrs.push(`data-stop="${esc(String(r.stopId))}"`);

    return `<tr ${dataAttrs.join(' ')}>${finalCells.map((v, i) => `<td${finalTitles[i] ? ` title="${esc(finalTitles[i])}"` : ''}>${v}</td>`).join('')}</tr>`;
  }

  function openModal(title, rowsOrHtml) {
    const m = document.getElementById(NS + 'modal');
    const t = document.getElementById(NS + 'modal-title');
    const b = document.getElementById(NS + 'modal-body');
    if (t) t.textContent = title || '';

    if (Array.isArray(rowsOrHtml)) {
      const rows = rowsOrHtml.slice();
      const selectable = /falsch einsortiert/i.test(title || '');

      state._modal = {
        rows,
        opts: { showPredict: true, selectable },
        title: title || '',
        selected: state._modal.selected instanceof Set ? state._modal.selected : new Set()
      };

      if (b) b.innerHTML = buildTableShell(selectable);

      const tbody = document.getElementById(NS + 'vt-body');
      const wrap = document.getElementById(NS + 'vt-wrap');

      function renderAll() {
        if (!tbody) return;
        tbody.innerHTML = rows.map(r => rowHtml(r, selectable)).join('');
      }

      const mailBtn = document.getElementById(NS + 'mail-selected');
      if (mailBtn) mailBtn.style.display = selectable ? '' : 'none';

      if (tbody && selectable) {
        tbody.addEventListener('change', ev => {
          const cb = ev.target.closest('input[type="checkbox"][data-sel="1"]');
          if (!cb) return;
          const key = String(cb.dataset.key || '');
          if (!key) return;
          if (cb.checked) state._modal.selected.add(key);
          else state._modal.selected.delete(key);
        }, { passive: true });
      }

      const thead = wrap?.querySelector('thead');
      if (thead) {
        thead.addEventListener('click', ev => {
          const th = ev.target.closest('th');
          if (!th) return;
          const col = Number(th.dataset.col || 0);

          Array.from(thead.querySelectorAll('th')).forEach(x => x.classList.remove(NS + 'sort-asc', NS + 'sort-desc'));
          const asc = !(th.dataset.dir === 'asc');
          th.dataset.dir = asc ? 'asc' : 'desc';
          th.classList.add(asc ? NS + 'sort-asc' : NS + 'sort-desc');

          const offset = selectable ? 1 : 0;
          const getKey = r => {
            switch (col - offset) {
              case 0: return r.__pid;
              case 1: return r.__addr;
              case 2: return r.__driver;
              case 3: return r.__tourNum;
              case 4: return r.__systempartner || '';
              case 5: return r.__status || '';
              case 6: return r.__delivTs;
              case 7: return r.__codesStr;
              case 8: return r.__serviceCode || '';
              case 9: return r.__predFromTs || 0;
              default: return '';
            }
          };

          rows.sort((a, b) => {
            const A = getKey(a), B = getKey(b);
            if (typeof A === 'number' && typeof B === 'number') return asc ? (A - B) : (B - A);
            return asc ? collator.compare(String(A), String(B)) : collator.compare(String(B), String(A));
          });

          state._modal.rows = rows;
          renderAll();
        }, { passive: true });
      }

      renderAll();

      if (tbody) {
        tbody.addEventListener('click', ev => {
          const tr = ev.target.closest('tr');
          if (!tr) return;
          if (ev.target.closest('.' + NS + 'eye')) return;
          toggleStopDetailInline(tr);
        }, { passive: true });
      }
    } else {
      if (b) b.innerHTML = rowsOrHtml || '';
      const mailBtn = document.getElementById(NS + 'mail-selected');
      if (mailBtn) mailBtn.style.display = 'none';
    }

    if (m) m.style.display = 'flex';
  }

  function hideModal() {
    const m = document.getElementById(NS + 'modal');
    if (m) m.style.display = 'none';
  }

  function setLoading(on) {
    isLoading = !!on;
    const el = document.getElementById(NS + 'loading');
    if (el) el.classList.toggle('on', on);
  }

  function dimButtons(on) {
    document.querySelectorAll('.' + NS + 'btn-sm[data-action="refreshApi"]').forEach(b => b.classList.toggle(NS + 'dim', !!on));
  }

  function render() {
    const list = document.getElementById(NS + 'list');
    if (!list) return;
    list.innerHTML = '';
    const d = document.createElement('div');
    d.className = NS + 'empty';
    d.textContent = isLoading ? 'Lade Daten …' : 'Daten geladen.';
    list.appendChild(d);
    updateFooterStatus();
  }

  function setKpis(values) {
    for (const [key, value] of Object.entries(values || {})) {
      const el = document.getElementById(NS + 'kpi-' + key);
      if (el) el.textContent = String(Number(value || 0));
    }
  }

  function addEvent(ev) {
    state.events.push({
      id: state.nextId++,
      title: ev.title || 'Ereignis',
      meta: ev.meta || '',
      sev: ev.sev || 'info',
      ts: ev.ts || Date.now(),
      read: !!ev.read
    });
  }

  function buildTableRowsAndCounts(prioRows, exp12Rows, exp18Rows) {
    const prioDeliveries = prioRows.filter(isDelivery).filter(isPRIO);
    const prioAll = prioDeliveries;
    const prioOpen = prioDeliveries.filter(r => !delivered(r));

    const expRows = [...exp12Rows, ...exp18Rows].filter(r => r.__isExpressSvc);

    const seen = new Set();
    const expDeliveries = expRows
      .filter(isDelivery)
      .filter(r => {
        const id = parcelId(r);
        if (!id || seen.has(id)) return false;
        seen.add(id);
        return true;
      });

    const expAll = expDeliveries;
    const expOpen = expDeliveries.filter(r => !delivered(r));

    const expLate11 = expDeliveries
      .filter(rowHasExpress12BySvc)
      .filter(r => {
        const ft = fromTime(r);
        if (!ft) return false;
        return (ft.getHours() > 11) || (ft.getHours() === 11 && ft.getMinutes() >= 1);
      });

    return { prioAll, prioOpen, expAll, expOpen, expLate11 };
  }

  function metricRows(kind, status) {
    if (status === 'norollout') {
      return state._noRolloutList.filter(row => {
        const product = norm(row.product).toLowerCase();
        if (kind === 'prio') return /priority|prio/.test(product);
        if (kind === 'express12') return /12:00/.test(product);
        if (kind === 'express18') return /18:00/.test(product);
        return !/priority|prio/.test(product);
      });
    }

    let rows;
    if (kind === 'prio') rows = state._prioAllList.slice();
    else if (kind === 'express12') rows = state._expAllList.filter(rowHasExpress12BySvc);
    else if (kind === 'express18') rows = state._expAllList.filter(rowHasExpress18BySvc);
    else rows = state._expAllList.slice();

    if (status === 'delivered') return rows.filter(delivered);
    if (status === 'open') return rows.filter(r => !delivered(r));
    return rows;
  }

  function updateKpisOverview() {
    const kinds = ['prio', 'express', 'express12', 'express18'];
    const values = {};
    for (const kind of kinds) {
      const count = status => {
        const rows = metricRows(kind, status);
        return kind === 'prio' ? rows.length : groupRowsByStop(rows).length;
      };
      values['all-' + kind] = count('all');
      values['delivered-' + kind] = count('delivered');
      values['open-' + kind] = count('open');
      values['norollout-' + kind] = metricRows(kind, 'norollout').length;
    }
    setKpis({
      ...values
    });
    for (const status of ['norollout', 'all', 'delivered', 'open']) {
      const total = Number(values[status + '-prio'] || 0) + Number(values[status + '-express'] || 0);
      const badge = document.getElementById(NS + 'total-' + status);
      if (badge) badge.textContent = String(total);
    }
  }

  function groupRowsByStop(rows) {
    const map = new Map();
    for (const r of rows) {
      const key =
        (r.stopId != null ? String(r.stopId) :
         (r.id != null ? String(r.id) :
          `${r.__addr}#${r.__tourNum || ''}`));

      let g = map.get(key);
      if (!g) {
        g = { ...r };
        g.__pkgCount = 1;
        g.__deliveryId = r.id != null ? r.id : null;
        map.set(key, g);
      } else {
        g.__pkgCount++;
      }
    }
    return Array.from(map.values());
  }

  function buildTwoUseTableHtml(rows) {
    let dateLabel = '';
    let importLabel = '';
    let sourceLabel = '';
    try {
      const saved = JSON.parse(String(GM_getValue('pmTwoUseLastDateRange', '') || '{}'));
      const from = parseLocalIsoDate(saved.from);
      const to = parseLocalIsoDate(saved.to);
      if (from && to) dateLabel = ` · ${localDateParts(from).label} bis ${localDateParts(to).label}`;
      const expected = Number(GM_getValue('pmTwoUseExpectedCount', 0) || 0);
      if (expected) importLabel = ` · eingelesen ${state._noRolloutList.length} von ${expected}`;
      const source = String(GM_getValue('pmTwoUseImportSource', '') || '');
      if (source) sourceLabel = ` · Quelle: ${source}`;
    } catch {}
    const columns = [
      ['', 'ticketSelect'],
      ['', 'scanEye'],
      ['Status', 'status'],
      ['Letzter Scan', 'lastScan'],
      ['Scanzeitpunkt', 'scanTime'],
      ['Produkt', 'product'],
      ['PKNR', 'pknr'],
      ['VD', 'vd'],
      ['Route', 'route'],
      ['Ziel PLZ', 'targetZip'],
      ['Ticker', 'ticker']
    ];
    const isCreateTicketRow = row => /tick(?:er|et)\s*erstellen/i.test(String(row?.ticker || ''));
    const safeHttpUrl = raw => {
      try {
        const decoder = document.createElement('textarea');
        decoder.innerHTML = String(raw || '');
        let action = decoder.value
          .replace(/&amp;/gi, '&')
          .replace(/\\u0026/gi, '&')
          .replace(/\\x26/gi, '&')
          .replace(/\\\//g, '/');
        try {
          const decoded = decodeURIComponent(action);
          if (/https?:\/\//i.test(decoded)) action += ` ${decoded}`;
        } catch {}
        const absolute = action.match(/https?:\/\/[^'"\s<>]+/i)?.[0]?.replace(/[),;]+$/, '');
        const relative = action.match(/["'](\/[^"']+)["']/)?.[1];
        const candidate = absolute || relative || action;
        const base = relative ? 'https://2use-render-prod.dpdit.de/' : location.href;
        const url = new URL(candidate, base);
        return /^(https?):$/.test(url.protocol) ? url.href : '';
      } catch {
        return '';
      }
    };
    const renderCell = (row, key) => {
      const value = row[key] || '—';
      if (key === 'ticketSelect') {
        const parcelNumber = String(row.pknr || '').replace(/\D+/g, '');
        return isCreateTicketRow(row) && parcelNumber
          ? `<td data-sort-value="0" style="width:30px;min-width:30px;text-align:center"><input type="checkbox" data-two-use-ticket="1" data-pknr="${esc(parcelNumber)}" title="Für gemeinsame Ticketerstellung auswählen"></td>`
          : '<td data-sort-value="1" style="width:30px;min-width:30px"></td>';
      }
      if (key === 'scanEye') {
        return row.pknr
          ? `<td style="width:34px;min-width:34px;padding:3px;text-align:center"><button type="button" class="${NS}eye" data-scan-popup="${esc(row.pknr)}" title="Lokale Paketauskunft als Popup öffnen">👁</button></td>`
          : '<td></td>';
      }
      if (key === 'status') {
        return `<td data-sort-value="${esc(row.status || '')}" title="${esc(row.status || 'Status')}" style="width:28px;min-width:28px;padding:0;text-align:center"><span aria-label="${esc(row.status || 'Status')}" style="display:inline-block;width:10px;height:26px;background:#e4003b;vertical-align:middle"></span></td>`;
      }
      if (key === 'pknr' && row.pknr) {
        const parcelNumber = String(row.pknr).replace(/\D+/g, '');
        const url = `https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${encodeURIComponent(parcelNumber)}`;
        return `<td data-sort-value="${esc(value)}"><button type="button" data-pm-external="${esc(url)}" title="${esc(value)}\n${esc(url)}" style="display:flex;width:100%;justify-content:space-between;gap:8px;padding:0;border:0;background:transparent;color:#111;font:inherit;text-align:left;cursor:pointer"><span>${esc(value)}</span><span style="color:#00a6d6;font-size:10px">▼</span></button></td>`;
      }
      if (key === 'ticker') {
        const tickerText = String(row.ticker || '');
        const ticketId = tickerText.match(/DPD\s*360\D*(\d{6,})/i)?.[1] || '';
        const createTicket = /tick(?:er|et)\s*erstellen/i.test(tickerText);
        const url = ticketId
          ? `https://dpd360.dpd.de/ticket/ticket_edit.aspx?ticketid=${encodeURIComponent(ticketId)}&view%27`
          : createTicket
            ? 'https://dpd360.dpd.de/ops/express_ticker.aspx'
            : safeHttpUrl(row.tickerUrl);
        if (url) {
          return `<td data-sort-value="${esc(value)}"><button type="button" data-pm-external="${esc(url)}"${createTicket ? ` data-express-ticket="${esc(String(row.pknr || '').replace(/\D+/g, ''))}"` : ''} title="${esc(value)}\n${esc(url)}" style="display:flex;width:100%;justify-content:space-between;gap:8px;padding:0;border:0;background:transparent;color:#111;font:inherit;text-align:left;cursor:pointer"><span>${esc(value)}</span><span style="color:#00a6d6;font-size:10px">▼</span></button></td>`;
        }
      }
      return `<td data-sort-value="${esc(value)}" title="${esc(value)}">${esc(value)}</td>`;
    };
    const body = rows.map(row => `<tr>${columns.map(([, key]) => renderCell(row, key)).join('')}</tr>`).join('');
    return `
      <div style="margin:0 0 8px;color:#475569;font:600 12px system-ui">
        2Use · Einrollung VD – Eingang ED (ED-Sicht)${dateLabel} · Depots 107, 195, 295 · alle Produkte · nur Fehlende${importLabel}${sourceLabel}
      </div>
      <div style="max-height:min(70vh,720px);overflow:auto;border-top:1px solid #aaa">
        <table class="${NS}tbl" data-two-use-table="1" style="font:11px Arial,sans-serif;border-collapse:collapse;background:#fff">
          <thead><tr>${columns.map(([label, key], index) => `<th${key === 'ticketSelect' ? ' style="width:30px;min-width:30px;text-align:center"' : key === 'scanEye' ? ' style="width:34px;min-width:34px"' : key === 'status' ? ' style="width:28px;min-width:28px;text-align:center"' : ''}>${key === 'ticketSelect' ? '<input type="checkbox" data-two-use-ticket-all="1" title="Alle Ticket-erstellen-Zeilen auswählen oder abwählen">' : key === 'scanEye' ? '' : `<button type="button" data-two-use-sort="${index}" title="Auf- oder absteigend sortieren" style="display:flex;width:100%;align-items:center;justify-content:space-between;gap:5px;padding:0;border:0;background:transparent;color:inherit;font:inherit;font-weight:700;cursor:pointer"><span>${esc(label)}</span><span data-sort-indicator style="color:#64748b">↕</span></button>`}</th>`).join('')}</tr></thead>
          <tbody>${body || `<tr><td colspan="${columns.length}">Keine fehlenden Sendungen gefunden.</td></tr>`}</tbody>
        </table>
      </div>`;
  }

  async function toggleStopDetailInline(tr) {
    if (!tr || !(lastOkRequest || getJwtAuthHeader())) return;
    const delId = tr.getAttribute('data-delivery');
    if (!delId) return;

    const tbody = tr.parentNode;
    if (!tbody) return;

    const next = tr.nextElementSibling;
    if (next && next.classList.contains(NS + 'detail-row')) {
      next.remove();
      return;
    }

    const detailRow = document.createElement('tr');
    detailRow.className = NS + 'detail-row';
    const colSpan = tr.children.length || 1;
    const td = document.createElement('td');
    td.colSpan = colSpan;
    td.textContent = 'Lade Stopp-Details …';
    detailRow.appendChild(td);
    tbody.insertBefore(detailRow, tr.nextSibling);

    try {
      let detail = deliveryDetailsCache.get(delId);
      if (!detail) {
        const headers = buildHeaders(lastOkRequest?.headers || {});
        const origin = (lastOkRequest?.url?.origin) || location.origin;
        const url = `${origin}/dispatcher/api/delivery/${encodeURIComponent(delId)}`;
        const res = await fetch(url, { credentials: 'include', headers });
        if (!res.ok) {
          td.textContent = `Fehler beim Laden (HTTP ${res.status})`;
          return;
        }
        detail = await res.json();
        deliveryDetailsCache.set(delId, detail);
      }

      const parcels = Array.isArray(detail.parcels) ? detail.parcels : [];
      const rowsHtml = parcels.map(p => {
        const svc = p.serviceCode || '';
        const isExpSvc = EXPRESS_SERVICE_WHITELIST.has(String(svc));
        const els = Array.isArray(p.elements) ? p.elements.join(', ') : (p.elements || '');
        const prio = p.priority || '';
        let psn = String(p.parcelNumber || '').replace(/\D+/g, '');
        if (psn.length === 13) psn = '0' + psn;

        const psnCell = psn
          ? `<button class="${NS}eye" title="Lokale Paketauskunft als Popup öffnen" data-psn="${esc(psn)}">👁</button><a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${esc(psn)}" target="_blank" rel="noopener">${esc(psn)}</a>`
          : '—';

        return `
          <tr class="${isExpSvc ? NS + 'row-express' : ''}">
            <td>${psnCell}</td>
            <td>${esc(svc || '—')}</td>
            <td>${esc(prio || '—')}</td>
            <td>${esc(els || '—')}</td>
            <td>${isExpSvc ? 'EXPRESS' : 'Normal'}</td>
          </tr>`;
      }).join('');

      const addr = [
        detail.addressStreet,
        detail.addressHouseno,
        detail.addressPcode,
        detail.addressCity
      ].filter(Boolean).join(' ');

      td.innerHTML = `
        <div class="${NS}detail-inner">
          <div style="margin-bottom:4px;">
            <b>Adresse:</b> ${esc(addr || '—')} ·
            <b>Tour:</b> ${esc(detail.tour || '—')} ·
            <b>Pakete:</b> ${parcels.length}
          </div>
          <table>
            <thead>
              <tr>
                <th>PSN</th>
                <th>Servicecode</th>
                <th>Priority</th>
                <th>Elements</th>
                <th>Typ</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml || '<tr><td colspan="5">Keine Pakete gefunden.</td></tr>'}
            </tbody>
          </table>
        </div>`;
    } catch {
      td.textContent = 'Fehler beim Laden der Stopp-Details.';
    }
  }

  function showMetricList(kind, status) {
    const labels = {
      prio: 'PRIO',
      express: 'Express',
      express12: 'Express 12',
      express18: 'Express 18'
    };
    const statusLabels = {
      norollout: 'kein Eingang',
      all: 'in Ausrollung',
      delivered: 'zugestellt',
      open: 'offen'
    };
    const rows = metricRows(kind, status);
    if (status === 'norollout') {
      if (state._twoUseError) {
        state._pendingTwoUseKind = kind;
        openModal('2Use – Kein Eingang', `
          <div class="${NS}empty">
            <b>2Use-Daten konnten nicht geladen werden.</b><br><br>${esc(state._twoUseError)}<br><br>
            <button class="${NS}btn-sm" data-action="connectTwoUse">2Use-Daten laden</button>
          </div>`);
        return;
      }
      openModal(`${labels[kind] || 'Sendungen'} – kein Eingang · ${rows.length}`, buildTwoUseTableHtml(rows));
      return;
    }
    const displayRows = kind === 'prio' ? rows : groupRowsByStop(rows);
    openModal(`${labels[kind] || 'Sendungen'} – ${statusLabels[status] || status} · ${displayRows.length}`, displayRows);
  }

  function showStatusTotal(status) {
    const labels = {
      norollout: 'Kein Eingang',
      all: 'Ausrollung',
      delivered: 'Zugestellt',
      open: 'Offen'
    };
    if (status === 'norollout') {
      if (!state._noRolloutList.length && state._twoUseError) {
        openModal('Gesamt – Kein Eingang', `<div class="${NS}empty"><b>2Use-Daten konnten nicht vollständig geladen werden.</b><br><br>${esc(state._twoUseError)}</div>`);
        return;
      }
      openModal(`Gesamt – Kein Eingang · ${state._noRolloutList.length}`, buildTwoUseTableHtml(state._noRolloutList.slice()));
      return;
    }
    const prioRows = metricRows('prio', status);
    const expressRows = groupRowsByStop(metricRows('express', status));
    const rows = [...prioRows, ...expressRows];
    openModal(`Gesamt – ${labels[status] || status} · ${rows.length}`, rows);
  }

  function showExpLate11() {
    const rows = state._expLate11List.slice();
    const grouped = groupRowsByStop(rows);
    openModal(`Express 12 – zu knapp / falsch einsortiert (>11:01 geplant) · ${grouped.length}`, grouped);
  }

  async function refreshViaApi_SAFE() {
    const [prioRows, exp12Rows, exp18Rows] = await Promise.all([
      fetchPagedFast(buildUrlPrio),
      fetchPagedFast((b, p) => buildUrlElements(b, p, '023')),
      fetchPagedFast((b, p) => buildUrlElements(b, p, '010'))
    ]);

    const prioN = expandAndNorm(prioRows);
    const exp12N = expandAndNorm(exp12Rows);
    const exp18N = expandAndNorm(exp18Rows);

    const { prioAll, prioOpen, expAll, expOpen, expLate11 } = buildTableRowsAndCounts(prioN, exp12N, exp18N);

    state._prioAllList = prioAll.slice();
    state._prioOpenList = prioOpen.slice();
    state._expAllList = expAll.slice();
    state._expOpenList = expOpen.slice();
    state._expLate11List = expLate11.slice();

    updateKpisOverview();

    addEvent({
      title: 'Aktualisiert',
      meta: `PRIO: ${prioAll.length}/${prioOpen.length} offen · EXPRESS: ${expAll.length}/${expOpen.length} offen · >11:01: ${expLate11.length}`,
      sev: 'info',
      read: true,
      ts: Date.now()
    });
  }

  function findPickupDeliveryTrigger() {
    const selectors = ['[role="tab"]', '.mat-mdc-tab', '.mat-tab-label', '.mat-mdc-tab-link', '.mat-tab-link', 'button', 'a', 'div'];
    for (const sel of selectors) {
      const nodes = Array.from(document.querySelectorAll(sel));
      for (const el of nodes) {
        const txt = norm(el.textContent || '');
        if (!/abholung.*zustellung|zustellung.*abholung/i.test(txt)) continue;
        return el.closest('[role="tab"], .mat-mdc-tab, .mat-tab-label, .mat-mdc-tab-link, .mat-tab-link, button, a') || el;
      }
    }
    return null;
  }

  function isLikelyActiveTab(el) {
    if (!el) return false;
    const aria = String(el.getAttribute?.('aria-selected') || '').toLowerCase();
    const cls = String(el.className || '');
    if (aria === 'true') return true;
    if (/active|selected|mdc-tab--active|mat-mdc-tab-active|mat-tab-label-active/i.test(cls)) return true;
    try {
      const cs = getComputedStyle(el);
      const bg = String(cs.backgroundColor || '');
      const color = String(cs.color || '');
      if (bg === 'rgb(225, 6, 50)' || bg === 'rgb(229, 0, 54)' || color === 'rgb(255, 255, 255)') return true;
    } catch {}
    return false;
  }

  function fireRealClick(el) {
    if (!el) return false;
    try { el.scrollIntoView({ block: 'center', inline: 'center' }); } catch {}
    const evOpts = { bubbles: true, cancelable: true, view: window };
    try { el.dispatchEvent(new PointerEvent('pointerdown', evOpts)); } catch {}
    try { el.dispatchEvent(new MouseEvent('mousedown', evOpts)); } catch {}
    try { el.dispatchEvent(new PointerEvent('pointerup', evOpts)); } catch {}
    try { el.dispatchEvent(new MouseEvent('mouseup', evOpts)); } catch {}
    try { el.dispatchEvent(new MouseEvent('click', evOpts)); } catch {}
    try { if (typeof el.click === 'function') el.click(); } catch {}
    return true;
  }

  async function ensurePickupDeliveryActive() {
    const trigger = findPickupDeliveryTrigger();
    if (!trigger) return false;
    if (isLikelyActiveTab(trigger)) return true;
    fireRealClick(trigger);
    await sleep(1200);
    return true;
  }

  async function waitForPickupRequest(timeoutMs = 12000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      if (lastOkRequest?.url?.href?.includes('/dispatcher/api/pickup-delivery')) return true;
      await sleep(250);
    }
    return false;
  }

  async function ensurePickupDeliverySourceReady() {
    if (lastOkRequest?.url?.href?.includes('/dispatcher/api/pickup-delivery')) return true;
    await ensurePickupDeliveryActive();
    const ok = await waitForPickupRequest(12000);
    if (ok) return true;

    const auth = getJwtAuthHeader();
    if (auth) {
      lastOkRequest = { url: buildBasePickupUrlFromToday(), headers: { authorization: auth } };
      return true;
    }
    return false;
  }

  async function initialOpenRefresh() {
    await ensurePickupDeliverySourceReady();
    await fullRefresh(false);
  }

  async function fullRefresh(manual) {
    if (isBusy) return;
    try {
      isBusy = true;
      state._dispatcherLoading = true;
      state._footerMessage = '';
      updateFooterStatus();
      setLoading(true);
      dimButtons(true);

      const ready = await ensurePickupDeliverySourceReady();
      if (!ready) {
        addEvent({
          title: 'Hinweis',
          meta: 'Abholung und Zustellung konnte nicht initialisiert werden.',
          sev: 'warn',
          read: true
        });
        return;
      }

      await Promise.all([
        buildTourDriverMap(),
        loadTourPartnerMap()
      ]);

      await Promise.all([
        refreshViaApi_SAFE(),
        refreshTwoUseMissing().catch(error => {
          console.error('2Use:', error);
          state._twoUseError = String(error?.message || error);
          updateKpisOverview();
          addEvent({
            title: '2Use nicht geladen',
            meta: state._twoUseError,
            sev: 'warn',
            read: true
          });
        })
      ]);
      state.lastRefreshAt = Date.now();
      render();
    } catch (e) {
      console.error(e);
      addEvent({ title: 'Fehler', meta: String(e?.message || e), sev: 'warn', read: true });
    } finally {
      setLoading(false);
      dimButtons(false);
      isBusy = false;
      state._dispatcherLoading = false;
      state.lastRefreshAt = Date.now();
      if (!state._twoUseLoading && !state._twoUseError) {
        state._footerMessage = `Dispatcher-Daten aktualisiert: ${formatDateTime(state.lastRefreshAt)}`;
      }
      updateFooterStatus();
      render();
    }
  }

  async function mailSelectedLate11() {
    const modal = state._modal || {};
    const rows = Array.isArray(modal.rows) ? modal.rows : [];
    const selectedKeys = modal.selected instanceof Set ? modal.selected : new Set();
    if (!rows.length) { alert('Keine Daten.'); return; }
    if (selectedKeys.size === 0) { alert('Keine Zeilen markiert.'); return; }

    const selected = rows.filter(r => {
      const key = String(r.stopId ?? r.id ?? r.__pid ?? '');
      return selectedKeys.has(key);
    });

    const byPartner = new Map();
    for (const r of selected) {
      const p = norm(r.__systempartner || '—');
      if (!byPartner.has(p)) byPartner.set(p, []);
      byPartner.get(p).push(r);
    }

    for (const [partner, list] of byPartner) {
      if (!partner || partner === '—') {
        alert('Mindestens eine Auswahl hat keinen Systempartner.');
        continue;
      }

      const rec = await getPartnerMailRecord(partner);
      const toRaw = rec?.to || '';
      const ccRaw = rec?.cc || '';
      const alias = rec?.alias || partner;

      const toL = normalizeEmailList(toRaw);
      const ccL = normalizeEmailList(ccRaw);

      if (toL.valid.length === 0) {
        alert(`Für Systempartner "${partner}" ist keine gültige E-Mail-Adresse hinterlegt.`);
        continue;
      }

      const subject = `Express 12 – falsch einsortiert (>11:01) – ${alias} – ${new Date().toLocaleDateString('de-DE')}`;
      const html = buildLate11MailHtml(alias, list);
      const ok = await copyHtmlToClipboard(html);
      openMailto(subject, toL.valid.join(','), ccL.valid.join(','));
      if (ok) {
        alert(`Mail-Entwurf für "${alias}" geöffnet.\nHTML ist in der Zwischenablage – im Mail-Body STRG+V.`);
      } else {
        alert(`Mail-Entwurf für "${alias}" geöffnet.\nKopieren fehlgeschlagen – bitte Tabelle manuell kopieren.`);
      }
    }
  }

  function buildLate11MailHtml(partner, rows) {
    const escH = s => String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
    const stamp = new Date().toLocaleString('de-DE');

    const bodyRows = rows.map(r => {
      const psn = escH(r.__pid || '—');
      const tour = escH(String(r.__tourNum || r.tour || '—'));
      const addr = escH(r.__addr || '—');
      const driver = escH(r.__driver || '—');
      const pred = escH(r.__predRangeStr || '—');
      const svc = escH((r.__serviceCodes || []).join(' ') || r.__serviceCode || '—');
      const status = escH(r.__status || '—');
      return `
        <tr>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${psn}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${tour}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${addr}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${driver}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${status}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${svc}</td>
          <td style="padding:6px 8px;border:1px solid #e5e7eb;">${pred}</td>
        </tr>`;
    }).join('');

    return `
      <div style="font:13px/1.45 -apple-system,Segoe UI,Arial,sans-serif;color:#111;">
        <div style="margin:0 0 10px 0;color:#334155">
          <b>${escH(partner)}</b> – Express 12 „falsch einsortiert“ (geplant > 11:01)<br/>
          Stand: ${escH(stamp)}
        </div>
        <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;min-width:700px;">
          <thead>
            <tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">PSN</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Tour</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Adresse</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Fahrer</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Status</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Service</th>
              <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Predict</th>
            </tr>
          </thead>
          <tbody>${bodyRows}</tbody>
        </table>
      </div>`;
  }

  function boot() {
    mountUI();
    if (!twoUseAutoTimer) {
      twoUseAutoTimer = setInterval(() => {
        const panel = document.getElementById(NS + 'panel');
        if (!document.hidden && panel && getComputedStyle(panel).display !== 'none' && Date.now() - lastTwoUseStartedAt > 60000) {
          connectTwoUse(true);
        }
      }, 30000);
    }
    setTimeout(() => {
      if (!(getSetting('depotSuffix') || '')) {
        const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
        if (g) {
          setSetting('depotSuffix', g);
          addEvent({ title: 'Einstellungen', meta: `Depotkennung aus Fahrzeugübersicht: ${g}`, sev: 'info', read: true });
        }
      }
    }, 1200);

    document.addEventListener('visibilitychange', () => {
      if (!document.hidden && document.getElementById(NS + 'panel') && getComputedStyle(document.getElementById(NS + 'panel')).display !== 'none') {
        queueRefreshFromSource();
      }
    });
  }

})();
