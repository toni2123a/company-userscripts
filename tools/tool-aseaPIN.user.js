// ==UserScript==
// @name         ASEA PIN Freigabe
// @namespace    http://tampermonkey.net/
// @version      6.2
// @description  Eingangsmengenabgleich: Tour-Bubbles + QR-Popup, Excel-Import und Mehrfachauswahl (Button/Contextmenü).
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js 
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/scanmonitor\.cgi.*$/
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
  'use strict';

  const DEPOT = (window.location.hostname.match(/scanserver-d(\d{7})\.ssw\.dpdit\.de/i) || [])[1] || '0000000';
  const STORAGE_KEY = 'spTourConfigScanmonitor_v5';
  const COLLAPSE_KEY = STORAGE_KEY + '_collapsed';

  // ---------- QR-Code Helfer ----------

  async function fetchQrContentForTour(tour) {
    const base = window.location.origin + '/lso/jcrp_ws/scanpocket/qrcode/clearance-granted?tour=';
    const url = base + encodeURIComponent(tour);
    const resp = await fetch(url, { credentials: 'include' });

    if (!resp.ok) throw new Error('HTTP-Status ' + resp.status + ' (' + resp.statusText + ')');

    let json;
    try { json = await resp.json(); } catch (e) { throw new Error('Antwort ist kein gültiges JSON.'); }

    if (json && json.qrCode && typeof json.qrCode.code === 'string' && json.qrCode.code.trim() !== '') return json.qrCode.code;
    if (json && typeof json.code === 'string' && json.code.trim() !== '') return json.code;

    console.error('QR-API Antwort für Tour', tour, json);
    if (json && (json.error || json.message)) throw new Error('Server meldet: ' + (json.error || json.message));
    throw new Error('Server liefert keinen QR-Code für diese Tour (kein Feld "code" in der Antwort).');
  }

  function buildQrImageUrl(content) {
    return 'https://barcodeapi.org/api/qr/' + encodeURIComponent(content);
  }

  // ---------- Bild-Kopieren (PNG) ----------

  async function copyCanvas(canvas) {
    if (!navigator.clipboard || !window.ClipboardItem) return false;
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
    if (!blob) return false;
    try {
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      return true;
    } catch (e) {
      return false;
    }
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
    } catch (e) {
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
        } catch (e) {
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

  // ---------- Popups ----------

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

      // NEU: Kopieren (als Bild)
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
        } catch (e) {
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

    // Text geändert: nur "Kopieren"
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
      } catch (e) {
        btnCopy.textContent = old;
        btnCopy.disabled = false;
        alert('Kopieren fehlgeschlagen.');
      }
    };
  }

  // ---------- Config-Helfer ----------

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

  function saveConfig(list) {
    try { window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list)); }
    catch (e) { console.error('Konfig speichern fehlgeschlagen:', e); }
  }

  function getSelectedSystempartnerNames(doc) {
    const sel = doc.querySelector('select[name="systempartner"]');
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

  // ---------- UI ----------

  function initInDocument(doc) {
    if (!doc.body || doc.getElementById('tm-sp-panel')) return;

    const css = `
#tm-sp-panel{position:fixed;top:70px;right:10px;width:360px;max-height:80vh;overflow:auto;background:#f5f5f5;border:1px solid #ccc;padding:6px 8px;font-family:Arial,sans-serif;font-size:12px;z-index:9999;box-sizing:border-box;}
#tm-sp-panel h3{margin:0 0 4px 0;font-size:13px;}
#tm-sp-panel button{padding:2px 6px;font-size:11px;margin:2px 2px;cursor:pointer;}
#tm-sp-info{margin:4px 0;}
.tm-tour-bubble{display:inline-block;padding:3px 8px;margin:2px 4px 2px 0;border-radius:12px;background:#7d7d7d;color:#fff;cursor:pointer;white-space:nowrap;border:1px solid transparent;}
.tm-tour-bubble:hover{filter:brightness(1.1);}
.tm-tour-bubble.tm-selected{background:#4a4a4a;border-color:#000;}
#tm-collapse-btn{position:absolute;top:0;right:0;width:28px;height:28px;padding:0;margin:0;line-height:28px;text-align:center;font-weight:bold;}
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

    function applyCollapsedState(isCollapsed) {
      panel.classList.toggle('tm-collapsed', !!isCollapsed);
      try { window.localStorage.setItem(COLLAPSE_KEY, isCollapsed ? '1' : '0'); } catch (e) {}
    }

    let collapsed = false;
    try { collapsed = (window.localStorage.getItem(COLLAPSE_KEY) === '1'); } catch (e) {}
    applyCollapsedState(collapsed);

    btnCollapse.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      applyCollapsedState(!panel.classList.contains('tm-collapsed'));
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

  // ---------- Starter ----------

  setInterval(() => {
    try {
      const docs = [document];

      if (window.frames && window.frames.length) {
        for (let i = 0; i < window.frames.length; i++) {
          const f = window.frames[i];
          try { if (f.document) docs.push(f.document); } catch (e) {}
        }
      }

      for (const doc of docs) {
        if (!doc || !doc.body) continue;
        if (/Eingangsmengenabgleich/.test(doc.body.textContent || '')) {
          const sel = doc.querySelector('select[name="systempartner"]');
          if (sel) initInDocument(doc);
        }
      }
    } catch (e) {
      console.error(e);
    }
  }, 800);

})();
