// ==UserScript==
// @name         Zerberus → DPD360  
// @namespace    https://dpd.de/
// @version      1.3c
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-Zerberus.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-Zerberus.user.js
// @description  Verknüpft die Avisdaten aus DPD 360 und zeigt diese auf der Website an. Eimalige Anmeldung (Tag) in dpd360 erforderlich
// @author       Thiemo Schöler
// @match        https://zerberus-dpd-02.csf.blujaysolutions.net/*
// @grant        GM_xmlhttpRequest
// @connect      dpd360.dpd.de
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const SEL_LABELS = 'td, label, div, span, dt';
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ---------- Finder ----------
  function findLabelEl(root) {
    const nodes = root.querySelectorAll(SEL_LABELS);
    for (const n of nodes) {
      if (n.textContent && n.textContent.trim() === 'Benutzerfeld 2') return n;
    }
    return null;
  }

  function extractParcelFromContext(labelEl) {
    if (!labelEl) return null;

    if (labelEl.matches('td')) {
      const row = labelEl.closest('tr');
      if (row) {
        const cells = Array.from(row.children);
        const idx = cells.indexOf(labelEl);
        if (idx > -1 && cells[idx + 1]) {
          const txt = cells[idx + 1].innerText || cells[idx + 1].textContent || '';
          const m = txt.replace(/\s+/g, ' ').match(/\b\d{10,}\b/);
          if (m) return m[0];
        }
      }
    }

    const siblings = [
      labelEl.nextElementSibling,
      labelEl.parentElement ? labelEl.parentElement.nextElementSibling : null
    ].filter(Boolean);

    for (const sib of siblings) {
      const txt = (sib.innerText || sib.textContent || '').replace(/\s+/g, ' ').trim();
      const m = txt.match(/\b\d{10,}\b/);
      if (m) return m[0];
    }

    const container =
      labelEl.closest('.row, .form-group, table, tbody, .content, .panel, .box, .boxContent') || document;
    const all = container.querySelectorAll('td, div, span');
    let after = false;
    for (const el of all) {
      const t = (el.textContent || '').trim();
      if (!after) {
        if (t === 'Benutzerfeld 2') after = true;
        continue;
      }
      const m = t.match(/\b\d{10,}\b/);
      if (m) return m[0];
    }

    const parentTxt = (labelEl.parentElement?.textContent || '').replace(/\s+/g, ' ').trim();
    const rest = parentTxt.replace(/^Benutzerfeld 2\s*/i, '');
    const m = rest.match(/\b\d{10,}\b/);
    return m ? m[0] : null;
  }

  // ---------- Styles & UI ----------
  function ensureStyles() {
    if (document.getElementById('dpd360-flyout-style')) return;
    const css = `
      .dpd360-flyout {
        position: fixed;
        top: 110px;
        right: 28px;
        width: 420px;
        max-height: 72vh;
        z-index: 9999;
      }
      .dpd360-flyout details {
        border: 1px solid #cfd4da;
        border-radius: 8px;
        background: #f8f9fa;
        box-shadow: 0 6px 18px rgba(0,0,0,0.08);
        overflow: hidden;
      }
      .dpd360-flyout summary {
        cursor: pointer;
        list-style: none;
        padding: 10px 12px;
        font-weight: 600;
        background: #eef1f4;
        border-bottom: 1px solid #dde1e5;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .dpd360-flyout summary::-webkit-details-marker { display: none; }
      .dpd360-flyout .dpd360-body {
        padding: 12px;
        font-size: 13px;
        color: #222;
        overflow: auto;
        max-height: 64vh;
      }
      .dpd360-row { margin: 4px 0 10px 0; color: #555; }
      .dpd360-row code { font-size: 12px; }
      .dpd360-section { margin-top: 10px; }
      .dpd360-section .title {
        font-weight: 600; margin-bottom: 4px; display:flex; align-items:center; justify-content:space-between; gap:8px;
      }
      .dpd360-actions { display:flex; gap:8px; flex-wrap:wrap; }
      .dpd360-btn {
        display:inline-block; padding:6px 10px; font-size:12px; border-radius:6px;
        text-decoration:none; border:1px solid #1e4fd7; background:#2563eb; color:#fff;
      }
      .dpd360-btn:hover { filter: brightness(0.95); }
      .dpd360-badge {
        display:inline-block; padding:2px 6px; font-size:12px; border-radius:6px;
        background:#e9ecef; color:#333;
      }
      @media (max-width: 1280px) { .dpd360-flyout { width: 360px; right: 16px; top: 100px; } }
      @media (max-width: 980px)  { .dpd360-flyout { position: fixed; width: 90vw; left: 5vw; right: auto; top: 90px; } }
    `;
    const style = document.createElement('style');
    style.id = 'dpd360-flyout-style';
    style.textContent = css;
    document.head.appendChild(style);
  }

  function createFlyout() {
    ensureStyles();
    const host = document.createElement('div');
    host.className = 'dpd360-flyout';
    host.innerHTML = `
      <details id="dpd360-flyout-details">
        <summary>
          <span>DPD360 Informationen</span>
          <span class="dpd360-badge" id="dpd360-parcel-mini"></span>
        </summary>
        <div class="dpd360-body" id="dpd360-flyout-body">
          <em>Suche läuft…</em>
        </div>
      </details>
    `;
    document.body.appendChild(host);
    return {
      setParcelBadge: (txtHtml) => (host.querySelector('#dpd360-parcel-mini').innerHTML = txtHtml),
      setBody: (html) => (host.querySelector('#dpd360-flyout-body').innerHTML = html),
      open: () => host.querySelector('#dpd360-flyout-details').setAttribute('open', '')
    };
  }

  // ---------- Helpers ----------
  // (A) Volle Adresse (für Google-Suche)
  function buildAddressQueryFromAvis(node) {
    if (!node) return null;
    const raw = node.innerHTML
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/&amp;/gi, '&')
      .replace(/&nbsp;/gi, ' ');
    const lines = raw
      .split('\n')
      .map(s => s.replace(/<[^>]+>/g, '').trim())
      .filter(Boolean)
      .filter(s => !/^Tel\.:/i.test(s) && s !== '---');
    if (!lines.length) return null;
    return lines.join(', ');
  }

  // (B) Nur Straße+Hausnr + PLZ/Ort (für Google Maps) – strikt 3. & 4. Zeile
  function buildMapsQueryFromAvis(node) {
    if (!node) return null;

    // in echte Zeilen zerlegen – ohne Zeilen rauszufiltern,
    // damit die Indizes (0: Name, 1: Name2, 2: Straße, 3: PLZ/Ort, …) stabil bleiben
    const lines = node.innerHTML
      .split(/<br\s*\/?>/i)
      .map(s => s
        .replace(/<[^>]+>/g, '')   // Tags weg
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/\u00A0/g, ' ')   // NBSP → normal
        .trim()
      );

    // Erwartetes Muster:
    // 0: Name
    // 1: Name2
    // 2: Straße + Hausnr
    // 3: (ggf. "DE - ") + PLZ + Ort
    const street = lines[2] || '';
    let plzOrt = (lines[3] || '').replace(/^[A-Z]{2}\s*[-–]\s*/i, '').trim(); // "DE - " entfernen

    if (!street || !plzOrt) return null;
    return `${street}, ${plzOrt}`;
  }

  // ---------- Fetch ----------
  function fetchDPD360(parcelNo, ui) {
    const dpdUrl = `https://dpd360.dpd.de/order/order_view.aspx?parcelno=${encodeURIComponent(parcelNo)}`;
    ui.setParcelBadge(`<a href="${dpdUrl}" target="_blank">${parcelNo}</a>`);

    GM_xmlhttpRequest({
      method: 'GET',
      url: dpdUrl,
      headers: { Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8' },
      timeout: 15000,
      onload: (resp) => {
        if (resp.status !== 200) {
          ui.setBody(`
            <div class="dpd360-row" style="color:#c00;">Fehler: ${resp.status} ${resp.statusText}</div>
            <div><a href="${dpdUrl}" target="_blank">DPD360 manuell öffnen</a></div>
          `);
          return;
        }
        try {
          const parser = new DOMParser();
          const doc = parser.parseFromString(resp.responseText, 'text/html');

          const labLabelAddress = doc.querySelector('#ContentPlaceHolder1_labLabelAddress');
          const labAvisAddress  = doc.querySelector('#ContentPlaceHolder1_labAvisAddress');
          const labRegAddress   = doc.querySelector('#ContentPlaceHolder1_labRegistrationaddress');
          const cloudHeadline   = doc.querySelector('#ContentPlaceHolder1_labCloudUserID_Headline');
          const accountLink     = doc.querySelector('#ContentPlaceHolder1_hplAccountView');
          const accountHref     = accountLink ? new URL(accountLink.getAttribute('href'), dpdUrl).href : null;

          const avisFull  = buildAddressQueryFromAvis(labAvisAddress);   // für Google-Suche (alle Zeilen)
          const mapsQuery = buildMapsQueryFromAvis(labAvisAddress);      // exakt 3.+4. Zeile

          const mapsUrl = mapsQuery ? `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(mapsQuery)}` : null;
          const webUrl  = avisFull  ? `https://www.google.com/search?q=${encodeURIComponent(avisFull)}` : null;

          const section = (title, node, extraRightHtml = '') =>
            node
              ? `<div class="dpd360-section">
                   <div class="title">
                     <span>${title}</span>
                     ${extraRightHtml || ''}
                   </div>
                   <div>${node.innerHTML}</div>
                 </div>`
              : '';

          const avisActions = (mapsUrl || webUrl)
            ? `<div class="dpd360-actions">
                 ${mapsUrl ? `<a class="dpd360-btn" href="${mapsUrl}" target="_blank" title="In Google Maps öffnen">Maps</a>` : ''}
                 ${webUrl  ? `<a class="dpd360-btn" href="${webUrl}"  target="_blank" title="Google Suche">Google</a>` : ''}
               </div>`
            : '';

          ui.setBody(`
            <div class="dpd360-row">
              Sendungsnr.: <code>${parcelNo}</code> &nbsp;|&nbsp;
              <a href="${dpdUrl}" target="_blank">DPD360 öffnen</a>
            </div>
            ${section('Client (nicht Verkäufer)', labLabelAddress)}
            ${section('Avis-Adresse', labAvisAddress, avisActions)}
            ${section('Registrierungsadresse', labRegAddress)}
            ${cloudHeadline ? `<div class="dpd360-section">${cloudHeadline.textContent.trim()}</div>` : ''}
            ${accountHref ? `<div class="dpd360-section"><a href="${accountHref}" target="_blank">Konto anzeigen</a></div>` : ''}
          `);
        } catch (e) {
          ui.setBody(`
            <div class="dpd360-row" style="color:#c00;">Parsing-Fehler: ${String(e)}</div>
            <div><a href="${dpdUrl}" target="_blank">DPD360 manuell öffnen</a></div>
          `);
        }
      },
      onerror: () => ui.setBody(`<div class="dpd360-row" style="color:#c00;">Netzwerkfehler.</div>`),
      ontimeout: () => ui.setBody(`<div class="dpd360-row" style="color:#c00;">Zeitüberschreitung.</div>`)
    });
  }

  // ---------- Waiter ----------
  async function waitForLabel(timeoutMs = 15000) {
    const start = Date.now();
    let label = findLabelEl(document);
    if (label) return label;

    let resolver;
    const p = new Promise((res) => (resolver = res));
    const mo = new MutationObserver(() => {
      const el = findLabelEl(document);
      if (el) {
        mo.disconnect();
        resolver(el);
      }
    });
    mo.observe(document.body, { childList: true, subtree: true, characterData: true });

    while (Date.now() - start < timeoutMs) {
      await sleep(300);
      const el = findLabelEl(document);
      if (el) {
        mo.disconnect();
        return el;
      }
    }
    mo.disconnect();
    return null;
  }

  // ---------- Main ----------
  (async function main() {
    const labelEl = await waitForLabel(15000);
    if (!labelEl) return;
    const parcelNo = extractParcelFromContext(labelEl);
    if (!parcelNo) return;

    const ui = createFlyout();
    // ui.open(); // optional: Panel direkt geöffnet starten
    fetchDPD360(parcelNo, ui);
  })();

})();
