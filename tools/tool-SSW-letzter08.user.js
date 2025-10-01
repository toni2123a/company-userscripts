// ==UserScript==
// @name         ASEA inkl. letzter Lagerscan (mit Standort-Setting)
// @namespace    https://github.com/toni2123a/company-userscripts
// @version      1.4.0
// @description  Zeigt den letzten Eintrag aus den Scanserver-Daten. Standortnummer wird auf der Katalog-Seite gespeichert.
// @match        https://toni2123a.github.io/company-userscripts/*
// @include      /^https?:\/\/scanserver-d00\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
  'use strict';

  const KEY = 'asea_standort'; // GM-Storage-Key für die Standortnummer (z.B. "0157")

  // ---------- UTIL ----------
  function setStandort(num) { try { GM_setValue(KEY, String(num).trim()); } catch(_) {} }
  function getStandort()    { try { return GM_getValue(KEY) || ''; } catch(_) { return ''; } }

  function extractStandortFromHost(host) {
    const m = String(host).match(/scanserver-d00(\d{4})\.ssw\.dpdit\.de/i);
    return m ? m[1] : '';
  }

  function css(s) { GM_addStyle(s); }
  function el(tag, attrs = {}, children = []) {
    const n = document.createElement(tag);
    Object.entries(attrs).forEach(([k,v]) => {
      if (k === 'class') n.className = v;
      else if (k === 'text') n.textContent = v;
      else n.setAttribute(k, v);
    });
    children.forEach(c => n.appendChild(c));
    return n;
  }

  // ---------- UI auf der Katalog-Seite ----------
  function mountSettingsOnCatalog() {
    // Sucht linke Tool-Spalte; wenn nicht vorhanden, oben einsetzen
    const anchor = document.querySelector('#tools-head')?.parentElement || document.body;

    css(`
      .asea-card{border:1px solid #e6e6e6;border-radius:10px;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.06);padding:1rem;margin-bottom:1rem}
      .asea-card h3{margin:.25rem 0 1rem;font-size:1.05rem}
      .asea-row{display:flex;gap:.5rem;align-items:center;flex-wrap:wrap}
      .asea-input{padding:.55rem .7rem;border:1px solid #cfd6e4;border-radius:8px;font-size:1rem;min-width:120px}
      .asea-btn{padding:.6rem 1rem;border-radius:8px;background:#005bbb;color:#fff;font-weight:700;border:none;cursor:pointer}
      .asea-btn:hover{background:#004099}
      .asea-ok{color:#0a7c2f;font-weight:600;margin-top:.5rem}
      .asea-hint{color:#555;font-size:.92rem;margin-top:.35rem}
    `);

    const card = el('div', {class: 'asea-card'});
    const h3   = el('h3', {text: 'Standortnummer festlegen'});
    const row  = el('div', {class: 'asea-row'});
    const inp  = el('input', {class: 'asea-input', type: 'text', maxlength: '4', placeholder: 'z. B. 0157', value: getStandort()});
    const btn  = el('button', {class: 'asea-btn', type: 'button', text: 'Speichern'});
    const ok   = el('div', {class: 'asea-ok', text: ''});
    const hint = el('div', {class: 'asea-hint', text: 'Diese Nummer wird vom Tool auf den DPD-Seiten verwendet. (Beispiel: scanserver-d00' + (getStandort()||'XXXX') + '.ssw.dpdit.de)'});

    btn.addEventListener('click', () => {
      const val = (inp.value || '').trim();
      if (!/^\d{4}$/.test(val)) {
        ok.textContent = 'Bitte 4-stellige Standortnummer eingeben.';
        ok.style.color = '#8a1c1c';
        return;
      }
      setStandort(val);
      ok.textContent = 'Gespeichert: ' + val;
      ok.style.color = '#0a7c2f';
      hint.textContent = 'Diese Nummer wird vom Tool auf den DPD-Seiten verwendet. (Beispiel: scanserver-d00' + val + '.ssw.dpdit.de)';
    });

    row.appendChild(inp);
    row.appendChild(btn);
    card.appendChild(h3);
    card.appendChild(row);
    card.appendChild(ok);
    card.appendChild(hint);

    // ganz oben in der Tool-Liste einfügen
    anchor.parentNode.insertBefore(card, anchor.nextSibling);
  }

  // ---------- Laufzeit auf der DPD-Seite ----------
  async function runOnDpdPage() {
    // 1) Standort ermitteln
    const fromHost = extractStandortFromHost(location.hostname);
    let standort = getStandort() || fromHost;

    // Falls nichts gespeichert & Host liefert auch nichts: kurze Abfrage
    if (!standort) {
      const input = prompt('Bitte 4-stellige Standortnummer eingeben (z. B. 0157):', '');
      if (input && /^\d{4}$/.test(input)) {
        standort = input;
        setStandort(standort);
      } else {
        console.warn('[ASEA] Keine gültige Standortnummer – Tool beendet.');
        return;
      }
    }

    // 2) Hinweis falls Host-Nummer ≠ gespeicherte Nummer
    if (fromHost && standort && fromHost !== standort) {
      console.info('[ASEA] Host ('+fromHost+') ≠ gespeicherte Nummer ('+standort+'). Verwende gespeicherte Nummer.');
    }

    // 3) Beispiel: hier deine eigentliche Logik (Scanserver-Daten abrufen usw.)
    //    Falls du GM_xmlhttpRequest brauchst, kannst du die URL mit der "standort" zusammensetzen:
    //    const url = `https://scanserver-d00${standort}.ssw.dpdit.de/cgi-bin/report_inbound_ofd.cgi?...`;
    //    GM_xmlhttpRequest({ method:'GET', url, onload: (r)=>{ /* ... */ } });

    // Demo: unauffälliger Hinweis in der Seite
    try {
      const tag = document.createElement('div');
      tag.textContent = 'ASEA-Tool aktiv (Standort: ' + standort + ')';
      tag.style.cssText = 'position:fixed;bottom:10px;right:10px;background:#005bbb;color:#fff;padding:6px 10px;border-radius:8px;font:12px/1.2 Arial, sans-serif;z-index:99999;opacity:.9';
      document.body.appendChild(tag);
      setTimeout(()=> tag.remove(), 5000);
    } catch(_) {}
  }

  // ---------- Router ----------
  const isCatalog = /:\/\/toni2123a\.github\.io\/company-userscripts\//.test(location.href);
  const isDpd     = /scanserver-d00\d{4}\.ssw\.dpdit\.de/i.test(location.hostname);
  if (isCatalog)  mountSettingsOnCatalog();
  if (isDpd)      runOnDpdPage();
})();
