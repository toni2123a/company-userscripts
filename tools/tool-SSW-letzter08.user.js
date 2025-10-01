// ==UserScript==
// @name         ASEA inkl. letzter Lagerscan (mit Standort-Setting)
// @namespace    https://github.com/toni2123a/company-userscripts
// @version      1.4.2
// @description  Zeigt den letzten Eintrag aus den Scanserver-Daten. Standortnummer wird auf der Katalog-Seite gespeichert und im GM-Storage gespiegelt.
// @match        https://toni2123a.github.io/company-userscripts/*
// @include      /^https?:\/\/scanserver-d00\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
  'use strict';
  const KEY = 'asea_standort';

  function setStandort(v) {
    const val = String(v || '').trim();
    try { GM_setValue(KEY, val); } catch(_) {}
    try { localStorage.setItem(KEY, val); } catch(_) {}
  }
  function getStandortGM() { try { return GM_getValue(KEY) || ''; } catch(_) { return ''; } }
  function getStandortLS() { try { return localStorage.getItem(KEY) || ''; } catch(_) { return ''; } }
  function getStandort()   { return getStandortGM() || getStandortLS(); }

  function extractNumFromHost(host) {
    const m = String(host).match(/scanserver-d00(\d{4})\.ssw\.dpdit\.de/i);
    return m ? m[1] : '';
  }

  // -------- Katalog-Seite: Button abgreifen & GM-Storage spiegeln --------
  function initCatalogBridge() {
    // wartet bis das Einstellungsfeld vorhanden ist
    let tries = 0;
    const iv = setInterval(() => {
      tries++;
      const inp = document.getElementById('aseaNum');
      const btn = document.getElementById('aseaSave');
      const msg = document.getElementById('aseaMsg');
      if (!inp || !btn) { if (tries > 40) clearInterval(iv); return; }

      clearInterval(iv);

      // initiale Synchronisation: falls GM-Wert existiert, zeige ihn; sonst LS übernehmen
      const gm = getStandortGM();
      if (gm) {
        inp.value = gm;
        try { localStorage.setItem(KEY, gm); } catch(_) {}
        if (msg) { msg.textContent = 'Gespeichert: ' + gm; msg.style.color = '#0a7c2f'; }
      } else {
        const ls = getStandortLS();
        if (ls) setStandort(ls); // spiegelt LS -> GM
      }

      btn.addEventListener('click', () => {
        const v = (inp.value || '').trim();
        if (!/^\d{4}$/.test(v)) {
          if (msg) { msg.textContent = 'Bitte 4-stellige Standortnummer eingeben.'; msg.style.color = '#8a1c1c'; }
          return;
        }
        setStandort(v); // schreibt GM + LS
        if (msg) { msg.textContent = 'Gespeichert: ' + v; msg.style.color = '#0a7c2f'; }
      });
    }, 125);
  }

  // -------- DPD-Seite: Standort nutzen --------
  function runOnDpdPage() {
    const fromHost = extractNumFromHost(location.hostname);
    let standort = getStandort() || fromHost;

    if (!standort) {
      const input = prompt('Bitte 4-stellige Standortnummer eingeben (z. B. 0157):', '');
      if (input && /^\d{4}$/.test(input)) { standort = input; setStandort(standort); }
      else { console.warn('[ASEA] Keine gültige Standortnummer – Tool beendet.'); return; }
    }

    // >>> HIER deine eigentliche Logik mit "standort" einbauen <<<
    // Beispiel:
    // const url = `https://scanserver-d00${standort}.ssw.dpdit.de/cgi-bin/report_inbound_ofd.cgi?...`;
    // GM_xmlhttpRequest({ method:'GET', url, onload: r => { /* ... */ } });

    // kleiner Hinweis
    try {
      const tag = document.createElement('div');
      tag.textContent = 'ASEA-Tool aktiv (Standort: ' + standort + ')';
      tag.style.cssText = 'position:fixed;bottom:10px;right:10px;background:#005bbb;color:#fff;padding:6px 10px;border-radius:8px;font:12px/1.2 Arial, sans-serif;z-index:99999;opacity:.9';
      document.body.appendChild(tag);
      setTimeout(()=> tag.remove(), 5000);
    } catch(_) {}
  }

  const isCatalog = /:\/\/toni2123a\.github\.io\/company-userscripts\//.test(location.href);
  const isDpd     = /scanserver-d00\d{4}\.ssw\.dpdit\.de/i.test(location.hostname);

  if (isCatalog) initCatalogBridge();
  if (isDpd)     runOnDpdPage();
})();
