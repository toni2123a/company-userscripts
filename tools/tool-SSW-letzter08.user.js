// ==UserScript==
// @name         ASEA inkl. letzter Lagerscan (mit Standort-Setting)
// @namespace    https://github.com/toni2123a/company-userscripts
// @version      1.4.3
// @description  Zeigt den letzten Eintrag aus den Scanserver-Daten. Standortnummer wird auf der Katalog-Seite gespeichert (GM-Storage + localStorage) und auf DPD-Seiten genutzt.
// @match        https://toni2123a.github.io/company-userscripts/*
// @include      /^https?:\/\/scanserver-d00\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-asea.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-asea.user.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addStyle
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
  'use strict';
  const KEY = 'asea_standort';

  // --- Storage Helpers ---
  function setStandort(v) {
    const val = String(v || '').trim();
    try { GM_setValue(KEY, val); console.info('[ASEA] GM_setValue ->', val); } catch(e) { console.warn('[ASEA] GM_setValue fehlgeschlagen', e); }
    try { localStorage.setItem(KEY, val); console.info('[ASEA] localStorage ->', val); } catch(e) { console.warn('[ASEA] localStorage fehlgeschlagen', e); }
  }
  function getGM() { try { return GM_getValue(KEY) || ''; } catch(_) { return ''; } }
  function getLS() { try { return localStorage.getItem(KEY) || ''; } catch(_) { return ''; } }
  function getStandortAny() {
    const gm = getGM();
    if (gm) return gm;
    const ls = getLS();
    return ls || '';
  }
  function extractFromHost(host) {
    const m = String(host).match(/scanserver-d00(\d{4})\.ssw\.dpdit\.de/i);
    return m ? m[1] : '';
  }

  // --- Katalog-Seite: UI anbinden & GM-Storage spiegeln ---
  function initCatalogBridge() {
    let tries = 0;
    const iv = setInterval(() => {
      tries++;
      const inp = document.getElementById('aseaNum');
      const btn = document.getElementById('aseaSave');
      const msg = document.getElementById('aseaMsg');
      if (!inp || !btn) { if (tries > 60) clearInterval(iv); return; }
      clearInterval(iv);

      // initial sync
      const gm = getGM();
      if (gm) {
        inp.value = gm;
        try { localStorage.setItem(KEY, gm); } catch(_) {}
        if (msg) { msg.textContent = 'Gespeichert: ' + gm; msg.style.color = '#0a7c2f'; }
        console.info('[ASEA] Katalog: GM->UI sync', gm);
      } else {
        const ls = getLS();
        if (ls) { setStandort(ls); if (msg) { msg.textContent = 'Gespeichert: ' + ls; msg.style.color = '#0a7c2f'; } }
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
    }, 100);
  }

  // --- DPD-Seite: Nummer verwenden ---
  function runOnDpdPage() {
    const fromHost = extractFromHost(location.hostname);
    let standort = getStandortAny() || fromHost;

    if (!standort) {
      const input = prompt('Bitte 4-stellige Standortnummer eingeben (z. B. 0157):', '');
      if (input && /^\d{4}$/.test(input)) { standort = input; setStandort(standort); }
      else { console.warn('[ASEA] Keine gültige Standortnummer – Tool beendet.'); return; }
    }
    console.info('[ASEA] Verwende Standort =', standort, '(Host:', fromHost || '—', ')');

    // >>> HIER deine eigentliche Logik (z. B. Request):
    // const url = `https://scanserver-d00${standort}.ssw.dpdit.de/cgi-bin/report_inbound_ofd.cgi?...`;
    // GM_xmlhttpRequest({ method:'GET', url, onload: r => { /* ... */ } });

    // dezenter Hinweis
    try {
      const tag = document.createElement('div');
      tag.textContent = 'ASEA aktiv – Standort: ' + standort;
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
