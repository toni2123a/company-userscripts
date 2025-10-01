// ==UserScript==
// @name         ASEA incl. letzter Lagerscan
// @namespace    http://tampermonkey.net/
// @version      1.4.4
// @description  Zeigt den letzten Lagerscan aus den Scanserver-Daten an.
// @author       Thiemo Schöler
// Läuft auf allen passenden scanserver-d001#### Hosts:
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// Für XHR:
// @grant        GM_xmlhttpRequest
// @connect      *
// ==/UserScript==
(function() {
  'use strict';

  // 4-stellige Nummer aus dem Hostnamen holen (z. B. 0157)
  const m = location.hostname.match(/scanserver-d001(\d{4})\.ssw\.dpdit\.de/i);
  if (!m) { console.warn('ASEA: keine Standortnummer im Host gefunden.'); return; }
  const standort = m[1];

  console.log('🔄 ASEA aktiv, Standort:', standort);

  function getCurrentDate() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const day = String(d.getDate()).padStart(2,'0');
    return `${y}${m}${day}`;
  }
  const currentDate = getCurrentDate();

  // API-URL mit dynamischem Host + Datum
  const host = `scanserver-d001${standort}.ssw.dpdit.de`;
  const apiUrl = `https://${host}/cgi-bin/pa.cgi?_url=file&_passwd=87654321&_disp=3&_pivotxx=0&_rastert=4&_rasteryt=0&_rasterx=0&_rastery=0&_pivot=0&_pivotbp=0&_sortby=time&_dca=0&_tabledef=time%7C&_arg59=dpd&_arg9=1%2C${currentDate}%2C000000%2C${currentDate}%2C235959&_DateConnectToggle=1&_DateDefault=-1.0&_arg9connect=on&_DateFrom=20250125&_TimeFrom=000000&_DateTo=${currentDate}&_TimeTo=235959&_arg3=08&_csv=5`;

  console.log('🌐 API-URL:', apiUrl);

  GM_xmlhttpRequest({
    method: 'GET',
    url: apiUrl,
    onload: function(response) {
      const rows = (response.responseText || '').trim().split('\n');
      if (rows.length < 2) { console.warn('⚠️ Keine Daten.'); return; }
      const [time, count] = rows[rows.length-1].split(';').map(s => s.trim());

      let table = document.getElementById('scanserver-tabelle');
      if (!table) {
        table = document.createElement('table');
        table.id = 'scanserver-tabelle';
        table.style.cssText = 'position:fixed;top:10px;right:10px;background:#fff;border:1px solid #000;padding:10px;z-index:1000;font-family:Arial, sans-serif;box-shadow:2px 2px 10px rgba(0,0,0,.5);border-collapse:collapse';
        document.body.appendChild(table);
        const head = table.insertRow();
        const h1 = head.insertCell(0), h2 = head.insertCell(1);
        h1.innerHTML = '<b>Scanart</b>'; h2.innerHTML = '<b>Uhrzeit</b>';
        h1.style.padding = h2.style.padding = '5px';
        h1.style.borderBottom = h2.style.borderBottom = '1px solid #000';
      }
      while (table.rows.length > 1) table.deleteRow(1);
      const r = table.insertRow(); const c1 = r.insertCell(0), c2 = r.insertCell(1);
      c1.textContent = 'Lager'; c2.textContent = time;
      c1.style.padding = c2.style.padding = '5px';
    }
  });
})();
