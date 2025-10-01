// ==UserScript==
// @name         ASEA incl. letzter Lagerscan
// @namespace    http://tampermonkey.net/
// @version      1.4
// @description  Zeigt den letzten Eintrag aus den Scanserver-Daten an, jetzt mit dynamischem Datum und "Lager"-Bezeichnung.
// @author       Thiemo Schöler
// @match        *://scanserver-d0010157.ssw.dpdit.de/cgi-bin/report_inbound_ofd.cgi*
// @grant        GM_xmlhttpRequest
// ==/UserScript==

(function() {
    'use strict';

    console.log("🔄 Tampermonkey-Script läuft...");

    // Funktion, um das aktuelle Datum im YYYYMMDD-Format zu erhalten
    function getCurrentDate() {
        let today = new Date();
        let year = today.getFullYear();
        let month = String(today.getMonth() + 1).padStart(2, '0'); // Monat beginnt bei 0
        let day = String(today.getDate()).padStart(2, '0');
        return `${year}${month}${day}`;
    }

    // Aktuelles Datum für die API abrufen
    let currentDate = getCurrentDate();
    console.log(`📆 Aktuelles Datum: ${currentDate}`);

    // Dynamische API-URL mit aktuellem Datum
    let apiUrl = `https://scanserver-d0010157.ssw.dpdit.de/cgi-bin/pa.cgi?_url=file&_passwd=87654321&_disp=3&_pivotxx=0&_rastert=4&_rasteryt=0&_rasterx=0&_rastery=0&_pivot=0&_pivotbp=0&_sortby=time&_dca=0&_tabledef=time%7C&_arg59=dpd&_arg9=1%2C${currentDate}%2C000000%2C${currentDate}%2C235959&_DateConnectToggle=1&_DateDefault=-1.0&_arg9connect=on&_DateFrom=20250125&_TimeFrom=000000&_DateTo=${currentDate}&_TimeTo=235959&_arg3=08&_csv=5`;

    console.log(`🌐 API-URL generiert: ${apiUrl}`);

    // CSV-Daten abrufen
    GM_xmlhttpRequest({
        method: "GET",
        url: apiUrl,
        onload: function(response) {
            console.log("📥 API-Antwort erhalten:", response.responseText);

            // CSV in ein Array von Objekten umwandeln
            let rows = response.responseText.trim().split("\n");
            let data = rows.slice(1).map(row => { // Erste Zeile ignorieren (Überschrift)
                let [time, count] = row.split(";");
                return { time: time.trim(), count: count.trim() };
            });

            // Letzten Eintrag holen
            if (data.length > 0) {
                let letzterEintrag = data[data.length - 1];
                console.log("📊 Letzter Eintrag:", letzterEintrag);

                // Tabelle erstellen oder bestehende finden
                let table = document.getElementById("scanserver-tabelle");
                if (!table) {
                    table = document.createElement("table");
                    table.id = "scanserver-tabelle";
                    table.style.position = "fixed";
                    table.style.top = "10px";  // Oben positionieren
                    table.style.right = "10px"; // Rechts positionieren
                    table.style.backgroundColor = "white";
                    table.style.border = "1px solid black";
                    table.style.padding = "10px";
                    table.style.zIndex = "1000";
                    table.style.fontFamily = "Arial, sans-serif";
                    table.style.boxShadow = "2px 2px 10px rgba(0,0,0,0.5)";
                    table.style.borderCollapse = "collapse";
                    document.body.appendChild(table);

                    // Tabellenkopf hinzufügen
                    let headerRow = table.insertRow();
                    let header1 = headerRow.insertCell(0);
                    let header2 = headerRow.insertCell(1);
                    header1.innerHTML = "<b>Scanart</b>";
                    header2.innerHTML = "<b>Uhrzeit</b>";
                    header1.style.padding = "5px";
                    header2.style.padding = "5px";
                    header1.style.borderBottom = "1px solid black";
                    header2.style.borderBottom = "1px solid black";
                }

                // Vorherige Daten entfernen, bevor neue Zeilen eingefügt werden
                while (table.rows.length > 1) {
                    table.deleteRow(1);
                }

                // Neue Zeilen mit den neuesten Daten füllen
                let newRow = table.insertRow();
                let cell1 = newRow.insertCell(0);
                let cell2 = newRow.insertCell(1);
                cell1.innerHTML = "Lager"; // Geändert von "Letzter Scan" zu "Lager"
                cell2.innerHTML = letzterEintrag.time;
                cell1.style.padding = "5px";
                cell2.style.padding = "5px";
            } else {
                console.warn("⚠️ Keine Daten gefunden.");
            }
        }
    });

})();
