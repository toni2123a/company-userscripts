// ==UserScript==
// @name         ASEA 🚀 Optimierte Barcode- & Google-Links
// @namespace    http://tampermonkey.net/
// @version      5.1
// @description  Scanserver ASEA Erweiterung. Fügt Barcode-, Google Maps- & Google Search-Spalten hinzu (schneller & stabiler)
// @author       Thiemo Schöler
// Läuft auf allen Standorten (z. B. 0157, 0160, ...):
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // Standortnummer (4-stellig) aus dem Host extrahieren – nur Info/Debug
    const m = location.hostname.match(/scanserver-d001(\d{4})\.ssw\.dpdit\.de/i);
    const standort = m ? m[1] : '????';
    console.log("🚀 Skript gestartet – Standort:", standort);

    function updateTable() {
        let table = document.querySelector('.restable tbody');
        if (!table) {
            console.log("❌ Tabelle nicht gefunden. Warte...");
            return;
        }

        console.log("✅ Tabelle gefunden. Überprüfe Spalten...");

        let headerRow = table.parentNode.querySelector("tr.tableheader");

        // Falls die neuen Spalten noch nicht existieren → Füge sie hinzu
        if (headerRow && !headerRow.dataset.extended) {
            ["Barcode-Link", "Google Maps", "Google Search"].forEach(text => {
                let newHeader = document.createElement("th");
                newHeader.textContent = text;
                headerRow.appendChild(newHeader);
            });
            headerRow.dataset.extended = "true"; // Verhindert doppeltes Einfügen
            console.log("✅ Neue Spalten in der Kopfzeile hinzugefügt!");
        }

        let rows = table.querySelectorAll("tr:not(.tableheader)"); // Alle Datenzeilen holen
        rows.forEach(row => {
            if (row.dataset.checked) return; // Falls schon verarbeitet → Überspringen
            let cells = row.cells;

            let plzCell = cells[7];  // Spalte 8: "Empfänger PLZ"
            let paketCell = cells[0]; // Spalte 1: "Paketscheinnummer"
            let soCodeCell = cells[3]; // Spalte 4: "SOCode"
            let scanartCell = cells[1]; // Spalte 2: "Letzte Scanart"
            let streetCell = cells[9]; // Spalte 10: "Straße"
            let nameCell = cells[10];  // Spalte 11: "Name1"

            // 🔹 Link aus Spalte 2 entfernen (nur Text behalten)
            if (scanartCell?.querySelector('a')) {
                scanartCell.innerHTML = scanartCell.innerText;
            }

            // 1️⃣ Barcode-Link
            if (plzCell && paketCell && soCodeCell) {
                let plzValue = plzCell.innerText.trim().substring(0, 5);
                let paketNr = paketCell.innerText.trim();
                let soCode = soCodeCell.innerText.trim();
                let finalBarcode = `%00${plzValue}${paketNr}${soCode}276`;
                let barcodeUrl = `https://barcodeapi.org/api/128/${encodeURIComponent(finalBarcode)}`;

                let barcodeCell = document.createElement("td");
                barcodeCell.innerHTML = `<a href="${barcodeUrl}" target="_blank" style="color:blue; text-decoration:underline;">🔗 Barcode</a>`;
                row.appendChild(barcodeCell);
            }

            // 2️⃣ Google Maps Link
            if (plzCell && streetCell) {
                let plz = plzCell.innerText.trim().substring(0, 5);
                let street = streetCell.innerText.trim().replace(/\s+/g, '+');
                let googleMapsURL = `https://www.google.com/maps/search/?api=1&query=${plz}+${street}`;

                let mapsCell = document.createElement("td");
                mapsCell.innerHTML = `<a href="${googleMapsURL}" target="_blank" style="color:blue; text-decoration:underline;">📍 Google Maps</a>`;
                row.appendChild(mapsCell);
            }

            // 3️⃣ Google Search Link
            if (plzCell && streetCell && nameCell) {
                let plz = plzCell.innerText.trim().substring(0, 5);
                let street = streetCell.innerText.trim().replace(/\s+/g, '+');
                let name = nameCell.innerText.trim().replace(/\s+/g, '+');
                let googleSearchURL = `https://www.google.com/search?q=${plz}+${street}+${name}`;

                let searchCell = document.createElement("td");
                searchCell.innerHTML = `<a href="${googleSearchURL}" target="_blank" style="color:blue; text-decoration:underline;">🔍 Google Search</a>`;
                row.appendChild(searchCell);
            }

            // Markiere die Zeile als bearbeitet
            row.dataset.checked = "true";
        });

        console.log("✅ Tabelle erfolgreich aktualisiert!");
    }

    // MutationObserver für dynamische Änderungen (z. B. nach Filterung)
    const observer = new MutationObserver(() => {
        console.log("🔄 Änderung erkannt! Aktualisiere...");
        updateTable();
    });

    // Starte das Skript nur, wenn die Tabelle da ist
    const tableCheckInterval = setInterval(() => {
        let table = document.querySelector('.restable tbody');
        if (table) {
            clearInterval(tableCheckInterval);
            observer.observe(table, { childList: true, subtree: true });
            updateTable();
        }
    }, 500);
})();
