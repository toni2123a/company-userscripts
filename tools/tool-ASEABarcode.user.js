// ==UserScript==
// @name         ASEA 🚀 Nur Barcode als Popup
// @namespace    http://tampermonkey.net/
// @version      6.2
// @description  Scanserver ASEA: Nur Barcode-Spalte (Symbol bleibt klein), Klick öffnet großes Overlay. Google-Links entfernt.
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-ASEA_Erweiterung.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-ASEA_Erweiterung.user.js
// @author
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @grant        none
// ==/UserScript==



(function () {

    'use strict';



    // --- Overlay einmalig anlegen ---

    function setupOverlay() {

        if (document.getElementById('barcodeOverlay')) return;

        const overlay = document.createElement('div');

        overlay.id = 'barcodeOverlay';

        Object.assign(overlay.style, {

            position: 'fixed', inset: '0', display: 'none',

            alignItems: 'center', justifyContent: 'center',

            background: 'rgba(0,0,0,0.7)', zIndex: '9999'

        });



        const img = document.createElement('img');

        img.id = 'barcodeOverlayImg';

        Object.assign(img.style, {

            maxWidth: '90%', maxHeight: '90%',

            boxShadow: '0 0 20px #fff', border: '4px solid #fff', background: '#fff'

        });



        overlay.appendChild(img);

        document.body.appendChild(overlay);



        // Klick/ESC schließt

        overlay.addEventListener('click', () => overlay.style.display = 'none');

        document.addEventListener('keydown', (e) => {

            if (e.key === 'Escape') overlay.style.display = 'none';

        });

    }



    function bigBarcode(finalCode) {

        return `https://barcodeapi.org/api/128/${encodeURIComponent(finalCode)}`;

    }



    function openOverlay(url) {

        const overlay = document.getElementById('barcodeOverlay');

        const img = document.getElementById('barcodeOverlayImg');

        if (!overlay || !img) return;

        img.src = url;

        overlay.style.display = 'flex';

    }



    function updateTable() {

        const tbody = document.querySelector('.restable tbody');

        if (!tbody) return;



        const headerRow = tbody.parentNode.querySelector('tr.tableheader');

        if (headerRow && !headerRow.dataset.extended) {

            // Nur 1 Spalte: Barcode

            const th = document.createElement('th');

            th.textContent = "III";

            headerRow.appendChild(th);

            headerRow.dataset.extended = "true";

        }



        const rows = tbody.querySelectorAll("tr:not(.tableheader)");

        rows.forEach(row => {

            if (row.dataset.checked) return;



            const cells = row.cells;

            const plzCell     = cells[7];  // Empfänger PLZ

            const paketCell   = cells[0];  // Paketscheinnummer

            const soCodeCell  = cells[3];  // SOCode

            const scanartCell = cells[1];  // Letzte Scanart



            // Link in "Letzte Scanart" entfernen (nur Text)

            if (scanartCell?.querySelector('a')) {

                scanartCell.innerHTML = scanartCell.innerText;

            }



            // Barcode-Symbol mit Overlay

            if (plzCell && paketCell && soCodeCell) {

                const plzValue = plzCell.innerText.trim().substring(0, 5).replace(/\D/g,'');

                const paketNr  = paketCell.innerText.trim();

                const soCode   = soCodeCell.innerText.trim();

                if (plzValue && paketNr && soCode) {

                    const finalBarcode = `%00${plzValue}${paketNr}${soCode}276`;

                    const url = bigBarcode(finalBarcode);



                    const td = document.createElement('td');

                    const a = document.createElement('a');

                    a.href = '#';

                    a.textContent = 'III';

                    a.style.textDecoration = 'underline';

                    a.addEventListener('click', (e) => {

                        e.preventDefault();

                        e.stopPropagation();

                        openOverlay(url);

                    });

                    td.appendChild(a);

                    row.appendChild(td);

                }

            }



            row.dataset.checked = "true";

        });

    }



    // MutationObserver für dynamische Änderungen

    const observer = new MutationObserver(updateTable);



    // Start

    const iv = setInterval(() => {

        const tbody = document.querySelector('.restable tbody');

        if (tbody) {

            clearInterval(iv);

            setupOverlay();

            observer.observe(tbody, { childList: true, subtree: true });

            updateTable();

        }

    }, 400);



})();
