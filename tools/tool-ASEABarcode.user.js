// ==UserScript==
// @name         ASEA 🚀 Nur Barcode als Popup
// @namespace    http://tampermonkey.net/
// @version      6.3
// @description  Scanserver ASEA: Nur Barcode-Spalte (Symbol bleibt klein), Klick öffnet großes Overlay inkl. Kopieren-Button.
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-ASEABarcode.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-ASEABarcode.user.js
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/report_inbound_ofd\.cgi.*$/
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // ---------- Overlay einmalig anlegen ----------

    function setupOverlay() {
        if (document.getElementById('barcodeOverlay')) return;

        const overlay = document.createElement('div');
        overlay.id = 'barcodeOverlay';
        Object.assign(overlay.style, {
            position: 'fixed',
            inset: '0',
            display: 'none',
            alignItems: 'center',
            justifyContent: 'center',
            background: 'rgba(0,0,0,0.7)',
            zIndex: '9999'
        });

        const box = document.createElement('div');
        Object.assign(box.style, {
            background: '#fff',
            padding: '12px',
            borderRadius: '4px',
            textAlign: 'center',
            boxShadow: '0 0 20px #000'
        });

        const img = document.createElement('img');
        img.id = 'barcodeOverlayImg';
        Object.assign(img.style, {
            maxWidth: '90vw',
            maxHeight: '70vh',
            marginBottom: '8px'
        });

        const btnCopy = document.createElement('button');
        btnCopy.textContent = 'Kopieren';
        btnCopy.style.margin = '4px';

        const btnClose = document.createElement('button');
        btnClose.textContent = 'Schließen';
        btnClose.style.margin = '4px';

        box.appendChild(img);
        box.appendChild(document.createElement('br'));
        box.appendChild(btnCopy);
        box.appendChild(btnClose);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        // ---- Schließen ----
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) overlay.style.display = 'none';
        });
        btnClose.onclick = () => overlay.style.display = 'none';
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') overlay.style.display = 'none';
        });

        // ---- Kopieren als Bild ----
        btnCopy.onclick = async () => {
            if (!img.src) return;

            try {
                const image = new Image();
                image.crossOrigin = 'anonymous';
                image.src = img.src;

                await image.decode();

                const canvas = document.createElement('canvas');
                canvas.width = image.naturalWidth;
                canvas.height = image.naturalHeight;

                const ctx = canvas.getContext('2d');
                ctx.fillStyle = '#fff';
                ctx.fillRect(0, 0, canvas.width, canvas.height);
                ctx.drawImage(image, 0, 0);

                const blob = await new Promise(res => canvas.toBlob(res, 'image/png'));
                if (!blob || !navigator.clipboard || !window.ClipboardItem) {
                    alert('Kopieren vom Browser blockiert.');
                    return;
                }

                await navigator.clipboard.write([
                    new ClipboardItem({ 'image/png': blob })
                ]);

                btnCopy.textContent = 'Kopiert ✓';
                setTimeout(() => btnCopy.textContent = 'Kopieren', 1200);

            } catch (e) {
                alert('Kopieren fehlgeschlagen.');
            }
        };
    }

    // ---------- Barcode URL ----------

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

    // ---------- Tabelle erweitern ----------

    function updateTable() {
        const tbody = document.querySelector('.restable tbody');
        if (!tbody) return;

        const headerRow = tbody.parentNode.querySelector('tr.tableheader');
        if (headerRow && !headerRow.dataset.extended) {
            const th = document.createElement('th');
            th.textContent = "III";
            headerRow.appendChild(th);
            headerRow.dataset.extended = "true";
        }

        const rows = tbody.querySelectorAll("tr:not(.tableheader)");
        rows.forEach(row => {
            if (row.dataset.checked) return;

            const cells = row.cells;
            const plzCell     = cells[7];
            const paketCell   = cells[0];
            const soCodeCell  = cells[3];
            const scanartCell = cells[1];

            if (scanartCell?.querySelector('a')) {
                scanartCell.innerHTML = scanartCell.innerText;
            }

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

    // ---------- Observer ----------

    const observer = new MutationObserver(updateTable);

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
