// ==UserScript==
// @name         ASEA PIN Freigabe 
// @namespace    http://tampermonkey.net/
// @version      5.6
// @description  Eingangsmengenabgleich: Tour-Bubbles + QR-Popup, Excel-Import und Mehrfachauswahl (Button/Contextmenü).
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @match        https://scanserver-d001\d{4}\.ssw.dpdit.de/cgi-bin/scanmonitor.cgi*
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/scanmonitor\.cgi.*$/
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
    'use strict';

    const STORAGE_KEY = 'spTourConfigScanmonitor_v5';

    // ---------- QR-Code Helfer ----------

    async function fetchQrContentForTour(tour) {
        const base = window.location.origin +
            '/lso/jcrp_ws/scanpocket/qrcode/clearance-granted?tour=';
        const url = base + encodeURIComponent(tour);
        const resp = await fetch(url, { credentials: 'include' });

        if (!resp.ok) {
            throw new Error('HTTP-Status ' + resp.status + ' (' + resp.statusText + ')');
        }

        let json;
        try {
            json = await resp.json();
        } catch (e) {
            throw new Error('Antwort ist kein gültiges JSON.');
        }

        // Standard-Struktur: { qrCode: { data:{...}, code:"..." } }
        if (json && json.qrCode && typeof json.qrCode.code === 'string' && json.qrCode.code.trim() !== '') {
            return json.qrCode.code;
        }

        // Fallback: direktes "code"
        if (json && typeof json.code === 'string' && json.code.trim() !== '') {
            return json.code;
        }

        console.error('QR-API Antwort für Tour', tour, json);

        if (json && (json.error || json.message)) {
            throw new Error('Server meldet: ' + (json.error || json.message));
        }

        throw new Error('Server liefert keinen QR-Code für diese Tour (kein Feld "code" in der Antwort).');
    }

    function buildQrImageUrl(content) {
        return 'https://barcodeapi.org/api/qr/' + encodeURIComponent(content);
    }

    async function showQrPopup(doc, tour) {
        try {
            const content = await fetchQrContentForTour(tour);
            const imgUrl = buildQrImageUrl(content);

            const overlay = doc.createElement('div');
            Object.assign(overlay.style, {
                position: 'fixed',
                inset: '0',
                background: 'rgba(0,0,0,0.6)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                zIndex: '999999'
            });

            const box = doc.createElement('div');
            Object.assign(box.style, {
                background: '#fff',
                padding: '12px',
                borderRadius: '4px',
                textAlign: 'center',
                minWidth: '320px',
                boxShadow: '0 0 15px rgba(0,0,0,0.5)',
                fontFamily: 'Arial, sans-serif',
                fontSize: '12px'
            });

            const title = doc.createElement('div');
            title.textContent = 'MDE Freigabe PIN (Zustellung) – Tour ' + tour;
            title.style.marginBottom = '6px';
            box.appendChild(title);

            const img = doc.createElement('img');
            img.src = imgUrl;
            Object.assign(img.style, {
                maxWidth: '360px',
                maxHeight: '360px',
                marginBottom: '6px'
            });
            box.appendChild(img);

            const info = doc.createElement('div');
            info.textContent = 'Tour: ' + tour + ' | Depot: 0010157';
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
                    'Tour: ' + tour + ' | Depot: 0010157' +
                    '</body></html>'
                );
                w.document.close();
                w.focus();
                w.print();
            };
            box.appendChild(btnPrint);

            overlay.appendChild(box);
            overlay.addEventListener('click', e => {
                if (e.target === overlay) overlay.remove();
            });

            doc.body.appendChild(overlay);
        } catch (e) {
            alert('Fehler beim Laden des QR-Codes für Tour ' + tour + ':\n' + e.message);
        }
    }

    async function showMultiQrPopup(doc, tours) {
        if (!tours || !tours.length) return;

        const overlay = doc.createElement('div');
        Object.assign(overlay.style, {
            position: 'fixed',
            inset: '0',
            background: 'rgba(0,0,0,0.6)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: '999999'
        });

        const box = doc.createElement('div');
        Object.assign(box.style, {
            background: '#fff',
            padding: '12px',
            borderRadius: '4px',
            textAlign: 'center',
            minWidth: '340px',
            maxWidth: '90vw',
            maxHeight: '90vh',
            overflow: 'auto',
            boxShadow: '0 0 15px rgba(0,0,0,0.5)',
            fontFamily: 'Arial, sans-serif',
            fontSize: '12px'
        });

        const title = doc.createElement('div');
        title.textContent = 'MDE Freigabe PIN – mehrere Touren (' + tours.join(', ') + ')';
        title.style.marginBottom = '6px';
        box.appendChild(title);

        const grid = doc.createElement('div');
        Object.assign(grid.style, {
            display: 'flex',
            flexWrap: 'wrap',
            justifyContent: 'center',
            gap: '10px'
        });
        box.appendChild(grid);

        const info = doc.createElement('div');
        info.textContent = 'Touren: ' + tours.join(', ') + ' | Depot: 0010157';
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

        overlay.appendChild(box);
        overlay.addEventListener('click', e => {
            if (e.target === overlay) overlay.remove();
        });

        doc.body.appendChild(overlay);

        const imgUrls = {};

        // QRs nacheinander laden
        for (const tour of tours) {
            const card = doc.createElement('div');
            Object.assign(card.style, {
                border: '1px solid #ccc',
                borderRadius: '4px',
                padding: '4px',
                textAlign: 'center',
                minWidth: '150px'
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
                const img = doc.createElement('img');
                img.src = url;
                Object.assign(img.style, {
                    maxWidth: '160px',
                    maxHeight: '160px'
                });
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
                html += '<div>Tour ' + t + ' | Depot: 0010157</div>';
                html += '<img src="' + url + '"><br>';
                html += '</div>';
            });
            html += '</body></html>';
            w.document.write(html);
            w.document.close();
            w.focus();
            w.print();
        };
    }

    // ---------- Config-Helfer ----------

    function normalizeName(name) {
        return name.trim().toLowerCase();
    }

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
        try {
            window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
        } catch (e) {
            console.error('Konfig speichern fehlgeschlagen:', e);
        }
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

    // ---------- UI in einem bestimmten Dokument aufbauen ----------

    function initInDocument(doc) {
        if (!doc.body || doc.getElementById('tm-sp-panel')) return;

        const css = 
        #tm-sp-panel {
            position: fixed;
            top: 70px;
            right: 10px;
            width: 360px;
            max-height: 80vh;
            overflow: auto;
            background: #f5f5f5;
            border: 1px solid #ccc;
            padding: 6px 8px;
            font-family: Arial, sans-serif;
            font-size: 12px;
            z-index: 9999;
        }
        #tm-sp-panel h3 {
            margin: 0 0 4px 0;
            font-size: 13px;
        }
        #tm-sp-panel button {
            padding: 2px 6px;
            font-size: 11px;
            margin: 2px 2px;
            cursor: pointer;
        }
        #tm-sp-info {
            margin: 4px 0;
        }
        .tm-tour-bubble {
            display:inline-block;
            padding:3px 8px;
            margin:2px 4px 2px 0;
            border-radius:12px;
            background:#7d7d7d; /* grau */
            color:#fff;
            cursor:pointer;
            white-space:nowrap;
            border:1px solid transparent;
        }
        .tm-tour-bubble:hover {
            filter:brightness(1.1);
        }
        .tm-tour-bubble.tm-selected {
            background:#4a4a4a;
            border-color:#000;
        }
        #tm-settings-panel {
            margin-top:6px;
            border-top:1px solid #ccc;
            padding-top:4px;
            display:none;
        }
        #tm-settings-panel table {
            border-collapse: collapse;
            width:100%;
        }
        #tm-settings-panel th,
        #tm-settings-panel td {
            border:1px solid #ddd;
            padding:2px 3px;
            font-size:11px;
        }
        #tm-settings-panel th {
            background:#eee;
        }
        #tm-settings-panel input[type="text"],
        #tm-excel-import {
            width:100%;
            box-sizing:border-box;
            font-size:11px;
        }
        #tm-excel-import {
            height:80px;
            resize:vertical;
        };
        const style = doc.createElement('style');
        style.textContent = css;
        doc.head.appendChild(style);

        const panel = doc.createElement('div');
        panel.id = 'tm-sp-panel';

        const title = doc.createElement('h3');
        title.textContent = 'MDE Freigabe PIN (Zustellung)';
        panel.appendChild(title);

        const btnSettings = doc.createElement('button');
        btnSettings.textContent = 'Einstellungen';
        panel.appendChild(btnSettings);

        const infoSpan = doc.createElement('div');
        infoSpan.id = 'tm-sp-info';
        infoSpan.textContent = 'Kein Systempartner ausgewählt.';
        panel.appendChild(infoSpan);

        const bubblesContainer = doc.createElement('div');
        bubblesContainer.id = 'tm-sp-bubbles';
        panel.appendChild(bubblesContainer);

        // Mehrfachauswahl + Button
        const multiDiv = doc.createElement('div');
        multiDiv.style.margin = '4px 0';

        const multiCheckbox = doc.createElement('input');
        multiCheckbox.type = 'checkbox';
        multiCheckbox.id = 'tm-multi-select';

        const multiLabel = doc.createElement('label');
        multiLabel.htmlFor = 'tm-multi-select';
        multiLabel.textContent = ' Mehrfachauswahl';

        const multiShowBtn = doc.createElement('button');
        multiShowBtn.textContent = 'Anzeigen';
        multiShowBtn.style.marginLeft = '4px';

        multiDiv.appendChild(multiCheckbox);
        multiDiv.appendChild(multiLabel);
        multiDiv.appendChild(multiShowBtn);
        panel.appendChild(multiDiv);

        const settingsPanel = doc.createElement('div');
        settingsPanel.id = 'tm-settings-panel';
        panel.appendChild(settingsPanel);

        doc.body.appendChild(panel);

        const selectedTours = new Set();

        function clearSelection() {
            selectedTours.clear();
            const bubbles = bubblesContainer.querySelectorAll('.tm-tour-bubble.tm-selected');
            bubbles.forEach(b => b.classList.remove('tm-selected'));
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
                if (entry && entry.tours) {
                    entry.tours.forEach(t => toursSet.add(t));
                }
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

                if (selectedTours.has(tour)) {
                    bub.classList.add('tm-selected');
                }

                // Linksklick
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

                // Rechtsklick: wie Button, aber für „Power-User“
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

        function renderSettingsPanel() {
            const cfg = loadConfig();
            settingsPanel.innerHTML = '';

            const caption = doc.createElement('div');
            caption.textContent = 'Systempartner / Touren – Einstellungen';
            caption.style.fontWeight = 'bold';
            caption.style.marginBottom = '4px';
            settingsPanel.appendChild(caption);

            const table = doc.createElement('table');
            const thead = doc.createElement('thead');
            const trh = doc.createElement('tr');
            ['Systempartner', 'Tournummern', 'Aktion'].forEach(txt => {
                const th = doc.createElement('th');
                th.textContent = txt;
                trh.appendChild(th);
            });
            thead.appendChild(trh);
            table.appendChild(thead);

            const tbody = doc.createElement('tbody');

            cfg.forEach((item, index) => {
                const tr = doc.createElement('tr');

                const tdName = doc.createElement('td');
                const inpName = doc.createElement('input');
                inpName.type = 'text';
                inpName.value = item.name;
                inpName.disabled = true;
                tdName.appendChild(inpName);

                const tdTours = doc.createElement('td');
                const inpTours = doc.createElement('input');
                inpTours.type = 'text';
                inpTours.value = item.tours.join(', ');
                inpTours.disabled = true;
                tdTours.appendChild(inpTours);

                const tdAct = doc.createElement('td');
                const btnEdit = doc.createElement('button');
                btnEdit.textContent = 'Bearbeiten';
                const btnDel = doc.createElement('button');
                btnDel.textContent = 'Löschen';
                tdAct.appendChild(btnEdit);
                tdAct.appendChild(doc.createTextNode(' '));
                tdAct.appendChild(btnDel);

                let editing = false;

                btnEdit.addEventListener('click', () => {
                    if (!editing) {
                        editing = true;
                        inpName.disabled = false;
                        inpTours.disabled = false;
                        btnEdit.textContent = 'Speichern';
                    } else {
                        const name = inpName.value.trim();
                        const toursArr = inpTours.value
                            .split(/[,\s]+/)
                            .map(t => t.trim())
                            .filter(Boolean);
                        if (!name || !toursArr.length) {
                            alert('Bitte Name und mindestens eine Tournummer angeben.');
                            return;
                        }
                        cfg[index] = { name, tours: toursArr };
                        saveConfig(cfg);
                        editing = false;
                        inpName.disabled = true;
                        inpTours.disabled = true;
                        btnEdit.textContent = 'Bearbeiten';
                        updateBubbles();
                    }
                });

                btnDel.addEventListener('click', () => {
                    if (!confirm('Diesen Systempartner löschen?')) return;
                    cfg.splice(index, 1);
                    saveConfig(cfg);
                    renderSettingsPanel();
                    updateBubbles();
                });

                tr.appendChild(tdName);
                tr.appendChild(tdTours);
                tr.appendChild(tdAct);
                tbody.appendChild(tr);
            });

            const trNew = doc.createElement('tr');

            const tdNewName = doc.createElement('td');
            const inpNewName = doc.createElement('input');
            inpNewName.type = 'text';
            inpNewName.placeholder = 'Systempartner-Name';
            tdNewName.appendChild(inpNewName);

            const tdNewTours = doc.createElement('td');
            const inpNewTours = doc.createElement('input');
            inpNewTours.type = 'text';
            inpNewTours.placeholder = 'Touren, z.B. 627, 623, 638';
            tdNewTours.appendChild(inpNewTours);

            const tdNewAct = doc.createElement('td');
            const btnNewSave = doc.createElement('button');
            btnNewSave.textContent = 'Speichern';
            btnNewSave.disabled = true;
            tdNewAct.appendChild(btnNewSave);

            function checkNewRow() {
                const hasName = inpNewName.value.trim().length > 0;
                const toursArr = inpNewTours.value
                    .split(/[,\s]+/)
                    .map(t => t.trim())
                    .filter(Boolean);
                btnNewSave.disabled = !(hasName && toursArr.length > 0);
            }

            inpNewName.addEventListener('input', checkNewRow);
            inpNewTours.addEventListener('input', checkNewRow);

            btnNewSave.addEventListener('click', () => {
                const name = inpNewName.value.trim();
                const toursArr = inpNewTours.value
                    .split(/[,\s]+/)
                    .map(t => t.trim())
                    .filter(Boolean);
                if (!name || !toursArr.length) {
                    alert('Bitte Name und mindestens eine Tournummer angeben.');
                    return;
                }
                const cfgNow = loadConfig();
                cfgNow.push({ name, tours: toursArr });
                saveConfig(cfgNow);
                renderSettingsPanel();
                updateBubbles();
            });

            trNew.appendChild(tdNewName);
            trNew.appendChild(tdNewTours);
            trNew.appendChild(tdNewAct);
            tbody.appendChild(trNew);

            table.appendChild(tbody);
            settingsPanel.appendChild(table);

            // --- Import aus Excel-Liste (2 Spalten: Systempartner, Tour) ---
            const importTitle = doc.createElement('div');
            importTitle.textContent = 'Import aus Excel-Liste (2 Spalten: Systempartner, Tour):';
            importTitle.style.marginTop = '6px';
            importTitle.style.fontWeight = 'bold';
            settingsPanel.appendChild(importTitle);

            const importHint = doc.createElement('div');
            importHint.textContent = 'In Excel Bereich A:B markieren, kopieren und hier einfügen. Überschrift wird ignoriert.';
            importHint.style.fontSize = '10px';
            importHint.style.marginBottom = '2px';
            settingsPanel.appendChild(importHint);

            const ta = doc.createElement('textarea');
            ta.id = 'tm-excel-import';
            settingsPanel.appendChild(ta);

            const importBtn = doc.createElement('button');
            importBtn.textContent = 'Importieren';
            importBtn.style.marginTop = '4px';
            settingsPanel.appendChild(importBtn);

            importBtn.addEventListener('click', () => {
                const text = ta.value;
                if (!text.trim()) {
                    alert('Bitte erst aus Excel einfügen (Strg+V).');
                    return;
                }

                const lines = text.split(/\r?\n/);
                let cfgNow = loadConfig();
                let addedCount = 0;

                lines.forEach(line => {
                    const trimmed = line.trim();
                    if (!trimmed) return;

                    const parts = line.split(/\t|;/);
                    if (parts.length < 2) return;

                    const name = parts[0].trim();
                    let tour = parts[1].trim();

                    if (!name || !tour || name.toLowerCase().startsWith('systempartner')) {
                        return; // Überschrift oder leer
                    }

                    tour = tour.replace(/[^\dA-Za-z]/g, '');
                    if (!tour) return;

                    const norm = normalizeName(name);
                    let entry = cfgNow.find(p => normalizeName(p.name) === norm);
                    if (!entry) {
                        entry = { name: name, tours: [] };
                        cfgNow.push(entry);
                    }

                    if (!entry.tours.includes(tour)) {
                        entry.tours.push(tour);
                        addedCount++;
                    }
                });

                if (!addedCount) {
                    alert('Es konnten keine neuen Touren importiert werden (evtl. alles schon vorhanden?).');
                    return;
                }

                saveConfig(cfgNow);
                alert('Import abgeschlossen. Neue Touren: ' + addedCount);
                ta.value = '';
                renderSettingsPanel();
                updateBubbles();
            });
        }

        const sel = doc.querySelector('select[name="systempartner"]');
        if (sel) {
            sel.addEventListener('change', updateBubbles);
        }

        btnSettings.addEventListener('click', () => {
            if (settingsPanel.style.display === 'none' || !settingsPanel.style.display) {
                renderSettingsPanel();
                settingsPanel.style.display = 'block';
            } else {
                settingsPanel.style.display = 'none';
            }
        });

        // Mehrfachauswahl an/aus
        multiCheckbox.addEventListener('change', () => {
            if (!multiCheckbox.checked) {
                clearSelection(); // => alle Markierungen weg
            }
        });

        // Button "Anzeigen"
        multiShowBtn.addEventListener('click', () => {
            if (!multiCheckbox.checked) {
                alert('Bitte zuerst „Mehrfachauswahl“ aktivieren.');
                return;
            }
            const list = selectedTours.size ? Array.from(selectedTours) : [];
            if (!list.length) {
                alert('Bitte mindestens eine Tour markieren.');
                return;
            }
            showMultiQrPopup(doc, list);
        });

        updateBubbles();
    }

    // ---------- Starter: richtigen Frame immer wieder prüfen ----------

    setInterval(() => {
        try {
            const docs = [];
            docs.push(document);

            if (window.frames && window.frames.length) {
                for (let i = 0; i < window.frames.length; i++) {
                    const f = window.frames[i];
                    try {
                        if (f.document) docs.push(f.document);
                    } catch (e) { /* anderer Origin */ }
                }
            }

            for (const doc of docs) {
                if (!doc || !doc.body) continue;
                if (/Eingangsmengenabgleich/.test(doc.body.textContent || '')) {
                    const sel = doc.querySelector('select[name="systempartner"]');
                    if (sel) {
                        initInDocument(doc);
                    }
                }
            }
        } catch (e) {
            console.error(e);
        }
    }, 800);

})();
ChatGPT:

Nachgedacht für 20s
Ja, das geht sauber über die URL 👍
Du kannst dir die Depotnummer einmal aus location.hostname ziehen und dann überall verwenden, statt sie hart reinzuschreiben.

1. Depotnummer aus der URL holen
Ganz oben im Script (direkt nach 'use strict';) fügst du das hier ein:

    'use strict';

    // Depotnummer aus der URL lesen, z.B.
    // https://scanserver-d0010157.ssw.dpdit.de/...
    const DEPOT = (window.location.hostname.match(/scanserver-d(\d{7})\.ssw\.dpdit\.de/i) || [])[1] || '0000000';
Das Regex (\d{7}) holt dir die 7-stellige Depotnummer (z.B. 0010157).

Wenn das mal nicht matcht, ist der Fallback 0000000 (kannst du nach Wunsch ändern, z.B. 'UNBEKANNT').

2. Alle festen „Depot: 0010157“ durch Variable ersetzen
In deinem Code kommt 0010157 aktuell an mehreren Stellen vor. Diese Stellen änderst du wie folgt:

a) In showQrPopup
Alt:

            const info = doc.createElement('div');
            info.textContent = 'Tour: ' + tour + ' | Depot: 0010157';
Neu:

            const info = doc.createElement('div');
            info.textContent = 'Tour: ' + tour + ' | Depot: ' + DEPOT;
Und beim Drucken:

Alt:

                    '<h3>MDE Freigabe PIN (Zustellung)</h3>' +
                    '<img src="' + imgUrl + '"><br>' +
                    'Tour: ' + tour + ' | Depot: 0010157' +
Neu:

                    '<h3>MDE Freigabe PIN (Zustellung)</h3>' +
                    '<img src="' + imgUrl + '"><br>' +
                    'Tour: ' + tour + ' | Depot: ' + DEPOT +
b) In showMultiQrPopup
Alt:

        const info = doc.createElement('div');
        info.textContent = 'Touren: ' + tours.join(', ') + ' | Depot: 0010157';
Neu:

        const info = doc.createElement('div');
        info.textContent = 'Touren: ' + tours.join(', ') + ' | Depot: ' + DEPOT;
Und beim Sammel-Druck:

Alt:

                html += '<div style="page-break-inside:avoid;margin-bottom:20px;">';
                html += '<div>Tour ' + t + ' | Depot: 0010157</div>';
                html += '<img src="' + url + '"><br>';
Neu:

                html += '<div style="page-break-inside:avoid;margin-bottom:20px;">';
                html += '<div>Tour ' + t + ' | Depot: ' + DEPOT + '</div>';
                html += '<img src="' + url + '"><br>';
Damit wird das Depot automatisch aus der URL wie
https://scanserver-d0010157.ssw.dpdit.de/cgi-bin/...
gezogen, und das Script funktioniert für alle Standorte/Depots ohne Anpassung.

Wenn du willst, kann ich dir auch die komplett fertige, angepasste Version des Scripts in einem Block zusammenstellen.

Du:
bitte erstell mir den kompletten code
ChatGPT:

Nachgedacht für 10s
Hier ist der komplette, angepasste Code mit dynamischer Depot-Erkennung aus der URL (DEPOT wird aus scanserver-d0010157... zu 0010157):

// ==UserScript==
// @name         ASEA PIN Freigabe 
// @namespace    http://tampermonkey.net/
// @version      5.7
// @description  Eingangsmengenabgleich: Tour-Bubbles + QR-Popup, Excel-Import und Mehrfachauswahl (Button/Contextmenü).
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-aseaPIN.user.js
// @match        https://scanserver-d001\d{4}\.ssw.dpdit.de/cgi-bin/scanmonitor.cgi*
// @include      /^https?:\/\/scanserver-d001\d{4}\.ssw\.dpdit\.de\/cgi-bin\/scanmonitor\.cgi.*$/
// @grant        none
// @run-at       document-end
// ==/UserScript==

(function () {
    'use strict';

    // Depotnummer aus Hostname holen, z.B. scanserver-d0010157.ssw.dpdit.de -> 0010157
    const DEPOT = (window.location.hostname.match(/scanserver-d(\d{7})\.ssw\.dpdit\.de/i) || [])[1] || '0000000';

    const STORAGE_KEY = 'spTourConfigScanmonitor_v5';

    // ---------- QR-Code Helfer ----------

    async function fetchQrContentForTour(tour) {
        const base = window.location.origin +
            '/lso/jcrp_ws/scanpocket/qrcode/clearance-granted?tour=';
        const url = base + encodeURIComponent(tour);
        const resp = await fetch(url, { credentials: 'include' });

        if (!resp.ok) {
            throw new Error('HTTP-Status ' + resp.status + ' (' + resp.statusText + ')');
        }

        let json;
        try {
            json = await resp.json();
        } catch (e) {
            throw new Error('Antwort ist kein gültiges JSON.');
        }

        // Standard-Struktur: { qrCode: { data:{...}, code:"..." } }
        if (json && json.qrCode && typeof json.qrCode.code === 'string' && json.qrCode.code.trim() !== '') {
            return json.qrCode.code;
        }

        // Fallback: direktes "code"
        if (json && typeof json.code === 'string' && json.code.trim() !== '') {
            return json.code;
        }

        console.error('QR-API Antwort für Tour', tour, json);

        if (json && (json.error || json.message)) {
            throw new Error('Server meldet: ' + (json.error || json.message));
        }

        throw new Error('Server liefert keinen QR-Code für diese Tour (kein Feld "code" in der Antwort).');
    }

    function buildQrImageUrl(content) {
        return 'https://barcodeapi.org/api/qr/' + encodeURIComponent(content);
    }

    async function showQrPopup(doc, tour) {
        try {
            const content = await fetchQrContentForTour(tour);
            const imgUrl = buildQrImageUrl(content);

            const overlay = doc.createElement('div');
            Object.assign(overlay.style, {
                position: 'fixed',
                inset: '0',
                background: 'rgba(0,0,0,0.6)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                zIndex: '999999'
            });

            const box = doc.createElement('div');
            Object.assign(box.style, {
                background: '#fff',
                padding: '12px',
                borderRadius: '4px',
                textAlign: 'center',
                minWidth: '320px',
                boxShadow: '0 0 15px rgba(0,0,0,0.5)',
                fontFamily: 'Arial, sans-serif',
                fontSize: '12px'
            });

            const title = doc.createElement('div');
            title.textContent = 'MDE Freigabe PIN (Zustellung) – Tour ' + tour;
            title.style.marginBottom = '6px';
            box.appendChild(title);

            const img = doc.createElement('img');
            img.src = imgUrl;
            Object.assign(img.style, {
                maxWidth: '360px',
                maxHeight: '360px',
                marginBottom: '6px'
            });
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

            overlay.appendChild(box);
            overlay.addEventListener('click', e => {
                if (e.target === overlay) overlay.remove();
            });

            doc.body.appendChild(overlay);
        } catch (e) {
            alert('Fehler beim Laden des QR-Codes für Tour ' + tour + ':\n' + e.message);
        }
    }

    async function showMultiQrPopup(doc, tours) {
        if (!tours || !tours.length) return;

        const overlay = doc.createElement('div');
        Object.assign(overlay.style, {
            position: 'fixed',
            inset: '0',
            background: 'rgba(0,0,0,0.6)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: '999999'
        });

        const box = doc.createElement('div');
        Object.assign(box.style, {
            background: '#fff',
            padding: '12px',
            borderRadius: '4px',
            textAlign: 'center',
            minWidth: '340px',
            maxWidth: '90vw',
            maxHeight: '90vh',
            overflow: 'auto',
            boxShadow: '0 0 15px rgba(0,0,0,0.5)',
            fontFamily: 'Arial, sans-serif',
            fontSize: '12px'
        });

        const title = doc.createElement('div');
        title.textContent = 'MDE Freigabe PIN – mehrere Touren (' + tours.join(', ') + ')';
        title.style.marginBottom = '6px';
        box.appendChild(title);

        const grid = doc.createElement('div');
        Object.assign(grid.style, {
            display: 'flex',
            flexWrap: 'wrap',
            justifyContent: 'center',
            gap: '10px'
        });
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

        overlay.appendChild(box);
        overlay.addEventListener('click', e => {
            if (e.target === overlay) overlay.remove();
        });

        doc.body.appendChild(overlay);

        const imgUrls = {};

        // QRs nacheinander laden
        for (const tour of tours) {
            const card = doc.createElement('div');
            Object.assign(card.style, {
                border: '1px solid #ccc',
                borderRadius: '4px',
                padding: '4px',
                textAlign: 'center',
                minWidth: '150px'
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
                const img = doc.createElement('img');
                img.src = url;
                Object.assign(img.style, {
                    maxWidth: '160px',
                    maxHeight: '160px'
                });
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
    }

    // ---------- Config-Helfer ----------

    function normalizeName(name) {
        return name.trim().toLowerCase();
    }

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
        try {
            window.localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
        } catch (e) {
            console.error('Konfig speichern fehlgeschlagen:', e);
        }
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

    // ---------- UI in einem bestimmten Dokument aufbauen ----------

    function initInDocument(doc) {
        if (!doc.body || doc.getElementById('tm-sp-panel')) return;

        const css = `
        #tm-sp-panel {
            position: fixed;
            top: 70px;
            right: 10px;
            width: 360px;
            max-height: 80vh;
            overflow: auto;
            background: #f5f5f5;
            border: 1px solid #ccc;
            padding: 6px 8px;
            font-family: Arial, sans-serif;
            font-size: 12px;
            z-index: 9999;
        }
        #tm-sp-panel h3 {
            margin: 0 0 4px 0;
            font-size: 13px;
        }
        #tm-sp-panel button {
            padding: 2px 6px;
            font-size: 11px;
            margin: 2px 2px;
            cursor: pointer;
        }
        #tm-sp-info {
            margin: 4px 0;
        }
        .tm-tour-bubble {
            display:inline-block;
            padding:3px 8px;
            margin:2px 4px 2px 0;
            border-radius:12px;
            background:#7d7d7d; /* grau */
            color:#fff;
            cursor:pointer;
            white-space:nowrap;
            border:1px solid transparent;
        }
        .tm-tour-bubble:hover {
            filter:brightness(1.1);
        }
        .tm-tour-bubble.tm-selected {
            background:#4a4a4a;
            border-color:#000;
        }
        #tm-settings-panel {
            margin-top:6px;
            border-top:1px solid #ccc;
            padding-top:4px;
            display:none;
        }
        #tm-settings-panel table {
            border-collapse: collapse;
            width:100%;
        }
        #tm-settings-panel th,
        #tm-settings-panel td {
            border:1px solid #ddd;
            padding:2px 3px;
            font-size:11px;
        }
        #tm-settings-panel th {
            background:#eee;
        }
        #tm-settings-panel input[type="text"],
        #tm-excel-import {
            width:100%;
            box-sizing:border-box;
            font-size:11px;
        }
        #tm-excel-import {
            height:80px;
            resize:vertical;
        }`;
        const style = doc.createElement('style');
        style.textContent = css;
        doc.head.appendChild(style);

        const panel = doc.createElement('div');
        panel.id = 'tm-sp-panel';

        const title = doc.createElement('h3');
        title.textContent = 'MDE Freigabe PIN (Zustellung)';
        panel.appendChild(title);

        const btnSettings = doc.createElement('button');
        btnSettings.textContent = 'Einstellungen';
        panel.appendChild(btnSettings);

        const infoSpan = doc.createElement('div');
        infoSpan.id = 'tm-sp-info';
        infoSpan.textContent = 'Kein Systempartner ausgewählt.';
        panel.appendChild(infoSpan);

        const bubblesContainer = doc.createElement('div');
        bubblesContainer.id = 'tm-sp-bubbles';
        panel.appendChild(bubblesContainer);

        // Mehrfachauswahl + Button
        const multiDiv = doc.createElement('div');
        multiDiv.style.margin = '4px 0';

        const multiCheckbox = doc.createElement('input');
        multiCheckbox.type = 'checkbox';
        multiCheckbox.id = 'tm-multi-select';

        const multiLabel = doc.createElement('label');
        multiLabel.htmlFor = 'tm-multi-select';
        multiLabel.textContent = ' Mehrfachauswahl';

        const multiShowBtn = doc.createElement('button');
        multiShowBtn.textContent = 'Anzeigen';
        multiShowBtn.style.marginLeft = '4px';

        multiDiv.appendChild(multiCheckbox);
        multiDiv.appendChild(multiLabel);
        multiDiv.appendChild(multiShowBtn);
        panel.appendChild(multiDiv);

        const settingsPanel = doc.createElement('div');
        settingsPanel.id = 'tm-settings-panel';
        panel.appendChild(settingsPanel);

        doc.body.appendChild(panel);

        const selectedTours = new Set();

        function clearSelection() {
            selectedTours.clear();
            const bubbles = bubblesContainer.querySelectorAll('.tm-tour-bubble.tm-selected');
            bubbles.forEach(b => b.classList.remove('tm-selected'));
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
                if (entry && entry.tours) {
                    entry.tours.forEach(t => toursSet.add(t));
                }
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

                if (selectedTours.has(tour)) {
                    bub.classList.add('tm-selected');
                }

                // Linksklick
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

                // Rechtsklick: wie Button, aber für „Power-User“
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

        function renderSettingsPanel() {
            const cfg = loadConfig();
            settingsPanel.innerHTML = '';

            const caption = doc.createElement('div');
            caption.textContent = 'Systempartner / Touren – Einstellungen';
            caption.style.fontWeight = 'bold';
            caption.style.marginBottom = '4px';
            settingsPanel.appendChild(caption);

            const table = doc.createElement('table');
            const thead = doc.createElement('thead');
            const trh = doc.createElement('tr');
            ['Systempartner', 'Tournummern', 'Aktion'].forEach(txt => {
                const th = doc.createElement('th');
                th.textContent = txt;
                trh.appendChild(th);
            });
            thead.appendChild(trh);
            table.appendChild(thead);

            const tbody = doc.createElement('tbody');

            cfg.forEach((item, index) => {
                const tr = doc.createElement('tr');

                const tdName = doc.createElement('td');
                const inpName = doc.createElement('input');
                inpName.type = 'text';
                inpName.value = item.name;
                inpName.disabled = true;
                tdName.appendChild(inpName);

                const tdTours = doc.createElement('td');
                const inpTours = doc.createElement('input');
                inpTours.type = 'text';
                inpTours.value = item.tours.join(', ');
                inpTours.disabled = true;
                tdTours.appendChild(inpTours);

                const tdAct = doc.createElement('td');
                const btnEdit = doc.createElement('button');
                btnEdit.textContent = 'Bearbeiten';
                const btnDel = doc.createElement('button');
                btnDel.textContent = 'Löschen';
                tdAct.appendChild(btnEdit);
                tdAct.appendChild(doc.createTextNode(' '));
                tdAct.appendChild(btnDel);

                let editing = false;

                btnEdit.addEventListener('click', () => {
                    if (!editing) {
                        editing = true;
                        inpName.disabled = false;
                        inpTours.disabled = false;
                        btnEdit.textContent = 'Speichern';
                    } else {
                        const name = inpName.value.trim();
                        const toursArr = inpTours.value
                            .split(/[,\s]+/)
                            .map(t => t.trim())
                            .filter(Boolean);
                        if (!name || !toursArr.length) {
                            alert('Bitte Name und mindestens eine Tournummer angeben.');
                            return;
                        }
                        cfg[index] = { name, tours: toursArr };
                        saveConfig(cfg);
                        editing = false;
                        inpName.disabled = true;
                        inpTours.disabled = true;
                        btnEdit.textContent = 'Bearbeiten';
                        updateBubbles();
                    }
                });

                btnDel.addEventListener('click', () => {
                    if (!confirm('Diesen Systempartner löschen?')) return;
                    cfg.splice(index, 1);
                    saveConfig(cfg);
                    renderSettingsPanel();
                    updateBubbles();
                });

                tr.appendChild(tdName);
                tr.appendChild(tdTours);
                tr.appendChild(tdAct);
                tbody.appendChild(tr);
            });

            const trNew = doc.createElement('tr');

            const tdNewName = doc.createElement('td');
            const inpNewName = doc.createElement('input');
            inpNewName.type = 'text';
            inpNewName.placeholder = 'Systempartner-Name';
            tdNewName.appendChild(inpNewName);

            const tdNewTours = doc.createElement('td');
            const inpNewTours = doc.createElement('input');
            inpNewTours.type = 'text';
            inpNewTours.placeholder = 'Touren, z.B. 627, 623, 638';
            tdNewTours.appendChild(inpNewTours);

            const tdNewAct = doc.createElement('td');
            const btnNewSave = doc.createElement('button');
            btnNewSave.textContent = 'Speichern';
            btnNewSave.disabled = true;
            tdNewAct.appendChild(btnNewSave);

            function checkNewRow() {
                const hasName = inpNewName.value.trim().length > 0;
                const toursArr = inpNewTours.value
                    .split(/[,\s]+/)
                    .map(t => t.trim())
                    .filter(Boolean);
                btnNewSave.disabled = !(hasName && toursArr.length > 0);
            }

            inpNewName.addEventListener('input', checkNewRow);
            inpNewTours.addEventListener('input', checkNewRow);

            btnNewSave.addEventListener('click', () => {
                const name = inpNewName.value.trim();
                const toursArr = inpNewTours.value
                    .split(/[,\s]+/)
                    .map(t => t.trim())
                    .filter(Boolean);
                if (!name || !toursArr.length) {
                    alert('Bitte Name und mindestens eine Tournummer angeben.');
                    return;
                }
                const cfgNow = loadConfig();
                cfgNow.push({ name, tours: toursArr });
                saveConfig(cfgNow);
                renderSettingsPanel();
                updateBubbles();
            });

            trNew.appendChild(tdNewName);
            trNew.appendChild(tdNewTours);
            trNew.appendChild(tdNewAct);
            tbody.appendChild(trNew);

            table.appendChild(tbody);
            settingsPanel.appendChild(table);

            // --- Import aus Excel-Liste (2 Spalten: Systempartner, Tour) ---
            const importTitle = doc.createElement('div');
            importTitle.textContent = 'Import aus Excel-Liste (2 Spalten: Systempartner, Tour):';
            importTitle.style.marginTop = '6px';
            importTitle.style.fontWeight = 'bold';
            settingsPanel.appendChild(importTitle);

            const importHint = doc.createElement('div');
            importHint.textContent = 'In Excel Bereich A:B markieren, kopieren und hier einfügen. Überschrift wird ignoriert.';
            importHint.style.fontSize = '10px';
            importHint.style.marginBottom = '2px';
            settingsPanel.appendChild(importHint);

            const ta = doc.createElement('textarea');
            ta.id = 'tm-excel-import';
            settingsPanel.appendChild(ta);

            const importBtn = doc.createElement('button');
            importBtn.textContent = 'Importieren';
            importBtn.style.marginTop = '4px';
            settingsPanel.appendChild(importBtn);

            importBtn.addEventListener('click', () => {
                const text = ta.value;
                if (!text.trim()) {
                    alert('Bitte erst aus Excel einfügen (Strg+V).');
                    return;
                }

                const lines = text.split(/\r?\n/);
                let cfgNow = loadConfig();
                let addedCount = 0;

                lines.forEach(line => {
                    const trimmed = line.trim();
                    if (!trimmed) return;

                    const parts = line.split(/\t|;/);
                    if (parts.length < 2) return;

                    const name = parts[0].trim();
                    let tour = parts[1].trim();

                    if (!name || !tour || name.toLowerCase().startsWith('systempartner')) {
                        return; // Überschrift oder leer
                    }

                    tour = tour.replace(/[^\dA-Za-z]/g, '');
                    if (!tour) return;

                    const norm = normalizeName(name);
                    let entry = cfgNow.find(p => normalizeName(p.name) === norm);
                    if (!entry) {
                        entry = { name: name, tours: [] };
                        cfgNow.push(entry);
                    }

                    if (!entry.tours.includes(tour)) {
                        entry.tours.push(tour);
                        addedCount++;
                    }
                });

                if (!addedCount) {
                    alert('Es konnten keine neuen Touren importiert werden (evtl. alles schon vorhanden?).');
                    return;
                }

                saveConfig(cfgNow);
                alert('Import abgeschlossen. Neue Touren: ' + addedCount);
                ta.value = '';
                renderSettingsPanel();
                updateBubbles();
            });
        }

        const sel = doc.querySelector('select[name="systempartner"]');
        if (sel) {
            sel.addEventListener('change', updateBubbles);
        }

        btnSettings.addEventListener('click', () => {
            if (settingsPanel.style.display === 'none' || !settingsPanel.style.display) {
                renderSettingsPanel();
                settingsPanel.style.display = 'block';
            } else {
                settingsPanel.style.display = 'none';
            }
        });

        // Mehrfachauswahl an/aus
        multiCheckbox.addEventListener('change', () => {
            if (!multiCheckbox.checked) {
                clearSelection(); // => alle Markierungen weg
            }
        });

        // Button "Anzeigen"
        multiShowBtn.addEventListener('click', () => {
            if (!multiCheckbox.checked) {
                alert('Bitte zuerst „Mehrfachauswahl“ aktivieren.');
                return;
            }
            const list = selectedTours.size ? Array.from(selectedTours) : [];
            if (!list.length) {
                alert('Bitte mindestens eine Tour markieren.');
                return;
            }
            showMultiQrPopup(doc, list);
        });

        updateBubbles();
    }

    // ---------- Starter: richtigen Frame immer wieder prüfen ----------

    setInterval(() => {
        try {
            const docs = [];
            docs.push(document);

            if (window.frames && window.frames.length) {
                for (let i = 0; i < window.frames.length; i++) {
                    const f = window.frames[i];
                    try {
                        if (f.document) docs.push(f.document);
                    } catch (e) { /* anderer Origin */ }
                }
            }

            for (const doc of docs) {
                if (!doc || !doc.body) continue;
                if (/Eingangsmengenabgleich/.test(doc.body.textContent || '')) {
                    const sel = doc.querySelector('select[name="systempartner"]');
                    if (sel) {
                        initInDocument(doc);
                    }
                }
            }
        } catch (e) {
            console.error(e);
        }
    }, 800);

})();
Gefällt dir diese Persönlichkeit?






Keine ausgewähltKeine ausgewählt
ChatGPT kann Fehler machen. Überprüfe wichtige Informationen. Siehe .
