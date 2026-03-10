// ==UserScript==
// @name         DPD Depot Portal - IHP Button
// @namespace    http://tampermonkey.net/
// @version      1.2
// @description  Fügt einen IHP Button neben der Lager-Beschreibung hinzu, wenn Scanart 8 und ZC/DC 24 gefunden wird
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-depotportal.ihp.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-depotportal.ihp.user.js
// @author       Thiemo Schöler
// @match        https://depotportal.dpd.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    const IHP_URL_BASE = 'https://dpdgroup.eu.wizyvision.app/posts/3/search-results?q=';

    function findParcelNumber() {
        const parcelNoLink = document.querySelector('[href*="parcelNo="]');
        if (parcelNoLink) {
            const match = parcelNoLink.href.match(/parcelNo=(\d+)/);
            if (match) return match[1];
        }

        const parcelHeader = document.querySelector('dpd-tracking h4');
        if (parcelHeader) {
            const match = parcelHeader.textContent.match(/\d{14}/);
            if (match) return match[0];
        }

        const h4List = document.querySelectorAll('h4');
        for (const h4 of h4List) {
            const match = h4.textContent.match(/\d{14}/);
            if (match) return match[0];
        }

        const parcelInput = document.querySelector('input[type="text"], input[placeholder*="Paket"], input[name*="parcel"], input[name*="tracking"]');
        if (parcelInput && parcelInput.value) {
            return parcelInput.value.trim();
        }

        const urlParams = new URLSearchParams(window.location.search);
        const parcelParam = urlParams.get('parcel') || urlParams.get('nummer') || urlParams.get('q');
        if (parcelParam) {
            return parcelParam.trim();
        }

        return null;
    }

    function findLagerRow() {
        const rows = document.querySelectorAll('tr');

        for (const row of rows) {
            const cells = row.querySelectorAll('td');
            if (cells.length < 5) continue;

            const scanartCell = cells[2];
            const zcDcCell = cells[4];
            const beschreibungCell = cells[3];

            const scanart = scanartCell.textContent.trim();
            const zcDc = zcDcCell.textContent.trim();
            const beschreibungDiv = beschreibungCell.querySelector('div');
            const beschreibung = beschreibungDiv ? beschreibungDiv.textContent.trim() : beschreibungCell.textContent.trim();

            if (scanart === '8' && (zcDc === '24' || zcDc === '024')) {
                if (beschreibung === 'Lager') {
                    return { row, zcDc: zcDc };
                }
            }
        }
        return null;
    }

    function addIHPButton(row, zcDc) {
        const cells = row.querySelectorAll('td');
        const targetCell = cells[3];

        if (targetCell.querySelector('.ihp-btn')) return;

        const beschreibungDiv = targetCell.querySelector('div');

        const btnContainer = document.createElement('span');
        btnContainer.style.display = 'inline-flex';
        btnContainer.style.alignItems = 'center';
        btnContainer.style.verticalAlign = 'middle';
        btnContainer.style.marginLeft = '12px';
        btnContainer.style.lineHeight = '1';

        const btn = document.createElement('a');
        btn.className = 'ihp-btn';
        btn.href = '#';
        btn.title = 'Im Wizyvision öffnen';
        btn.style.display = 'inline-flex';
        btn.style.alignItems = 'center';
        btn.style.justifyContent = 'center';
        btn.style.padding = '2px 10px';
        btn.style.cursor = 'pointer';
        btn.style.backgroundColor = '#e63946';
        btn.style.color = 'white';
        btn.style.textDecoration = 'none';
        btn.style.borderRadius = '4px';
        btn.style.fontSize = '11px';
        btn.style.fontWeight = '600';
        btn.style.fontFamily = 'system-ui, -apple-system, sans-serif';
        btn.style.letterSpacing = '0.3px';
        btn.style.boxShadow = '0 1px 3px rgba(0,0,0,0.2)';
        btn.style.transition = 'all 0.2s ease';
        btn.style.whiteSpace = 'nowrap';
        btn.style.height = 'auto';
        btn.style.minHeight = '0';
        btn.innerHTML = '<span style="margin-right:4px">📦</span>IHP';

        btn.addEventListener('mouseenter', () => {
            btn.style.backgroundColor = '#d62839';
            btn.style.transform = 'translateY(-1px)';
            btn.style.boxShadow = '0 2px 6px rgba(0,0,0,0.25)';
        });

        btn.addEventListener('mouseleave', () => {
            btn.style.backgroundColor = '#e63946';
            btn.style.transform = 'translateY(0)';
            btn.style.boxShadow = '0 1px 3px rgba(0,0,0,0.2)';
        });

        btn.addEventListener('click', (e) => {
            e.preventDefault();
            const parcelNumber = findParcelNumber();
            if (parcelNumber) {
                const url = `${IHP_URL_BASE}${parcelNumber}`;
                window.open(url, '_blank');
            } else {
                alert('Paketnummer konnte nicht ermittelt werden');
            }
        });

        btnContainer.appendChild(btn);

        if (beschreibungDiv) {
            beschreibungDiv.parentNode.insertBefore(btnContainer, beschreibungDiv.nextSibling);
        } else {
            targetCell.appendChild(btnContainer);
        }
    }

    function init() {
        const lagerRow = findLagerRow();
        if (lagerRow) {
            addIHPButton(lagerRow.row, lagerRow.zcDc);
        }
    }

    const observer = new MutationObserver(() => {
        init();
    });

    observer.observe(document.body, {
        childList: true,
        subtree: true
    });

    init();
})();
