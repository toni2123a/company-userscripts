// ==UserScript==
// @name         DPD Content Check + Lost&Found Mail abkürzung
// @namespace    bodo-scripts
// @version      1.3
// @description  Neue Spalte mit L&F-Mail-Button + auf Lost&Found beide "E-Mail versenden"-Buttons automatisch klicken
// @match        https://dpdgroup.eu.wizyvision.app/*
// @match        https://lostandfound.dpdgroup.com/*
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.l&f.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.l&f.user.js
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const HOST_DPD = 'dpdgroup.eu.wizyvision.app';
    const HOST_LF  = 'lostandfound.dpdgroup.com';

    function waitForBody(callback) {
        if (document.body) {
            callback();
        } else {
            const obs = new MutationObserver(() => {
                if (document.body) {
                    obs.disconnect();
                    callback();
                }
            });
            obs.observe(document.documentElement, { childList: true, subtree: true });
        }
    }

    // ---------------------------------------------------------
    // TEIL 1: DPD-Posts-Seite -> neue Spalte + Button
    // ---------------------------------------------------------
    function isPostsPage() {
        return location.hostname === HOST_DPD && location.pathname.startsWith('/posts');
    }

    function getWVFromRow(row) {
        const link = row.querySelector('a[data-testid="LinkComp"]');
        if (!link) return null;
        const txt = (link.textContent || '').trim();
        const m = txt.match(/WV-\d+/);
        return m ? m[0] : txt || null;
    }

    function addHeaderColumn() {
        const headerRow = document.querySelector('thead.MuiTableHead-root tr');
        if (!headerRow) return;
        if (headerRow.dataset.linkedInfoHeader === '1') return;

        const ths = Array.from(headerRow.querySelectorAll('th'));
        if (!ths.length) return;

        const erfasstTh = ths.find(th => th.textContent.trim().toUpperCase() === 'ERFASST');
        if (!erfasstTh) return;

        const newTh = document.createElement('th');
        newTh.className = erfasstTh.className;
        newTh.textContent = 'VERKNÜPFTE INFOS';

        headerRow.insertBefore(newTh, erfasstTh);
        headerRow.dataset.linkedInfoHeader = '1';
    }

    function addBodyCells() {
        const rows = document.querySelectorAll('tbody.MuiTableBody-root tr');
        if (!rows.length) return;

        rows.forEach(row => {
            if (row.dataset.linkedInfoAdded === '1') return;

            const cells = row.querySelectorAll('td[data-testid="StickyHeadTable"]');
            if (!cells.length) return;

            const erfasstCell = cells[cells.length - 1];

            const newTd = document.createElement('td');
            newTd.className = erfasstCell.className;
            newTd.setAttribute('data-testid', 'LinkedInfoCell');

            const btn = document.createElement('button');
btn.textContent = 'L&F-Mail';

// Optisch schöner Button
btn.style.cursor = 'pointer';
btn.style.fontSize = '11px';
btn.style.padding = '3px 8px';
btn.style.border = '1px solid #2b6cb0';
btn.style.background = '#4299e1';
btn.style.color = 'white';
btn.style.borderRadius = '6px';
btn.style.display = 'inline-block';
btn.style.marginTop = '4px';
btn.style.fontWeight = '500';
btn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.1)';
btn.style.transition = 'all 0.15s ease';

// Hover Effekt
btn.addEventListener('mouseover', () => {
    btn.style.background = '#2b6cb0';
});
btn.addEventListener('mouseout', () => {
    btn.style.background = '#4299e1';
});

            btn.addEventListener('click', (ev) => {
                ev.stopPropagation();
                const wv = getWVFromRow(row);
                if (!wv) {
                    alert('WV-Nummer nicht gefunden.');
                    return;
                }
                const url = 'https://lostandfound.dpdgroup.com/search?ids='
                    + encodeURIComponent(wv)
                    + '&autoMail=1';
                window.open(url, '_blank');
            });

            newTd.appendChild(btn);
            erfasstCell.parentNode.insertBefore(newTd, erfasstCell);

            row.dataset.linkedInfoAdded = '1';
        });
    }

    function runOnDpdPage() {
        if (!isPostsPage()) return;

        const headerRow = document.querySelector('thead.MuiTableHead-root tr');
        const body = document.querySelector('tbody.MuiTableBody-root');
        if (!headerRow || !body) return;

        addHeaderColumn();
        addBodyCells();
    }

    function initDpd() {
        let lastPath = location.pathname;

        const mo = new MutationObserver(() => {
            runOnDpdPage();
        });
        mo.observe(document.body, { childList: true, subtree: true });

        setInterval(() => {
            if (location.pathname !== lastPath) {
                lastPath = location.pathname;
                runOnDpdPage();
            }
        }, 1000);

        runOnDpdPage();
    }

    // ---------------------------------------------------------
    // TEIL 2: Lost&Found-Seiten
    // ---------------------------------------------------------
   // Router für Lost&Found: reagiert auf Seitenwechsel /search -> /mailPdf
let lfRouterStarted = false;

function initLostAndFoundRouter() {
    if (lfRouterStarted) return;
    lfRouterStarted = true;

    let lastPath = '';

    function checkRoute() {
        if (location.hostname !== HOST_LF) return;

        const currentPath = location.pathname;
        if (currentPath === lastPath) return;
        lastPath = currentPath;

        console.log('[TM] L&F: Route gewechselt zu', currentPath);

        if (currentPath.startsWith('/search')) {
            initLfSearch();
        } else if (currentPath.startsWith('/mailPdf')) {
            initLfMailPdf();
        }
    }

    // direkt einmal + dann alle 500ms
    checkRoute();
    setInterval(checkRoute, 500);
}


    // /search?ids=WV-...&autoMail=1  -> ersten Button anklicken
    function initLfSearch() {
        const params = new URLSearchParams(location.search);
        const autoMail = params.get('autoMail');
        if (autoMail !== '1') return;

        const wv = (params.get('ids') || '').trim();
        if (!wv) return;

        function clickEmailButton() {
            const selector = 'button[data-cy="sendEmail_' + wv + '"]';
            let btn = document.querySelector(selector);

            if (!btn) {
                btn = document.querySelector('button[data-cy^="sendEmail_"]');
            }

            if (btn) {
                console.log('[TM] search: klicke sendEmail-Button', btn);
                btn.click();
                return true;
            }
            console.log('[TM] search: noch kein sendEmail-Button gefunden');
            return false;
        }

        // Polling, falls das Ergebnis träge lädt
        let tries = 0;
        const maxTries = 50; // ~10 Sekunden
        const timer = setInterval(() => {
            tries++;
            if (clickEmailButton() || tries >= maxTries) {
                clearInterval(timer);
            }
        }, 200);
    }

    // /mailPdf/...  -> "E-Mail versenden" + "Ok" klicken
 // /mailPdf/...  -> "E-Mail versenden" + "Ok" genau 1x klicken
// /mailPdf/...  -> "E-Mail versenden" genau 1x klicken, OK bleibt sichtbar
function initLfMailPdf() {
    console.log('[TM] mailPdf: init auf', location.href);

    if (document.body.dataset.tmMailPdfDone === '1') {
        console.log('[TM] mailPdf: bereits ausgeführt, breche ab.');
        return;
    }
    document.body.dataset.tmMailPdfDone = '1';

    let mailClicked = false;

    function tryMailClick() {
        if (mailClicked) return;

        // Nur den echten Senden-Button klicken (type="submit")
        const mailBtn = document.querySelector('button[id^="mailIds-"][type="submit"]');
        if (mailBtn) {
            console.log('[TM] mailPdf: klicke Mail-Button', mailBtn.id, mailBtn);
            mailBtn.click();
            mailClicked = true;

            // Wenn Mail geklickt wurde, Observer später automatisch stoppen
            setTimeout(() => {
                observer.disconnect();
                console.log('[TM] mailPdf: Observer gestoppt (Mail gesendet).');
            }, 2000);

        } else {
            console.log('[TM] mailPdf: Mail-Button (type=submit) noch nicht da');
        }
    }

    // 1x direkt probieren
    tryMailClick();

    // Danach DOM-Änderungen überwachen
    const observer = new MutationObserver(() => {
        tryMailClick();
    });

    observer.observe(document.body, { childList: true, subtree: true });

    // Sicherheits-Timeout: nach 20s Observer stoppen
    setTimeout(() => {
        if (!mailClicked) {
            console.log('[TM] mailPdf: Timeout, stoppe Observer.');
            observer.disconnect();
        }
    }, 20000);
}

    // ---------------------------------------------------------
    // Einstieg
    // ---------------------------------------------------------
  waitForBody(() => {
    if (location.hostname === HOST_DPD) {
        initDpd();
    } else if (location.hostname === HOST_LF) {
        initLostAndFoundRouter();
    }
});

})();
