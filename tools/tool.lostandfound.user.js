// ==UserScript==
// @name         DPD Content Check + Lost&Found
// @namespace    bodo-scripts
// @version      2.0
// @description  Neue Spalte mit L&F-Mail-Button + Status Änderung auf abgeschlossen
// @match        https://dpdgroup.eu.wizyvision.app/*
// @match        https://lostandfound.dpdgroup.com/*
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.lostandfound.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool.lostandfound.user.js
// @run-at       document-start
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const HOST_DPD = 'dpdgroup.eu.wizyvision.app';
    const HOST_LF  = 'lostandfound.dpdgroup.com';

    // ---------------------------------------------------------
    // Hilfsfunktionen
    // ---------------------------------------------------------
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
    //  TEIL 1: DPD – Liste + Detail
    // ---------------------------------------------------------

   function isDpdListPage() {
    // alle Listen-Ansichten, z.B. /posts/3?statusId=7
    return location.hostname === HOST_DPD &&
           location.pathname.startsWith('/posts') &&
           !location.pathname.includes('/view/');
}

function isDpdDetailPage() {
    // Detailseite: /posts/.../view/...
    return location.hostname === HOST_DPD &&
           location.pathname.startsWith('/posts') &&
           location.pathname.includes('/view/');
}


    function getWVFromRow(row) {
        const link = row.querySelector('a[data-testid="LinkComp"]');
        if (!link) return null;
        const txt = (link.textContent || '').trim();
        const m = txt.match(/WV-\d+/);
        return m ? m[0] : txt || null;
    }

    // ---- Spalten-Header in der Liste ----
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

    // ---- Zeilen in der Liste ergänzen (Buttons) ----
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

            // --- Button 1: L&F-Mail ---
            const mailBtn = document.createElement('button');
            mailBtn.textContent = 'L&F-Mail';

            mailBtn.style.cursor = 'pointer';
            mailBtn.style.fontSize = '11px';
            mailBtn.style.padding = '3px 8px';
            mailBtn.style.border = '1px solid #2b6cb0';
            mailBtn.style.background = '#4299e1';
            mailBtn.style.color = 'white';
            mailBtn.style.borderRadius = '6px';
            mailBtn.style.display = 'inline-block';
            mailBtn.style.marginTop = '4px';
            mailBtn.style.fontWeight = '500';
            mailBtn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.1)';
            mailBtn.style.transition = 'all 0.15s ease';

            mailBtn.addEventListener('mouseover', () => {
                mailBtn.style.background = '#2b6cb0';
            });
            mailBtn.addEventListener('mouseout', () => {
                mailBtn.style.background = '#4299e1';
            });

            mailBtn.addEventListener('click', (ev) => {
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

            newTd.appendChild(mailBtn);

            // --- Button 2: Status "Geschlossen" setzen ---
            const finishBtn = document.createElement('button');
            finishBtn.textContent = 'Abschließen';

            finishBtn.style.cursor = 'pointer';
            finishBtn.style.fontSize = '11px';
            finishBtn.style.padding = '3px 8px';
            finishBtn.style.border = '1px solid #38a169';
            finishBtn.style.background = '#48bb78';
            finishBtn.style.color = 'white';
            finishBtn.style.borderRadius = '6px';
            finishBtn.style.display = 'inline-block';
            finishBtn.style.marginTop = '4px';
            finishBtn.style.marginLeft = '4px';
            finishBtn.style.fontWeight = '500';
            finishBtn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.1)';
            finishBtn.style.transition = 'all 0.15s ease';

            finishBtn.addEventListener('mouseover', () => {
                finishBtn.style.background = '#38a169';
            });
            finishBtn.addEventListener('mouseout', () => {
                finishBtn.style.background = '#48bb78';
            });

       finishBtn.addEventListener('click', (ev) => {
    ev.stopPropagation();

    const detailLink = row.querySelector('a[data-testid="LinkComp"]');
    if (!detailLink) {
        alert('Detail-Link nicht gefunden.');
        return;
    }

    // Post-ID aus dem Link holen: /posts/3/view/<ID>?...
    const href = detailLink.href;
    const m = href.match(/\/view\/([^/?]+)/);
    if (!m) {
        alert('Konnte Post-ID nicht aus dem Link lesen.');
        return;
    }
    const postId = m[1];

    // merken, dass dieser Post auto-geschlossen werden soll
    try {
        sessionStorage.setItem('tmAutoFinish:' + postId, '1');
    } catch (e) {
        console.log('[TM] sessionStorage-Fehler:', e);
    }

    // wie normaler Klick -> React-Routing bleibt schnell
    detailLink.click();
});


            newTd.appendChild(finishBtn);

            erfasstCell.parentNode.insertBefore(newTd, erfasstCell);
            row.dataset.linkedInfoAdded = '1';
        });
    }

    function runOnDpdList() {
        if (!isDpdListPage()) return;

        const headerRow = document.querySelector('thead.MuiTableHead-root tr');
        const body = document.querySelector('tbody.MuiTableBody-root');
        if (!headerRow || !body) return;

        addHeaderColumn();
        addBodyCells();
    }

    // ---- Detailseite: Status automatisch auf "Geschlossen" setzen ----
  function initDpdDetailAutoFinish() {
    if (!isDpdDetailPage()) return;

    // aktuelle Post-ID aus URL holen
    const m = location.pathname.match(/\/view\/([^/?]+)/);
    if (!m) return;
    const postId = m[1];

    // nur ausführen, wenn in der Liste vorher "Abschließen" geklickt wurde
    let shouldFinish = false;
    try {
        shouldFinish = sessionStorage.getItem('tmAutoFinish:' + postId) === '1';
    } catch (e) {
        console.log('[TM] DPD Detail: sessionStorage-Fehler', e);
    }
    if (!shouldFinish) return;

    console.log('[TM] DPD Detail: autoFinish für Post', postId, 'aktiv auf', location.href);

    // Schlüssel wieder löschen, damit Ablauf nur 1x pro Post läuft
    try {
        sessionStorage.removeItem('tmAutoFinish:' + postId);
    } catch (e) {
        console.log('[TM] DPD Detail: Konnte tmAutoFinish nicht entfernen', e);
    }

    let tabClicked = false;
    let dropdownOpened = false;
    let dropdownOpenedAt = 0;
    let statusSet = false;
    let tries = 0;
    const maxTries = 50;   // ~10 Sekunden bei 200 ms Intervall

    const timer = setInterval(() => {
        tries++;
        console.log('[TM] DPD Detail: Versuch', tries);

        // 1. Tab "Einzelheiten" sicher aktiv
        const tabBtn = document.querySelector('button[data-testid="PostViewCard"]');
        if (tabBtn && !tabClicked) {
            ['mousedown', 'mouseup', 'click'].forEach(type => {
                tabBtn.dispatchEvent(new MouseEvent(type, {
                    bubbles: true,
                    cancelable: true,
                    view: window
                }));
            });
            tabClicked = true;
            console.log('[TM] DPD Detail: Tab "Einzelheiten" angeklickt');
        }

        // 2. Dropdown öffnen (mit echten MouseEvents)
        if (!dropdownOpened) {
            const p = document.querySelector('p[data-testid="StatusField"]');
            const trigger = p ? p.closest('div[role="button"][aria-haspopup="listbox"]') : null;

            if (trigger) {
                console.log('[TM] DPD Detail: öffne Status-Dropdown', trigger);
                ['mousedown', 'mouseup', 'click'].forEach(type => {
                    trigger.dispatchEvent(new MouseEvent(type, {
                        bubbles: true,
                        cancelable: true,
                        view: window
                    }));
                });
                dropdownOpened = true;
                dropdownOpenedAt = Date.now();
            } else {
                console.log('[TM] DPD Detail: Dropdown-Trigger noch nicht gefunden');
            }
            return; // erst im nächsten Tick nach Optionen schauen
        }

        // 3. Nach kurzer Wartezeit "Geschlossen" auswählen
        if (!statusSet && Date.now() - dropdownOpenedAt >= 200) {
            const menu = document.querySelector('#menu-statusId');
            if (!menu) {
                console.log('[TM] DPD Detail: Menu #menu-statusId noch nicht im DOM');
                return;
            }

            const allItems = menu.querySelectorAll('li[data-testid="StatusField"]');
            console.log('[TM] DPD Detail: gefundene Optionen:', allItems.length);

            const target = Array.from(allItems).find(el =>
                el.getAttribute('data-value') === '8' ||
                (el.textContent || '').trim() === 'Geschlossen'
            );

            if (target) {
    console.log('[TM] DPD Detail: klicke Option "Geschlossen"', target);
    ['mousedown', 'mouseup', 'click'].forEach(type => {
        target.dispatchEvent(new MouseEvent(type, {
            bubbles: true,
            cancelable: true,
            view: window
        }));
    });

    statusSet = true;
    clearInterval(timer);
    console.log('[TM] DPD Detail: Status gesetzt, Haupt-Timer gestoppt.');

    // Neuer Poll: warten, bis der Status-Text wirklich "Geschlossen" ist
    let backTries = 0;
    const backMaxTries = 50; // ~10 Sekunden

    const backTimer = setInterval(() => {
        backTries++;

        const statusTextEl = document.querySelector('p[data-testid="StatusField"]');
        const statusText = statusTextEl ? statusTextEl.textContent.trim() : '';
        console.log('[TM] DPD Detail: BackCheck', backTries, 'Status=', statusText);

        if (statusText === 'Geschlossen') {
            const backBtn = document.querySelector('button[data-testid="BackButton"]');
            if (backBtn) {
                console.log('[TM] DPD Detail: Status = Geschlossen, klicke Back-Button', backBtn);
                ['mousedown', 'mouseup', 'click'].forEach(type => {
                    backBtn.dispatchEvent(new MouseEvent(type, {
                        bubbles: true,
                        cancelable: true,
                        view: window
                    }));
                });
            } else {
                console.log('[TM] DPD Detail: Back-Button (data-testid="BackButton") nicht gefunden');
            }
            clearInterval(backTimer);
            return;
        }

        if (backTries >= backMaxTries) {
            console.log('[TM] DPD Detail: BackCheck Timeout, Back-Button wird nicht geklickt.');
            clearInterval(backTimer);
        }
    }, 200);

    return;
}
 else {
                console.log('[TM] DPD Detail: Option "Geschlossen" noch nicht gefunden');
            }
        }

        if (tries >= maxTries) {
            console.log('[TM] DPD Detail: Abbruch nach maxTries, Status nicht gesetzt.');
            clearInterval(timer);
        }
    }, 200);
}



    // ---- Router für DPD (Liste vs. Detail) ----
   let dpdRouterStarted = false;
function initDpdRouter() {
    if (dpdRouterStarted) return;
    dpdRouterStarted = true;

    let lastPath = '';

    function checkRoute() {
        if (location.hostname !== HOST_DPD) return;

        const currentPath = location.pathname;
        if (currentPath === lastPath) return;
        lastPath = currentPath;

        console.log('[TM] DPD: Route gewechselt zu', currentPath);

        if (isDpdListPage()) {
            runOnDpdList();
        } else if (isDpdDetailPage()) {
            initDpdDetailAutoFinish();
        }
    }

    checkRoute();
    setInterval(checkRoute, 500);

    const mo = new MutationObserver(() => {
        if (isDpdListPage()) {
            runOnDpdList();
        }
    });
    mo.observe(document.body, { childList: true, subtree: true });
}

    // ---------------------------------------------------------
    //  TEIL 2: Lost&Found – Mail automatisch auslösen
    // ---------------------------------------------------------

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
                console.log('[TM] L&F /search: klicke sendEmail-Button', btn);
                btn.click();
                return true;
            }
            console.log('[TM] L&F /search: noch kein sendEmail-Button gefunden');
            return false;
        }

        let tries = 0;
        const maxTries = 50; // ~10s
        const timer = setInterval(() => {
            tries++;
            if (clickEmailButton() || tries >= maxTries) {
                clearInterval(timer);
            }
        }, 200);
    }

    // /mailPdf/...  -> "E-Mail versenden" genau 1x klicken, OK-Popup bleibt
    function initLfMailPdf() {
        console.log('[TM] L&F mailPdf: init auf', location.href);

        if (document.body.dataset.tmMailPdfDone === '1') {
            console.log('[TM] L&F mailPdf: bereits ausgeführt, breche ab.');
            return;
        }
        document.body.dataset.tmMailPdfDone = '1';

        let mailClicked = false;

        function tryMailClick() {
            if (mailClicked) return;

            const mailBtn = document.querySelector('button[id^="mailIds-"][type="submit"]');
            if (mailBtn) {
                console.log('[TM] L&F mailPdf: klicke Mail-Button', mailBtn.id, mailBtn);
                mailBtn.click();
                mailClicked = true;

                setTimeout(() => {
                    observer.disconnect();
                    console.log('[TM] L&F mailPdf: Observer gestoppt (Mail gesendet).');
                }, 2000);
            } else {
                console.log('[TM] L&F mailPdf: Mail-Button (type=submit) noch nicht da');
            }
        }

        const observer = new MutationObserver(() => {
            tryMailClick();
        });

        tryMailClick();
        observer.observe(document.body, { childList: true, subtree: true });

        setTimeout(() => {
            if (!mailClicked) {
                console.log('[TM] L&F mailPdf: Timeout, Observer gestoppt.');
                observer.disconnect();
            }
        }, 20000);
    }

    // Router für Lost&Found (/search vs. /mailPdf)
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

        checkRoute();
        setInterval(checkRoute, 500);
    }

    // ---------------------------------------------------------
    // Einstieg
    // ---------------------------------------------------------
  waitForBody(() => {
    if (location.hostname === HOST_DPD) {
        initDpdRouter();
    } else if (location.hostname === HOST_LF) {
        initLostAndFoundRouter();
    }
});

})();
