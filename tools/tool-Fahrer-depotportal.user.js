// ==UserScript==
// @name         Anzeige der Fahrernamen im Depotportal 
// @namespace    bodo.tools
// @version      1.6
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-Fahrer-depotportal.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-Fahrer-depotportal.user.js
// @description  Button neben "Beleg" -> Fahrer aus DPD360 nur bei Klick laden
// @match        https://depotportal.dpd.com/*/tracking/parcels/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @connect      dpd360.dpd.de
// @run-at       document-idle
// ==/UserScript==

(function () {
  "use strict";

  const LOGP = "[Fahrer-Lookup]";

  // ---- GM_addStyle Fallback ----
  function addStyle(css) {
    if (typeof GM_addStyle === "function") {
      try { GM_addStyle(css); return; } catch (e) { console.warn(LOGP, "GM_addStyle fail:", e); }
    }
    const s = document.createElement("style");
    s.type = "text/css";
    s.appendChild(document.createTextNode(css));
    document.head.appendChild(s);
  }

  addStyle(`
    .bodo-driver-btn{margin-left:.5rem;padding:.2rem .45rem;font:500 11px/1 system-ui;border:1px solid #bbb;border-radius:4px;background:#fff;cursor:pointer}
    .bodo-driver-btn[disabled]{opacity:.6;cursor:wait}
    .bodo-driver-badge{display:inline-block;margin-left:.4rem;padding:.2rem .45rem;border-radius:999px;border:1px solid #d0d0d0;background:#f6f6f6;font:600 11px/1.2 system-ui}
    .bodo-driver-error{color:#a40000;font-weight:700}
  `);

  const DPD_BASE = "https://dpd360.dpd.de/order/order_view.aspx?parcelno=";
  const DIGITS = /\b\d{14}\b/;

  function getParcelNumber() {
    for (const el of document.querySelectorAll("h4")) {
      const m = (el.textContent || "").replace(/\s+/g," ").trim().match(DIGITS);
      if (m) return m[0];
    }
    return "";
  }

  function findBelegAnchors(root = document) {
    const anchors = [];
    for (const a of root.querySelectorAll("a")) {
      const text = (a.textContent || "").trim();
      if ((text === "Beleg" || /\bBeleg\b/.test(text)) && !a.dataset.bodoDriverInitialized) {
        anchors.push(a);
      }
    }
    return anchors;
  }

  function injectForAnchor(anchor) {
    const parcelNo = getParcelNumber();
    if (!parcelNo) { console.warn(LOGP, "keine Paketnummer gefunden"); return; }

    anchor.dataset.bodoDriverInitialized = "1";

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "bodo-driver-btn";
    btn.textContent = "Fahrer";

    const badge = document.createElement("span");
    badge.className = "bodo-driver-badge";
    badge.textContent = "—";

    anchor.insertAdjacentElement("afterend", btn);
    btn.insertAdjacentElement("afterend", badge);

    btn.addEventListener("click", () => {
      const old = btn.textContent;
      btn.disabled = true;
      btn.textContent = "…";
      badge.classList.remove("bodo-driver-error");
      badge.textContent = "…";

      fetchDriver(parcelNo)
        .then(name => { badge.textContent = name || (badge.classList.add("bodo-driver-error"), "nicht gefunden"); })
        .catch(err => { console.error(LOGP, err); badge.classList.add("bodo-driver-error"); badge.textContent = "Fehler"; })
        .finally(() => { btn.disabled = false; btn.textContent = old; });
    });
  }

  function fetchDriver(parcelNo) {
    const url = DPD_BASE + encodeURIComponent(parcelNo);
    return new Promise((resolve, reject) => {
      if (typeof GM_xmlhttpRequest !== "function") {
        return reject(new Error("GM_xmlhttpRequest nicht verfügbar (anderen Manager/Grants prüfen)"));
      }
      GM_xmlhttpRequest({
        method: "GET",
        url,
        withCredentials: true,
        timeout: 15000,
        onload: (res) => {
          if (res.status !== 200) return reject(new Error("HTTP " + res.status));
          const doc = new DOMParser().parseFromString(res.responseText, "text/html");
          let name = doc.querySelector("#ContentPlaceHolder1_labDriver")?.textContent?.trim() || "";
          if (!name) {
            const cands = [...doc.querySelectorAll("span,div,b,strong")].filter(e => /Fahrer/i.test(e.textContent||""));
            for (const c of cands) {
              const t = c.parentElement?.querySelector("span,div")?.textContent?.trim();
              if (t && !/Fahrer/i.test(t)) { name = t; break; }
            }
          }
          resolve(name);
        },
        onerror: () => reject(new Error("Netzwerkfehler")),
        ontimeout: () => reject(new Error("Timeout"))
      });
    });
  }

  function scan(root = document) { findBelegAnchors(root).forEach(injectForAnchor); }

  function setupObserver() {
    const obs = new MutationObserver(m => { for (const x of m) if (x.addedNodes.length) { scan(document); break; } });
    obs.observe(document.body, { childList: true, subtree: true });
  }

  // Start
  console.log(LOGP, "init");
  scan();
  setupObserver();
})();
