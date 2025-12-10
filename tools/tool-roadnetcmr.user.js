// ==UserScript==
// @name         Roadnet – Anhänge Filter + Frachtbrief PDF Tabs 
// @namespace    https://roadnet.dpdgroup.com/
// @version      2.0.0
// @description  Filter für Anhänge anhand sichtbarer (nicht .hidden) Labels inkl. Negativ-Filter (!CMR) + Frachtbrief: PDFs im Tab, Multi ohne ZIP.
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-roadnetcmr.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-roadnetcmr.user.js
// @match        https://roadnet.dpdgroup.com/execution/route_legs*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(() => {
  "use strict";

  // =========================
  // Teil A: Anhänge-Filter
  // =========================
  const COL_TH_SELECTOR =
    "th#frm_route_legs\\:tbl\\:col_attachmentTypes, th[id$=':col_attachmentTypes']";
  const STATE = { value: "" };

  const norm = (s) => (s ?? "").replace(/\s+/g, " ").trim();
  const clean = (s) => norm(String(s || "").replace(/\u00A0/g, " ").replace(/\?\!/g, ""));

  const styleId = "tm-att-filter-style-20";
  if (!document.getElementById(styleId)) {
    const st = document.createElement("style");
    st.id = styleId;
    st.textContent = `
      .tm-hidden { display: none !important; }
      .tm-cmr-rowbtn{
        cursor:pointer; margin-right:6px; border:1px solid #999; background:#f5f5f5;
        border-radius:2px; height:18px; line-height:16px; padding:0 6px; font-size:12px;
      }
    `;
    document.head.appendChild(st);
  }

  function isVisibleEl(el) {
    if (!el) return false;
    if (el.classList?.contains("hidden")) return false;
    const cs = getComputedStyle(el);
    if (cs.display === "none" || cs.visibility === "hidden" || cs.opacity === "0") return false;
    return true;
  }

  function attachmentsVisibleText(cell) {
    if (!cell) return "";
    const spans = Array.from(cell.querySelectorAll("span")).filter(isVisibleEl);

    const parts = [];
    for (const sp of spans) {
      const t = clean(sp.textContent);
      if (t) parts.push(t);
    }

    // fallback für Icons/Tooltips (nur sichtbare Elemente)
    const attrs = ["title", "aria-label", "alt"];
    const visNodes = Array.from(cell.querySelectorAll("*")).filter(isVisibleEl);
    for (const el of visNodes) {
      for (const a of attrs) {
        const v = clean(el.getAttribute?.(a));
        if (v) parts.push(v);
      }
    }

    return clean([...new Set(parts)].join(" "));
  }

  function parseQuery(q) {
    // OR mit | ; pro OR-Block: AND-Terms via Leerzeichen.
    // Negation: term beginnt mit !
    const s = clean(q).toLowerCase();
    if (!s) return [];

    return s
      .split("|")
      .map((group) =>
        group
          .trim()
          .split(/\s+/)
          .filter(Boolean)
          .map((t) => ({ neg: t.startsWith("!"), term: t.replace(/^!+/, "") }))
          .filter((x) => x.term.length > 0)
      )
      .filter((group) => group.length > 0);
  }

  function findFilterCtx() {
    const th = document.querySelector(COL_TH_SELECTOR);
    if (!th) return null;

    const baseId = th.id.replace(/:col_.+$/, "");
    const root = document.getElementById(baseId) || th.closest(".ui-datatable");
    if (!root) return null;

    const tbody = document.getElementById(`${baseId}_data`);
    if (!tbody) return null;

    const headerRow = th.closest("tr");
    const headerTable = th.closest("table");
    const colIndex = headerRow ? Array.from(headerRow.querySelectorAll("th")).indexOf(th) : -1;

    const filterRow =
      headerTable?.querySelector("tr.ui-column-filter-row") ||
      headerTable?.querySelector("tr.ui-filter-row") ||
      null;

    return { th, baseId, root, tbody, headerTable, headerRow, filterRow, colIndex };
  }

  function getFilterHostCell(ctx) {
    if (ctx.filterRow && ctx.colIndex >= 0) {
      const cells = Array.from(ctx.filterRow.children);
      if (cells[ctx.colIndex]) return cells[ctx.colIndex];
    }
    return ctx.th;
  }

  function getAttachmentsCell(tr, ctx) {
    const tds = tr.querySelectorAll("td");
    return (ctx.colIndex >= 0 && tds[ctx.colIndex]) ? tds[ctx.colIndex] : null;
  }

  function buildGroups(tbody) {
    const rows = Array.from(tbody.querySelectorAll("tr"));
    const groups = [];
    let current = null;

    for (const tr of rows) {
      const ri = tr.getAttribute("data-ri");
      if (ri !== null) {
        current = { ri, data: tr, extras: [] };
        groups.push(current);
      } else if (current) {
        current.extras.push(tr);
      }
    }
    return groups;
  }

  function matches(text, parsed) {
    if (parsed.length === 0) return true;

    return parsed.some((andGroup) => {
      return andGroup.every(({ neg, term }) => {
        const has = text.includes(term);
        return neg ? !has : has;
      });
    });
  }

  function applyFilter(ctx) {
    const parsed = parseQuery(STATE.value);
    const groups = buildGroups(ctx.tbody);

    for (const g of groups) {
      const cell = getAttachmentsCell(g.data, ctx);
      const txt = attachmentsVisibleText(cell).toLowerCase();

      const show = matches(txt, parsed);

      g.data.classList.toggle("tm-hidden", !show);
      for (const ex of g.extras) ex.classList.toggle("tm-hidden", !show);
    }
  }

  function ensureFilterUI(ctx) {
    const host = getFilterHostCell(ctx);
    if (!host || host.querySelector("input.tm-attachments-filter")) return;

    const wrap = document.createElement("div");
    wrap.style.display = "flex";
    wrap.style.alignItems = "center";
    wrap.style.gap = "6px";
    wrap.style.width = "100%";

    const input = document.createElement("input");
    input.type = "text";
    input.className = "tm-attachments-filter ui-inputfield ui-inputtext ui-widget";
    input.placeholder = "Anhänge… (CMR | POD | !CMR)";
    input.value = STATE.value;
    input.style.width = "100%";
    input.style.boxSizing = "border-box";
    input.style.height = "20px";

    const clear = document.createElement("button");
    clear.type = "button";
    clear.textContent = "×";
    clear.title = "Filter löschen";
    clear.style.height = "20px";
    clear.style.padding = "0 8px";
    clear.style.cursor = "pointer";
    clear.style.border = "1px solid #999";
    clear.style.background = "#f5f5f5";
    clear.style.borderRadius = "2px";

    input.addEventListener("input", () => {
      STATE.value = input.value;
      applyFilter(ctx);
    });

    clear.addEventListener("click", () => {
      STATE.value = "";
      input.value = "";
      applyFilter(ctx);
    });

    if (ctx.filterRow && host.closest("tr") === ctx.filterRow) host.textContent = "";

    wrap.appendChild(input);
    wrap.appendChild(clear);
    host.appendChild(wrap);
  }

  // =========================
  // Teil B: Frachtbrief -> PDFs im Tab, Multi ohne ZIP
  // =========================
  const FRACHTBRIEF_BTN_ID = "frm_route_legs:j_id_41_0_2_3_1_3_4_1";
  const MAX_TABS = 25;
  const REVOKE_AFTER_MS = 60_000;
  const DL_TIMEOUT_MS = 20_000;

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  const tabQueue = [];
  function openBlankTabNow() {
    const w = window.open("about:blank", "_blank");
    if (w) { try { w.document.title = "Frachtbrief…"; } catch {} }
    return w;
  }
  function queueTabs(n) {
    const want = Math.min(n, MAX_TABS);
    let opened = 0;
    for (let i = 0; i < want; i++) {
      const w = openBlankTabNow();
      if (!w) break;
      tabQueue.push(w);
      opened++;
    }
    return opened;
  }
  function pushPdfToNextTab(blob) {
    const url = URL.createObjectURL(blob);

    while (tabQueue.length) {
      const w = tabQueue.shift();
      if (!w || w.closed) continue;
      try {
        w.location.href = url;
        setTimeout(() => { try { URL.revokeObjectURL(url); } catch {} }, REVOKE_AFTER_MS);
        return true;
      } catch {}
    }

    // Fallback (kann geblockt werden)
    const w2 = window.open(url, "_blank");
    if (w2) {
      setTimeout(() => { try { URL.revokeObjectURL(url); } catch {} }, REVOKE_AFTER_MS);
      return true;
    }
    return false;
  }

  function isPdfLike(filename, mime) {
    const fn = (filename || "").toLowerCase();
    const mt = (mime || "").toLowerCase();
    return mt.includes("pdf") || fn.endsWith(".pdf") || fn.includes("cmr") || fn.includes("frachtbrief");
  }

  function getFrachtbriefButton() {
    return document.getElementById(FRACHTBRIEF_BTN_ID)
      || Array.from(document.querySelectorAll("button")).find(b => norm(b.textContent).includes("Frachtbrief erzeugen"))
      || null;
  }

  function rowIsChecked(tr) {
    const cb = tr.querySelector("input[type='checkbox']");
    if (cb && cb.checked) return true;
    const box = tr.querySelector(".ui-chkbox-box");
    return !!(box && box.classList.contains("ui-state-active"));
  }

  function clickRowCheckbox(tr, wantChecked) {
    const cb = tr.querySelector("input[type='checkbox']");
    const box = tr.querySelector(".ui-chkbox-box") || cb;
    if (!box) return false;
    const isChecked = rowIsChecked(tr);
    if (wantChecked !== isChecked) box.click();
    return true;
  }

  function clearAll(tbody) {
    for (const tr of Array.from(tbody.querySelectorAll("tr[data-ri]"))) {
      if (rowIsChecked(tr)) clickRowCheckbox(tr, false);
    }
  }

  function getSelectedDataRows(tbody) {
    return Array.from(tbody.querySelectorAll("tr[data-ri]")).filter(rowIsChecked);
  }

  // Download waiter queue (hooked into PrimeFaces.download)
  const dlWaiters = [];
  function nextDownloadPromise() {
    return new Promise(resolve => dlWaiters.push(resolve));
  }

  function patchPrimefacesDownload() {
    if (window.__tm_pf_dl_patched) return;
    if (!window.PrimeFaces || typeof PrimeFaces.download !== "function") return;

    const orig = PrimeFaces.download.bind(PrimeFaces);

    PrimeFaces.download = function(url, mime, filename, cookieName) {
      const xhr = new XMLHttpRequest();
      xhr.open("GET", url, true);
      xhr.responseType = "blob";
      xhr.onload = function() {
        const blob = xhr.response;
        const ct = (blob && blob.type) ? blob.type : (mime || "");

        const resolver = dlWaiters.shift();
        if (resolver) resolver({ ok: true, url, filename, mime: ct });

        if (isPdfLike(filename, ct)) {
          const pushed = pushPdfToNextTab(blob);
          if (!pushed && typeof window.download === "function") {
            // Popup geblockt => fallback Download
            window.download(blob, filename, ct);
          }
        } else if (typeof window.download === "function") {
          window.download(blob, filename, ct);
        } else {
          // not expected, but keep original path
          orig(url, mime, filename, cookieName);
        }

        try {
          const ctxPath = (PrimeFaces.settings && PrimeFaces.settings.contextPath) ? PrimeFaces.settings.contextPath : "/";
          PrimeFaces.setCookie(cookieName, "true", { path: ctxPath || "/" });
        } catch {}
      };
      xhr.onerror = function() {
        const resolver = dlWaiters.shift();
        if (resolver) resolver({ ok: false, url, filename, mime });
        // fallback to original
        orig(url, mime, filename, cookieName);
      };
      xhr.send();
    };

    window.__tm_pf_dl_patched = true;
    console.log("[TM] PrimeFaces.download gepatcht: PDFs -> Tabs (Tabs müssen vorher geöffnet sein).");
  }

  async function runMulti(selectedRows, btn, tbody) {
    const opened = queueTabs(selectedRows.length);
    if (opened === 0) {
      alert("Popups sind blockiert. Bitte Popups für roadnet.dpdgroup.com erlauben.");
      return;
    }
    if (opened < selectedRows.length) selectedRows = selectedRows.slice(0, opened);

    for (const tr of selectedRows) {
      clearAll(tbody);
      clickRowCheckbox(tr, true);

      const dl = nextDownloadPromise();

      window.__tm_bypass_once = true;
      btn.click();
      window.__tm_bypass_once = false;

      await Promise.race([dl, sleep(DL_TIMEOUT_MS)]);
      await sleep(30);
    }
  }

  function attachFrachtbriefMultiHandler(ctx) {
    const btn = getFrachtbriefButton();
    if (!btn || btn.__tm_attached) return;

    btn.addEventListener("click", async (ev) => {
      if (window.__tm_bypass_once) return;

      const selected = getSelectedDataRows(ctx.tbody);

      if (selected.length <= 1) {
        // Single: Tab vorab öffnen, damit PDF sicher im Tab landet
        if (selected.length === 1) {
          const opened = queueTabs(1);
          if (opened === 0) console.warn("[TM] Popup blockiert => PDF wird vermutlich heruntergeladen.");
        }
        return; // PrimeFaces normal laufen lassen
      }

      // Multi: sonst gibt's ZIP -> wir übernehmen
      ev.preventDefault();
      ev.stopImmediatePropagation();
      await runMulti(selected, btn, ctx.tbody);

    }, true);

    btn.__tm_attached = true;
    console.log("[TM] Frachtbrief Multi: kein ZIP, PDFs einzeln.");
  }

  // =========================
  // Main runner (beides)
  // =========================
  function runAll() {
    const ctx = findFilterCtx();
    if (!ctx) return;

    // Filter
    ensureFilterUI(ctx);
    applyFilter(ctx);

    // Frachtbrief
    patchPrimefacesDownload();
    attachFrachtbriefMultiHandler(ctx);
  }

  const obs = new MutationObserver(() => {
    if (obs._raf) cancelAnimationFrame(obs._raf);
    obs._raf = requestAnimationFrame(runAll);
  });
  obs.observe(document.documentElement, { childList: true, subtree: true });

  window.addEventListener("load", runAll);
  runAll();
})();
