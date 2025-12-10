// ==UserScript==
// @name         Roadnet – Anhänge Filter 
// @namespace    https://roadnet.dpdgroup.com/
// @version      1.8.0
// @description  Filter für Anhänge anhand sichtbarer (nicht .hidden) Labels. Unterstützt Negativ-Filter: !CMR.
// @match        https://roadnet.dpdgroup.com/execution/route_legs*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(() => {
  "use strict";

  const COL_TH_SELECTOR =
    "th#frm_route_legs\\:tbl\\:col_attachmentTypes, th[id$=':col_attachmentTypes']";
  const STATE = { value: "" };

  const norm = (s) => (s ?? "").replace(/\s+/g, " ").trim();
  const clean = (s) => norm(String(s || "").replace(/\u00A0/g, " ").replace(/\?\!/g, ""));

  const styleId = "tm-att-filter-style-18";
  if (!document.getElementById(styleId)) {
    const st = document.createElement("style");
    st.id = styleId;
    st.textContent = `.tm-hidden { display: none !important; }`;
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

  function findCtx() {
    const th = document.querySelector(COL_TH_SELECTOR);
    if (!th) return null;

    const baseId = th.id.replace(/:col_.+$/, "");
    const root = document.getElementById(baseId) || th.closest(".ui-datatable");
    if (!root) return null;

    // bei dir korrekt: baseId_data
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
        current.extras.push(tr); // paired/helper-row
      }
    }
    return groups;
  }

  function matches(text, parsed) {
    // parsed: Array<OR-group>, OR-group: Array<{neg,term}> (AND)
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

  function ensureUI(ctx) {
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

  function run() {
    const ctx = findCtx();
    if (!ctx) return;
    ensureUI(ctx);
    applyFilter(ctx);
  }

  const obs = new MutationObserver(() => {
    if (obs._raf) cancelAnimationFrame(obs._raf);
    obs._raf = requestAnimationFrame(run);
  });
  obs.observe(document.documentElement, { childList: true, subtree: true });

  window.addEventListener("load", run);
  run();
})();
