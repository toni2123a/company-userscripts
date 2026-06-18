// ==UserScript==
// @name         Roadnet – Frühschicht Übersicht Entladung
// @namespace    bodo.dpd.custom
// @version      1.0.2
// @description  Zeigt nur den kleinen ÜbEntl-Bubble und die Übersicht „Frühschicht – Übersicht Entladung“. Keine Roadnet-Zusammenfassung, keine LTS-Funktion, kein Bridge-Export.
// @match        https://roadnet.dpdgroup.com/execution/transport_units*
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool_roadnet_uebersicht.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool_roadnet_uebersicht.user.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  if (window.__RN_FRUEHSCHICHT_ENTLADUNG_ONLY_RUNNING) return;
  window.__RN_FRUEHSCHICHT_ENTLADUNG_ONLY_RUNNING = true;

  const BTN_OV_ID = 'rn-tu-openbtn-uebentl';

  function debounce(fn, ms) {
    let t = null;
    return (...args) => {
      clearTimeout(t);
      t = setTimeout(() => fn(...args), ms);
    };
  }

function ensureEntladungShortcutButton() {
  if (!location.href.includes('/execution/transport_units')) return null;
  if (window.__RN_OV_ENTL_RUNNING) return document.getElementById(BTN_OV_ID);
  window.__RN_OV_ENTL_RUNNING = true;

  const OVERLAY_ID = 'rnOvEntlOverlay';
  const STYLE_ID = 'rnOvEntlStyle';
  const OPEN_KEY = 'rnOvEntlOverlayOpen';

  const n = (s) => String(s ?? '').replace(/\s+/g, ' ').trim();
  const nk = (s) =>
    n(s)
      .toLowerCase()
      .replace(/\u00a0/g, ' ')
      .replace(/[^\p{L}\p{N}%# ]/gu, '')
      .replace(/\s+/g, ' ')
      .trim();

  const esc = (s) =>
    String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

  const fmtStamp = () => {
    const d = new Date();
    const p = (x) => String(x).padStart(2, '0');
    return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  };

  const fmtWB = (raw) => {
    const digits = n(raw).replace(/\D/g, '');
    if (digits.length === 8) return `${digits.slice(0, 4)}-${digits.slice(4)}`;
    return n(raw);
  };

  const fmtPct = (raw) => {
    const t = n(raw).replace(/\./g, '').replace(',', '.').replace(/[^\d.\-]/g, '');
    if (!t) return '';
    const v = parseFloat(t);
    if (!Number.isFinite(v)) return '';
    return `${Number.isInteger(v) ? String(v) : v.toFixed(1).replace('.', ',')}%`;
  };

  const parseRoadnetDateTime = (raw) => {
    const t = n(raw);
    if (!t) return null;

    const m = t.match(/(\d{1,2})\.(\d{1,2})\.(\d{2,4})(?:\s*,?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) {
      const year = Number(m[3].length === 2 ? '20' + m[3] : m[3]);
      const d = new Date(year, Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0), 0);
      return Number.isFinite(d.getTime()) ? d : null;
    }

    const d2 = new Date(t);
    return Number.isFinite(d2.getTime()) ? d2 : null;
  };

  const minutesSinceText = (raw) => {
    const start = parseRoadnetDateTime(raw);
    if (!start) return '—';

    const diff = Math.floor((Date.now() - start.getTime()) / 60000);
    if (!Number.isFinite(diff)) return '—';
    if (diff < 0) return '0 Min';
    return `${diff} Min`;
  };

  const isVisibleEl = (el) => {
    if (!el) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  };

  const isOverlayVisible = () => {
    const el = document.getElementById(OVERLAY_ID);
    return !!el && getComputedStyle(el).display !== 'none';
  };

  const setOverlayOpen = (open) => {
    try { sessionStorage.setItem(OPEN_KEY, open ? '1' : '0'); } catch {}
  };
  const getOverlayOpen = () => {
    try { return sessionStorage.getItem(OPEN_KEY) === '1'; } catch { return false; }
  };

  function ensureEntladungTab() {
    const tabs = Array.from(document.querySelectorAll('a,button,li,span,div'))
      .filter((el) => n(el.textContent) === 'Entladung' && isVisibleEl(el));
    if (!tabs.length) return false;

    const activeRe = /ui-state-active|ui-tabs-active|ui-tabs-selected|active|selected|current/i;
    const active = tabs.some((el) =>
      activeRe.test(String(el.className || '')) ||
      activeRe.test(String(el.parentElement?.className || '')) ||
      !!el.closest?.('.ui-state-active, .ui-tabs-active, .ui-tabs-selected, .active, .selected, .current')
    );
    if (active) return true;

    const el = tabs.find((x) => x.tagName === 'A' || x.tagName === 'BUTTON') || tabs[0];
    try {
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      return true;
    } catch {
      try { el.click(); return true; } catch { return false; }
    }
  }

  function clickRefreshButton() {
    const cands = Array.from(document.querySelectorAll('button,a,span'))
      .filter(isVisibleEl)
      .filter((el) => {
        const t = n(el.textContent);
        const ti = n(el.getAttribute('title'));
        const ar = n(el.getAttribute('aria-label'));
        const cl = String(el.className || '');
        return /aktualisieren/i.test(t) || /refresh|reload|aktualisieren/i.test(ti) || /refresh|reload|aktualisieren/i.test(ar) || /refresh/i.test(cl);
      });

    const btn = cands.find((el) => el.tagName === 'BUTTON') || cands[0];
    if (!btn) return false;
    try {
      btn.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      btn.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
      btn.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      return true;
    } catch {
      try { btn.click(); return true; } catch { return false; }
    }
  }

  function tableRoot() {
    const direct = document.getElementById('frm_transport_units:tbl');
    if (direct && direct.classList.contains('ui-datatable')) return direct;

    const cands = Array.from(document.querySelectorAll('div.ui-datatable[id$=":tbl"]')).filter(isVisibleEl);
    if (!cands.length) return null;
    if (cands.length === 1) return cands[0];

    let best = cands[0];
    let bestRows = -1;
    for (const c of cands) {
      const rows = c.querySelectorAll('tbody tr').length;
      if (rows > bestRows) {
        bestRows = rows;
        best = c;
      }
    }
    return best;
  }

  function headerMap(root) {
    const heads = Array.from(root.querySelectorAll('.ui-datatable-scrollable-header thead th')).filter(isVisibleEl);
    const h = heads.length ? heads : Array.from(root.querySelectorAll('thead th'));
    const map = new Map();
    h.forEach((th, idx) => {
      const key = nk(n(th.innerText ?? th.textContent ?? ''));
      if (key && !map.has(key)) map.set(key, idx);
    });
    return map;
  }

  function pickIdx(hm, keys) {
    for (const k of keys) {
      const key = nk(k);
      if (hm.has(key)) return hm.get(key);
      for (const [hk, idx] of hm.entries()) {
        if (hk.includes(key)) return idx;
      }
    }
    return null;
  }

  function cellText(td) {
    if (!td) return '';
    const main = n(td.innerText ?? td.textContent ?? '');
    if (/[A-Za-zÄÖÜäöü]/.test(main)) return main;
    const attrs = [
      td.getAttribute('title'),
      td.getAttribute('aria-label'),
      td.getAttribute('data-original-title'),
      td.getAttribute('data-tooltip'),
      td.getAttribute('data-title')
    ].map(n).filter(Boolean);
    return attrs.find((x) => /[A-Za-zÄÖÜäöü]/.test(x)) || main;
  }

  function extractRows() {
    const root = tableRoot();
    if (!root) return [];

    const hm = headerMap(root);
    const idx = {
      status: pickIdx(hm, ['status']),
      carrier: pickIdx(hm, ['frachtführer', 'frachtfuehrer', 'carrier', 'dienstleister', 'transporteur', 'unternehmer', 'spedition']),
      from: pickIdx(hm, ['herkunft', 'abgangsort', 'abgang', 'von', 'name abgangsstandort', 'code abgangsstation']),
      wb: pickIdx(hm, ['wb nr', 'nummer transporteinheit', 'lts #', 'lts', 'nummer transporte', 'transporteinheit']),
      fuel: pickIdx(hm, ['auslastung', 'auslastung (%)', 'fuellstand', 'füllstand']),
      unlBeg: pickIdx(hm, ['entladung beginn', 'entladebeginn']),
      unlEnd: pickIdx(hm, ['entladung ende', 'entladeende']),
      traffic: pickIdx(hm, ['verkehrsart', 'art', 'transportart'])
    };

    const rowsA = Array.from(root.querySelectorAll('.ui-datatable-scrollable-body tbody tr')).filter((tr) => tr.querySelectorAll('td').length);
    const rowsB = Array.from(root.querySelectorAll('tbody tr')).filter((tr) => tr.querySelectorAll('td').length);
    const merged = [...rowsA, ...rowsB];

    const uniq = [];
    const seen = new Set();
    for (const tr of merged) {
      const sig = (tr.getAttribute('data-rk') || '') + '|' + n(tr.innerText).slice(0, 180);
      if (seen.has(sig)) continue;
      seen.add(sig);
      uniq.push(tr);
    }

    const out = [];
    for (const tr of uniq) {
      const tds = Array.from(tr.querySelectorAll('td'));
      if (!tds.length) continue;
      const get = (k) => {
        const i = idx[k];
        if (i == null || i < 0 || i >= tds.length) return '';
        return cellText(tds[i]);
      };
      const row = {
        status: get('status'),
        carrier: get('carrier'),
        from: get('from'),
        wb: get('wb'),
        fuel: get('fuel'),
        unlBeg: get('unlBeg'),
        unlEnd: get('unlEnd'),
        traffic: get('traffic')
      };
      if (Object.values(row).some((v) => n(v))) out.push(row);
    }
    return out;
  }

  const isPickup = (row) => {
    const t = nk(row.traffic);
    return t.includes('pickup') || t === 'pickup';
  };
  const isUnloaded = (row) => {
    const s = nk(row.status);
    return s.includes('entladen') || !!n(row.unlEnd);
  };
  const isAtGate = (row) => !!n(row.unlBeg) && !n(row.unlEnd);

  function ensureUI() {
    if (!document.getElementById(BTN_OV_ID)) {
      const btn = document.createElement('button');
      btn.id = BTN_OV_ID;
      btn.type = 'button';
      btn.textContent = 'ÜbEntl';
      btn.style.cssText = 'position:fixed;right:14px;top:168px;z-index:2147483647;cursor:pointer;border:1px solid rgba(255,255,255,.14);background:rgba(10,15,25,.62);color:#eaf2ff;border-radius:9px;padding:2px 5px;font:900 8.5px system-ui;line-height:1;letter-spacing:.2px;box-shadow:0 8px 22px rgba(0,0,0,.20)';
      btn.addEventListener('click', () => toggleOverlay());
      document.body.appendChild(btn);
    }

    if (!document.getElementById(STYLE_ID)) {
      const style = document.createElement('style');
      style.id = STYLE_ID;
      style.textContent = `
        #${OVERLAY_ID}{position:fixed;inset:0;z-index:2147483646;display:none;color:#eef6ff;background:radial-gradient(1000px 600px at 20% -10%,rgba(71,120,220,.27),rgba(8,12,20,0) 60%),linear-gradient(180deg,#0b1220,#070c16 55%,#060a13);font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;}
        .ovTop{height:64px;display:flex;align-items:center;justify-content:space-between;padding:10px 18px;border-bottom:1px solid rgba(255,255,255,.08);background:rgba(5,10,18,.35);}
        .ovTitle{font-weight:950;font-size:clamp(18px,1.6vw,34px)}
        .ovMeta{display:flex;gap:16px;align-items:center;color:rgba(234,242,255,.78);font-weight:800;font-size:clamp(12px,1vw,18px)}
        .ovClose{cursor:pointer;border:1px solid rgba(255,255,255,.14);background:rgba(255,255,255,.06);color:#fff;border-radius:14px;padding:8px 12px;font-weight:950;font-size:clamp(12px,1vw,18px)}
        .ovBody{height:calc(100vh - 64px);padding:16px 18px 18px;box-sizing:border-box;display:grid;grid-template-columns:1.05fr 2.2fr 2.2fr;grid-template-rows:auto auto 1fr;gap:14px;}
        .ovCard{border:1px solid rgba(255,255,255,.08);background:rgba(9,15,30,.55);border-radius:16px;box-shadow:0 22px 70px rgba(0,0,0,.32);overflow:hidden;min-width:0;}
        .ovKpi{grid-column:1 / 4;display:grid;grid-template-columns:repeat(5,1fr);gap:14px}
        .ovKInner{padding:16px}
        .ovLbl{color:rgba(234,242,255,.78);font-weight:900;font-size:clamp(12px,1.3vw,25px)}
        .ovVal{margin-top:6px;font-weight:1000;line-height:1;font-size:clamp(42px,6.2vw,140px)}
        .ovProg{grid-column:1 / 4;padding:12px 14px}
        .ovBar{height:16px;border-radius:999px;background:rgba(255,255,255,.10);overflow:hidden;border:1px solid rgba(255,255,255,.10)}
        .ovBarIn{height:100%;width:0;background:linear-gradient(90deg,rgba(34,197,94,.95),rgba(59,130,246,.95));}
        .ovProgText{display:flex;justify-content:space-between;margin-top:8px;font-weight:900;color:rgba(234,242,255,.78);font-size:clamp(11px,.95vw,16px)}
        .ovSide{grid-column:1;display:flex;flex-direction:column;min-height:0;}
        .ovMain1{grid-column:2;min-height:0;}
        .ovMain2{grid-column:3;min-height:0;}
        .ovHead{padding:14px 14px 12px;border-bottom:1px solid rgba(255,255,255,.08);display:flex;justify-content:space-between;align-items:center}
        .ovHTitle{font-weight:1000;font-size:clamp(16px,1.7vw,30px)}
        .ovHCount{font-weight:1000;font-size:clamp(16px,1.7vw,30px);color:rgba(234,242,255,.88)}
        .ovList{padding:14px;display:flex;flex-direction:column;gap:10px;overflow:auto;max-height:100%}
        .ovRow{display:flex;align-items:center;justify-content:space-between;padding:10px 12px;border-radius:14px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.08)}
        .ovRowN{font-weight:950;font-size:clamp(13px,1.2vw,22px);word-break:break-word;padding-right:10px}
        .ovRowC{font-weight:1000;font-size:clamp(16px,1.8vw,30px)}
        .ovTableWrap{height:100%;display:flex;flex-direction:column;min-height:0}
        .ovTableBody{overflow:auto;min-height:0;background:rgba(9,15,30,.10)}
        .ovTable{width:100%;border-collapse:separate;border-spacing:0;table-layout:fixed}
        .ovTable th{position:sticky;top:0;z-index:2;background:rgba(12,18,32,.96);border-bottom:1px solid rgba(255,255,255,.12);text-align:left;padding:9px 10px;font-weight:950;font-size:clamp(12px,1vw,19px);white-space:nowrap}
        .ovTable td{border-bottom:1px solid rgba(255,255,255,.08);padding:9px 10px;font-weight:900;font-size:clamp(11px,1vw,18px);line-height:1.08;word-break:break-word;vertical-align:middle}
        .ovZ0{background:rgba(255,255,255,0)} .ovZ1{background:rgba(255,255,255,.03)}
        .ovWB{white-space:nowrap;font-variant-numeric:tabular-nums}
        .ovFuel{white-space:nowrap}
        .ovDuration{white-space:nowrap;font-variant-numeric:tabular-nums;font-size:clamp(10px,.9vw,16px)!important}
        .ovCarrier{font-size:clamp(9px,.78vw,14px)!important;line-height:1.12!important;font-weight:850!important;overflow-wrap:anywhere;word-break:normal}
        .ovBadge{display:inline-flex;align-items:center;justify-content:center;padding:3px 7px;border-radius:999px;font-weight:1000;font-size:clamp(8px,.70vw,12px);border:1px solid rgba(255,255,255,.12);white-space:nowrap;max-width:100%;box-sizing:border-box}
        .ovBRed{background:rgba(239,68,68,.22);color:#ffd7d7;border-color:rgba(239,68,68,.40)}
        .ovBGreen{background:rgba(34,197,94,.20);color:#d6ffe5;border-color:rgba(34,197,94,.40)}
        .ovBGray{background:rgba(148,163,184,.16);color:rgba(234,242,255,.90);border-color:rgba(148,163,184,.28)}
      `;
      document.head.appendChild(style);
    }

    if (!document.getElementById(OVERLAY_ID)) {
      const wrap = document.createElement('div');
      wrap.id = OVERLAY_ID;
      wrap.innerHTML = `
        <div class="ovTop">
          <div class="ovTitle">Frühschicht – Übersicht Entladung</div>
          <div class="ovMeta">
            <div>Stand: <span id="ovStamp">—</span></div>
            <button class="ovClose" id="ovClose" type="button">Schließen</button>
          </div>
        </div>
        <div class="ovBody">
          <div class="ovKpi">
            <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Avisierte<br>Transporteinheiten</div><div class="ovVal" id="ovKpiTotal">0</div></div></div>
            <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Noch nicht da,<br>auf dem Weg zu uns</div><div class="ovVal" id="ovKpiNotHere">0</div></div></div>
            <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Angekommen<br>noch voll, auf dem Hof</div><div class="ovVal" id="ovKpiArrived">0</div></div></div>
            <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Am Tor,<br>beim Entladen</div><div class="ovVal" id="ovKpiGate">0</div></div></div>
            <div class="ovCard"><div class="ovKInner"><div class="ovLbl">Entladen,<br>bereits leer</div><div class="ovVal" id="ovKpiDone">0</div></div></div>
          </div>

          <div class="ovCard ovProg">
            <div class="ovBar"><div class="ovBarIn" id="ovBarIn"></div></div>
            <div class="ovProgText"><span id="ovProgText">Fortschritt: 0%</span><span></span></div>
          </div>

          <div class="ovCard ovSide">
            <div class="ovHead"><div class="ovHTitle">Status-Zählung</div><div class="ovHCount" id="ovStatusTotal">—</div></div>
            <div class="ovList" id="ovStatusList"></div>
          </div>

          <div class="ovCard ovMain1">
            <div class="ovTableWrap">
              <div class="ovHead"><div class="ovHTitle">Noch zu entladen</div><div class="ovHCount" id="ovOpenCount">0</div></div>
              <div class="ovTableBody" id="ovOpenBody"></div>
            </div>
          </div>

          <div class="ovCard ovMain2">
            <div class="ovTableWrap">
              <div class="ovHead"><div class="ovHTitle">Am Tor</div><div class="ovHCount" id="ovGateCount">0</div></div>
              <div class="ovTableBody" id="ovGateBody"></div>
            </div>
          </div>
        </div>
      `;
      document.body.appendChild(wrap);

      document.getElementById('ovClose').addEventListener('click', () => {
        const el = document.getElementById(OVERLAY_ID);
        if (el) el.style.display = 'none';
        setOverlayOpen(false);
        stopAutoScroll();
      });
    }
  }

  const badgeHtml = (status) => {
    const s = nk(status);
    if (s.includes('abgefahren')) return `<span class="ovBadge ovBRed">Abgefahren</span>`;
    if (s.includes('angekommen')) return `<span class="ovBadge ovBGreen">Angekommen</span>`;
    return `<span class="ovBadge ovBGray">${esc(status || '—')}</span>`;
  };

  function buildTable(rows, mode) {
    const cols = mode === 'gate'
      ? [
          { key: 'wb', label: 'WB Nr', cls: 'ovWB' },
          { key: 'from', label: 'Herkunft', cls: '' },
          { key: 'fuel', label: 'Auslastung', cls: 'ovFuel' },
          { key: 'entladeSeit', label: 'Entladung seit', cls: 'ovDuration' }
        ]
      : [
          { key: 'wb', label: 'WB Nr', cls: 'ovWB' },
          { key: 'from', label: 'Herkunft', cls: '' },
          { key: 'fuel', label: 'Auslastung', cls: 'ovFuel' },
          { key: 'status', label: 'Status', cls: '' },
          { key: 'carrier', label: 'Frachtf.', cls: 'ovCarrier' }
        ];

    const colStyle = mode === 'gate'
      ? '<colgroup><col style="width:20%"><col style="width:34%"><col style="width:18%"><col style="width:28%"></colgroup>'
      : '<colgroup><col style="width:20%"><col style="width:24%"><col style="width:18%"><col style="width:22%"><col style="width:16%"></colgroup>';

    return `
      <table class="ovTable">
        ${colStyle}
        <thead><tr>${cols.map((c) => `<th>${esc(c.label)}</th>`).join('')}</tr></thead>
        <tbody>
          ${rows.map((r, i) => {
            const z = i % 2 ? 'ovZ1' : 'ovZ0';
            return `<tr class="${z}">${cols.map((c) => {
              if (c.key === 'status') return `<td class="${c.cls}">${badgeHtml(r.status)}</td>`;
              if (c.key === 'entladeSeit') return `<td class="${c.cls}">${esc(minutesSinceText(r.unlBeg))}</td>`;
              if (c.key === 'wb') return `<td class="${c.cls}">${esc(fmtWB(r.wb) || '—')}</td>`;
              if (c.key === 'fuel') return `<td class="${c.cls}">${esc(fmtPct(r.fuel) || '—')}</td>`;
              return `<td class="${c.cls}">${esc(n(r[c.key]) || '—')}</td>`;
            }).join('')}</tr>`;
          }).join('')}
        </tbody>
      </table>
    `;
  }

  let rendering = false;
  function renderOverlay() {
    if (rendering) return;
    rendering = true;
    try {
      ensureUI();
      ensureEntladungTab();

      const stamp = document.getElementById('ovStamp');
      if (stamp) stamp.textContent = fmtStamp();

      const all = extractRows();
      const eligible = all.filter((r) => !isPickup(r));

      const total = eligible.length;

      // KPI nach Vorgabe:
      const notHere = eligible.filter((r) => nk(r.status).includes('abgefahren')).length;   // Noch nicht da
      const arrived = eligible.filter((r) => nk(r.status).includes('angekommen')).length;  // Angekommen auf dem Hof
      const gate = eligible.filter(isAtGate);                                              // Entladebeginn ✓ / Entladeende ✗
      const done = eligible.filter(isUnloaded).length;                                     // Entladeende ✓ oder Status "Entladen"

      document.getElementById('ovKpiTotal').textContent = String(total);
      document.getElementById('ovKpiNotHere').textContent = String(notHere);
      document.getElementById('ovKpiArrived').textContent = String(arrived);
      document.getElementById('ovKpiGate').textContent = String(gate.length);
      document.getElementById('ovKpiDone').textContent = String(done);

      // Fortschritt bleibt wie gehabt (done/total)
      const pct = total > 0 ? Math.round((done / total) * 100) : 0;
      document.getElementById('ovBarIn').style.width = `${pct}%`;
      document.getElementById('ovProgText').textContent = `Fortschritt: ${pct}% (${done}/${total})`;

      // Tabellen unten bleiben wie gehabt:
      const open = eligible.filter((r) => !isUnloaded(r));
      const openOnly = open.filter((r) => !isAtGate(r));

      const sc = new Map();
      for (const r of eligible) {
        const key = n(r.status) || '—';
        sc.set(key, (sc.get(key) || 0) + 1);
      }
      const statusEntries = Array.from(sc.entries()).sort((a, b) => b[1] - a[1]);
      document.getElementById('ovStatusList').innerHTML = statusEntries.length
        ? statusEntries.map(([k, v]) => `<div class="ovRow"><div class="ovRowN">${esc(k)}</div><div class="ovRowC">${v}</div></div>`).join('')
        : '<div style="padding:8px 6px;color:rgba(234,242,255,.75);font-weight:900;">Keine Daten.</div>';
      document.getElementById('ovStatusTotal').textContent = String(eligible.length);

      document.getElementById('ovOpenBody').innerHTML = buildTable(openOnly, 'open');
      document.getElementById('ovGateBody').innerHTML = buildTable(gate, 'gate');
      document.getElementById('ovOpenCount').textContent = String(openOnly.length);
      document.getElementById('ovGateCount').textContent = String(gate.length);
    } finally {
      rendering = false;
    }
  }

  const AUTO_SCROLL = { enabled: true, step: 0.6, intervalMs: 25, pauseMs: 1200 };
  let autoScrollTimer = null;
  let autoScrollState = new Map();

  function getAutoScrollHosts() {
    const hosts = [];
    const a = document.getElementById('ovStatusList');
    const b = document.getElementById('ovOpenBody');
    const c = document.getElementById('ovGateBody');
    if (a) hosts.push(a);
    if (b) hosts.push(b);
    if (c) hosts.push(c);
    return hosts;
  }

  function isScrollableHost(el) {
    return !!el && (el.scrollHeight - el.clientHeight) > 6;
  }

  function tickAutoScroll() {
    if (!AUTO_SCROLL.enabled) return;
    if (!isOverlayVisible()) return;

    const hosts = getAutoScrollHosts();
    const now = Date.now();

    for (const el of hosts) {
      if (!isScrollableHost(el)) {
        autoScrollState.delete(el);
        continue;
      }

      let st = autoScrollState.get(el);
      if (!st) {
        st = { dir: 1, nextMoveAt: now + AUTO_SCROLL.pauseMs };
        autoScrollState.set(el, st);
      }

      if (now < st.nextMoveAt) continue;

      const maxScroll = el.scrollHeight - el.clientHeight;
      let nextTop = el.scrollTop + st.dir * AUTO_SCROLL.step;

      if (nextTop <= 0) {
        nextTop = 0;
        st.dir = 1;
        st.nextMoveAt = now + AUTO_SCROLL.pauseMs;
      } else if (nextTop >= maxScroll) {
        nextTop = maxScroll;
        st.dir = -1;
        st.nextMoveAt = now + AUTO_SCROLL.pauseMs;
      }

      el.scrollTop = nextTop;
    }
  }

  function startAutoScroll() {
    stopAutoScroll();
    autoScrollTimer = setInterval(tickAutoScroll, AUTO_SCROLL.intervalMs);
  }

  function stopAutoScroll() {
    if (autoScrollTimer) clearInterval(autoScrollTimer);
    autoScrollTimer = null;
    autoScrollState = new Map();
  }

  function toggleOverlay() {
    ensureUI();
    ensureEntladungTab();
    const el = document.getElementById(OVERLAY_ID);
    if (!el) return;
    const vis = getComputedStyle(el).display !== 'none';
    el.style.display = vis ? 'none' : 'block';
    setOverlayOpen(!vis);
    if (!vis) {
      renderOverlay();
      startAutoScroll();
    } else {
      stopAutoScroll();
    }
  }

  const renderDebounced = (() => {
    const fn = () => { if (isOverlayVisible()) renderOverlay(); };
    return debounce(fn, 250);
  })();

  function installObserver() {
    const root = tableRoot();
    if (!root) return false;
    const mo = new MutationObserver(() => renderDebounced());
    mo.observe(root, { childList: true, subtree: true, characterData: true, attributes: true });
    return true;
  }

  ensureUI();

  setTimeout(() => {
    ensureEntladungTab();
    if (getOverlayOpen()) {
      const el = document.getElementById(OVERLAY_ID);
      if (el) el.style.display = 'block';
      renderOverlay();
      startAutoScroll();
    }
  }, 800);

  let tries = 0;
  const obsTry = setInterval(() => {
    tries++;
    if (installObserver()) clearInterval(obsTry);
    if (tries >= 40) clearInterval(obsTry);
  }, 500);

  setInterval(() => {
    if (!isOverlayVisible()) return;
    setOverlayOpen(true);
    if (!ensureEntladungTab()) return;
    clickRefreshButton();
    setTimeout(() => { if (isOverlayVisible()) renderOverlay(); }, 1400);
  }, 20000);

  setInterval(() => { if (isOverlayVisible()) renderOverlay(); }, 4000);

  return document.getElementById(BTN_OV_ID);
}

  ensureEntladungShortcutButton();
})();
