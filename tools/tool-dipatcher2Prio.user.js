// ==UserScript==
// @name         DPD Dispatcher – Prio Monitoring (v3.2.2)
// @namespace    bodo.dpd.custom
// @version      3.2.2
// @description  PRIO-Monitoring mit API-Replay, Tour-Gruppierung, Kommentar-Nachladen, Auto-Refresh (sicher), Tracking-Link, Loading-Hinweis.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const NS = 'pm-';
  const log = (...a)=>console.debug('[PrioMonitoring]', ...a);
  const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
  const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const fmt = ts => { try { return new Date(ts||Date.now()).toLocaleString('de-DE'); } catch { return String(ts||''); } };

  // ---------- global state ----------
  const state={events:[], nextId:1, _bootShown:false};
  let lastOkRequest=null;
  let autoEnabled = true, autoTimer = null, isBusy = false, isLoading = false;
  const commentCache = new Map();

  // ---------- styles ----------
  function ensureStyles(){
    if (document.getElementById(NS+'style')) return;
    const style=document.createElement('style'); style.id=NS+'style';
    style.textContent=`
    .${NS}fixed-wrap{position:fixed;top:8px;left:50%;transform:translateX(-50%);display:flex;gap:8px;z-index:99999}
    .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:8px 14px;border-radius:999px;font:600 13px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .${NS}badge{min-width:18px;height:18px;line-height:18px;text-align:center;background:#b00020;color:#fff;border-radius:9px;font:700 11px/18px system-ui;padding:0 6px;display:none}
    .${NS}panel{position:fixed;top:48px;left:50%;transform:translateX(-50%);width:min(760px,92vw);max-height:74vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100000}
    .${NS}header{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
    .${NS}actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
    .${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer}
    .${NS}kpis{display:flex;gap:10px;align-items:center}
    .${NS}kpi{background:#f5f5f5;border:1px solid rgba(0,0,0,.08);padding:4px 8px;border-radius:999px;font:600 12px system-ui}
    .${NS}kpi b{font-variant-numeric:tabular-nums}
    .${NS}filter{display:flex;gap:6px;align-items:center;font:600 12px system-ui}
    .${NS}filter select{padding:4px 8px;border-radius:8px;border:1px solid rgba(0,0,0,.15);background:#fff}
    .${NS}list{list-style:none;margin:0;padding:0}
    .${NS}item{padding:12px 14px;border-left:4px solid transparent;border-bottom:1px solid rgba(0,0,0,.06)}
    .${NS}title{font:700 14px system-ui;margin:0 0 6px;text-align:center}
    .${NS}meta{font:400 12px/1.35 system-ui;opacity:.85}
    .${NS}comment{margin-top:6px;font:500 12px/1.35 system-ui;background:#fafafa;border:1px solid rgba(0,0,0,.08);border-radius:8px;padding:6px 8px}
    .${NS}sev-error{background:#fff2f2;border-left-color:#b00020}
    .${NS}sev-warn{background:#fff7e6;border-left-color:#e67600}
    .${NS}sev-info{background:#f7f7f7;border-left-color:#bdbdbd}
    .${NS}empty{padding:14px 12px;opacity:.75;text-align:center;font:500 12px system-ui}
    .${NS}dim{opacity:.6;pointer-events:none}
    .${NS}chip{display:inline-flex;gap:6px;align-items:center;border:1px solid rgba(0,0,0,.12);background:#fff;padding:4px 8px;border-radius:999px;font:600 12px system-ui}
    .${NS}dot{width:8px;height:8px;border-radius:50%;background:#16a34a}
    .${NS}dot.off{background:#9ca3af}
    .${NS}plink{color:inherit;text-decoration:underline;font-weight:700}
    .${NS}gitems{margin-left:12px; display:none}
    .${NS}gopen > .${NS}gitems{display:block}
    .${NS}loading{display:none;padding:8px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:600 12px system-ui;background:#fffbe6}
    .${NS}loading.on{display:block}
    `;
    document.head.appendChild(style);
  }

  // ---------- UI mount ----------
  function mountUI(){
    ensureStyles();
    if (!document.body) return false;
    if (document.getElementById(NS+'wrap') && document.getElementById(NS+'panel')) return true;

    // Button
    const wrap=document.createElement('div'); wrap.id=NS+'wrap'; wrap.className=NS+'fixed-wrap';
    const btn=document.createElement('button'); btn.id=NS+'btn'; btn.className=NS+'btn'; btn.textContent='Prio Monitoring'; btn.type='button';
    const badge=document.createElement('span'); badge.id=NS+'badge'; badge.className=NS+'badge'; badge.textContent='0';
    wrap.appendChild(btn); wrap.appendChild(badge);
    document.body.appendChild(wrap);

    // Panel
    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel'; panel.style.display='none';
    panel.innerHTML=`
      <div class="${NS}header">
        <div class="${NS}kpis">
          <span class="${NS}kpi" id="${NS}chip-in">PRIO in Zustellung: <b id="${NS}kpi-in">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-open">PRIO offen: <b id="${NS}kpi-open">0</b></span>
        </div>
        <div class="${NS}actions">
          <div class="${NS}filter">
            <span>Kommentare:</span>
            <select id="${NS}filter-comment">
              <option value="all">Alle</option>
              <option value="with">nur mit</option>
              <option value="without">nur ohne</option>
            </select>
          </div>
          <span class="${NS}chip" id="${NS}auto-chip"><span class="${NS}dot" id="${NS}auto-dot"></span>Auto 60s</span>
          <button class="${NS}btn-sm" data-action="refreshApi">Aktualisieren (API)</button>
        </div>
      </div>
      <div id="${NS}loading" class="${NS}loading">Lade Daten …</div>
      <ul class="${NS}list" id="${NS}list"></ul>
      <div class="${NS}empty" id="${NS}note-capture">Hinweis: Einmal die normale Liste laden/suchen, dann „Aktualisieren (API)“.</div>
      <div class="${NS}empty" id="${NS}note-error" style="display:none"></div>
    `;
    document.body.appendChild(panel);

    const toggle = (force)=>{ const open = panel.style.display !== 'none'; const will = force!==undefined ? force : !open; panel.style.display = will ? '' : 'none'; };
    btn.addEventListener('click', ()=>toggle());
    panel.addEventListener('click', async (e)=>{
      const b=e.target.closest('.'+NS+'btn-sm'); if(!b) return;
      const a=b.dataset.action;
      if(a==='refreshApi'){ await fullRefresh().catch(console.error); }
    });
    document.getElementById(NS+'filter-comment').addEventListener('change', ()=>render());
    document.addEventListener('click',(e)=>{ if(panel.style.display==='none') return; if(panel.contains(e.target)||wrap.contains(e.target)) return; toggle(false); });

    // Auto toggle
    const autoDot = document.getElementById(NS+'auto-dot');
    const autoChip = document.getElementById(NS+'auto-chip');
    function setAutoUI(){ autoDot.classList.toggle('off', !autoEnabled); }
    autoChip.addEventListener('click', ()=>{ autoEnabled = !autoEnabled; setAutoUI(); scheduleAuto(); });
    setAutoUI();

    if (!state._bootShown) { addEvent({title:'Prio Monitoring bereit', meta:'Auto-Refresh 60s • Tour-Gruppierung • Kommentare integriert', sev:'info', read:true}); state._bootShown=true; }
    render();
    return true;
  }

  // self-heal (Button/Panel wieder anheften, falls SPA neu rendert)
  setInterval(()=>{ try { mountUI(); } catch{} }, 1500);
  new MutationObserver(()=>{ try { mountUI(); } catch{} }).observe(document.documentElement, {childList:true, subtree:true});

  // ---------- small helpers ----------
  function setLoading(on){
    isLoading = !!on;
    const el = document.getElementById(NS+'loading');
    if (el) el.classList.toggle('on', isLoading);
  }
  function dimButtons(on){
    document.querySelectorAll('.'+NS+'btn-sm').forEach(b=>b.classList.toggle(NS+'dim', !!on));
  }

  // ---------- rendering ----------
  function setKpis(inDel, open){
    const elIn   = document.getElementById(NS+'kpi-in');
    const elOpen = document.getElementById(NS+'kpi-open');
    const chipIn   = document.getElementById(NS+'chip-in');
    const chipOpen = document.getElementById(NS+'chip-open');
    const vIn = Number(inDel||0), vOpen = Number(open||0);
    if (elIn) elIn.textContent = String(vIn);
    if (elOpen) elOpen.textContent = String(vOpen);
    if (chipIn) chipIn.style.display   = vIn   === 0 ? 'none' : '';
    if (chipOpen) chipOpen.style.display = vOpen === 0 ? 'none' : '';
  }

  function render(){
    const list = document.getElementById(NS+'list');
    const badge = document.getElementById(NS+'badge');
    if (!list) return;

    // Liste leeren
    list.innerHTML='';

    // (Optional) Ungelesen-Badge – derzeit deaktiviert
    if (badge){ badge.style.display='none'; }

    // Kommentar-Filter
    const modeEl = document.getElementById(NS+'filter-comment');
    const mode = modeEl ? (modeEl.value || 'all') : 'all';
    let filtered = state.events.filter(ev=>{
      if (mode==='with')    return ev.hasComment;
      if (mode==='without') return !ev.hasComment;
      return true;
    });

    // Einzelereignisse sortieren (für Gruppen)
    const sevRank = s => (s==='error'?0 : s==='warn'?1 : 2);
    filtered.sort((a,b)=>{
      const sr = sevRank(a.sev) - sevRank(b.sev);
      if (sr !== 0) return sr;
      return (b.ts||0) - (a.ts||0);
    });

    if (filtered.length === 0){
      const d=document.createElement('div');
      d.className=NS+'empty';
      d.textContent= isLoading ? 'Lade Daten …' : 'Keine Ereignisse.';
      list.appendChild(d);
      return;
    }

    // --- Gruppierung nach Tour (nur echte Ereignisse error/warn) ---
    const groupsMap = new Map(); // tourKey -> { tour, items:[], err:0, warn:0 }
    for (const ev of filtered) {
      if (ev.sev === 'info') continue; // Info nicht gruppieren (verhindert "Tour —")
      const m = /Tour\s+(\d+)/i.exec(ev.meta || '');
      const tourKey = m ? m[1] : '—';
      if (!groupsMap.has(tourKey)) groupsMap.set(tourKey, { tour: tourKey, items: [], err: 0, warn: 0 });
      const g = groupsMap.get(tourKey);
      g.items.push(ev);
      if (ev.sev === 'error') g.err++;
      else if (ev.sev === 'warn') g.warn++;
    }
    let groups = Array.from(groupsMap.values()).filter(g => g.items.length > 0 && (g.err + g.warn) > 0);

    // Gruppen-Sortierung: kritische Gruppen zuerst, dann Tournummer ASC
    function tourNum(x){ return /^\d+$/.test(x) ? Number(x) : Number.POSITIVE_INFINITY; }
    groups.sort((a, b) => {
      const ba = (a.err > 0) ? 0 : 1;
      const bb = (b.err > 0) ? 0 : 1;
      if (ba !== bb) return ba - bb;
      return tourNum(a.tour) - tourNum(b.tour);
    });

    // Items innerhalb Gruppe: error > warn > info, ts desc
    const sortGroupItems = arr => arr.sort((a,b)=>{
      const sr = sevRank(a.sev) - sevRank(b.sev);
      if (sr !== 0) return sr;
      return (b.ts||0) - (a.ts||0);
    });

    // Rendern der Gruppen
    for (const g of groups){
      const liGroup = document.createElement('li');
      liGroup.className = `${NS}item ${NS}sev-info`;
      const expanded = g.err > 0; // kritische Gruppen offen
      if (expanded) liGroup.classList.add(NS+'gopen');

      const headTitle = document.createElement('div');
      headTitle.className = NS+'title';
      headTitle.textContent = `Tour ${g.tour} (${g.err} kritisch / ${g.warn} verspätet)`;

      const headHint = document.createElement('div');
      headHint.className = NS+'meta';
      headHint.textContent = 'Klicke zum Auf- oder Zuklappen';

      const itemsWrap = document.createElement('div');
      itemsWrap.className = NS+'gitems';
      if (expanded) itemsWrap.style.display = 'block';

      sortGroupItems(g.items);
      for (const ev of g.items){
        const row = document.createElement('div');
        row.className = `${NS}item ${ev.sev==='error'?NS+'sev-error':ev.sev==='warn'?NS+'sev-warn':NS+'sev-info'}`;

        const parcelTxt = ev.parcel && /^\d+$/.test(ev.parcel)
          ? `<a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${ev.parcel}" target="_blank" rel="noopener">${ev.parcel}</a>`
          : (ev.parcel ? esc(ev.parcel) : '');

        let titleHtml = esc(ev.title);
        if (parcelTxt) {
          const m = /(\d{8,})\s*$/.exec(ev.title||'');
          if (m && m[1] === ev.parcel) titleHtml = esc(ev.title.slice(0, ev.title.length - m[1].length)) + parcelTxt;
          else titleHtml += ' ' + parcelTxt;
        }

        row.innerHTML = `
          <div class="${NS}title">${titleHtml}</div>
          <div class="${NS}meta">${esc(ev.meta||'')} – ${fmt(ev.ts)}${ev.read?'':' • ungelesen'}</div>
          ${ev.comment ? `<div class="${NS}comment">${esc(ev.comment)}</div>` : ''}`;
        itemsWrap.appendChild(row);
      }

      liGroup.addEventListener('click', (e)=>{
        if (e.target.closest('a')) return;
        liGroup.classList.toggle(NS+'gopen');
        const show = liGroup.classList.contains(NS+'gopen');
        itemsWrap.style.display = show ? 'block' : 'none';
      });

      liGroup.appendChild(headTitle);
      liGroup.appendChild(headHint);
      liGroup.appendChild(itemsWrap);
      list.appendChild(liGroup);
    }

    // Falls gerade geladen wird und noch nichts gerendert: Hinweis
    if (isLoading && groups.length === 0) {
      const d=document.createElement('div');
      d.className=NS+'empty';
      d.textContent='Lade Daten …';
      list.appendChild(d);
    }
  }

  function addEvent(ev){
    const e={
      id:state.nextId++,
      title:ev.title||'Ereignis',
      meta:ev.meta||'',
      sev:ev.sev||'info',
      ts:ev.ts||Date.now(),
      read:!!ev.read,
      comment:ev.comment||'',
      hasComment: !!(ev.comment && ev.comment.trim()),
      parcel: ev.parcel || ''
    };
    state.events.push(e); render();
  }

  // ---------- network hook (clone request) ----------
  ;(function hook(){
    if (!window.__pm_fetch_hooked && window.fetch) {
      const orig=window.fetch;
      window.fetch=async function(i,init={}){ const res=await orig(i,init);
        try{
          const uStr=typeof i==='string'?i:(i&&i.url)||''; if(uStr.includes('/dispatcher/api/pickup-delivery') && res.ok){
            const u=new URL(uStr,location.origin); const h={}; const src=(init&&init.headers)||(i&&i.headers);
            if(src){ if(src.forEach) src.forEach((v,k)=>h[String(k).toLowerCase()]=String(v)); else if(Array.isArray(src)) src.forEach(([k,v])=>h[String(k).toLowerCase()]=String(v)); else Object.entries(src).forEach(([k,v])=>h[String(k).toLowerCase()]=String(v)); }
            if(!h['authorization']){ const m=document.cookie.match(/(?:^|;\\s*)dpd-register-jwt=([^;]+)/); if(m){ h['authorization']='Bearer '+decodeURIComponent(m[1]); } }
            lastOkRequest={url:u, headers:h}; const n=document.getElementById(NS+'note-capture'); if(n) n.style.display='none';
          }
        }catch{}
        return res;
      };
      window.__pm_fetch_hooked = true;
    }
    if (!window.__pm_xhr_hooked && window.XMLHttpRequest) {
      const X=window.XMLHttpRequest; const open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
      X.prototype.open=function(m,u){ this.__pm_url=(typeof u==='string')?new URL(u,location.origin):null; this.__pm_headers={}; return open.apply(this,arguments); };
      X.prototype.setRequestHeader=function(k,v){ try{this.__pm_headers[String(k).toLowerCase()]=String(v);}catch{} return setH.apply(this,arguments); };
      X.prototype.send=function(){ const onload=()=>{ try{
        if(this.__pm_url && this.__pm_url.href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300){
          if(!this.__pm_headers['authorization']){ const m=document.cookie.match(/(?:^|;\\s*)dpd-register-jwt=([^;]+)/); if(m){ this.__pm_headers['authorization']='Bearer '+decodeURIComponent(m[1]); } }
          lastOkRequest={url:this.__pm_url, headers:this.__pm_headers}; const n=document.getElementById(NS+'note-capture'); if(n) n.style.display='none';
        }}catch{} this.removeEventListener('load',onload); }; this.addEventListener('load',onload); return send.apply(this,arguments); };
      window.__pm_xhr_hooked = true;
    }
  })();

  // ---------- API helpers ----------
  function buildHeaders(h){ const H=new Headers(); try{
    if(h){ Object.entries(h).forEach(([k,v])=>{const key=k.toLowerCase(); if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)){ H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()), v);} }); }
    if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
  }catch{} return H; }
  function buildUrl(base, page){ const u=new URL(base.href); const q=u.searchParams; q.set('page',String(page)); q.set('pageSize','500'); if(!q.get('priority')) q.set('priority','prio'); u.search=q.toString(); return u; }
  function buildUrlByParcel(base, parcel){ const u=new URL(base.href); const q=u.searchParams; q.set('page','1'); q.set('pageSize','500'); q.set('parcelNumber', String(parcel)); if(!q.get('priority')) q.set('priority','prio'); u.search=q.toString(); return u; }

  async function fetchAll(){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size = 500, maxPages = 50;
    let page = 1, rows = [], totalKnown = null;

    while (page <= maxPages) {
      const u = buildUrl(lastOkRequest.url, page);
      const r = await fetch(u.toString(), { credentials:'include', headers });
      if (!r.ok) {
        if (page === 1) {
          const n=document.getElementById(NS+'note-error'); if (n){ n.style.display='block'; n.textContent=`Serverfehler ${r.status}`; }
          throw new Error(`HTTP ${r.status}`);
        }
        break;
      }
      const j = await r.json();
      const chunk = (j.items||j.content||j.data||j.results||j)||[];
      if (page === 1) {
        const t  = Number(j.totalElements||j.total||j.count);
        const tp = Number(j.totalPages);
        if (Number.isFinite(t) && t >= 0) totalKnown = t;
        if (Number.isFinite(tp) && tp > 0) totalKnown = totalKnown ?? tp * size;
      }
      rows.push(...chunk);
      if (chunk.length < size) break;
      page++;
      await sleep(40);
    }
    return { rows, total: totalKnown || rows.length };
  }

  // ---------- mapping + detection ----------
  const isPRIO     = r => String(r.priority||'').toUpperCase()==='PRIO';
  const parcelId   = r => r.parcelNumber || (Array.isArray(r.parcelNumbers)&&r.parcelNumbers[0]) || r.id || '';
  // geliefert? -> ausgeschlossen
  const delivered  = r => {
    const st = String(r.deliveryStatus || '').toUpperCase();
    return !!r.deliveredTime || st === 'DELIVERED' || st === 'DELIVERED_TO_PUDO';
  };
  const addCodes   = r => Array.isArray(r.additionalCodes)? r.additionalCodes.map(String): [];
  const isCritCode = r => addCodes(r).some(c=>/^(041|061|032)$/.test(String(c)));
  const tourOf  = r => r.tour ? String(r.tour) : '';
  const addrOf  = r => [r.street, r.houseno].filter(Boolean).join(' ');
  const placeOf = r => [r.postalCode, r.city].filter(Boolean).join(' ');
  const whoOf   = r => [r.name, r.name2].filter(Boolean).join(' – ');

  const pickComment = (r) => {
    const fields = [ r?.freeComment, r?.freeTextDc, r?.note, r?.comment ];
    for (const f of fields) {
      if (Array.isArray(f) && f.length) return f.filter(Boolean).join(' | ');
      if (typeof f === 'string' && f.trim()) return f.trim();
    }
    return '';
  };
  let commentColIdx = null;
  function findCommentColumnIndex() {
    if (commentColIdx != null) return commentColIdx;
    const headers = Array.from(document.querySelectorAll('thead th, [role="columnheader"]'));
    let idx = -1;
    headers.forEach((h, i) => {
      const t = (h.textContent || '').trim().toLowerCase();
      if (idx === -1 && /(kommentar|frei.?text|free.?text|notiz)/.test(t)) idx = i;
    });
    commentColIdx = idx >= 0 ? idx : null;
    return commentColIdx;
  }
  function commentFromDomByParcel(parcel) {
    if (!parcel) return '';
    const idx = findCommentColumnIndex();
    const rows = Array.from(document.querySelectorAll('tbody tr, [role="row"]')).filter(tr =>
      (tr.textContent || '').includes(parcel)
    );
    if (!rows.length) return '';
    for (const tr of rows) {
      if (idx != null) {
        const cells = Array.from(tr.querySelectorAll('td, [role="gridcell"]'));
        const cell = cells[idx];
        if (cell) {
          const div = cell.querySelector('div[title]');
          const val = (div?.getAttribute('title') || cell.textContent || '').trim();
          if (val) return val;
        }
      }
      const any = tr.querySelector('td div[title]');
      if (any) {
        const val = (any.getAttribute('title') || any.textContent || '').trim();
        if (val) return val;
      }
    }
    return '';
  }
  const commentOf = (r) => {
    const fromApi = pickComment(r);
    if (fromApi) return fromApi;
    return commentFromDomByParcel(parcelId(r));
  };

  const composeDateTime = (dateStr, timeStr) => {
    if (!dateStr || !timeStr) return null;
    const s = `${dateStr}T${String(timeStr).slice(0,8)}`;
    const d = new Date(s);
    return isNaN(d) ? null : d;
  };
  const predictStart = r => r.from2 ? composeDateTime(r.date, r.from2) : null;
  const predictEnd   = r => {
    if (r.to2) { const d=composeDateTime(r.date, r.to2); if (d) return d; }
    const st = predictStart(r); return st ? new Date(st.getTime() + 60*60*1000) : null;
  };
  const etaScantime  = r => r.etaScanTime ? composeDateTime(r.date, r.etaScanTime) : null;
  const isPredictOverdueNoETA = r => {
    if (delivered(r)) return false;
    const end = predictEnd(r); if (!end) return false;
    const now = new Date(); if (now <= end) return false;
    return !etaScantime(r);
  };

  // ---------- comments autoload ----------
  async function fetchMissingComments(){
    if(!lastOkRequest) { addEvent({title:'Hinweis', meta:'Kein API-Request zum Klonen vorhanden.', sev:'info', read:true}); render(); return; }
    const headers = buildHeaders(lastOkRequest.headers);
    const toFill = state.events.filter(e => !e.hasComment && e.parcel);
    if (toFill.length === 0) return;

    dimButtons(true);
    let filled = 0;
    for (const ev of toFill){
      const p = ev.parcel;
      if (commentCache.has(p)) {
        const c = commentCache.get(p);
        if (c) { ev.comment = c; ev.hasComment = true; filled++; render(); }
        continue;
      }
      try{
        const url = buildUrlByParcel(lastOkRequest.url, p);
        const res = await fetch(url.toString(), { credentials:'include', headers });
        if (res.ok){
          const j = await res.json();
          const rows = (j.items||j.content||j.data||j.results||j)||[];
          const r = rows.find(x => (parcelId(x) === p)) || rows[0];
          let c = '';
          if (r) c = pickComment(r) || commentFromDomByParcel(p);
          commentCache.set(p, c);
          if (c) { ev.comment = c; ev.hasComment = true; filled++; render(); }
        }
      }catch(e){ /* ignore */ }
      await sleep(80);
    }
    if (filled>0) addEvent({title:'Kommentare', meta:`Nachgeladen: ${filled}/${toFill.length}`, sev:'info', read:true});
    dimButtons(false);
    render();
  }

  // ---------- refresh logic ----------
  async function refreshViaApi(){
    const { rows, total } = await fetchAll();

    let prioIn = 0, prioOpen = 0, prioDelivered = 0, prioTotal = 0;
    const reds=[], oranges=[];
    for(const r of rows){
      if(!isPRIO(r)) continue;
      prioTotal++;
      if (delivered(r)) { prioDelivered++; continue; } else { prioIn++; }
      if (isCritCode(r)) reds.push(r);
      else if (isPredictOverdueNoETA(r)) oranges.push(r);
    }
    prioOpen = 0;
    setKpis(prioIn, prioOpen);

    // neue Events sammeln und erst am Ende tauschen
    const newEvents = [];
    const push = (ev)=> newEvents.push({
      id: ++state.nextId,
      title: ev.title, meta: ev.meta, sev: ev.sev || 'info',
      ts: ev.ts || Date.now(), read: !!ev.read,
      comment: ev.comment || '', hasComment: !!(ev.comment && ev.comment.trim()),
      parcel: ev.parcel || ''
    });

    const mkWnd = (ps,pe)=> (ps&&pe) ? `${ps.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'})}–${pe.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'})}` : '';

    for (const r of reds){
      const ps=predictStart(r), pe=predictEnd(r);
      push({
        title:`PRIO kritisch (Code): ${parcelId(r) || 'Sendung'}`,
        meta:`${tourOf(r)?'Tour '+tourOf(r)+' • ':''}${addrOf(r)}${placeOf(r)?' ('+placeOf(r)+')':''}${whoOf(r)?' • '+whoOf(r):''} • Zusatzcode ${addCodes(r).join(',')}${mkWnd(ps,pe)?' • '+mkWnd(ps,pe):''}`,
        sev:'error', comment: commentOf(r), parcel: parcelId(r)
      });
    }
    for (const r of oranges){
      const ps=predictStart(r), pe=predictEnd(r);
      push({
        title:`PRIO verspätet (Predict): ${parcelId(r) || 'Sendung'}`,
        meta:`${tourOf(r)?'Tour '+tourOf(r)+' • ':''}${addrOf(r)}${placeOf(r)?' ('+placeOf(r)+')':''}${whoOf(r)?' • '+whoOf(r):''} • Zeitfenster überschritten • ${mkWnd(ps,pe) || 'Predict'}`,
        sev:'warn', comment: commentOf(r), parcel: parcelId(r)
      });
    }
    push({
      title:'Aktualisiert',
      meta:`Rot: ${reds.length} • Orange: ${oranges.length} • Geprüft: ${(total ? `${rows.length}/${total}` : `${rows.length}`)} • PRIO: ${prioTotal} • in Zustellung: ${prioIn} • geliefert: ${prioDelivered}`,
      sev:'info', read:true
    });

    state.events = newEvents;
  }

  async function fullRefresh(){
    if (isBusy) return;
    try{
      if (!lastOkRequest) {
        addEvent({title:'Hinweis', meta:'Kein API-Request erkannt. Bitte einmal die normale Suche ausführen, danach funktioniert Auto 60s.', sev:'info', read:true});
        render();
        return;
      }
      isBusy = true;
      setLoading(true);
      dimButtons(true);
      addEvent({title:'Aktualisiere (API)…', meta:'Replay aktiver Filter (pageSize=500)', sev:'info', read:true}); 
      render();

      await refreshViaApi();
      await fetchMissingComments();
    } catch(e){
      console.error(e);
      addEvent({title:'Fehler (API)', meta:String(e && e.message || e), sev:'warn'});
    } finally{
      setLoading(false);
      dimButtons(false);
      isBusy = false;
      render();
    }
  }

  // ---------- auto refresh ----------
  function scheduleAuto(){
    try{
      if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
      if (!autoEnabled) return;
      if (document.hidden) return;
      if (!lastOkRequest) return; // nur wenn Vorlage vorhanden

      autoTimer = setInterval(()=>{
        if (!lastOkRequest) return;
        fullRefresh().catch(()=>{});
      }, 60_000);
    }catch{}
  }
  document.addEventListener('visibilitychange', ()=>scheduleAuto());

  // ---------- boot ----------
  (function boot(){
    mountUI();
    scheduleAuto();
  })();

})();
