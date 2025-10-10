// ==UserScript==
// @name         DPD Dispatcher – Express_12 Monitor (v1.0.3)
// @namespace    bodo.dpd.custom
// @version      1.0.3
// @description  EXPRESS_12 Monitoring: findet EXPRESS_12 auch ohne UI-Filter. Lädt ALLE Seiten, testet mehrere priority/service Varianten, plus Fallback ohne Filter. Warn: Fenster >=12:00; Info: Erinnerung ab 11:00. Panel versetzt (Mitte+Offset).
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const NS = 'e12-';
  const OFFSET_PX = 260; // Position zwischen PRIO (Mitte) und rechts
  const log = (...a)=>console.debug('[Express_12]', ...a);
  const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
  const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const fmt = ts => { try { return new Date(ts||Date.now()).toLocaleString('de-DE'); } catch { return String(ts||''); } };

  const state={events:[], nextId:1, _bootShown:false};
  let lastOkRequest=null;
  let autoEnabled = true, autoTimer = null, isBusy = false, isLoading = false;
  const commentCache = new Map();

  // --------------------------------- Styles / UI ---------------------------------
  function ensureStyles(){
    if (document.getElementById(NS+'style')) return;
    const style=document.createElement('style'); style.id=NS+'style';
    style.textContent=`
    .${NS}fixed-wrap{position:fixed; top:8px; left:calc(50% + ${OFFSET_PX}px); transform:translateX(0); display:flex; gap:8px; z-index:99999}
    .${NS}btn{border:1px solid rgba(0,0,0,.12); background:#fff; padding:8px 14px; border-radius:999px; font:600 13px system-ui; cursor:pointer; box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .${NS}badge{min-width:18px; height:18px; line-height:18px; text-align:center; background:#0b61a4; color:#fff; border-radius:9px; font:700 11px/18px system-ui; padding:0 6px; display:none}
    .${NS}panel{position:fixed; top:48px; left:calc(50% + ${OFFSET_PX}px); transform:translateX(0); width:min(760px,92vw); max-height:74vh; overflow:auto; background:#fff; border:1px solid rgba(0,0,0,.12); box-shadow:0 12px 28px rgba(0,0,0,.18); border-radius:12px; z-index:100000}
    .${NS}header{display:grid; grid-template-columns:1fr auto; gap:10px; align-items:center; padding:10px 12px; border-bottom:1px solid rgba(0,0,0,.08); font:700 13px system-ui}
    .${NS}actions{display:flex; gap:8px; align-items:center; flex-wrap:wrap}
    .${NS}btn-sm{border:1px solid rgba(0,0,0,.12); background:#f7f7f7; padding:6px 10px; border-radius:8px; font:600 12px system-ui; cursor:pointer}
    .${NS}kpis{display:flex; gap:10px; align-items:center}
    .${NS}kpi{background:#f5f5f5; border:1px solid rgba(0,0,0,.08); padding:4px 8px; border-radius:999px; font:600 12px system-ui}
    .${NS}kpi b{font-variant-numeric:tabular-nums}
    .${NS}filter{display:flex; gap:6px; align-items:center; font:600 12px system-ui}
    .${NS}filter select{padding:4px 8px; border-radius:8px; border:1px solid rgba(0,0,0,.15); background:#fff}
    .${NS}list{list-style:none; margin:0; padding:0}
    .${NS}item{padding:12px 14px; border-left:4px solid transparent; border-bottom:1px solid rgba(0,0,0,.06)}
    .${NS}title{font:700 14px system-ui; margin:0 0 6px; text-align:center}
    .${NS}meta{font:400 12px/1.35 system-ui; opacity:.85}
    .${NS}comment{margin-top:6px; font:500 12px/1.35 system-ui; background:#fafafa; border:1px solid rgba(0,0,0,.08); border-radius:8px; padding:6px 8px}
    .${NS}sev-error{background:#fff2f2; border-left-color:#b00020}
    .${NS}sev-warn{background:#fff7e6; border-left-color:#e67600}
    .${NS}sev-info{background:#f7f7f7; border-left-color:#bdbdbd}
    .${NS}empty{padding:14px 12px; opacity:.75; text-align:center; font:500 12px system-ui}
    .${NS}dim{opacity:.6; pointer-events:none}
    .${NS}chip{display:inline-flex; gap:6px; align-items:center; border:1px solid rgba(0,0,0,.12); background:#fff; padding:4px 8px; border-radius:999px; font:600 12px system-ui}
    .${NS}dot{width:8px; height:8px; border-radius:50%; background:#16a34a}
    .${NS}dot.off{background:#9ca3af}
    .${NS}plink{color:inherit; text-decoration:underline; font-weight:700}
    .${NS}loading{display:none; padding:8px 12px; border-bottom:1px solid rgba(0,0,0,.08); font:600 12px system-ui; background:#fffbe6}
    .${NS}loading.on{display:block}
    `;
    document.head.appendChild(style);
  }

  function mountUI(){
    ensureStyles();
    if (!document.body) return false;
    if (document.getElementById(NS+'wrap') && document.getElementById(NS+'panel')) return true;

    const wrap=document.createElement('div'); wrap.id=NS+'wrap'; wrap.className=NS+'fixed-wrap';
    const btn=document.createElement('button'); btn.id=NS+'btn'; btn.className=NS+'btn'; btn.textContent='Express_12 Monitor'; btn.type='button';
    const badge=document.createElement('span'); badge.id=NS+'badge'; badge.className=NS+'badge'; badge.textContent='0';
    wrap.appendChild(btn); wrap.appendChild(badge);
    document.body.appendChild(wrap);

    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel'; panel.style.display='none';
    panel.innerHTML=`
      <div class="${NS}header">
        <div class="${NS}kpis">
          <span class="${NS}kpi" id="${NS}chip-in">EXPRESS_12 in Zustellung: <b id="${NS}kpi-in">0</b></span>
          <span class="${NS}kpi" id="${NS}chip-open">Erinnerungen (>=11:00): <b id="${NS}kpi-rem">0</b></span>
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
      if(b.dataset.action==='refreshApi'){ await fullRefresh().catch(console.error); }
    });
    document.getElementById(NS+'filter-comment').addEventListener('change', ()=>render());
    document.addEventListener('click',(e)=>{ if(panel.style.display==='none') return; if(panel.contains(e.target)||wrap.contains(e.target)) return; toggle(false); });

    const autoDot = document.getElementById(NS+'auto-dot');
    const autoChip = document.getElementById(NS+'auto-chip');
    function setAutoUI(){ autoDot.classList.toggle('off', !autoEnabled); }
    autoChip.addEventListener('click', ()=>{ autoEnabled = !autoEnabled; setAutoUI(); scheduleAuto(); });
    setAutoUI();

    if (!state._bootShown) { addEvent({title:'Express_12 Monitor bereit', meta:'Fenster versetzt • Auto-Refresh 60s • Kommentare', sev:'info', read:true}); state._bootShown=true; }
    render();
    return true;
  }

  setInterval(()=>{ try { mountUI(); } catch{} }, 1500);
  new MutationObserver(()=>{ try { mountUI(); } catch{} }).observe(document.documentElement, {childList:true, subtree:true});

  function setLoading(on){
    isLoading = !!on;
    const el = document.getElementById(NS+'loading');
    if (el) el.classList.toggle('on', isLoading);
  }
  function dimButtons(on){
    document.querySelectorAll('.'+NS+'btn-sm').forEach(b=>b.classList.toggle(NS+'dim', !!on));
  }
  function setBadge(n){
    const b = document.getElementById(NS+'badge'); if(!b) return;
    const v = Math.max(0, Number(n||0));
    b.textContent = String(v);
    b.style.display = v > 0 ? 'inline-block' : 'none';
  }
  function setKpis(inDel, rem){
    const elIn   = document.getElementById(NS+'kpi-in');
    const elRem  = document.getElementById(NS+'kpi-rem');
    const chipIn   = document.getElementById(NS+'chip-in');
    const chipRem  = document.getElementById(NS+'chip-open');
    const vIn = Number(inDel||0), vRem = Number(rem||0);
    if (elIn) elIn.textContent = String(vIn);
    if (elRem) elRem.textContent = String(vRem);
    if (chipIn) chipIn.style.display   = vIn   === 0 ? 'none' : '';
    if (chipRem) chipRem.style.display = vRem === 0 ? 'none' : '';
  }

  function render(){
    const list = document.getElementById(NS+'list');
    if (!list) return;
    list.innerHTML='';

    const modeEl = document.getElementById(NS+'filter-comment');
    const mode = modeEl ? (modeEl.value || 'all') : 'all';
    let filtered = state.events.filter(ev=>{
      if (mode==='with')    return ev.hasComment;
      if (mode==='without') return !ev.hasComment;
      return true;
    });

    const sevRank = s => (s==='error'?0 : s==='warn'?1 : 2);
    filtered.sort((a,b)=>{ const sr = sevRank(a.sev) - sevRank(b.sev); if (sr!==0) return sr; return (b.ts||0) - (a.ts||0); });

    if (filtered.length === 0){
      const d=document.createElement('div');
      d.className=NS+'empty';
      d.textContent= isLoading ? 'Lade Daten …' : 'Keine Ereignisse.';
      list.appendChild(d);
      return;
    }

    for(const ev of filtered){
      const row=document.createElement('li');
      row.className = `${NS}item ${ev.sev==='error'?NS+'sev-error':ev.sev==='warn'?NS+'sev-warn':NS+'sev-info'}`;

      const parcelTxt = ev.parcel && /^\d+$/.test(ev.parcel)
        ? `<a class="${NS}plink" href=https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${ev.parcel} target="_blank" rel="noopener">${ev.parcel}</a>`
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
      list.appendChild(row);
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

  // --------------------------------- Network hook ---------------------------------
  ;(function hook(){
    if (!window.__e12_fetch_hooked && window.fetch) {
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
      window.__e12_fetch_hooked = true;
    }
    if (!window.__e12_xhr_hooked && window.XMLHttpRequest) {
      const X=window.XMLHttpRequest; const open=X.prototype.open, send=X.prototype.send, setH=X.prototype.setRequestHeader;
      X.prototype.open=function(m,u){ this.__e12_url=(typeof u==='string')?new URL(u,location.origin):null; this.__e12_headers={}; return open.apply(this,arguments); };
      X.prototype.setRequestHeader=function(k,v){ try{this.__e12_headers[String(k).toLowerCase()]=String(v);}catch{} return setH.apply(this,arguments); };
      X.prototype.send=function(){ const onload=()=>{ try{
        if(this.__e12_url && this.__e12_url.href.includes('/dispatcher/api/pickup-delivery') && this.status>=200 && this.status<300){
          if(!this.__e12_headers['authorization']){ const m=document.cookie.match(/(?:^|;\\s*)dpd-register-jwt=([^;]+)/); if(m){ this.__e12_headers['authorization']='Bearer '+decodeURIComponent(m[1]); } }
          lastOkRequest={url:this.__e12_url, headers:this.__e12_headers}; const n=document.getElementById(NS+'note-capture'); if(n) n.style.display='none';
        }}catch{} this.removeEventListener('load',onload); }; this.addEventListener('load',onload); return send.apply(this,arguments); };
      window.__e12_xhr_hooked = true;
    }
  })();

  // --------------------------------- API helpers ---------------------------------
  function buildHeaders(h){ const H=new Headers(); try{
    if(h){ Object.entries(h).forEach(([k,v])=>{const key=k.toLowerCase(); if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)){ H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()), v);} }); }
    if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
  }catch{} return H; }

  // Kandidaten für API-Parameter, falls UI nichts setzt / abweicht
  const PRIORITY_KEYS = [
    // häufigster Fall
    {k:'priority', v:'EXPRESS_12'},
    {k:'priority', v:'express_12'},
    {k:'priority', v:'EXPRESS12'},
    {k:'priority', v:'express12'},
    // mögliche Alternativen
    {k:'serviceType', v:'express_12'},
    {k:'service',     v:'express_12'},
    {k:'product',     v:'express_12'}
  ];

  function stripPriorityParams(q){
    // alles entfernen, was den Priority-Filter beeinflussen könnte
    ['priority','serviceType','service','product'].forEach(p=>q.delete(p));
    // mögliche Spaltenfilterreste (je nach UI) entfernen
    ['prio','prio12','priorityId','priority_id'].forEach(p=>q.delete(p));
  }

  function cloneBaseUrl(base){
    const u=new URL(base.href);
    return u;
  }

  function buildUrlVariant(base, page, variantIdx){
    const u = cloneBaseUrl(base);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize','500');
    stripPriorityParams(q);
    const cand = PRIORITY_KEYS[variantIdx];
    if (cand) q.set(cand.k, cand.v);
    u.search = q.toString();
    return u;
  }

  function buildUrlByParcelVariant(base, parcel, variantIdx){
    const u = cloneBaseUrl(base);
    const q = u.searchParams;
    q.set('page','1');
    q.set('pageSize','500');
    q.set('parcelNumber', String(parcel));
    stripPriorityParams(q);
    const cand = PRIORITY_KEYS[variantIdx];
    if (cand) q.set(cand.k, cand.v);
    u.search = q.toString();
    return u;
  }

  // Lädt alle Seiten – testet mehrere Varianten; bei 0 Treffern Fallback (ohne Priority)
  async function fetchAll(){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size = 500, maxPages = 50;

    async function runVariant(variantIdx, useFallbackNoPriority=false){
      let page=1, rows=[], totalKnown=null;
      while(page<=maxPages){
        const u = useFallbackNoPriority
          ? (()=>{ const t=cloneBaseUrl(lastOkRequest.url); const q=t.searchParams; q.set('page', String(page)); q.set('pageSize','500'); stripPriorityParams(q); t.search=q.toString(); return t; })()
          : buildUrlVariant(lastOkRequest.url, page, variantIdx);
        const r = await fetch(u.toString(), { credentials:'include', headers });
        if(!r.ok){ if(page===1) throw new Error(`HTTP ${r.status}`); break; }
        const j = await r.json();
        const chunk = (j.items||j.content||j.data||j.results||j)||[];
        if (page===1){
          const t  = Number(j.totalElements||j.total||j.count);
          const tp = Number(j.totalPages);
          if (Number.isFinite(t) && t >= 0) totalKnown = t;
          if (Number.isFinite(tp) && tp > 0) totalKnown = totalKnown ?? tp * size;
        }
        rows.push(...chunk);
        if (chunk.length < size) break;
        page++; await sleep(40);
      }
      return {rows, total: totalKnown || rows.length};
    }

    // 1) Varianten mit Priority versuchen – abbrechen, sobald EXPRESS_12 im Ergebnis auftaucht
    for (let i=0;i<PRIORITY_KEYS.length;i++){
      const {rows,total} = await runVariant(i,false);
      const hasExp = rows.some(r=>isEXP12(r));
      if (hasExp || rows.length>0) return {rows,total};
    }

    // 2) Fallback: ohne Priority-Parameter (UI-Query pur), danach clientseitig filtern
    return await runVariant(0,true);
  }

  // --------------------------------- Mapping / Regeln ---------------------------------
  const isEXP12    = r => {
    const p = String(r.priority||r.product||r.serviceType||r.service||'').toUpperCase();
    return p === 'EXPRESS_12' || p === 'EXPRESS12';
  };
  const parcelId   = r => r.parcelNumber || (Array.isArray(r.parcelNumbers)&&r.parcelNumbers[0]) || r.id || '';
  const delivered  = r => {
    const st = String(r.deliveryStatus || '').toUpperCase();
    return !!r.deliveredTime || st === 'DELIVERED' || st === 'DELIVERED_TO_PUDO';
  };
  const addCodes   = r => Array.isArray(r.additionalCodes)? r.additionalCodes.map(String): [];
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
  const commentOf = (r) => pickComment(r) || commentFromDomByParcel(parcelId(r));

  const composeDateTime = (dateStr, timeStr) => {
    if (!dateStr || !timeStr) return null;
    const s = `${dateStr}T${String(timeStr).slice(0,8)}`;
    const d = new Date(s);
    return isNaN(d) ? null : d;
  };
  const predictStart = r => r.from2 ? composeDateTime(r.date, r.from2) : null;
  const predictEnd   = r => r.to2 ? composeDateTime(r.date, r.to2) : (r.from2 ? composeDateTime(r.date, r.from2) : null);

  function isEndAtOrAfterNoon(r){
    const end = predictEnd(r);
    if (!end) return false; // nur Fälle mit Zeitfenster
    const noon = new Date(end); noon.setHours(12,0,0,0);
    return end.getTime() >= noon.getTime(); // 11:00–12:00 = true; 10:59–11:59 = false
  }
  function needs11Reminder(now, r){
    if (delivered(r)) return false;
    const t = now || new Date();
    const eleven = new Date(t); eleven.setHours(11,0,0,0);
    return t.getTime() >= eleven.getTime();
  }

  // --------------------------------- Comments autoload ---------------------------------
  async function fetchMissingComments(){
    if(!lastOkRequest) { addEvent({title:'Hinweis', meta:'Kein API-Request zum Klonen vorhanden.', sev:'info', read:true}); render(); return; }
    const headers = buildHeaders(lastOkRequest.headers);
    const targets = state.events.filter(e => !e.hasComment && e.parcel);
    if (!targets.length) return;

    dimButtons(true);
    let filled = 0;
    for (const ev of targets){
      const p = ev.parcel;
      if (commentCache.has(p)) {
        const c = commentCache.get(p);
        if (c) { ev.comment = c; ev.hasComment = true; filled++; render(); }
        continue;
      }
      // Varianten für Einzelabruf testen
      let c='';
      for (let i=0;i<PRIORITY_KEYS.length && !c;i++){
        try{
          const url = buildUrlByParcelVariant(lastOkRequest.url, p, i);
          const res = await fetch(url.toString(), { credentials:'include', headers });
          if (res.ok){
            const j = await res.json();
            const rows = (j.items||j.content||j.data||j.results||j)||[];
            const r = rows.find(x => (parcelId(x) === p)) || rows[0];
            if (r) c = pickComment(r) || commentFromDomByParcel(p);
          }
        }catch(e){ /* ignore */ }
        await sleep(50);
      }
      commentCache.set(p, c);
      if (c) { ev.comment = c; ev.hasComment = true; filled++; render(); }
      await sleep(60);
    }
    if (filled>0) addEvent({title:'Kommentare', meta:`Nachgeladen: ${filled}/${targets.length}`, sev:'info', read:true});
    dimButtons(false);
    render();
  }

  // --------------------------------- Refresh logic ---------------------------------
  async function refreshViaApi(){
    const { rows, total } = await fetchAll();

    let inDel = 0, totalExp = 0, deliveredCount = 0;
    const afterNoon=[], remindList=[];
    const now = new Date();

    for(const r of rows){
      if(!isEXP12(r)) continue;
      totalExp++;
      if (delivered(r)) { deliveredCount++; continue; } else { inDel++; }

      if (isEndAtOrAfterNoon(r)) afterNoon.push(r);
      if (needs11Reminder(now, r)) remindList.push(r);
    }

    setBadge(afterNoon.length + remindList.length);
    setKpis(inDel, remindList.length);

    const newEvents = [];
    const push = (ev)=> newEvents.push({
      id: ++state.nextId,
      title: ev.title, meta: ev.meta, sev: ev.sev || 'info',
      ts: ev.ts || Date.now(), read: !!ev.read,
      comment: ev.comment || '', hasComment: !!(ev.comment && ev.comment.trim()),
      parcel: ev.parcel || ''
    });

    const mkWnd = (ps,pe)=> {
      if(!ps && !pe) return '';
      const f = d => d?.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}) || '';
      if (ps && pe) return `${f(ps)}–${f(pe)}`;
      return f(pe||ps);
    };

    for (const r of afterNoon){
      const ps=predictStart(r), pe=predictEnd(r);
      push({
        title:`EXPRESS_12 (Fenster ≥ 12:00): ${parcelId(r) || 'Sendung'}`,
        meta:`${tourOf(r)?'Tour '+tourOf(r)+' • ':''}${addrOf(r)}${placeOf(r)?' ('+placeOf(r)+')':''}${whoOf(r)?' • '+whoOf(r):''}${mkWnd(ps,pe)?' • '+mkWnd(ps,pe):''}${addCodes(r).length? ' • Zusatzcode '+addCodes(r).join(','):''}`,
        sev:'warn', comment: commentOf(r), parcel: parcelId(r)
      });
    }

    if (remindList.length){
      for (const r of remindList){
        const ps=predictStart(r), pe=predictEnd(r);
        push({
          title:`Erinnerung (>=11:00) nicht zugestellt: ${parcelId(r) || 'Sendung'}`,
          meta:`${tourOf(r)?'Tour '+tourOf(r)+' • ':''}${addrOf(r)}${placeOf(r)?' ('+placeOf(r)+')':''}${whoOf(r)?' • '+whoOf(r):''}${mkWnd(ps,pe)?' • '+mkWnd(ps,pe):''}`,
          sev:'info', comment: commentOf(r), parcel: parcelId(r)
        });
      }
    }

    push({
      title:'Aktualisiert',
      meta:`Nach 12: ${afterNoon.length} • Erinnerungen (>=11:00): ${remindList.length} • Geprüft: ${(total ? `${rows.length}/${total}` : `${rows.length}`)} • EXPRESS_12: ${totalExp} • in Zustellung: ${inDel} • geliefert: ${deliveredCount}`,
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
      addEvent({title:'Aktualisiere (API)…', meta:'Alle Seiten, mehrere Priority-Varianten + Fallback', sev:'info', read:true});
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

  function scheduleAuto(){
    try{
      if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
      if (!autoEnabled) return;
      if (document.hidden) return;
      if (!lastOkRequest) return;
      autoTimer = setInterval(()=>{ if (!lastOkRequest) return; fullRefresh().catch(()=>{}); }, 60_000);
    }catch{}
  }
  document.addEventListener('visibilitychange', ()=>scheduleAuto());

  (function boot(){
    mountUI();
    scheduleAuto();
  })();

})();
