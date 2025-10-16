// ==UserScript==
// @name         DPD Dispatcher – Partner-Report Mailer (v5.1.0, IndexedDB)
// @namespace    bodo.dpd.custom
// @version      5.1.0
// @description  Gesamtübersicht an Verteiler + pro Systempartner Detail-Mail; Empfänger lokal speichern (IndexedDB), Export/Import; EML-Download optional; per-Zeile Senden-Button.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        GM_xmlhttpRequest
// @connect      10.14.7.169
// ==/UserScript==

(function(){
  'use strict';

  const NS='fvpr-', PANEL_ID='#fvpr-panel', OFFSET_PX=-240, REFRESH_MS=60_000, RENDER_DEBOUNCE=300;
  const GATEWAY_DEFAULT = 'http://10.14.7.169/mail.php'; // XAMPP-Endpoint
  const GATEWAY_API_KEY = 'fvpr-SECRET-123';           // frei wählen


  // ---------- Utils ----------
  const norm=s=>String(s||'').replace(/\s+/g,' ').trim();
  const parsePct=s=>{ if(s==null)return null; const t=String(s).replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'').trim(); if(!t)return null; const v=parseFloat(t); return Number.isFinite(v)?v:null; };
  const parseIntDe=s=>{ if(s==null)return null; const t=String(s).replace(/\./g,'').replace(',','.').replace(/[^\d.\-]/g,'').trim(); if(!t)return null; const v=Math.round(parseFloat(t)); return Number.isFinite(v)?v:null; };
  const fmtPct=v=>Number.isFinite(v)?v.toFixed(1).replace('.',','):'—';
  const fmtInt=v=>v==null?'—':String(Math.round(v||0)).replace(/\B(?=(\d{3})+(?!\d))/g,'.');
  const todayStr=()=>new Date().toLocaleDateString('de-DE');
  const pad2=n=>String(n).padStart(2,'0');
  const timeHM=()=>{ const d=new Date(); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; };
  const dateStamp=()=>{ const d=new Date(); return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`; };
  const sum=(vs,proj)=>vs.reduce((a,v)=>a+(proj(v)||0),0);
  const avg=(vs,proj)=>{ const arr=vs.map(proj).filter(x=>x!=null); return arr.length?arr.reduce((a,b)=>a+b,0)/arr.length:0; };
  const groupByPartner=rows=>{ const m=new Map(); for(const r of rows){ if(!m.has(r.partner)) m.set(r.partner,[]); m.get(r.partner).push(r); } return m; };
  const qsaMain=sel=>Array.from(document.querySelectorAll(sel)).filter(el=>!el.closest(PANEL_ID));
  const DEBUG = localStorage.getItem('fvpr-debug') === '1';

  // ---------- IndexedDB ----------
  const IDB_NAME='fvpr_db', IDB_VER=1;
  function idbOpen(){
    return new Promise((res,rej)=>{
      const req=indexedDB.open(IDB_NAME, IDB_VER);
      req.onupgradeneeded=()=>{
        const db=req.result;
        if(!db.objectStoreNames.contains('partners')) db.createObjectStore('partners',{keyPath:'name'});
        if(!db.objectStoreNames.contains('settings')){
          const s=db.createObjectStore('settings',{keyPath:'id'});
          s.put({id:'global', distTo:'', distCc:'', subjectPrefix:'Aktueller Tour.Report', signature:'', httpGateway:''});
        }
      };
      req.onsuccess=()=>res(req.result);
      req.onerror =()=>rej(req.error);
    });
  }
  async function idbGet(store,key){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readonly').objectStore(store).get(key); r.onsuccess=()=>res(r.result||null); r.onerror=()=>rej(r.error); }); }
  async function idbPut(store,val){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readwrite').objectStore(store).put(val); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
  async function idbDel(store,key){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readwrite').objectStore(store).delete(key); r.onsuccess=()=>res(true); r.onerror=()=>rej(r.error); }); }
  async function idbAll(store){ const db=await idbOpen(); return new Promise((res,rej)=>{ const r=db.transaction(store,'readonly').objectStore(store).getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error); }); }
  async function ensureSettingsRecord(){
  const def = {
    id:'global',
    // nur diese drei sind im UI sichtbar
    distTo:'', distCc:'', subjectPrefix:'Aktueller Tour.Report',
    // interne Felder (nicht im UI, aber von deliverMail genutzt)
    signature:'',                         // falls du später wieder willst
    httpGateway: GATEWAY_DEFAULT,         // aus Konstante
    apiKey:      GATEWAY_API_KEY          // aus Konstante
  };
  const cur = await idbGet('settings','global');
  if (!cur) { await idbPut('settings', def); return def; }
  // Backfill fehlende Keys, aber vorhandene Werte nicht überschreiben
  for (const k of Object.keys(def)) if (!(k in cur)) cur[k] = def[k];
  await idbPut('settings', cur);
  return cur;
}
async function exportDb(){
  try{
    const settings = await idbGet('settings','global');
    const partners = await idbAll('partners');

    const data = {
      version: '3.1.0',
      exportedAt: new Date().toISOString(),
      // nur die 3 sichtbaren Felder exportieren
      settings: {
        subjectPrefix: settings?.subjectPrefix || 'Aktueller Tour.Report',
        distTo:        settings?.distTo || '',
        distCc:        settings?.distCc || ''
      },
      partners: partners || []
    };

    const json = JSON.stringify(data, null, 2);
    downloadFile(`fvpr_export_${dateStamp()}.json`, 'application/json', json);
    toast('Export erstellt');
  }catch(e){
    console.error('[fvpr] Export fehlgeschlagen', e);
    toast('Export fehlgeschlagen', false);
  }
}

function getGridViewport(){
  // nimm das erste Grid/Datagrid, das die Fahrzeugübersicht enthält
  const grid = qsaMain('[role="grid"], .Datagrid__Root-sc-')[0] || document.querySelector('[role="grid"]');
  if(!grid) return null;
  // typischerweise ist das scrollbare Element selbst das Grid oder sein Parent
  const el = grid.closest('[class*="Datagrid"]') || grid;
  return el;
}



    function toast(msg, ok=true){
  const el=document.createElement('div');
  el.style.cssText='position:fixed;right:16px;bottom:16px;padding:10px 14px;border-radius:10px;font:600 13px system-ui;color:#fff;z-index:2147483647;' +
                   (ok?'background:#16a34a':'background:#b91c1c');
  el.textContent=msg;
  document.body.appendChild(el);
  setTimeout(()=>{ el.style.opacity='0'; el.style.transition='opacity .3s'; setTimeout(()=>el.remove(),300); }, 1400);
}
let cfgDirty=false;

    // Diagnose
    async function dumpDiag(){
  const C = findColumns();
  console.log('[fvpr] Spalten-Map:', C);

  const { ok, rows } = await readAllRows();
  if(!ok){ console.warn('[fvpr] Keine Rows'); alert('Keine Daten gefunden'); return; }

  // Rohwerte aus dem DOM für die ersten 25 Zeilen zeigen – inkl. der Zelleninhalte der Stopps-Spalte
  const ths = qsaMain('thead th,[role="columnheader"]');
  const headerText  = i => (ths[i]?.textContent || '').replace(/\s+/g,' ').trim();
  const headerTitle = i => (ths[i]?.querySelector('input[title], [title]')?.getAttribute('title') || '').trim();

  console.log('[fvpr] Header[stopsTotal]: text="%s", title="%s"', headerText(C.stopsTotal), headerTitle(C.stopsTotal));

  const trs = qsaMain('tbody tr,[role="row"]');
  const sample = [];
  for (let idx=0; idx<Math.min(25, trs.length); idx++){
    const tds = Array.from(trs[idx].querySelectorAll('td,[role="gridcell"]'));
    const cell = tds[C.stopsTotal];
    sample.push({
      i: idx,
      partnerCell: (tds[C.sys]?.textContent||'').trim(),
      stopsCellText: (cell?.textContent || '').trim(),
      stopsParsed: parseIntDe(cell?.textContent || '')
    });
  }
  console.table(sample);

  // Auch die bereits geparsten Rows ausgeben (was ins Aggregat geht):
  console.table(rows.slice(0,25).map((r,i)=>({i, partner:r.partner, tour:r.tour, stops:r.stops, open:r.open, pkgs:r.pkgs})));

  // Summen so wie das Script sie berechnet:
  const sumStops = rows.reduce((a,r)=>a+(r.stops||0),0);
  const sumOpen  = rows.reduce((a,r)=>a+(r.open||0),0);
  console.log('[fvpr] Summen (alle eingesammelten Zeilen): stops=%d, open=%d, rows=%d', sumStops, sumOpen, rows.length);
  alert('Daten-Check in der Konsole (F12) → Reiter „Konsole“ / console.table() ansehen.');
}


  // ---------- Spaltenerkennung & Daten ----------
 function findColumns(){
  const ths = qsaMain('thead th,[role="columnheader"]');
  if (!ths.length) return null;

  // Helfer: Text/Titel je Spalte
  const headerText  = i => (ths[i]?.textContent || '').replace(/\s+/g,' ').trim().toLowerCase();
  const headerTitle = i => (ths[i]?.querySelector('input[title], [title]')?.getAttribute('title') || '').trim().toLowerCase();

  // Alle Header einmal in Array sammeln, damit wir mehrfach filtern können
  const H = Array.from({length: ths.length}, (_,i)=>({
    i,
    text:  headerText(i),
    title: headerTitle(i),
  }));

  // Helper: nach Titel exakt, sonst Text enthält alle Wörter
  const byTitleExact = (s) => H.find(h => h.title === s)?.i ?? -1;
  const byTextAll    = (...need) => (H.find(h => need.every(w => h.text.includes(w)))?.i ?? -1);

  // 1) Systempartner / Tour / Driver / ETA (wie gehabt, tolerant)
  const sys    = H.find(h => /\bsystempartner\b/.test(h.text))?.i ?? -1;
  const tour   = H.find(h => /^\s*tour\s*$/.test(h.text) || /\btour\s*nr\b/.test(h.text))?.i ?? -1;
  const driver = H.find(h => /(zustellername|fahrername|fahrer|driver)/.test(h.text))?.i ?? -1;
  const eta    = H.find(h => /(eta\s*%$|\beta\s*%$)/.test(h.text))?.i ?? -1;

  // 2) **Zustell**-Spalten: zuerst exakter TITLE, dann Fallback auf TEXT
  const stopsTotal =
      byTitleExact('zustellstopps gesamt') >= 0 ? byTitleExact('zustellstopps gesamt')
    : byTextAll('zustellstopps','gesamt')        >= 0 ? byTextAll('zustellstopps','gesamt')
    : byTextAll('stopp','gesamt'); // letzter Fallback (Notlösung)

  const stopsOpen =
      byTitleExact('offene zustellstopps') >= 0 ? byTitleExact('offene zustellstopps')
    : byTextAll('offene','zustellstopps')   >= 0 ? byTextAll('offene','zustellstopps')
    : byTextAll('offene','stopps');

  const pkgsTotal =
      byTitleExact('geplante zustellpakete') >= 0 ? byTitleExact('geplante zustellpakete')
    : byTextAll('geplante','zustellpakete')   >= 0 ? byTextAll('geplante','zustellpakete')
    : byTextAll('pakete','gesamt');

  // 3) Rest wie gehabt
  const obstacles  = H.find(h => /\bzustellhindernisse\b/.test(h.text) || /\bhinderniss?e?\b/.test(h.text))?.i ?? -1;
  const pickupOpen =
      byTitleExact('offene abholstopps') >= 0 ? byTitleExact('offene abholstopps')
    : byTextAll('offene','abholstopp')   >= 0 ? byTextAll('offene','abholstopp')
    : byTextAll('abhol','offen');

  const cols = { sys, tour, driver, eta, stopsTotal, stopsOpen, pkgsTotal, obstacles, pickupOpen };
  if (cols.sys < 0) return null;

  // Debug: einmal sehen, was gewählt wurde
  console.debug('[fvpr] erkannte Spalten:', cols, H.map(h=>({i:h.i,text:h.text,title:h.title})));
  return cols;
}

    async function readAllRows(maxSteps=100){
  const vp = getGridViewport();
  // falls kein Scroll-Container gefunden -> Fallback: bisherige Logik
  if(!vp || vp.scrollHeight <= vp.clientHeight) return readRows();

  const seen = new Set();
  const acc  = [];

  // zum Anfang
  vp.scrollTop = 0;

  // kleine Wartehilfe
  const sleep = ms => new Promise(r=>setTimeout(r, ms));

  let steps = 0, lastCount = -1, stagnation = 0;
  while(steps++ < maxSteps){
    // aktuelle Sicht auslesen
    const {ok, rows} = readRows();
    if(ok){
      for(const r of rows){
        const key = `${r.partner}||${r.tour}||${r.driver}`;
        if(!seen.has(key)){
          seen.add(key);
          acc.push(r);
        }
      }
    }

    // Abbruchkriterium: Ende erreicht oder keine neuen mehr gefunden
    const atEnd = Math.ceil(vp.scrollTop + vp.clientHeight) >= vp.scrollHeight;
    if(acc.length === lastCount) stagnation++; else { stagnation = 0; lastCount = acc.length; }
    if(atEnd || stagnation >= 3) break;

    // weiter scrollen
    vp.scrollTop = Math.min(vp.scrollTop + vp.clientHeight * 0.9, vp.scrollHeight);
    await sleep(60); // DOM nachladen lassen
  }
  return { ok: acc.length>0, rows: acc };
}

  function readRows(){
    const C=findColumns(); if(!C) return {ok:false,rows:[]};
    const trs=qsaMain('tbody tr,[role="row"]').filter(tr=>{
      const cells=tr.querySelectorAll('td,[role="gridcell"]');
      const needed=Math.max(...Object.values(C).filter(v=>typeof v==='number'&&v>=0));
      return cells && cells.length>needed;
    });
    const out=[];
    for(const tr of trs){
      const tds=Array.from(tr.querySelectorAll('td,[role="gridcell"]'));
      const partner=norm(tds[C.sys]?.textContent||''); if(!partner) continue;
      out.push({
        partner,
        tour:      C.tour>=0?norm(tds[C.tour]?.textContent):'',
        driver:    C.driver>=0?norm((tds[C.driver]?.querySelector('div[title]')?.getAttribute('title'))||tds[C.driver]?.textContent):'',
        eta:       C.eta>=0?parsePct(tds[C.eta]?.textContent):null,
        stops:     C.stopsTotal>=0?parseIntDe(tds[C.stopsTotal]?.textContent):null,
        open:      C.stopsOpen>=0?parseIntDe(tds[C.stopsOpen]?.textContent):null,
        pkgs:      C.pkgsTotal>=0?parseIntDe(tds[C.pkgsTotal]?.textContent):null,
        obstacles: C.obstacles>=0?parseIntDe(tds[C.obstacles]?.textContent):null,
        pOpen:     C.pickupOpen>=0?parseIntDe(tds[C.pickupOpen]?.textContent):null,
      });
    }
    return {ok:true,rows:out};
  }

  // ---------- Styles ----------
  function ensureStyles(){
    if(document.getElementById(NS+'style')) return;
    const s=document.createElement('style'); s.id=NS+'style';
    s.textContent=`
      .${NS}wrap{position:fixed;top:8px;left:calc(50% + ${OFFSET_PX}px);display:flex;gap:8px;z-index:2147483647}
      .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:8px 14px;border-radius:999px;font:600 13px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
      .${NS}panel{position:fixed;top:48px;left:calc(50% + ${OFFSET_PX}px);width:min(1150px,96vw);max-height:76vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:2147483646}
      .${NS}hdr{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
      .${NS}pill{display:inline-flex;gap:6px;align-items:center;flex-wrap:wrap}
      .${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer;margin-left:6px}
      .${NS}tbl{width:100%;border-collapse:collapse}
      .${NS}tbl thead th{position:sticky;top:0;z-index:1;padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.12);font:700 12px system-ui;text-align:right;white-space:nowrap;background:#ffe2e2;color:#8b0000}
      .${NS}tbl th:first-child,.${NS}tbl td:first-child{text-align:left}
      .${NS}tbl tbody tr{cursor:pointer}
      .${NS}tbl tbody tr:hover{background:#f8fafc}
      .${NS}tbl tbody td{padding:8px 10px;border-bottom:1px solid rgba(0,0,0,.06);font:500 12px system-ui;text-align:right;white-space:nowrap}
      .${NS}tbl tbody td.${NS}act{ text-align:center }
      .${NS}tbl tfoot td{padding:8px 10px;border-top:1px solid rgba(0,0,0,.12);font:700 12px system-ui;background:#e0f2ff;color:#003366;text-align:right;white-space:nowrap}
      .${NS}tbl tfoot td:first-child{text-align:left}
      .${NS}empty{padding:12px;text-align:center;opacity:.7}
      .${NS}cfg{padding:10px;border-top:1px solid rgba(0,0,0,.06);background:#fafafa}
      .${NS}cfg input, .${NS}cfg textarea{width:100%;padding:6px;border:1px solid #e5e7eb;border-radius:8px}
      .${NS}cfg textarea{min-height:90px}
      .${NS}row{display:grid;grid-template-columns:1fr 2fr;gap:10px;margin:6px 0}
      .${NS}row3{display:grid;grid-template-columns:1fr 2fr 1fr;gap:10px;margin:6px 0}
      .${NS}modal{position:fixed;inset:0;background:rgba(0,0,0,.35);display:flex;align-items:center;justify-content:center;z-index:2147483647}
      .${NS}modal-box{background:#fff;min-width:min(560px,96vw);max-width:96vw;border-radius:12px;box-shadow:0 20px 40px rgba(0,0,0,.25);padding:14px}
      .${NS}modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:10px}
      .${NS}sep{width:1px;height:24px;background:rgba(0,0,0,.1);margin:0 2px;border-radius:1px}
      .${NS}note{opacity:.75;font-size:12px}
      .${NS}mini{font-size:12px;opacity:.8}
      .${NS}iconbtn{padding:4px 8px;border:1px solid rgba(0,0,0,.15);border-radius:999px;background:#fff;cursor:pointer}
       ${!DEBUG ? `.${NS}pill [data-act="diag"]{display:none!important}` : ''}


    `;
    document.head.appendChild(s);
  }

  // ---------- UI ----------
  async function mountUI(){
    ensureStyles();
    if(document.getElementById(NS+'wrap')) return;

    const wrap=document.createElement('div'); wrap.id=NS+'wrap'; wrap.className=NS+'wrap';
    const btn=document.createElement('button'); btn.className=NS+'btn'; btn.textContent='Partner-Report';
    wrap.append(btn); document.body.appendChild(wrap);

    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel'; panel.style.display='none';
    panel.setAttribute('id',PANEL_ID.slice(1));
    panel.innerHTML=`
      <div class="${NS}hdr">
        <div>Auswertung – Systempartner (Fahrzeugübersicht) <span class="${NS}mini">[Stand: ${todayStr()} ${timeHM()}]</span></div>
        <div class="${NS}pill">
          <button class="${NS}btn-sm" data-act="refresh">Aktualisieren</button>
          <span class="${NS}sep"></span>
          <button class="${NS}btn-sm" data-act="send-dist">Gesamt an Verteiler (Mail)</button>
          <button class="${NS}btn-sm" data-act="eml-dist">Gesamt als EML</button>
          <span class="${NS}sep"></span>
          <button class="${NS}btn-sm" data-act="send-partner">Pro Partner (Mail)</button>
          <button class="${NS}btn-sm" data-act="eml-partner">Pro Partner (EML)</button>
          <span class="${NS}sep"></span>
          <button class="${NS}btn-sm" data-act="settings">Einstellungen</button>
          <button class="${NS}btn-sm" data-act="diag">Daten-Check</button>
        </div>
      </div>
      <div id="${NS}content"></div>
      <div class="${NS}cfg" style="display:none" id="${NS}cfgbox">
  <h4 style="margin:0 0 6px 0;font:700 14px system-ui">Globale Einstellungen</h4>
  <div class="${NS}row"><label>Betreff-Prefix</label><input id="${NS}cfg-subj" type="text"></div>
  <div class="${NS}row"><label>Verteiler „An“</label><input id="${NS}cfg-to" type="text" placeholder="kommagetrennt: a@b.de, c@d.de"></div>
  <div class="${NS}row"><label>Verteiler „CC“</label><input id="${NS}cfg-cc" type="text" placeholder="optional"></div>

  <button class="${NS}btn-sm" data-act="cfg-save">Speichern</button>
  <button class="${NS}btn-sm" data-act="cfg-hide">Schließen</button>
  <button class="${NS}btn-sm" data-act="export">Export</button>
  <button class="${NS}btn-sm" data-act="import">Import</button>
</div>

        <div class="${NS}note">Per-Partner: In der Tabelle eine Zeile anklicken oder „✉︎“ drücken, um An/CC/Alias zu pflegen bzw. direkt zu senden.</div>
        <input type="file" id="${NS}impfile" accept="application/json" style="display:none">
      </div>
    `;
    document.body.appendChild(panel);

     document.addEventListener('keydown', (e)=>{
  if(e.ctrlKey && e.altKey && e.key.toLowerCase()==='d'){
    const now = !(localStorage.getItem('fvpr-debug') === '1');
    localStorage.setItem('fvpr-debug', now ? '1' : '0');
    document.querySelectorAll('[data-act="diag"]')
      .forEach(b => b.style.display = now ? '' : 'none');
    toast(`Debug-Button ${now ? 'sichtbar' : 'versteckt'}`);
  }
});

    const toggle = async (force)=>{
  const open = panel.style.display!=='none';
  const will = (force!==undefined ? force : !open);
  panel.style.display = will ? '' : 'none';
  if (will) { await fillCfg(); wireCfgInputs(); }
};

    btn.addEventListener('click',()=>toggle());
    document.addEventListener('click',e=>{ if(panel.style.display==='none') return; if(panel.contains(e.target)||wrap.contains(e.target)) return; toggle(false); });

    // Zeilenklick = Dialog
    document.getElementById(NS+'content').addEventListener('click', e=>{
      const sendBtn=e.target.closest('button[data-sp]');
      if(sendBtn){ sendSinglePartner(sendBtn.dataset.sp); e.stopPropagation(); return; }
      const tr=e.target.closest('tbody tr'); if(!tr) return;
      const partner=tr.getAttribute('data-partner'); if(partner) openPartnerDialog(partner);
    });

    // <- Das ist der "Panel-Handler"
panel.addEventListener('click', async e => {
  const b = e.target.closest('button[data-act]');
  if (!b) return;

  if (b.dataset.act === 'refresh') render(true);
  if (b.dataset.act==='settings'){
  const box=document.getElementById(NS+'cfgbox');
  box.style.display = box.style.display==='none' ? '' : 'none';
  if (box.style.display!=='none'){ await fillCfg(); wireCfgInputs(); }
}

  if (b.dataset.act === 'cfg-save') await saveCfgFromUI();
  if (b.dataset.act === 'cfg-hide') document.getElementById(NS+'cfgbox').style.display='none';
  if (b.dataset.act === 'send-dist')  await sendSummaryToDist();
  if (b.dataset.act === 'eml-dist')   await emlSummaryToDist();
  if (b.dataset.act === 'send-partner') await sendPerPartner();
  if (b.dataset.act === 'eml-partner')  await emlPerPartner();
  if (b.dataset.act === 'export') await exportDb();
  if (b.dataset.act === 'import') document.getElementById(NS+'impfile').click();
  if (b.dataset.act === 'diag') {
  await dumpDiag();
  return;


}


  // <- hier kannst du zusätzliche Aktionen anhängen, z.B. den Test-Button:
  if (b.dataset.act === 'cfg-test') {
    const g = await idbGet('settings','global');
    console.log('settings/global =', g);
    alert('Siehe Konsole (F12) → settings/global');
  }
});


    document.getElementById(NS+'impfile').addEventListener('change', async (ev)=>{
      const f=ev.target.files?.[0]; if(!f) return;
      const text=await f.text();
      try{
        const data=JSON.parse(text);
        if(data.settings) await idbPut('settings',{id:'global', ...data.settings});
        if(Array.isArray(data.partners)){
          const db=await idbOpen();
          const tx=db.transaction('partners','readwrite'); const st=tx.objectStore('partners');
          const keys=await new Promise(r=>{ const k=st.getAllKeys(); k.onsuccess=()=>r(k.result||[]); });
          await Promise.all(keys.map(k=>new Promise(r=>{ const d=st.delete(k); d.onsuccess=r; })));
          for(const p of data.partners){ st.put(p); }
          await new Promise(r=>{ tx.oncomplete=r; });
        }
        alert('Import erfolgreich.');
        fillCfg(); render(true);
      }catch(e){ console.error(e); alert('Import fehlgeschlagen (ungültiges JSON).'); }
    });

    render();
      await fillCfg();
      wireCfgInputs();

  }

  async function fillCfg(){
  const g = await ensureSettingsRecord();
  const setVal = (id, val) => { const el = document.getElementById(id); if (el) el.value = val; };
  setVal(NS+'cfg-subj', g.subjectPrefix || 'Aktueller Tour.Report');
  setVal(NS+'cfg-to',   g.distTo || '');
  setVal(NS+'cfg-cc',   g.distCc || '');
  // keine weiteren setVal-Aufrufe (kein cfg-sign / cfg-gw / cfg-key)
}


    function wireCfgInputs(){
  cfgDirty = false;
  const ids = [NS+'cfg-subj', NS+'cfg-to', NS+'cfg-cc']; // nur diese drei
  for (const id of ids){
    const i = document.getElementById(id);
    if (!i) continue;
    i.oninput  = ()=>{ cfgDirty=true; };
    i.onchange = ()=>{ cfgDirty=true; };
  }
  const box = document.getElementById(NS+'cfgbox');
  box.onkeydown = (ev)=>{
    if((ev.ctrlKey||ev.metaKey) && ev.key.toLowerCase()==='s'){
      ev.preventDefault();
      document.querySelector(`button[data-act="cfg-save"]`)?.click();
    }
  };
}



  async function saveCfgFromUI(){
  try{
    const cur = await ensureSettingsRecord(); // holt aktuelle inkl. interner Felder
    const rec = {
      id:'global',
      // nur UI-Felder überschreiben:
      subjectPrefix: (document.getElementById(NS+'cfg-subj').value || '').trim(),
      distTo:        (document.getElementById(NS+'cfg-to').value   || '').trim(),
      distCc:        (document.getElementById(NS+'cfg-cc').value   || '').trim(),
      // interne Felder unverändert übernehmen:
      signature:  cur.signature || '',
      httpGateway: cur.httpGateway || GATEWAY_DEFAULT,
      apiKey:      cur.apiKey || GATEWAY_API_KEY,
    };

    await idbPut('settings', rec);

    // Plausibilisierung (nur die drei vergleichen)
    const back = await idbGet('settings','global');
    if(!back){ toast('Speichern fehlgeschlagen', false); return; }
    const same =
      back.subjectPrefix===rec.subjectPrefix &&
      back.distTo===rec.distTo &&
      back.distCc===rec.distCc;

    if(!same){
      console.error('Mismatch after save', {rec, back});
      toast('Speichern fehlgeschlagen (Mismatch)', false);
      return;
    }

    cfgDirty=false;
    toast('Einstellungen gespeichert');
    await fillCfg(); // UI aktualisieren
    wireCfgInputs();
  }catch(e){
    console.error(e);
    toast('Fehler beim Speichern', false);
  }
}

  // ---------- Partner-Dialog ----------
  async function openPartnerDialog(partner){
    const cur=await idbGet('partners', partner) || {name:partner,to:'',cc:'',alias:''};
    const ov=document.createElement('div'); ov.className=NS+'modal';
    ov.innerHTML=`
      <div class="${NS}modal-box">
        <h3 style="margin:0 0 8px 0;font:700 16px system-ui">Einstellungen – ${partner}</h3>
        <div class="${NS}row"><label>Alias (optional)</label><input id="${NS}pd-alias" type="text" value="${cur.alias||''}"></div>
        <div class="${NS}row"><label>An</label><input id="${NS}pd-to" type="text" value="${cur.to||''}" placeholder="a@b.de, c@d.de"></div>
        <div class="${NS}row"><label>CC</label><input id="${NS}pd-cc" type="text" value="${cur.cc||''}" placeholder="optional"></div>
        <div class="${NS}modal-actions">
          <button class="${NS}btn-sm" data-act="send">✉︎ Senden</button>
          <button class="${NS}btn-sm" data-act="save">Speichern</button>
          <button class="${NS}btn-sm" data-act="clear">Löschen</button>
          <button class="${NS}btn-sm" data-act="close">Schließen</button>
        </div>
      </div>`;
    document.body.appendChild(ov);
    ov.addEventListener('click', async e=>{
      if(e.target===ov) ov.remove();
      const b=e.target.closest('button[data-act]'); if(!b) return;
      if(b.dataset.act==='close') ov.remove();
      if(b.dataset.act==='clear'){ await idbDel('partners', partner); ov.remove(); }
      if(b.dataset.act==='save'){
        const to=ov.querySelector('#'+NS+'pd-to').value.trim();
        const cc=ov.querySelector('#'+NS+'pd-cc').value.trim();
        const alias=ov.querySelector('#'+NS+'pd-alias').value.trim();
        await idbPut('partners',{name:partner,to,cc,alias});
        alert('Gespeichert.');
      }
      if(b.dataset.act==='send'){
        const to=ov.querySelector('#'+NS+'pd-to').value.trim();
        const cc=ov.querySelector('#'+NS+'pd-cc').value.trim();
        const alias=ov.querySelector('#'+NS+'pd-alias').value.trim();
        await idbPut('partners',{name:partner,to,cc,alias});
        await sendSinglePartner(partner);
      }
    });
  }

  // ---------- Mail/EML/Clipboard ----------
  async function copyHtmlToClipboard(html){ try{
    if(navigator.clipboard&&window.ClipboardItem){
      const item=new ClipboardItem({'text/html':new Blob([html],{type:'text/html'})});
      await navigator.clipboard.write([item]);
    } else {
      const d=document.createElement('div'); d.style.position='fixed'; d.style.left='-99999px'; d.innerHTML=html;
      document.body.appendChild(d); const r=document.createRange(); r.selectNodeContents(d);
      const sel=window.getSelection(); sel.removeAllRanges(); sel.addRange(r); document.execCommand('copy'); sel.removeAllRanges(); d.remove();
    }
    return true;
  }catch(e){ console.error(e); return false; } }
  function openMailto(subject,to='',cc=''){
    const href=`mailto:${encodeURIComponent(to||'')}?subject=${encodeURIComponent(subject)}${cc?`&cc=${encodeURIComponent(cc)}`:''}`;
    window.open(href,'_blank');
  }
  function qpEncodeUtf8(str){
    const utf8=new TextEncoder().encode(str);
    let out='',lineLen=0; const hex=b=>b.toString(16).toUpperCase().padStart(2,'0');
    for(const b of utf8){ const ch=String.fromCharCode(b); const safe=(b===0x09||b===0x20)||(b>=33&&b<=60)||(b>=62&&b<=126); const token=(safe&&b!==0x3D)?ch:'='+hex(b); if(lineLen+token.length>73){ out+='=\r\n'; lineLen=0; } out+=token; lineLen+=token.length; }
    return out;
  }
  function buildEml(subject,html){
    const boundary='=_fvpr_'+Math.random().toString(36).slice(2);
    const headers=['From: ','To: ',`Subject: ${subject}`,'MIME-Version: 1.0',`Date: ${new Date().toUTCString()}`,`Content-Type: multipart/alternative; boundary="${boundary}"`].join('\r\n');
    const body=`--${boundary}\r\nContent-Type: text/plain; charset="UTF-8"\r\nContent-Transfer-Encoding: 7bit\r\n\r\nBitte HTML-Anzeige aktivieren.\r\n\r\n--${boundary}\r\nContent-Type: text/html; charset="UTF-8"\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n${qpEncodeUtf8(html)}\r\n\r\n--${boundary}--\r\n`;
    return headers+'\r\n\r\n'+body;
  }
  function downloadFile(filename,mime,content){
    const a=document.createElement('a'); a.href=URL.createObjectURL(new Blob([content],{type:mime})); a.download=filename; a.click(); setTimeout(()=>URL.revokeObjectURL(a.href),0);
  }

  // ---------- HTML Renderer ----------
  function partnerHtml(partner,list,signature){
    const rowsHtml=list.map(r=>`
      <tr>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${partner}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;">${r.tour||'—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${r.driver||'—'}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.stops)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pkgs)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.open)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.obstacles)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(r.eta)}</td>
        <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(r.pOpen)}</td>
      </tr>`).join('');
    const totals={ tours:list.length, etaAvg:avg(list,r=>r.eta), stops:sum(list,r=>r.stops), pkgs:sum(list,r=>r.pkgs), open:sum(list,r=>r.open), obstacles:sum(list,r=>r.obstacles), pOpen:sum(list,r=>r.pOpen) };
    const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';
    return `
      <div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
        <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
        <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
          <thead><tr style="background:#ffe2e2;color:#8b0000;font-weight:700;">
            <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">SP</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">Tour</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Fahrername</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">offen</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA</th>
            <th style="padding:6px 8px;border:1px solid #e5e7eb;">offene Abholstops</th>
          </tr></thead>
          <tbody>${rowsHtml}
            <tr style="background:#e0f2ff;color:#003366;font-weight:700;">
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt</td>
              <td style="padding:8px;border:1px solid #e5e7eb;"></td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Touren: ${fmtInt(totals.tours)}</td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">ETA Ø: ${fmtPct(totals.etaAvg)}</td>
              <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
            </tr>
          </tbody>
        </table>
        ${signatureHtml}
      </div>`;
  }
  function summaryHtml(per, totals, signature){
    const head=`<thead><tr>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">Systempartner</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Touren</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">ETA % (Ø)</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Stopps gesamt</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Offene Stopps</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Pakete gesamt</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Zustellhindernisse</th>
      <th style="padding:6px 8px;border:1px solid #e5e7eb;">Offene Abholstops</th>
    </tr></thead>`;
    const body=per.map(p=>`<tr>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:left;">${p.partner}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.tours)}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(p.etaAvg)}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.stops)}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.open)}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.pkgs)}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.obstacles)}</td>
      <td style="padding:6px 8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(p.pOpen)}</td>
    </tr>`).join('');
   const foot=`<tfoot><tr style="background:#e0f2ff;color:#00366;font-weight:700;">
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:left;">Gesamt</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.tours)}</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtPct(totals.etaAvg)}</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.stops)}</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.open)}</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pkgs)}</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.obstacles)}</td>
      <td style="padding:8px;border:1px solid #e5e7eb;text-align:right;">${fmtInt(totals.pOpen)}</td>
    </tr></tfoot>`;
    const signatureHtml = signature ? `<div style="margin-top:10px">${signature}</div>` : '';
    return `<div style="font:14px/1.5 system-ui,Segoe UI,Arial,sans-serif;">
      <div style="margin:0 0 6px 0;color:#334155">Stand: ${todayStr()} ${timeHM()}</div>
      <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;">
        ${head}<tbody>${body}</tbody>${foot}
      </table>${signatureHtml}</div>`;
  }

  // ---------- Aggregation ----------
  async function getAggregates(){
    const {ok,rows}=await readAllRows();   // <- sammelt alle virtuellen Zeilen
     if(!ok||rows.length===0) return null;
    const groups=groupByPartner(rows);
    const per=[];
    for(const [partner,list] of groups){
      per.push({ partner, list,
        tours:list.length, etaAvg:avg(list,r=>r.eta), stops:sum(list,r=>r.stops),
        open:sum(list,r=>r.open), pkgs:sum(list,r=>r.pkgs), obstacles:sum(list,r=>r.obstacles), pOpen:sum(list,r=>r.pOpen),
      });
    }
    per.sort((a,b)=>a.partner.localeCompare(b.partner,'de'));
    const totals={
      tours:per.reduce((a,p)=>a+(p.tours||0),0),
      etaAvg:avg(rows,r=>r.eta), stops:per.reduce((a,p)=>a+(p.stops||0),0),
      open:per.reduce((a,p)=>a+(p.open||0),0), pkgs:per.reduce((a,p)=>a+(p.pkgs||0),0),
      obstacles:per.reduce((a,p)=>a+(p.obstacles||0),0), pOpen:per.reduce((a,p)=>a+(p.pOpen||0),0),
    };
    return {per, totals};
  }

  // ---------- Versand gesamt / partner ----------
  async function sendSummaryToDist(){
    const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
    const g=await ensureSettingsRecord();
    const subject=`${g.subjectPrefix||'Aktueller Tour.Report'} – Gesamt – ${todayStr()}`;
    const html=summaryHtml(agg.per, agg.totals, g.signature||'');
    await deliverMail({subject, html, to:g.distTo, cc:g.distCc});
  }
  async function emlSummaryToDist(){
    const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
    const g=await ensureSettingsRecord();
    const subject=`${g.subjectPrefix||'Aktueller Tour.Report'} – Gesamt – ${todayStr()}`;
    const html=summaryHtml(agg.per, agg.totals, g.signature||'');
    const eml=buildEml(subject,html);
    downloadFile(`TourReport_Gesamt_${dateStamp()}.eml`,'message/rfc822',eml);
  }
  async function sendPerPartner(){
    const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
    const g=await ensureSettingsRecord();
    for(const p of agg.per){ await sendSinglePartner(p.partner, {agg,g}); }
  }
  async function emlPerPartner(){
    const agg=await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
    const g=await ensureSettingsRecord();
    const stamp=dateStamp();
    let count=0;
    for(const p of agg.per){
      const ov=await idbGet('partners', p.partner);
      const alias=ov?.alias || p.partner;
      const subject=`${g.subjectPrefix||'Aktueller Tour.Report'} – ${alias} – ${todayStr()}`;
      const html=partnerHtml(p.partner, p.list, g.signature||'');
      const eml=buildEml(subject, html);
      const fname=`TourReport_${alias.replace(/[^A-Za-z0-9._-]+/g,'_').slice(0,80)}_${stamp}.eml`;
      downloadFile(fname,'message/rfc822',eml);
      count++;
    }
    alert(`${count} EML-Dateien erzeugt.`);
  }
  async function sendSinglePartner(partner, preload){
    const agg=preload?.agg || await getAggregates(); if(!agg){ alert('Keine Daten gefunden.'); return; }
    const g=preload?.g || await ensureSettingsRecord();
    const p=agg.per.find(x=>x.partner===partner); if(!p){ alert('Partner nicht gefunden.'); return; }
    const ov=await idbGet('partners', partner);
    const alias=ov?.alias || partner;
    const subject=`${g.subjectPrefix||'Aktueller Tour.Report'} – ${alias} – ${todayStr()}`;
    const html=partnerHtml(partner, p.list, g.signature||'');
    await deliverMail({subject, html, to:(ov?.to||g.distTo||''), cc:(ov?.cc||g.distCc||'')});
  }

  function splitEmails(raw){
  // trennt an Komma/Semikolon/Whitespace
  return (raw||'')
    .split(/[,;\s]+/)
    .map(s=>s.trim())
    .filter(Boolean);
}
function isEmail(s){
  // simple & robust (PHPMailer ist toleranter, aber das reicht)
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s);
}
function normalizeEmailList(raw){
  const arr = splitEmails(raw);
  const valid = [];
  const invalid = [];
  const seen = new Set();
  for(const a of arr){
    const low = a.toLowerCase();
    if(seen.has(low)) continue;
    seen.add(low);
    (isEmail(a) ? valid : invalid).push(a);
  }
  return { valid, invalid };
}




  // ---------- Abstraktion Versand (Mailto / optional HTTP-Gateway) ----------
async function deliverMail({subject, html, to, cc}){
  // Listen säubern
  const toL = normalizeEmailList(to);
  const ccL = normalizeEmailList(cc);

  if (toL.valid.length===0){
    toast('Keine gültige Empfängeradresse (An). Bitte prüfen.', false);
    return;
  }
  if (toL.invalid.length || ccL.invalid.length){
    console.warn('[fvpr] Ignoriere ungültige Adressen', {toInvalid:toL.invalid, ccInvalid:ccL.invalid});
    toast(`Ungültige Adressen ignoriert: ${[...toL.invalid, ...ccL.invalid].join(', ')}`, false);
  }

  const g   = await ensureSettingsRecord();
  const url = (g.httpGateway || GATEWAY_DEFAULT).trim();
  const key = (g.apiKey || GATEWAY_API_KEY || '').trim();

  // 1) Tampermonkey-XHR zuerst (Mixed-Content/CORS-sicher)
  if (typeof GM_xmlhttpRequest === 'function' && /^https?:\/\//i.test(url)) {
    try {
      const res = await new Promise((resolve, reject)=>{
        GM_xmlhttpRequest({
          method: 'POST',
          url,
          headers: { 'Content-Type':'application/json', 'X-Api-Key': key },
          data: JSON.stringify({
            subject,
            html,
            to: toL.valid.join(','),
            cc: ccL.valid.join(',')
          }),
          onload: r => resolve(r),
          onerror: e => reject(e),
          ontimeout: () => reject(new Error('timeout')),
          timeout: 10000
        });
      });
      const ok = res.status>=200 && res.status<300;
      let body = null; try{ body = JSON.parse(res.responseText||''); }catch{}
      if (!ok || !body || body.ok!==true) throw new Error(`Gateway-Fehler ${res.status}: ${res.responseText}`);
      toast('Mail über Gateway gesendet');
      return;
    } catch (e){
      console.error('[fvpr] GM gateway error', e);
      toast('Gateway nicht erreichbar – Fallback Outlook', false);
    }
  }

  // 2) Optional: fetch wenn Gateway per HTTPS verfügbar
  if (/^https:\/\//i.test(url)) {
    try {
      const r = await fetch(url, {
        method:'POST',
        headers:{ 'Content-Type':'application/json', 'X-Api-Key': key },
        body: JSON.stringify({
          subject,
          html,
          to: toL.valid.join(','),
          cc: ccL.valid.join(',')
        }),
        mode:'cors',
        keepalive:true
      });
      const t = await r.text(); let j=null; try{ j=JSON.parse(t);}catch{}
      if(!r.ok||!j||j.ok!==true) throw new Error(`HTTP ${r.status}: ${t}`);
      toast('Mail über Gateway gesendet');
      return;
    } catch(e){
      console.error('[fvpr] fetch gateway error', e);
      toast('Gateway nicht erreichbar – Fallback Outlook', false);
    }
  }

  // 3) Fallback Outlook (mit bereinigten Adressen)
  await copyHtmlToClipboard(html);
  openMailto(subject, toL.valid.join(','), ccL.valid.join(','));
  alert('Entwurf geöffnet. HTML ist in der Zwischenablage – Strg+V drücken.');
}


  // ---------- Render Tabelle ----------
let renderTimer = null;
async function render(force=false){
  clearTimeout(renderTimer);

  const run = async ()=>{
    const cont = document.getElementById(NS+'content');
    if (!cont) return;

    // Alle Zeilen einsammeln (auch virtuelle/gescrolte):
    const { ok, rows } = await readAllRows();
    if (!ok || rows.length === 0){
      cont.innerHTML = `<div class="${NS}empty">Keine Daten gefunden (Tab „Fahrzeugübersicht“ – Tabelle sichtbar?).</div>`;
      return;
    }

    const groups = groupByPartner(rows);
    const per = [];
    for (const [partner, list] of groups){
      per.push({
        partner, list,
        tours: list.length,
        etaAvg: avg(list, r=>r.eta),
        stops:  sum(list, r=>r.stops),
        open:   sum(list, r=>r.open),
        pkgs:   sum(list, r=>r.pkgs),
        obstacles: sum(list, r=>r.obstacles),
        pOpen:  sum(list, r=>r.pOpen),
      });
    }
    per.sort((a,b)=>a.partner.localeCompare(b.partner,'de'));

    const totals = {
      tours: per.reduce((a,p)=>a+(p.tours||0),0),
      etaAvg: avg(rows, r=>r.eta),
      stops: per.reduce((a,p)=>a+(p.stops||0),0),
      open:  per.reduce((a,p)=>a+(p.open||0),0),
      pkgs:  per.reduce((a,p)=>a+(p.pkgs||0),0),
      obstacles: per.reduce((a,p)=>a+(p.obstacles||0),0),
      pOpen: per.reduce((a,p)=>a+(p.pOpen||0),0),
    };

    const head = `<thead><tr>
      <th>Systempartner</th><th>Touren</th><th>ETA % (Ø)</th><th>Stopps gesamt</th>
      <th>Offene Stopps</th><th>Pakete gesamt</th><th>Zustellhindernisse</th><th>Offene Abholstops</th>
      <th>Aktion</th>
    </tr></thead>`;

    const body = per.map(p=>`<tr data-partner="${p.partner.replace(/"/g,'&quot;')}">
      <td style="text-align:left">${p.partner}</td><td>${fmtInt(p.tours)}</td><td>${fmtPct(p.etaAvg)}</td>
      <td>${fmtInt(p.stops)}</td><td>${fmtInt(p.open)}</td><td>${fmtInt(p.pkgs)}</td>
      <td>${fmtInt(p.obstacles)}</td><td>${fmtInt(p.pOpen)}</td>
      <td class="${NS}act"><button class="${NS}iconbtn" title="Mail an Partner" data-sp="${p.partner}">✉︎</button></td>
    </tr>`).join('');

    const foot = `<tfoot><tr>
      <td>Gesamt</td><td>${fmtInt(totals.tours)}</td><td>${fmtPct(totals.etaAvg)}</td>
      <td>${fmtInt(totals.stops)}</td><td>${fmtInt(totals.open)}</td><td>${fmtInt(totals.pkgs)}</td>
      <td>${fmtInt(totals.obstacles)}</td><td>${fmtInt(totals.pOpen)}</td><td></td>
    </tr></tfoot>`;

    cont.innerHTML = `<table class="${NS}tbl">${head}<tbody>${body}</tbody>${foot}</table>`;
  };

  if (force) await run(); else renderTimer = setTimeout(run, RENDER_DEBOUNCE);
}

  // ---------- Boot ----------
  async function init(){
    if(document.getElementById(NS+'wrap')) return;
    ensureStyles();
    await ensureSettingsRecord();
    await mountUI();
    render(true);
  }
  const kick=()=>{ try{ init(); }catch(e){ console.error(e); } };
  setInterval(kick,1500);
  new MutationObserver(kick).observe(document.documentElement,{childList:true,subtree:true});
  setInterval(()=>{ try{ render(); }catch{} },REFRESH_MS);
})();
