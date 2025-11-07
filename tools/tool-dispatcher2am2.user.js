// ==UserScript==
// @name         DPD Dispatcher – Prio/Express12 Monitoring
// @namespace    bodo.dpd.custom
// @version      5.3.0
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-dispatcher2am2.user.js
// @description  PRIO/EXPRESS12: KPIs & Listen. Status DE (DOM bevorzugt), sortierbare Tabellen, Zusatzcode, Predict, Zustellzeit, Button „EXPRESS12 >11:01“. Panel bleibt offen; PSN mit Auge-Button öffnet Scanserver.
// @match        https://dispatcher2-de.geopost.com/*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  const NS = 'pm-';
  const sleep = (ms)=>new Promise(r=>setTimeout(r,ms));
  const esc = s => String(s||'').replace(/[&<>"']/g, ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  const fmt = ts => { try { return new Date(ts||Date.now()).toLocaleString('de-DE'); } catch { return String(ts||''); } };

  const state={
      filterExpress: 'all',   // 'all' | '12' | '18'
    events:[], nextId:1, _bootShown:false,
    _prioAllList:[], _prioOpenList:[],
    _expAllList:[],  _expOpenList:[],
    _expLate11List:[],
    _modal:{rows:[], opts:{}, title:''}
  };
  let lastOkRequest=null;
  let autoEnabled = true, autoTimer = null, isBusy = false, isLoading = false;
  const commentCache = new Map(), statusCache = new Map();

  // ---------- styles ----------
  function ensureStyles(){
    if (document.getElementById(NS+'style')) return;
    const style=document.createElement('style'); style.id=NS+'style';
    style.textContent=`
    .${NS}fixed-wrap{position:fixed;top:8px;left:50%;transform:translateX(-50%);display:flex;gap:8px;z-index:99999}
    .${NS}btn{border:1px solid rgba(0,0,0,.12);background:#fff;padding:8px 14px;border-radius:999px;font:600 13px system-ui;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.08)}
    .${NS}badge{min-width:18px;height:18px;line-height:18px;text-align:center;background:#b00020;color:#fff;border-radius:9px;font:700 11px/18px system-ui;padding:0 6px;display:none}
    .${NS}panel{position:fixed;top:48px;left:50%;transform:translateX(-50%);width:min(1100px,95vw);max-height:74vh;overflow:auto;background:#fff;border:1px solid rgba(0,0,0,.12);box-shadow:0 12px 28px rgba(0,0,0,.18);border-radius:12px;z-index:100000}
    .${NS}header{display:grid;grid-template-columns:1fr auto;gap:10px;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
    .${NS}kpis{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
    .${NS}kpi{background:#f5f5f5;border:1px solid rgba(0,0,0,.08);padding:4px 8px;border-radius:999px;font:600 12px system-ui}
    .${NS}kpi[data-action]{cursor:pointer}
    .${NS}actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
    .${NS}btn-sm{border:1px solid rgba(0,0,0,.12);background:#f7f7f7;padding:6px 10px;border-radius:8px;font:600 12px system-ui;cursor:pointer}
    .${NS}filter{display:flex;gap:6px;align-items:center;font:600 12px system-ui}
    .${NS}filter select{padding:4px 8px;border-radius:8px;border:1px solid rgba(0,0,0,.15);background:#fff}
    .${NS}list{list-style:none;margin:0;padding:0}
    .${NS}item{padding:12px 14px;border-left:4px solid transparent;border-bottom:1px solid rgba(0,0,0,.06)}
    .${NS}title{font:700 14px system-ui;margin:0 0 6px;text-align:center}
    .${NS}meta{font:400 12px/1.35 system-ui;opacity:.85}
    .${NS}comment{margin-top:6px;font:500 12px/1.35 system-ui;background:#fafafa;border:1px solid rgba(0,0,0,.08);border-radius:8px;padding:6px 8px}
    .${NS}sev-info{background:#f7f7f7;border-left-color:#bdbdbd}
    .${NS}empty{padding:14px 12px;opacity:.75;text-align:center;font:500 12px system-ui}
    .${NS}dim{opacity:.6;pointer-events:none}
    .${NS}chip{display:inline-flex;gap:6px;align-items:center;border:1px solid rgba(0,0,0,.12);background:#fff;padding:4px 8px;border-radius:999px;font:600 12px system-ui}
    .${NS}dot{width:8px;height:8px;border-radius:50%;background:#16a34a}
    .${NS}dot.off{background:#9ca3af}
    .${NS}plink{color:inherit;text-decoration:underline;font-weight:700}
    .${NS}loading{display:none;padding:8px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:600 12px system-ui;background:#fffbe6}
    .${NS}loading.on{display:block}
    .${NS}modal{position:fixed;inset:0;display:none;align-items:center;justify-content:center;background:rgba(0,0,0,.35);z-index:100001}
    .${NS}modal-inner{background:#fff;max-width:min(1200px,95vw);max-height:78vh;overflow:auto;border-radius:12px;box-shadow:0 12px 28px rgba(0,0,0,.2);border:1px solid rgba(0,0,0,.12)}
    .${NS}modal-head{display:flex;justify-content:space-between;align-items:center;padding:10px 12px;border-bottom:1px solid rgba(0,0,0,.08);font:700 13px system-ui}
    .${NS}modal-body{padding:8px 12px}
    .${NS}tbl{width:100%;border-collapse:collapse;font:12px system-ui}
    .${NS}tbl th, .${NS}tbl td{border-bottom:1px solid rgba(0,0,0,.08);padding:6px 8px;vertical-align:top}
    .${NS}tbl th{text-align:left;background:#fafafa;position:sticky;top:0;cursor:pointer;user-select:none}
    .${NS}sort-asc::after{content:" ▲";font-size:11px}
    .${NS}sort-desc::after{content:" ▼";font-size:11px}
    .${NS}eye{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border:1px solid rgba(0,0,0,.2);border-radius:6px;background:#fff;margin-right:6px;cursor:pointer;font-size:12px;line-height:1}
    .${NS}eye:hover{background:#f3f4f6}
    /* --- Toolbar oben, KPIs unten --- */
.${NS}header{
  display:grid;
  grid-template-columns: 1fr;
  gap:10px;
  align-items:start;
}

.${NS}toolbar{
  display:flex;
  flex-wrap:wrap;
  justify-content:space-between;
  align-items:center;
  gap:8px;
}

.${NS}group{
  display:flex;
  flex-wrap:wrap;
  align-items:center;
  gap:8px;
  background:#f9fafb;
  border:1px solid rgba(0,0,0,.08);
  border-radius:12px;
  padding:6px 8px;
}

.${NS}label{
  display:flex;
  align-items:center;
  gap:6px;
  font:600 12px system-ui;
}

.${NS}select{
  padding:4px 8px;
  border-radius:8px;
  border:1px solid rgba(0,0,0,.15);
  background:#fff;
}

.${NS}kpis{
  display:flex;
  flex-wrap:wrap;
  justify-content:flex-start;
  gap:8px;
  margin-top:4px;
}

.${NS}kpi{
  flex:1 1 auto;
  background:#f5f5f5;
  border:1px solid rgba(0,0,0,.08);
  padding:6px 10px;
  border-radius:999px;
  font:600 12px system-ui;
  white-space:nowrap;
}

    `;
    document.head.appendChild(style);
  }

  // ==== NEU: Settings (Passwort + Depotkennung) =================================
const LSKEY = 'pmSettings';
function loadSettings(){
  try { return Object.assign({ scanserverPass:'', depotSuffix:'' }, JSON.parse(localStorage.getItem(LSKEY)||'{}')); }
  catch { return { scanserverPass:'', depotSuffix:'' }; }
}
function saveSettings(s){
  try { localStorage.setItem(LSKEY, JSON.stringify(s)); } catch{}
}
function setSetting(k,v){
  const s = loadSettings(); s[k]=v; saveSettings(s);
}
function getSetting(k){ return loadSettings()[k]; }

// Depotkennung aus Fahrzeugübersicht „erraten“ (nimmt erste Zeile unter Kopf: Depot)
// Depotkennung aus Fahrzeugübersicht exakt ermitteln (nur Spalte "Depot")
// NUR im aktiven Fahrzeug-Grid suchen (kein globales Scannen mehr)
function guessDepotSuffixFromVehicleTable(root){
  const grid = root || findVehicleGridContainer();     // <— dein Grid-Finder
  if (!grid) return '';

  const norm = s => String(s||'').replace(/\s+/g,' ').trim();
  const counts  = new Map();
  const pick = (t) => {
    // bevorzugt Host-Muster: d0010###
    let m = t.match(/d0010(\d{3})/i) || t.match(/0010(\d{3})/) || t.match(/010(\d{3})/);
    if (m) return m[1];
    // isolierte 3-stellige Zahl als eigenes Token
    m = t.match(/\b(\d{3})\b/);
    if (m) return m[1];
    // gepaddet am Wortende (…00157 / 0157)
    m = t.match(/0{0,2}(\d{3})\b/);
    if (m) return m[1];
    return '';
  };

    // nach den Funktionsdefinitionen:
window.pm_findVehicleGridContainer = findVehicleGridContainer;
window.pm_guessDepotSuffixFromVehicleTable = guessDepotSuffixFromVehicleTable;
window.pm_guessDepot = () => guessDepotSuffixFromVehicleTable(findVehicleGridContainer());

  // exakt in diesem Grid nach Head „Depot“ suchen
  const ths = qsaDeep('thead th,[role="columnheader"]', grid);
  const iDepot = ths.findIndex(th => /^Depot$/i.test(norm(th.textContent||th.title||'')));
  if (iDepot < 0) return '';

  // Zeilen innerhalb dieses Grids
  const rows = qsaDeep('tbody tr,[role="row"]', grid)
                 .filter(r => qsaDeep('td,[role="gridcell"]', r).length)
                 .slice(0, 1000);

  for (const tr of rows){
    const tds = qsaDeep('td,[role="gridcell"]', tr);
    const raw = norm(
      tds[iDepot]?.getAttribute?.('aria-label') ||
      tds[iDepot]?.getAttribute?.('data-title') ||
      tds[iDepot]?.querySelector?.('[title]')?.getAttribute('title') ||
      tds[iDepot]?.innerText || tds[iDepot]?.textContent || ''
    );
    if (!raw) continue;
    const suf = pick(raw);
    if (suf && suf !== '000') counts.set(suf, (counts.get(suf)||0)+1);
  }

  if (!counts.size) return '';
  let best='', bestN=-1; for (const [k,n] of counts){ if (n>bestN){ best=k; bestN=n; } }
  return best; // z. B. "157"
}


// ==== ÄNDERUNG: Scanserver-URL baut aus Settings (mit Fallback Auto-Guess) ====
function getScanserverBase(){
  let suf = String(getSetting('depotSuffix')||'').replace(/\D+/g,'').slice(-3);
  if (!suf) {
    // letzter Versuch: NUR aktives Grid
    const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());

    if (g) { suf = g; setSetting('depotSuffix', g); }
  }
  if (!suf) suf = '157'; // Fallback
  return `https://scanserver-d0010${suf}.ssw.dpdit.de/cgi-bin/pa.cgi`;
}

function buildScanserverUrl(psnRaw){
  const pass = getSetting('scanserverPass')||'';
  if (!pass){
    addEvent({title:'Scanserver', meta:'Kein Passwort hinterlegt. Bitte in den Einstellungen setzen.', sev:'warn'});
    openSettingsModal(); // öffnet UI
    return '';
  }
  let psn = String(psnRaw||'').replace(/\D+/g,'');
  if (psn.length === 13) psn = '0' + psn;

  const base = getScanserverBase();
  const params = new URLSearchParams();
  params.set('_url','file');
  params.set('_passwd', pass);           // <-- aus Settings
  params.set('_disp','3');
  params.set('_pivotxx','0');
  params.set('_rastert','4');
  params.set('_rasteryt','0');
  params.set('_rasterx','0');
  params.set('_rastery','0');
  params.set('_pivot','0');
  params.set('_pivotbp','0');
  params.set('_sortby','date|time');
  params.set('_dca','0');
  params.set('_tabledef','psn|date|time|sa|tour|zc|sc|adr1|str|hno|plz1|city|dc|etafrom|etato');
  params.set('_arg59','dpd');
  params.set('_arg0a', psn);
  params.set('_arg0b', psn);
  params.set('_arg0',  psn + ',' + psn);
  params.set('_csv','0');
  return `${base}?${params.toString()}`;
}

function openScanserver(psn){
  const url = buildScanserverUrl(psn);
  if (url) window.open(url, '_blank', 'noopener');
}

// ==== NEU: Einstellungs-UI im Panel ==========================================
function openSettingsModal(){
  const s = loadSettings();
  const html = `
    <div style="display:grid;gap:10px;max-width:520px">
      <label style="display:grid;gap:6px;font:600 12px system-ui">
        Scanserver-Passwort
        <input id="${NS}inp-pass" type="password" placeholder="••••••••" value="${esc(s.scanserverPass||'')}"
               style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px"/>
      </label>
      <label style="display:grid;gap:6px;font:600 12px system-ui">
        Depotkennung (3-stellig, z. B. 157)
        <div style="display:flex;gap:8px;align-items:center">
          <input id="${NS}inp-depot" type="text" pattern="\\d{3}" maxlength="3" placeholder="157" value="${esc(String(s.depotSuffix||'').slice(-3))}"
                 style="padding:8px;border:1px solid rgba(0,0,0,.2);border-radius:8px;width:100px;text-align:center;font-weight:700;letter-spacing:.5px"/>
          <button class="${NS}btn-sm" data-action="guessDepot">Auto erkennen</button>
        </div>
        <div style="opacity:.7;font:12px system-ui">Host wird daraus gebaut: <code>scanserver-d0010<strong>${esc(String(s.depotSuffix||'').slice(-3)||'157')}</strong>.ssw.dpdit.de</code></div>
      </label>
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="${NS}btn-sm" data-action="saveSettings">Speichern</button>
      </div>
    </div>`;
  openModal('Einstellungen', html);
}



  
  
  // ---------- UI ----------
  function mountUI(){
    ensureStyles();
    if (!document.body) return false;
    if (document.getElementById(NS+'wrap') && document.getElementById(NS+'panel')) return true;

    const wrap=document.createElement('div'); wrap.id=NS+'wrap'; wrap.className=NS+'fixed-wrap';
    const btn=document.createElement('button'); btn.id=NS+'btn'; btn.className=NS+'btn'; btn.textContent='Prio/Express Monitoring'; btn.type='button';
    const badge=document.createElement('span'); badge.id=NS+'badge'; badge.className=NS+'badge'; badge.textContent='0';
    wrap.appendChild(btn); wrap.appendChild(badge);
    document.body.appendChild(wrap);

    const panel=document.createElement('div'); panel.id=NS+'panel'; panel.className=NS+'panel';
    panel.style.display='none'; // nur per Button öffnen
    panel.innerHTML=`
 <div class="${NS}header">
  <!-- Toolbar oben -->
  <div class="${NS}toolbar">
    <div class="${NS}group">
      <span class="${NS}label">Kommentare:</span>
      <select id="${NS}filter-comment" class="${NS}select">
        <option value="all">Alle</option>
        <option value="with">nur mit</option>
        <option value="without">nur ohne</option>
      </select>

      <span class="${NS}label">Express:</span>
      <select id="${NS}filter-express" class="${NS}select">
        <option value="all">alle</option>
        <option value="18">nur 18er</option>
        <option value="12">nur 12er</option>
      </select>
    </div>

    <div class="${NS}group">
      <button class="${NS}btn-sm" data-action="openSettings">Einstellungen</button>
      <span class="${NS}chip" id="${NS}auto-chip">
        <span class="${NS}dot" id="${NS}auto-dot"></span>Auto 60s
      </span>
      <button class="${NS}btn-sm" data-action="refreshApi">Aktualisieren (API)</button>
      <button class="${NS}btn-sm" data-action="showExpLate11">Falsch einsortierte EXPRESS12 >11:01</button>
    </div>
  </div>

  <!-- KPIs unten -->
  <div class="${NS}kpis">
    <span class="${NS}kpi" id="${NS}chip-prio-all"  data-action="showPrioAll">
      PRIO in Ausrollung: <b id="${NS}kpi-prio-all">0</b>
    </span>
    <span class="${NS}kpi" id="${NS}chip-prio-open" data-action="showPrioOpen">
      PRIO noch nicht zugestellt: <b id="${NS}kpi-prio-open">0</b>
    </span>
    <span class="${NS}kpi" id="${NS}chip-exp-all"  data-action="showExpAll">
      EXPRESS in Ausrollung: <b id="${NS}kpi-exp-all">0</b>
    </span>
    <span class="${NS}kpi" id="${NS}chip-exp-open" data-action="showExpOpen">
      EXPRESS noch nicht zugestellt: <b id="${NS}kpi-exp-open">0</b>
    </span>
  </div>
</div>


    `;
    document.body.appendChild(panel);

    const modal=document.createElement('div'); modal.id=NS+'modal'; modal.className=NS+'modal';
    modal.innerHTML=`
      <div class="${NS}modal-inner">
        <div class="${NS}modal-head">
          <div id="${NS}modal-title">Liste</div>
          <button class="${NS}btn-sm" data-action="closeModal">Schließen</button>
        </div>
        <div class="${NS}modal-body" id="${NS}modal-body"></div>
      </div>`;
    document.body.appendChild(modal);

    // Toggle nur über Hauptbutton
    btn.addEventListener('click', ()=>{ panel.style.display = (panel.style.display==='none') ? '' : 'none'; });

    panel.addEventListener('click', async (e)=>{
  const k1=e.target.closest('#'+NS+'chip-prio-all');  if(k1){showPrioAll(); return;}
  const k2=e.target.closest('#'+NS+'chip-prio-open'); if(k2){showPrioOpen(); return;}
  const k3=e.target.closest('#'+NS+'chip-exp-all');   if(k3){showExpAll(); return;}
  const k4=e.target.closest('#'+NS+'chip-exp-open');  if(k4){showExpOpen(); return;}

  const b = e.target.closest('.'+NS+'btn-sm');
  if(!b) return;
  const a = b.dataset.action;
  if (a === 'openSettings') { openSettingsModal(); return; }
  if(a==='refreshApi'){ await fullRefresh().catch(console.error); return; }
  if(a==='showExpLate11'){ showExpLate11(); return; }

  // --- NEU: Settings-Aktionen ---
  if(a==='openSettings'){ openSettingsModal(); return; }

  if(a==='guessDepot'){
    const g = guessDepotSuffixFromVehicleTable();
    if(g){
      const inp = document.getElementById(NS+'inp-depot');
      if (inp) inp.value = g;
    }
    return;
  }

  if(a==='saveSettings'){
    const pass = (document.getElementById(NS+'inp-pass')?.value || '');
    const dep  = (document.getElementById(NS+'inp-depot')?.value || '').replace(/\D+/g,'').slice(-3);
    saveSettings({ ...loadSettings(), scanserverPass: pass, depotSuffix: dep });
    addEvent({title:'Einstellungen', meta:'Gespeichert.', sev:'info', read:true});
    hideModal();
    return;
  }
});

const expSel = document.getElementById(NS+'filter-express');
if (expSel) {
  expSel.addEventListener('change', ()=>{
    state.filterExpress = expSel.value || 'all';
    // Falls ein Express-Modal offen ist, neu rendern:
    if (document.getElementById(NS+'modal')?.style.display === 'flex') {
      if (state._modal.title.startsWith('Express')) {
        // Entscheide anhand des Titels, welche Liste offen war
        if (/noch nicht zugestellt/i.test(state._modal.title)) showExpOpen();
        else if (/falsch einsortiert/i.test(state._modal.title)) showExpLate11();
        else showExpAll();
      }
    }
      updateKpisForCurrentState();
  });
}


    modal.addEventListener('click', (e)=>{
  // schließen (Hintergrund oder Button)
  if (e.target.dataset.action === 'closeModal' || e.target === modal) { hideModal(); return; }

  // Auge -> Scanserver (bleibt)
  const eye = e.target.closest('button.'+NS+'eye[data-psn]');
  if (eye){ openScanserver(String(eye.dataset.psn||'')); return; }

  // *** NEU: Buttons im Einstellungs-Dialog ***
  const btn = e.target.closest('button[data-action]');
  if (!btn) return;
  const a = btn.dataset.action;

  if (a === 'guessDepot') {
  const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
  if (g) {
    const inp = document.getElementById(NS+'inp-depot');
    if (inp) inp.value = g;
    addEvent({title:'Einstellungen', meta:`Depotkennung erkannt: ${g}`, sev:'info', read:true});
  } else {
    addEvent({title:'Einstellungen', meta:'Depotkennung konnte im aktiven Fahrzeug-Grid nicht ermittelt werden.', sev:'warn', read:true});
  }
  return;
}


  if (a === 'saveSettings') {
    const pass = (document.getElementById(NS+'inp-pass')?.value || '');
    const dep  = (document.getElementById(NS+'inp-depot')?.value || '').replace(/\D+/g,'').slice(-3);
    saveSettings({ ...loadSettings(), scanserverPass: pass, depotSuffix: dep });
    addEvent({title:'Einstellungen', meta:'Gespeichert.', sev:'info', read:true});
    hideModal();
    return;
  }
});


    const autoDot = document.getElementById(NS+'auto-dot');
    const autoChip = document.getElementById(NS+'auto-chip');
    function setAutoUI(){ autoDot.classList.toggle('off', !autoEnabled); }
    autoChip.addEventListener('click', ()=>{ autoEnabled = !autoEnabled; setAutoUI(); scheduleAuto(); });
    setAutoUI();

    if (!state._bootShown) { addEvent({title:'Bereit', meta:'Status DE aus Tabelle • Sortierbare Überschriften • Zusatzcode • Predict • Zustellzeit • Button EXPRESS12 >11:01 • Auge→Scanserver', sev:'info', read:true}); state._bootShown=true; }
    render();
    return true;
  }

  setInterval(()=>{ try { mountUI(); } catch{} }, 1500);
  new MutationObserver(()=>{ try { mountUI(); } catch{} }).observe(document.documentElement, {childList:true, subtree:true});

  // ---------- helpers ----------
// ---------- helpers ----------

// einmalige Helper
const wait = (ms)=>new Promise(r=>setTimeout(r,ms));
const norm = s => String(s||'').replace(/\s+/g,' ').trim();
async function waitFor(condFn, {timeout=12000, step=250}={}) {
  const t0 = Date.now();
  while (Date.now()-t0 < timeout){
    try{ if(await condFn()) return true; }catch{}
    await wait(step);
  }
  return false;
}

    // Prüfen, ob aktuell "Fahrzeugübersicht"-Tab aktiv ist
function isOnVehicleTab(){
  const tabs = qsaDeep('[role="tab"],li,button,div,span');
  const active = tabs.find(el =>
    (el.getAttribute('aria-selected')==='true' || /react-tabs__tab--selected|is-active|active/.test(el.className||'')) &&
    /fahrzeug(über|ue)sicht/i.test((el.textContent||'').trim())
  );
  return !!active;
}

// Sichtbaren Grid-Container finden (virtuelles Grid)
function findVehicleGridContainer(){
  // Kandidaten: DataGrid/AG-Grid/React-Virtualized sowie der größte scrollbare Bereich
  const cands = qsaDeep(
    '.MuiDataGrid-virtualScroller, .MuiDataGrid-main, .ReactVirtualized__Grid, ' +
    '.ag-center-cols-viewport, .ag-body-viewport, ' +
    '[role="grid"], .rt-table, .rt-tbody, [data-rttable="true"]'
  ).filter(el => el.offsetParent !== null);

  if (!cands.length) return null;

  // nimm das Element, das tatsächlich scrollt (overflow & scrollHeight>clientHeight)
    const scrollable = cands.filter(el => {
    const cs = getComputedStyle(el);
    const ov = cs.overflowY || cs.overflow;
    return (ov && /auto|scroll/i.test(ov)) && (el.scrollHeight > el.clientHeight);
  });
  const pool = scrollable.length ? scrollable : cands;

  // größter Scrollbereich gewinnt
  return pool.reduce((best, el) => ((el.scrollHeight||0) > ((best?.scrollHeight)||0) ? el : best), null);
}

// Tabellenkopf-Indexe für Tour/Zusteller finden (mehrere Varianten)
function detectColIndexes(root){
  const norm = s => String(s||'').replace(/\s+/g,' ').trim();

  // 1) Header *global* suchen (nicht nur unter root)
  const ths = qsaDeep('thead th,[role="columnheader"]').map(el => ({
    el, txt: norm(el.textContent||el.title||'')
  }));

  // Da es mehrere Head-Sets gibt (du hast 3 Duplikate gezeigt), nehmen wir das *letzte vollständige Set*
  const labels = ['Depot','Systempartner','Tour','Zustellername']; // bekannte Reihenfolge
  let startIdx = -1;
  for (let i=ths.length-1; i>=0; i--){
    if (ths[i].txt === labels[0]) { startIdx = i; break; }
  }
  let hdrSlice = startIdx>=0 ? ths.slice(startIdx, startIdx+labels.length+5) : ths;

  // Indizes aus Headertext
  const iTourFromHead = hdrSlice.findIndex(h => /\bTour(\s*nr|nummer)?\b/i.test(h.txt));
  const iDrvFromHead  = hdrSlice.findIndex(h => /^(Zusteller(\s*name)?|Fahrer)$/i.test(h.txt));
  let iTour = iTourFromHead >= 0 ? iTourFromHead : -1;
  let iDrv  = iDrvFromHead  >= 0 ? iDrvFromHead  : -1;

  // 2) Fallback: erste Datenreihe unter *root* nach aria/data/title prüfen
  if (iTour < 0 || iDrv < 0){
    const firstRow = qsaDeep('[role="row"], tbody tr', root).find(r => qsaDeep('[role="gridcell"], td', r).length);
    if (firstRow){
      const cells = qsaDeep('[role="gridcell"], td', firstRow);
      cells.forEach((c,i)=>{
        const label = norm(
          c.getAttribute('aria-label') || c.getAttribute('data-title') ||
          c.querySelector('[title]')?.getAttribute('title') ||
          c.innerText || c.textContent || ''
        );
        if (iTour < 0 && /\bTour(\s*nr|nummer)?\b/i.test(label)) iTour=i;
        if (iDrv  < 0 && /(Zusteller(\s*name)?|Fahrer)/i.test(label))    iDrv =i;
      });
    }
  }
  return { iTour, iDrv };
}


// Aus einer sichtbaren Tabelle Tour→Zusteller mapen
function collectMapFromTable(tbl, map){
  const norm = s => String(s||'').replace(/\s+/g,' ').trim();
  const { iTour, iDrv } = detectColIndexes(tbl);
  if (iTour < 0 || iDrv < 0) return 0;

  const rows = qsaDeep('[role="row"], tbody tr', tbl).filter(r => qsaDeep('[role="gridcell"], td', r).length);
  let added = 0;

  for (const tr of rows){
    const tds = qsaDeep('[role="gridcell"], td', tr);
    const get = (i)=> norm(
      tds[i]?.getAttribute?.('aria-label') || tds[i]?.getAttribute?.('data-title') ||
      tds[i]?.querySelector?.('[title]')?.getAttribute('title') ||
      tds[i]?.innerText || tds[i]?.textContent || ''
    );
    const tour = get(iTour).replace(/[^\dA-Za-z]/g,'');
    const drv  = get(iDrv);
    if (tour && drv && !map.has(tour)){ map.set(tour, drv); added++; }
  }
  return added;
}

// Durch das Grid scrollen und unterwegs sammeln (für Virtualization)
async function scrollGridAndCollect(map, {timeoutMs=18000, step=250}={}){
  let grid = findVehicleGridContainer();
  if (!grid) return 0;

  const start = Date.now();
  let totalAdded = 0, prev = -1;

  while (Date.now()-start < timeoutMs){
    // Tabellen/Gitter *unter dem aktuellen Viewport-Container*
    const scopes = [grid, grid.parentElement, document];
    for (const scope of scopes){
      const tablesHere = qsaDeep('table,[role="grid"]', scope).filter(el => el.offsetParent !== null);
      for (const tbl of tablesHere) totalAdded += collectMapFromTable(tbl, map);
    }

    // scroll vorwärts
    const maxY = grid.scrollHeight - grid.clientHeight;
    const next = Math.min((grid.scrollTop||0) + Math.max(80, grid.clientHeight*0.9), maxY);
    grid.scrollTop = next;

    if (next >= maxY || next === prev) break;
    prev = next;
    await new Promise(r=>setTimeout(r, step));
  }
  return totalAdded;
}


    // Startet Autoload, sobald Fahrzeugübersicht aktiv ist / wird
function startAutoloadTourDriverMap(){
  // Sofort probieren (die Seite startet in Fahrzeugübersicht)
  buildTourDriverMapAutoload();

  // Danach mit Observer nachladen, falls Grid später rendert
  const mo = new MutationObserver(()=> {
    if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size) return;
    buildTourDriverMapAutoload();
  });
  mo.observe(document.documentElement, {subtree:true, childList:true});
}

async function buildTourDriverMapAutoload(){
  try{
    if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size) return;

    const map = new Map();

    // 1) Alles, was gerade sichtbar ist, einsammeln
    qsaDeep('table,[role="grid"]').filter(el => el.offsetParent !== null)
      .forEach(tbl => collectMapFromTable(tbl, map));

    // 2) Grid komplett durchscrollen (virtuelle Zeilen laden)
    await scrollGridAndCollect(map);

    if (map.size){
      window.__pmTour2Driver = map;
      console.info('[PM] AutoLoad Tour→Zusteller: ', map.size);
    }
  } catch(e){
    console.warn('[PM] AutoLoad Fehler:', e);
  }
}

    setTimeout(()=>{
  const grid = findVehicleGridContainer();
  console.log('[PM] Grid:', grid);
  console.log('[PM] Sichtbare tables/grids:', qsaDeep('table,[role="grid"]').filter(el=>el.offsetParent!==null).length);
}, 2000);

    setTimeout(()=> {
  if (!(getSetting('depotSuffix')||'')) {
    const g = guessDepotSuffixFromVehicleTable(findVehicleGridContainer());
    if (g) { setSetting('depotSuffix', g); addEvent({title:'Einstellungen', meta:`Depotkennung aus Fahrzeugübersicht: ${g}`, sev:'info', read:true}); }
  }
}, 1200);


// ==== Shadow-DOM / iframe tolerant ====
function* nodesDeep(root=document){
  const stack=[root];
  while(stack.length){
    const n=stack.pop();
    yield n;
    const kids = n instanceof ShadowRoot ? n.children : (n.querySelectorAll ? n.children : []);
    for (let i=kids.length-1;i>=0;i--) stack.push(kids[i]);
    if (n.shadowRoot) stack.push(n.shadowRoot);
    if (n.tagName==='IFRAME' && n.contentDocument) stack.push(n.contentDocument);
  }
}
function qsaDeep(selector, root=document){
  const out=[];
  for (const n of nodesDeep(root)) if (n.querySelectorAll) out.push(...n.querySelectorAll(selector));
  return out;
}
function findByTextDeep(selector, re, root=document){
  const els = qsaDeep(selector, root);
  return els.find(el => re.test(norm(el.textContent||el.title||'')));
}
function clickDeep(el){
  if (!el) return;
  ['pointerdown','mousedown','mouseup','click'].forEach(t =>
    el.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window}))
  );
}
function getActiveTabEl(){
  const sel = [
    '[role="tab"][aria-selected="true"]',
    '.active,.is-active,.tab--active',
    '[data-active="true"]'
  ].join(',');
  return qsaDeep(sel)[0] || null;
}

// ---- Tour→Zusteller aus sichtbarer Tabelle bauen (deep/shadow-sicher)
async function tryBuildTourDriverMapFromDom(){
  if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size) return;
  const norm = s => String(s||'').replace(/\s+/g,' ').trim();

  // alle sichtbaren Tabellen/Gitter holen (Shadow DOM/iframes werden von qsaDeep abgedeckt)
  const visibleTables = qsaDeep('table,[role="grid"]').filter(el => el.offsetParent !== null);

  for (const tbl of visibleTables) {
    const ths = qsaDeep('thead th,[role="columnheader"]', tbl);
    const idx = (re)=> ths.findIndex(th => re.test(norm(th.textContent||th.title||'')));

    // Spaltenpositionen bestimmen (Tour / Zustellername)
    let iTour = idx(/\btour(\s*nr|nummer)?\b/i);
    let iDrv  = idx(/^(zusteller(\s*name)?|fahrer)$/i);

    // Fallback: manche Grids haben keine echten <th>, sondern aria-Label/Data-Title in den Zellen
    if (iTour < 0 || iDrv < 0) {
      const firstRow = qsaDeep('tbody tr,[role="row"]', tbl)
        .find(r => qsaDeep('td,[role="gridcell"]', r).length);
      if (firstRow) {
        const cells = qsaDeep('td,[role="gridcell"]', firstRow);
        cells.forEach((c, i) => {
          const label = norm(
            c.getAttribute('aria-label') || c.getAttribute('data-title') ||
            c.querySelector('[title]')?.getAttribute('title') || c.textContent || ''
          );
          if (iTour < 0 && /\btour(\s*nr|nummer)?\b/i.test(label)) iTour = i;
          if (iDrv  < 0 && /(zusteller(\s*name)?|fahrer)/i.test(label)) iDrv  = i;
        });
      }
    }

    if (iTour < 0 || iDrv < 0) continue;

    const map = new Map();
    const rows = qsaDeep('tbody tr,[role="row"]', tbl)
      .filter(r => qsaDeep('td,[role="gridcell"]', r).length);

    for (const tr of rows){
      const tds  = qsaDeep('td,[role="gridcell"]', tr);
      const get  = (i)=> norm(
        tds[i]?.getAttribute?.('aria-label') || tds[i]?.getAttribute?.('data-title') ||
        tds[i]?.querySelector?.('[title]')?.getAttribute('title') || tds[i]?.textContent || ''
      );
      const tour = get(iTour).replace(/[^\dA-Za-z]/g,'');
      const drv  = get(iDrv);
      if (tour && drv) map.set(tour, drv);
    }

    if (map.size){
      window.__pmTour2Driver = map;
      console.info('[PM] Map gebaut aus Tabelle:', map.size);
      return;
    }
  }
}

// ==== Auto-Navigation: Tab öffnen → Aktualisieren → mehrfach Map bauen → zurück ====
async function ensureTourDriverMapByAutoNav(){
  if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size) return true;

  console.info('[PM] AutoNav start …');
  const prevTab = getActiveTabEl();

  // 1) Tab "Fahrzeugübersicht" finden (Umlaut/ue)
  const tab =
    findByTextDeep('[role="tab"],li,a,button,div,span', /fahrzeug(über|ue)sicht/i) ||
    qsaDeep('[data-rttab="true"],[aria-controls],[id]').find(el =>
      /fahrzeug(über|ue)sicht/i.test((el.textContent||el.title||'').trim())
    );
  if (!tab){ console.warn('[PM] Tab "Fahrzeugübersicht" nicht gefunden'); return false; }
  clickDeep(tab);
  await wait(400);

  // 2) „Aktualisieren“ anstoßen (best effort)
  const refreshBtn =
    findByTextDeep('button,a', /aktualisier|refresh|reload/i) ||
    qsaDeep('button[title*="Aktualis" i],a[title*="Aktualis" i],button[aria-label*="Aktualis" i],a[aria-label*="Aktualis" i],button[class*="refresh" i],a[class*="reload" i]')[0];
  if (refreshBtn){ clickDeep(refreshBtn); console.info('[PM] Refresh geklickt.'); }

  // 3) Statt „sichtbar?“-Heuristik: wir POLLEN einfach 20x (≈6s) und versuchen jedes Mal die Map zu bauen
  for (let i=0;i<20;i++){
    await tryBuildTourDriverMapFromDom();
    if (window.__pmTour2Driver instanceof Map && window.__pmTour2Driver.size){
      console.info('[PM] Tour→Zusteller Map size:', window.__pmTour2Driver.size);
      if (prevTab && prevTab !== getActiveTabEl()) { clickDeep(prevTab); await wait(200); }
      return true;
    }
    await wait(300);
  }

  console.warn('[PM] Konnte Map nicht bauen (Timeout).');
  if (prevTab && prevTab !== getActiveTabEl()) { clickDeep(prevTab); await wait(200); }
  return false;
}

// ---- Fahrer-Auflösung nutzt die Map (driverOf im restlichen Code ruft das eh) ----


  function setLoading(on){ isLoading=!!on; const el=document.getElementById(NS+'loading'); if(el) el.classList.toggle('on',on); }
  function dimButtons(on){ document.querySelectorAll('.'+NS+'btn-sm').forEach(b=>b.classList.toggle(NS+'dim', !!on)); }
  function setBadge(n){ const b=document.getElementById(NS+'badge'); if(!b) return; const v=Math.max(0,Number(n||0)); b.textContent=String(v); b.style.display=v>0?'inline-block':'none'; }

  function setKpis({prioAll, prioOpen, expAll, expOpen}){
    const set = (id,val,chipId)=>{ const el=document.getElementById(id); const chip=document.getElementById(chipId); if(el) el.textContent=String(Number(val||0)); if(chip) chip.style.display=Number(val||0)===0?'none':''; };
    set(NS+'kpi-prio-all',  prioAll, NS+'chip-prio-all');
    set(NS+'kpi-prio-open', prioOpen, NS+'chip-prio-open');
    set(NS+'kpi-exp-all',   expAll,  NS+'chip-exp-all');
    set(NS+'kpi-exp-open',  expOpen, NS+'chip-exp-open');
  }

  function openModal(title, html, rows=null, opts=null){
    const m=document.getElementById(NS+'modal');
    const t=document.getElementById(NS+'modal-title');
    const b=document.getElementById(NS+'modal-body');
    if (t) t.textContent = title || '';
    if (rows) { state._modal.rows = rows.slice(); state._modal.opts = Object.assign({}, opts||{}); state._modal.title = title||''; }
    if (b) { b.innerHTML = html || ''; attachSortHandlers(); }
    if (m) m.style.display='flex';
  }
  function hideModal(){ const m=document.getElementById(NS+'modal'); if(m) m.style.display='none'; }

  // ---------- grid helpers ----------
  const isDelivery = r => String(r?.orderType || '').toUpperCase() === 'DELIVERY';
  const isPRIO = r => String(r?.priority||r?.prio||'').toUpperCase() === 'PRIO';
  const hasExpress12 = r => {
    const el = r?.elements;
    if (Array.isArray(el)) return el.map(String).includes('023');
    return typeof el === 'string' ? /\b023\b/.test(el) : false;
  };
  const hasExpress18 = r => {
  const el = r?.elements;
  if (Array.isArray(el)) return el.map(String).includes('010'); // 010 = EXPRESS 18 (bei euch)
  return typeof el === 'string' ? /\b010\b/.test(el) : false;
};
const hasExpressAny = r => hasExpress12(r) || hasExpress18(r);
const expressTypeOf = r => hasExpress12(r) ? '12' : (hasExpress18(r) ? '18' : '');

  const addCodes   = r => Array.isArray(r.additionalCodes)? r.additionalCodes.map(String): [];

  // ---- STATUS (DOM bevorzugt, Fallback Mapping)
  const apiStatus = r => (r.statusName || r.statusText || r.stateText || r.status || r.deliveryStatus || r.parcelStatus || '').toString().trim();
  let statusColIdx = null;
  function findStatusColumnIndex(){
    if (statusColIdx != null) return statusColIdx;
    const headers = Array.from(document.querySelectorAll('thead th, [role="columnheader"]'));
    let idx = -1;
    headers.forEach((h,i)=>{
      const t=(h.textContent||'').trim().toLowerCase();
      if (idx===-1 && /(status|stat)/.test(t)) idx=i;
    });
    statusColIdx = idx >= 0 ? idx : null;
    return statusColIdx;
  }
  function statusFromDomByParcel(parcel){
    if (!parcel) return '';
    if (statusCache.has(parcel)) return statusCache.get(parcel)||'';
    const idx = findStatusColumnIndex();
    const rows = Array.from(document.querySelectorAll('tbody tr, [role="row"]')).filter(tr =>
      (tr.textContent || '').includes(parcel)
    );
    for (const tr of rows){
      let val = '';
      if (idx!=null){
        const cells = Array.from(tr.querySelectorAll('td, [role="gridcell"]'));
        const cell = cells[idx];
        if (cell){
          const div = cell.querySelector('div[title]');
          val = (div?.getAttribute('title') || cell.textContent || '').trim();
        }
      }
      if (!val){
        const guess = Array.from(tr.querySelectorAll('*')).find(el=>{
          const s=(el.textContent||'').trim().toUpperCase();
          return /(ZUGESTELLT|ZUSTELLUNG|PROBLEM)/.test(s);
        });
        if (guess) val = (guess.textContent||'').trim();
      }
      if (val){ statusCache.set(parcel, val); return val; }
    }
    return '';
  }
  function mapEnToDeStatus(s){
    const u = String(s||'').toUpperCase();
    if (!u) return '';
    if (/DELIVERED|DELIVERED_TO_PUDO/.test(u)) return 'ZUGESTELLT';
    if (/OUT_FOR_DELIVERY|IN_DELIVERY|ON.?ROUTE|VEHICLE/.test(u)) return 'ZUSTELLUNG';
    if (/DELIVERY.*PROBLEM|PROBLEM|FAIL|NOT.?DELIVERED/.test(u)) return 'ZUSTELLUNG PROBLEM';
    if (/AT.?DEPOT|AT.?HUB/.test(u)) return 'IM DEPOT';
    if (/RETURN/.test(u)) return 'RETURNSENDUNG';
    if (/DELIVERY/.test(u)) return 'ZUSTELLUNG';
    return '';
  }
  const statusOf = r => {
    const p = parcelId(r);
    const fromDom = statusFromDomByParcel(p);
    if (fromDom) return fromDom;
    const mapped = mapEnToDeStatus(apiStatus(r));
    return mapped || '—';
  };

  const delivered = r => {
    const s = statusOf(r).toUpperCase();
    if (r.deliveredTime) return true;
    return /(ZUGESTELLT)/.test(s);
  };

  // ---- NEU: globale Map aus anderer Ansicht (falls schon per Console befüllt)
const tour2Driver = (() => window.__pmTour2Driver instanceof Map ? window.__pmTour2Driver : new Map());

// (deins) Tour ermitteln
const tourOf  = r => r.tour ? String(r.tour) : '';

// ---- ERSATZ: Fahrerauflösung -> erst API-Felder, dann Map per Tour
const driverOf = r => {
  const direct = r.driverName || r.driver || r.courierName || r.riderName || r.tourDriver || '';
  if (direct && direct.trim()) return direct.trim();
  const key = String(tourOf(r) || '').replace(/[^\dA-Za-z]/g,'');
  return (tour2Driver().get(key) || '').trim();
};


  const addrOf  = r => [r.street, r.houseno].filter(Boolean).join(' ');
  const placeOf = r => [r.postalCode, r.city].filter(Boolean).join(' ');

  const parcelId   = r => r.parcelNumber || (Array.isArray(r.parcelNumbers)&&r.parcelNumbers[0]) || r.id || '';

  const composeDateTime = (dateStr, timeStr) => { if (!dateStr || !timeStr) return null; const s = `${dateStr}T${String(timeStr).slice(0,8)}`; const d = new Date(s); return isNaN(d) ? null : d; };
  const fromTime = r => r.from2 ? composeDateTime(r.date, r.from2) : null;
  const toTime   = r => r.to2   ? composeDateTime(r.date, r.to2)   : null;
  const deliveredTime = r => r.deliveredTime ? new Date(r.deliveredTime) : null;

  // EXPRESS12 „falsch einsortiert >11:01“
  const isExpressLateAfter11 = r => { const ft = fromTime(r); if (!ft) return false; return (ft.getHours() > 11) || (ft.getHours()===11 && ft.getMinutes()>=1); };

  // Kommentare
  const pickComment = (r) => {
    const fields = [ r?.freeComment, r?.freeTextDc, r?.note, r?.comment ];
    for (const f of fields) {
      if (Array.isArray(f) && f.length) return f.filter(Boolean).join(' | ');
      if (typeof f === 'string' && f.trim()) return f.trim();
    }
    return '';
  };
  function buildUrlByParcel(base, parcel){
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page','1'); q.set('pageSize','500'); q.set('parcelNumber', String(parcel));
    q.delete('priority'); q.delete('elements');
    u.search = q.toString();
    return u;
  }
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
  const commentOf = (r) => { const fromApi = pickComment(r); if (fromApi) return fromApi; return commentFromDomByParcel(parcelId(r)); };

  // ---------- network hook ----------
  (function hook(){
    if (!window.__pm_fetch_hooked && window.fetch) {
      const orig = window.fetch;
      window.fetch = async function(i, init = {}) {
        const res = await orig(i, init);
        try {
          const uStr = typeof i === 'string' ? i : (i && i.url) || '';
          if (uStr.includes('/dispatcher/api/pickup-delivery') && res.ok) {
            const u = new URL(uStr, location.origin);
            const q = u.searchParams;
            if (!q.get('parcelNumber')) {
              const h = {};
              const src = (init && init.headers) || (i && i.headers);
              if (src) {
                if (src.forEach) src.forEach((v, k) => h[String(k).toLowerCase()] = String(v));
                else if (Array.isArray(src)) src.forEach(([k, v]) => h[String(k).toLowerCase()] = String(v));
                else Object.entries(src).forEach(([k, v]) => h[String(k).toLowerCase()] = String(v));
              }
              if (!h['authorization']) {
                const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                if (m) h['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
              }
              lastOkRequest = { url: u, headers: h };
                const n = document.getElementById(NS + 'note-capture'); if (n) n.style.display = 'none';

                // >>> NEU: Auto-Refresh aktivieren, sobald wir einen gültigen Request haben
                try { scheduleAuto(); } catch {}
                // Optional: initial gleich laden, statt 60s zu warten
                try { if (autoEnabled && !isBusy && !document.hidden) { fullRefresh().catch(()=>{}); } } catch {}

            }
          }
        } catch {}
        return res;
      };
      window.__pm_fetch_hooked = true;
    }

    if (!window.__pm_xhr_hooked && window.XMLHttpRequest) {
      const X = window.XMLHttpRequest;
      const open = X.prototype.open, send = X.prototype.send, setH = X.prototype.setRequestHeader;

      X.prototype.open = function(m, u) {
        this.__pm_url = (typeof u === 'string') ? new URL(u, location.origin) : null;
        this.__pm_headers = {};
        return open.apply(this, arguments);
      };
      X.prototype.setRequestHeader = function(k, v) {
        try { this.__pm_headers[String(k).toLowerCase()] = String(v); } catch {}
        return setH.apply(this, arguments);
      };
      X.prototype.send = function() {
        const onload = () => {
          try {
            if (this.__pm_url && this.__pm_url.href.includes('/dispatcher/api/pickup-delivery') &&
                this.status >= 200 && this.status < 300) {
              const q = this.__pm_url.searchParams;
              if (!q.get('parcelNumber')) {
                if (!this.__pm_headers['authorization']) {
                  const m = document.cookie.match(/(?:^|;\s*)dpd-register-jwt=([^;]+)/);
                  if (m) this.__pm_headers['authorization'] = 'Bearer ' + decodeURIComponent(m[1]);
                }
                lastOkRequest = { url: this.__pm_url, headers: this.__pm_headers };
                  const n = document.getElementById(NS + 'note-capture'); if (n) n.style.display = 'none';

                  // >>> NEU
                  try { scheduleAuto(); } catch {}
                  try { if (autoEnabled && !isBusy && !document.hidden) { fullRefresh().catch(()=>{}); } } catch {}

              }
            }
          } catch {}
          this.removeEventListener('load', onload);
        };
        this.addEventListener('load', onload);
        return send.apply(this, arguments);
      };
      window.__pm_xhr_hooked = true;
    }
  })();

  // ---------- API helpers ----------
  function buildUrlPrio(base, page){
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize', '500');
    q.set('priority', 'prio');
    q.delete('elements'); q.delete('parcelNumber');
    u.search = q.toString();
    return u;
  }
  function buildUrlElements(base, page, el){
    const u = new URL(base.href);
    const q = u.searchParams;
    q.set('page', String(page));
    q.set('pageSize', '500');
    q.set('elements', String(el));
    q.delete('priority'); q.delete('parcelNumber');
    u.search = q.toString();
    return u;
  }
  function buildHeaders(h){ const H=new Headers(); try{
    if(h){ Object.entries(h).forEach(([k,v])=>{const key=k.toLowerCase(); if(['authorization','accept','x-xsrf-token','x-csrf-token'].includes(key)){ H.set(key==='accept'?'Accept':key.replace(/(^.|-.)/g,s=>s.toUpperCase()), v);} }); }
    if(!H.has('Accept')) H.set('Accept','application/json, text/plain, */*');
  }catch{} return H; }
  async function fetchPaged(builder){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size = 500, maxPages = 50;
    let page = 1, rows = [];
    while (page <= maxPages) {
      const u = builder(lastOkRequest.url, page);
      const r = await fetch(u.toString(), { credentials:'include', headers });
      if (!r.ok) break;
      const j = await r.json();
      const chunk = (j.items||j.content||j.data||j.results||j)||[];
      rows.push(...chunk);
      if (chunk.length < size) break;
      page++;
      await sleep(40);
    }
    return rows;
  }
  async function fetchPagedWithTotal(builder){
    if(!lastOkRequest) throw new Error('Kein 200-OK Request zum Klonen gefunden.');
    const headers = buildHeaders(lastOkRequest.headers);
    const size = 500, maxPages = 50;
    let page = 1, rows = [], totalKnown = null;
    while (page <= maxPages) {
      const u = builder(lastOkRequest.url, page);
      const r = await fetch(u.toString(), { credentials:'include', headers });
      if (!r.ok) { if (page === 1) throw new Error(`HTTP ${r.status}`); break; }
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

  // ---------- table rendering + sorting ----------
  function formatHHMM(d){ return d ? d.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}) : ''; }
  function buildTable(rows, {showPredict=false}={}){
    const ths = [
      'Paketscheinnummer','Adresse','Fahrer','Tour','Status','Zustellzeit','Zusatzcode'
    ];
    if (showPredict) ths.push('Predict');
    const head = `<tr>${ths.map((h,i)=>`<th data-col="${i}">${h}</th>`).join('')}</tr>`;
    const body = rows.map(r=>{
      const p = parcelId(r);
      const eye = p ? `<button class="${NS}eye" title="Scanserver öffnen" data-psn="${esc(p)}">👁</button>` : '';
      const link = p ? `<a class="${NS}plink" href="https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${p}" target="_blank" rel="noopener">${p}</a>` : '—';
      const addr = [addrOf(r), placeOf(r)].filter(Boolean).join(' · ') || '—';
      const driver = driverOf(r) || '—';
      const tour = tourOf(r) || '—';
      const stat = statusOf(r);
      const dtime = deliveredTime(r) ? formatHHMM(deliveredTime(r)) : '—';
      const codes = addCodes(r).join(', ') || '—';
      const pred = (showPredict && !delivered(r)) ? (()=>{ const s=fromTime(r), e=toTime(r); return (s||e)?`${formatHHMM(s)}–${formatHHMM(e)}`:'—'; })() : (showPredict?'—':'');
      const tds = [`${eye}${link}`, esc(addr), esc(driver), esc(tour), esc(stat), esc(dtime), esc(codes)];
      if (showPredict) tds.push(esc(pred));
      return `<tr>${tds.map(v=>`<td>${v}</td>`).join('')}</tr>`;
    }).join('') || `<tr><td colspan="${showPredict?8:7}">Keine Einträge.</td></tr>`;
    return `<table class="${NS}tbl"><thead>${head}</thead><tbody>${body}</tbody></table>`;
  }

  function attachSortHandlers(){
    const body = document.getElementById(NS+'modal-body');
    const table = body && body.querySelector('table');
    if (!table) return;
    const ths = Array.from(table.querySelectorAll('th'));
    ths.forEach(th=>{
      th.addEventListener('click', ()=>{
        const col = Number(th.dataset.col||0);
        ths.forEach(x=>x.classList.remove(NS+'sort-asc',NS+'sort-desc'));
        const ascending = !(th.dataset.dir==='asc');
        th.dataset.dir = ascending ? 'asc' : 'desc';
        th.classList.add(ascending?NS+'sort-asc':NS+'sort-desc');
        const rows = state._modal.rows.slice();
        const opts = state._modal.opts||{};
        rows.sort((a,b)=>{
          const get = (r)=>{
            switch(col){
              case 0: return String(parcelId(r)||'');
              case 1: return [addrOf(r),placeOf(r)].filter(Boolean).join(' · ');
              case 2: return String(driverOf(r)||'');
              case 3: return Number(tourOf(r)||0);
              case 4: return String(statusOf(r)||'');
              case 5: return deliveredTime(r) ? deliveredTime(r).getTime() : 0;
              case 6: return (addCodes(r)||[]).join(', ');
              case 7: { const s=fromTime(r), e=toTime(r); return s?s.getTime():(e?e.getTime():0); }
            }
          };
          const va=get(a), vb=get(b);
          if (typeof va==='number' && typeof vb==='number') return ascending ? (va-vb) : (vb-va);
          return ascending ? String(va).localeCompare(String(vb),'de') : String(vb).localeCompare(String(va),'de');
        });
        state._modal.rows = rows;
        const html = buildTable(rows, opts);
        openModal(state._modal.title, html); // re-render
      });
    });
  }
  function filterByExpressSelection(rows){
  if (state.filterExpress === '12') return rows.filter(hasExpress12);
  if (state.filterExpress === '18') return rows.filter(hasExpress18);
  return rows;
}
  function getFilteredExpressCounts(){
  const f = state.filterExpress;
  const filt = f==='12' ? hasExpress12 : f==='18' ? hasExpress18 : null;

  const expAllList  = filt ? state._expAllList.filter(filt)  : state._expAllList;
  const expOpenList = filt ? state._expOpenList.filter(filt) : state._expOpenList;

  return { expAllCount: expAllList.length, expOpenCount: expOpenList.length };
}

function updateKpisForCurrentState(){
  const { expAllCount, expOpenCount } = getFilteredExpressCounts();
  setKpis({
    prioAll:  state._prioAllList.length,
    prioOpen: state._prioOpenList.length,
    expAll:   expAllCount,
    expOpen:  expOpenCount
  });
}

  function showPrioAll(){  const rows=state._prioAllList;  const html=buildTable(rows); openModal(`PRIO – in Ausrollung (alle) · ${rows.length}`, html, rows, {showPredict:false}); }
  function showPrioOpen(){ const rows=state._prioOpenList; const html=buildTable(rows,{showPredict:true}); openModal(`PRIO – noch nicht zugestellt · ${rows.length}`, html, rows, {showPredict:true}); }
  function showExpAll(){
  const src = state._expAllList;
  const rows = filterByExpressSelection(src);
  const sel = state.filterExpress==='12'?' (12)':
              state.filterExpress==='18'?' (18)':'';
  const html=buildTable(rows);
  openModal(`Express${sel} – in Ausrollung (alle) · ${rows.length}`, html, rows, {showPredict:false});
}

function showExpOpen(){
  const src = state._expOpenList;
  const rows = filterByExpressSelection(src);
  const sel = state.filterExpress==='12'?' (12)':
              state.filterExpress==='18'?' (18)':'';
  const html=buildTable(rows,{showPredict:true});
  openModal(`Express${sel} – noch nicht zugestellt · ${rows.length}`, html, rows, {showPredict:true});
}

function showExpLate11(){
  // bleibt nur für 12er
  const rows = state._expLate11List.slice();
  const html=buildTable(rows,{showPredict:true});
  openModal(`Express 12 – falsch einsortiert (>11:01 geplant) · ${rows.length}`, html, rows, {showPredict:true});
}


  // ---------- render ----------
  function render(){
    const list = document.getElementById(NS+'list'); if (!list) return; list.innerHTML='';
    const d=document.createElement('div'); d.className=NS+'empty'; d.textContent= isLoading ? 'Lade Daten …' : 'Aktualisiert.';
    list.appendChild(d);
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
      parcel: ev.parcel || '',
      kind: ev.kind || ''
    };
    state.events.push(e); render();
  }

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

  // ---------- refresh ----------
  // ---------- refresh ----------
async function refreshViaApi(){
  const { rows: prioRows } = await fetchPagedWithTotal(buildUrlPrio);
  const prioDeliveries = prioRows.filter(isDelivery).filter(isPRIO);
  const prioAll  = prioDeliveries;
  const prioOpen = prioDeliveries.filter(r=>!delivered(r));

  // Express 12 (023) + Express 18 (010) laden
  const exp12Rows = await fetchPaged((base,p)=>buildUrlElements(base,p,'023')); // Express 12
  const exp18Rows = await fetchPaged((base,p)=>buildUrlElements(base,p,'010')); // Express 18 (010)
  const expRows   = [...exp12Rows, ...exp18Rows];

  // DE-Duplizieren MUSS vor der Nutzung deklariert werden
  const seen = new Set();
  const expDeliveries = expRows
    .filter(isDelivery)
    .filter(hasExpressAny)
    .filter(r=>{
      const id = parcelId(r);
      if (!id || seen.has(id)) return false;
      seen.add(id);
      return true;
    });

  const expAll    = expDeliveries;
  const expOpen   = expDeliveries.filter(r=>!delivered(r));
  // >11:01 bleibt fachlich nur für 12er
  const expLate11 = expDeliveries.filter(hasExpress12).filter(isExpressLateAfter11);

  state._prioAllList   = prioAll.slice();
  state._prioOpenList  = prioOpen.slice();
  state._expAllList    = expAll.slice();
  state._expOpenList   = expOpen.slice();
  state._expLate11List = expLate11.slice();
  const { expAllCount, expOpenCount } = getFilteredExpressCounts();
  setKpis({
  prioAll:  prioAll.length,
  prioOpen: prioOpen.length,
  expAll:   expAllCount,
  expOpen:  expOpenCount
});

  state.events = [{
    id: ++state.nextId,
    title:'Aktualisiert',
    meta:`PRIO: in Ausrollung ${prioAll.length} • offen ${prioOpen.length} • EXPRESS: in Ausrollung ${expAll.length} • offen ${expOpen.length} • „>11:01“-Liste (12er): ${expLate11.length}`,
    sev:'info', read:true, ts:Date.now()
  }];
}

  async function fullRefresh(){
    if (isBusy) return;
    try{
      if (!lastOkRequest) { addEvent({title:'Hinweis', meta:'Kein API-Request erkannt. Bitte einmal die normale Suche ausführen, danach funktioniert Auto 60s.', sev:'info', read:true}); render(); return; }
      isBusy = true; setLoading(true); dimButtons(true);
      addEvent({title:'Aktualisiere (API)…', meta:'Replay aktiver Filter (pageSize=500)', sev:'info', read:true}); render();

      await refreshViaApi();
       await tryBuildTourDriverMapFromDom();   // <-- NEU

      await fetchMissingComments();
        document.addEventListener('visibilitychange', ()=>{
  scheduleAuto();
  if (!document.hidden && autoEnabled && lastOkRequest) { fullRefresh().catch(()=>{}); }
});

    } catch(e){
      console.error(e);
      addEvent({title:'Fehler (API)', meta:String(e && e.message || e), sev:'warn'});
    } finally{
      setLoading(false); dimButtons(false); isBusy = false; render();
    }
  }

  // ---------- auto ----------


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

  // ---------- boot ----------
  (function boot(){
  mountUI();

  // NEU: Autoload direkt aus der sichtbaren Fahrzeugübersicht
  startAutoloadTourDriverMap();

  // API-Auto-Refresh wie gehabt
  scheduleAuto();
})();


})();
