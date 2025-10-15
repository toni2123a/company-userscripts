// ==UserScript==
// @name         Eticor Prüfprotokoll (dynamische Boxen + sofortige Signatur)
// @namespace    http://tampermonkey.net/
// @version      1.5
// @description  Erzeugt strukturiertes Prüfprotokoll (PDF) mit dynamischen Boxen und zuverlässiger Signatur. Button unter "Aufgabe nicht eingehalten" auf Detailseiten.
// @match        https://www.eticor-portal.com/*
// @require      https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js
// @grant        GM_addStyle
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';
  const LOG = (...a)=>console.log('[TM-ETICOR]', ...a);

  GM_addStyle(`
    .tm-embed-wrap { margin-top: 10px; }
    .tm-embed-btn {
      display:inline-block; padding:8px 12px; background:#167FFC; color:#fff !important;
      border-radius:8px; font-weight:700; border:0; cursor:pointer;
      box-shadow:0 2px 8px rgba(0,0,0,.15);
    }
    #tm-pruef-modal { position: fixed; left: 50%; top: 50%; transform: translate(-50%,-50%);
      z-index:2147483600; background:#fff; border-radius:8px; box-shadow:0 10px 40px rgba(0,0,0,.35);
      padding:16px; width:760px; max-width:95%; display:none; }
    #tm-pruef-overlay { position: fixed; inset:0; background:rgba(0,0,0,.35); z-index:2147483599; display:none; }
    #tm-sign-canvas { border:1px solid #ccc; width:100%; height:160px; touch-action:none; background:#fff; }
    #tm-actions { display:flex; gap:8px; justify-content:flex-end; margin-top:10px; }
  `);

  /* ---------- Routing ---------- */
  const onUrlChange = (() => {
    let last = location.href;
    return () => { if (location.href !== last) { last = location.href; tryInsert(); } };
  })();
  const ps = history.pushState;
  history.pushState = function(){ ps.apply(this, arguments); onUrlChange(); };
  addEventListener('popstate', onUrlChange);
  addEventListener('hashchange', onUrlChange);

  const isDetail = () => /^https:\/\/www\.eticor-portal\.com\/delegations\/\d+(?:[/?#]|$)/.test(location.href);

  function tryInsert(){
    if (!isDetail()) return;
    const t = setInterval(() => {
      const target = findRightPanel();
      if (target) { clearInterval(t); embedButton(target); }
    }, 300);
    setTimeout(()=>clearInterval(t), 8000);
  }

  function findRightPanel(){
    const labels = Array.from(document.querySelectorAll('aside label, [role="radiogroup"] label, .MuiFormGroup-root label, main label'));
    const neg = labels.find(l => /Aufgabe\s+nicht\s+eingehalten/i.test(l.textContent || ''));
    if (neg) return neg.closest('.MuiFormGroup-root') || neg.parentElement || neg;
    const box = Array.from(document.querySelectorAll('aside, [class*="Right"], [class*="Sidebar"]'))
      .find(n => /Neuer\s+Prüfeintrag/i.test(n.innerText||''));
    return box || null;
  }

  function embedButton(container){
    if (document.getElementById('tm-embed-btn')) return;
    const wrap = document.createElement('div');
    wrap.className = 'tm-embed-wrap';
    const btn = document.createElement('button');
    btn.id = 'tm-embed-btn';
    btn.className = 'tm-embed-btn';
    btn.type = 'button';
    btn.textContent = 'Prüfprotokoll → PDF';
    btn.title = 'Erzeuge Prüfprotokoll (PDF) aus dieser Delegation';
    btn.addEventListener('click', openUi);
    wrap.appendChild(btn);
    if (container.nextSibling) container.parentNode.insertBefore(wrap, container.nextSibling);
    else container.parentNode.appendChild(wrap);
    LOG('Button eingebettet');
  }

  /* ---------- UI + Signatur ---------- */
  let resizeCanvasFn = null; // damit wir nach dem Öffnen resize sicher triggern

  function ensureModal() {
    if (document.getElementById('tm-pruef-modal')) return;
    const overlay = document.createElement('div'); overlay.id='tm-pruef-overlay';
    const modal = document.createElement('div'); modal.id='tm-pruef-modal';
    modal.innerHTML = `
      <h3 style="margin:0 0 10px 0;">Prüfprotokoll erzeugen & signieren</h3>
      <div id="tm-preview-area" style="border:1px solid #eee; padding:10px; max-height:160px; overflow:auto;"></div>
      <div style="margin-top:12px;">
        <label style="font-weight:600">Unterschrift (zeichnen):</label>
        <canvas id="tm-sign-canvas"></canvas>
        <div id="tm-actions">
          <button id="tm-clear-sign">Signatur löschen</button>
          <button id="tm-pdf">PDF erzeugen</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    document.body.appendChild(modal);
    overlay.addEventListener('click', closeUi);

    const c = modal.querySelector('#tm-sign-canvas');
    const ctx = c.getContext('2d');

    function resize() {
      const r=c.getBoundingClientRect(), dpr=window.devicePixelRatio||1;
      if (r.width === 0 || r.height === 0) return; // erst nach Öffnen sinnvoll
      c.width=Math.round(r.width*dpr); c.height=Math.round(r.height*dpr);
      ctx.setTransform(dpr,0,0,dpr,0,0); ctx.clearRect(0,0,r.width,r.height);
      ctx.lineWidth=2; ctx.lineCap='round'; ctx.strokeStyle='#111';
    }
    resizeCanvasFn = resize; // merken, damit wir beim Öffnen erneut ausführen

    const pos = e => { const r=c.getBoundingClientRect(); const p=e.touches? e.touches[0]: e; return {x:p.clientX-r.left, y:p.clientY-r.top}; };
    let drawing=false;

    // Pointer
    c.addEventListener('pointerdown',e=>{ if (!resizeCanvasFn) return; drawing=true; const p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); e.preventDefault(); });
    c.addEventListener('pointermove',e=>{ if(!drawing) return; const p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); e.preventDefault(); });
    c.addEventListener('pointerup',()=>drawing=false);
    c.addEventListener('pointerleave',()=>drawing=false);

    // Maus Fallback
    c.addEventListener('mousedown',e=>{ drawing=true; const p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); });
    c.addEventListener('mousemove',e=>{ if(!drawing) return; const p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); });
    addEventListener('mouseup',()=>drawing=false);

    // Touch Fallback
    c.addEventListener('touchstart',e=>{ drawing=true; const p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); e.preventDefault(); }, {passive:false});
    c.addEventListener('touchmove',e=>{ if(!drawing) return; const p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); e.preventDefault(); }, {passive:false});
    c.addEventListener('touchend',()=>drawing=false);

    addEventListener('resize', ()=>resize());
    modal.querySelector('#tm-clear-sign').addEventListener('click',()=>resize());
    modal.querySelector('#tm-pdf').addEventListener('click',()=>generatePdf(c.toDataURL('image/png')));
  }

  function openUi(){
    ensureModal();
    const d = extract();
    document.getElementById('tm-preview-area').innerHTML =
      `<div><b>ID:</b> ${esc(d.id||'-')}</div>
       <div><b>Fällig:</b> ${esc(d.dueDate||'-')}</div>
       <div style="margin-top:4px"><b>Titel:</b> ${esc(d.title||'-')}</div>
       <div style="margin-top:4px"><b>Status:</b> ${esc(d.status||'-')}</div>
       <div style="margin-top:4px"><b>Kommentar:</b> ${esc(d.comment||'-')}</div>`;
    document.getElementById('tm-pruef-overlay').style.display='block';
    const m = document.getElementById('tm-pruef-modal');
    m.style.display='block';

    // WICHTIG: Canvas erst jetzt sauber dimensionieren (fix für "erst nach 'löschen' zeichnen")
    requestAnimationFrame(()=>{ if (typeof resizeCanvasFn === 'function') resizeCanvasFn(); });
    setTimeout(()=>{ if (typeof resizeCanvasFn === 'function') resizeCanvasFn(); }, 50);
  }
  function closeUi(){ const o=document.getElementById('tm-pruef-overlay'); const m=document.getElementById('tm-pruef-modal'); if(o)o.style.display='none'; if(m)m.style.display='none'; }
  const esc = s => String(s).replace(/[&<>]/g, c=>({ '&':'&amp;','<':'&lt;','>':'&gt;' }[c]));

  /* ---------- Extraction ---------- */
  function clean(s){
    return String(s||'')
      .replace(/\s+/g,' ')
      .replace(/Mehr|Versteckte Elemente anzeigen/gi,'')
      .replace(/(\s*\/\s*){2,}/g,' / ')
      .replace(/\|\s*\|\s*/g,'|')
      .trim();
  }
function extract(){
  const d = {};

  // --- robuste ID-Erkennung ---
(function setId(){
  // 1) bevorzugt: genau "ID <zahl>"
  const nodes = Array.from(document.querySelectorAll(
    'span,div,p,li,h1,h2,h3,td,strong,b'
  ));
  // Helper: sichtbarer Text
  const getText = el => (el && el.textContent || '').replace(/\s+/g,' ').trim();

  // Kandidaten mit exakt "ID 123"
  const exact = nodes.map(n => getText(n))
    .filter(t => /^ID\s*\d+$/i.test(t));

  // Kandidaten mit beginnend "ID 123 ..." (falls exact nicht gefunden)
  const starts = nodes.map(n => getText(n))
    .filter(t => /^ID\s*\d+/i.test(t));

  let idText = exact[0] || starts[0] || '';

  // Falls nur ein längerer String vorhanden war, auf "ID 123" kürzen
  if (idText) {
    const m = idText.match(/ID\s*(\d+)/i);
    if (m) idText = `ID ${m[1]}`;
  }

  // Fallback auf URL
  if (!idText) idText = `ID ${location.pathname.split('/').pop()}`;

  d.id = idText;
})();


  // Breadcrumb/Pfad
  const crumbs = Array.from(document.querySelectorAll('nav, [class*="Breadcrumb"]'))
    .map(n=>n.innerText).join(' / ');
  d.department = clean(crumbs);

  // (optionales) Fälligkeitsdatum
  const due = document.querySelector('[data-testid="dueDatetext"] span');
  d.dueDate = (due && due.textContent) ? clean(due.textContent) : '';

  // Titel: nimm das H2 mit der Aufgabenformulierung ("Prüfen Sie …")
  let h2 = document.querySelector('h2.MuiTypography-h5');
  if (!h2) {
    h2 = Array.from(document.querySelectorAll('h2'))
      .find(h => (h.textContent || '').trim().length > 20);
  }
  d.title = clean(h2 ? h2.textContent : '');  // -> landet in "Aufgabe / Titel"

  // Status (rechter Radiobutton)
  const sel = document.querySelector('[role="radiogroup"] [aria-checked="true"]') ||
              document.querySelector('input[type="radio"]:checked');
  d.status = sel ? clean((sel.closest('label')||sel.parentElement||{}).innerText || sel.value || '') : '';

  // Kommentar
  const comment = document.querySelector('textarea');
  d.comment = clean(comment ? (comment.value || '') : '');

  // Durchgeführt von: gezielt das User-Label (Nachname, Vorname)
  let userEl = document.querySelector('span.MuiTypography-body1.css-ximsk5');
  if (!userEl) {
    userEl = Array.from(document.querySelectorAll('span.MuiTypography-body1, .MuiTypography-body1'))
      .find(el => /[A-Za-zÄÖÜäöüß-]+,\s*[A-Za-zÄÖÜäöüß-]+/.test((el.textContent||'').trim()));
  }
  d.user = userEl ? clean(userEl.textContent) : '';

  d.extractedAt = new Date().toLocaleString('de-DE');
  d.performedAt = new Date().toLocaleDateString('de-DE'); // immer heute

  return d;
}



  // --- PDF: dynamische Boxen mit Label-auf-eigener-Zeile ---
function generatePdf(signatureDataUrl){
  const { jsPDF } = window.jspdf;
  const d = extract();

  // Seiten-/Layout-Settings
  const M = { L:16, R:16, T:16, B:16 };
  const pageW = 210, pageH = 297, contentW = pageW - M.L - M.R;
  const lineH = 5, padTop = 6, padBetween = 4;

  const doc = new jsPDF({ unit:'mm', format:'a4' });
  doc.setLineWidth(0.6);

  // helpers
  const line = (x1,y1,x2,y2)=>doc.line(x1,y1,x2,y2);
  const title = (txt)=>{ doc.setFont('helvetica','bold'); doc.setFontSize(18); doc.text(txt, M.L, y); y+=3; line(M.L,y,pageW-M.R,y); y+=5; };
  const ensure = (need=20)=>{ if (y+need>pageH-M.B){ doc.addPage(); y=M.T; } };

  function labeledBox(labelTxt, contentTxt, opts={}){
    const { minH=16, fs=10, maxW=contentW-6, leftIndent=3 } = opts;
    const startY = y;
    doc.setFont('helvetica','bold'); doc.setFontSize(10);
    const labelY = startY + padTop;
    doc.text(labelTxt, M.L + leftIndent, labelY);

    doc.setFont('helvetica','normal'); doc.setFontSize(fs);
    const valueLines = doc.splitTextToSize(String(contentTxt||'-'), maxW);
    const valueY = labelY + lineH;   // eine Zeile unter Label
    doc.text(valueLines, M.L + leftIndent, valueY);

    const needed = Math.max(minH, (padTop + lineH + lineH*valueLines.length + 4));
    doc.rect(M.L, startY, contentW, needed);
    y = startY + needed + padBetween;
  }

  function pruefRowFixed(labelTxt, valueTxt, boxH = 16){
  const startY = y;
  const labelY = startY + 6;          // 1. Zeile: Label
  const valueY = labelY + 5;          // 2. Zeile: Wert (eine Zeile darunter)

  // Rahmen (fixe Höhe)
  doc.rect(M.L, startY, contentW, boxH);

  // Label + Wert (kein Umbruch nötig bei kurzen Texten)
  doc.setFont('helvetica','bold'); doc.setFontSize(10);
  doc.text(labelTxt, M.L + 3, labelY);
  doc.setFont('helvetica','normal'); doc.setFontSize(10);
  doc.text(String(valueTxt || '-'), M.L + 70, valueY);

  y = startY + boxH + padBetween;     // weiter unterhalb der Box
}



  // Start
  let y = M.T;
  title('Prüfprotokoll');

  // Meta
  ensure(22);
  const metaY = y;
  doc.rect(M.L, metaY, contentW, 22);
  doc.setFont('helvetica','bold'); doc.setFontSize(10);
  doc.text('Erstellt:', M.L+3, metaY+7);
  doc.setFont('helvetica','normal'); doc.text(d.extractedAt, M.L+35, metaY+7);
  doc.setFont('helvetica','bold'); doc.text('ID:', M.L+3, metaY+15);
  doc.setFont('helvetica','normal'); doc.text(d.id, M.L+35, metaY+15);
  y = metaY + 22 + padBetween;

  // Pfad
  ensure(18);
  labeledBox('Abteilung / Pfad:', d.department, { minH:18 });

  // Aufgabe / Titel (enthält jetzt den langen H2-Text)
  ensure(18);
  labeledBox('Aufgabe / Titel', d.title, { minH:18, fs:11 });

  // --- KEINE Beschreibung mehr ---

  // Prüfung (Kopf + Zeilen)
  ensure(10);
  const prStart = y;
  doc.rect(M.L, prStart, contentW, 8);
  doc.setFont('helvetica','bold'); doc.setFontSize(10);
  doc.text('Prüfung', M.L+3, prStart+6);
  y = prStart + 8 + padBetween;

  ensure(16); pruefRowFixed('Art der Prüfung', d.status || '-', 14);
ensure(16); pruefRowFixed('Datum der Prüfung', d.performedAt, 14);
ensure(16); pruefRowFixed('Durchgeführt von', d.user || '-', 14);

  // Kommentar
  ensure(40);
  labeledBox('Kommentar (Neuer Prüfeintrag)', d.comment || '-', { minH:40 });

  // Unterschrift
  ensure(36);
  const sigH = 32;
  const sigY = y;
  doc.rect(M.L, sigY, contentW*0.62, sigH);
  doc.setFont('helvetica','bold'); doc.setFontSize(10);
  doc.text('Unterschrift Prüfer', M.L+3, sigY + padTop);
  try{
    const p = doc.getImageProperties(signatureDataUrl);
    const w = contentW*0.62 - 20;
    const h = (p.height/p.width)*w;
    const imgY = sigY + padTop + lineH; // unter Label
    doc.addImage(signatureDataUrl, 'PNG', M.L+10, imgY, w, Math.min(20, h));
  }catch(e){
    doc.setFont('helvetica','normal'); doc.setFontSize(10);
    doc.text('_____________________________', M.L+10, sigY + padTop + lineH);
  }
  y = sigY + sigH + padBetween;

  // Fuß
  doc.setFont('helvetica','normal'); doc.setFontSize(9);
  doc.text('Erstellt aus Eticor-Delegation via Tampermonkey', M.L, pageH - M.B);

  const fname = `Pruefprotokoll_${(d.id||'Eticor').replace(/\s+/g,'')}_${new Date().toISOString().slice(0,10)}.pdf`;
  doc.save(fname);
  closeUi();
}

  /* ---------- Boot ---------- */
  function boot(){
    const mo = new MutationObserver(() => {
      if (isDetail() && !document.getElementById('tm-embed-btn')) {
        const t = findRightPanel(); if (t) embedButton(t);
      }
    });
    mo.observe(document.documentElement, { childList:true, subtree:true });
    tryInsert();
  }
  boot();
})();
