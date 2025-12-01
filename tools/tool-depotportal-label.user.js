// ==UserScript==
// @name         Depotportal – Paketschein
// @namespace    bodo.tools
// @version      1.08
// @description  Paketschein mit Empfänger/Absender, 1/1, festem QR zur Abstell-Okay-Seite, Barcode und Drucken / Zwischenablage / Abbrechen
// @updateURL    https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-depotportal-label.user.js
// @downloadURL  https://raw.githubusercontent.com/toni2123a/company-userscripts/main/tools/tool-depotportal-label.user.js
// @match        https://depotportal.dpd.com/*
// @match        http://depotportal.dpd.com/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // ================================================================
    // [1] Barcode und fester QR zur ASG-Seite
    // ================================================================
    function buildBarcodeValue(val) {
        // Nur Ziffern behalten und vorne ein % setzen
        let s = (val || '').toString().replace(/[^\d]/g, '');
        if (!s) return '';
        if (!s.startsWith('%')) {
            s = '%' + s;
        }
        return s;
    }

    function barcodeUrl(val) {
        const finalCode = buildBarcodeValue(val);
        if (!finalCode) return '';
        return `https://barcodeapi.org/api/128/${encodeURIComponent(finalCode)}`;
    }

    function asgQrUrl() {
        const url = 'https://www.dpd.com/de/de/support/abstell-okay-bei-dpd/';
        return 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&margin=0&data=' + encodeURIComponent(url);
    }

    // ================================================================
    // [2] Datum + Text-Helfer
    // ================================================================
    function printDate() {
        const d = new Date();
        return d.toLocaleDateString('de-DE', { year:'2-digit', month:'2-digit', day:'2-digit' }) +
               ' ' +
               d.toLocaleTimeString('de-DE', { hour:'2-digit', minute:'2-digit' });
    }
    function txt(el) {
        return (el && el.textContent ? el.textContent : '').replace(/\s+/g,' ').trim();
    }

    // ================================================================
    // [3] Paketscheinnummer + Service + PLZ
    // ================================================================
    function getParcelNumber() {
        // 1. Versuch: aus der URL (…pln=XXXXXXXXXXXXXX…)
        const mUrl = location.href.match(/[?&]pln=(\d{14})/);
        if (mUrl) return mUrl[1];

        // 2. Fallback: erste 14-stellige Zahl im Dokument
        const m = (document.body.innerText || '').match(/\b\d{14}\b/);
        return m ? m[0] : '';
    }

    function getServiceCode() {
        const t = document.querySelectorAll('table');
        for (const tb of t) {
            const th = tb.querySelectorAll('th');
            let idx = -1;
            for (let i=0;i<th.length;i++) if (/^Service$/i.test(txt(th[i]))) idx = i;
            if (idx === -1) continue;
            const rows = tb.querySelectorAll('tr');
            for (const r of rows) {
                const c = r.querySelectorAll('td');
                if (c.length > idx) return txt(c[idx]);
            }
        }
        return '';
    }

    function getPlz() {
        const m = (document.body.innerText||'').match(/\b\d{5}\b/);
        return m ? m[0] : '';
    }

    // ================================================================
    // [4] Empfänger + Absender
    // ================================================================
    function address(start, stops) {
        let raw = (document.body.innerText||'').replace(/\r/g,'');
        let p = raw.indexOf(start);
        if (p === -1) return '';
        p = raw.indexOf('\n',p);
        let end = raw.length;
        for (const s of stops) {
            const x = raw.indexOf(s,p);
            if (x !== -1 && x < end) end = x;
        }
        return raw.substring(p,end)
                  .split('\n')
                  .map(s => s.trim())
                  .filter(Boolean)
                  .filter(s => !/Details\s+(aus|ein)blenden/i.test(s))
                  .join('<br>');
    }

    // ================================================================
    // [5] Layout-HTML – Versender 5 cm hoch / 2,5 cm breit
    //      Empfänger editierbar + Feld UNTER Empfänger editierbar
    // ================================================================
    function labelHtml(data) {
        return `
<div style="width:100%;height:100%;box-sizing:border-box;display:flex;flex-direction:column;">

  <div style="flex:3;display:flex;flex-direction:column;">

    <div style="height:50mm;display:flex;border-bottom:1px solid #000;box-sizing:border-box;">
      <div style="flex:1;border-right:1px solid #000;padding:4px;box-sizing:border-box;">
        <div style="font-size:7px;margin-bottom:4px;">Empfänger:</div>
        <!-- Empfänger EDITIERBAR -->
        <div id="labelEmpf" contenteditable="true"
             style="font-size:12.5px;font-weight:bold;outline:none;min-height:18mm;word-wrap:break-word;white-space:normal;">
          ${data.emp}
        </div>
      </div>

      <!-- Versender 5 cm hoch, 2,5 cm breit -->
      <div style="width:25mm;height:50mm;position:relative;overflow:hidden;box-sizing:border-box;">
        <div style="
            position:absolute;
            top:9mm;
            right:1mm;
            transform:rotate(90deg);
            font-size:7px;
            line-height:1.3;
        ">
Absender:
${data.abs}
        </div>
      </div>
    </div>

    <div style="flex:1;display:grid;grid-template-columns:3fr 0.8fr 1.4fr;box-sizing:border-box;">
      <!-- FELD UNTER EMPFÄNGER: komplett editierbar -->
      <div style="border-right:1px solid #000;border-bottom:1px solid #000;">
        <div id="labelNote" contenteditable="true"
             style="width:100%;height:100%;box-sizing:border-box;padding:3px;font-size:14px;white-space:pre-wrap;word-wrap:break-word;outline:none;">
        </div>
      </div>
      <div style="border-right:1px solid #000;border-bottom:1px solid #000;display:flex;align-items:center;justify-content:center;font-size:8px;line-height:1.2;">
        Lieferung<br>1 / 1
      </div>
      <div style="border-bottom:1px solid #000;display:flex;align-items:flex-start;justify-content:center;">
        <img src="${data.qr}" style="margin-top:4px;max-width:96%;max-height:96%;">
      </div>
    </div>

  </div>

  <div style="flex:2;padding-top:4px;box-sizing:border-box;">
    <div style="font-size:20px;font-weight:bold;letter-spacing:2px;text-align:left;padding-left:8px;margin-bottom:2px;">
      ${data.psn.replace(/(\d{4})(?=\d)/g,'$1 ')}
    </div>
    <div style="font-size:9px;text-align:center;margin-bottom:4px;">
      ${data.service}
    </div>
    <div style="font-size:8px;text-align:center;margin-bottom:2px;">
      ${data.date}
    </div>
    <div style="text-align:center;margin-top:2px;">
      <img src="${data.bc}" style="width:92%;max-width:92%;height:110px;object-fit:contain;">
    </div>
  </div>

</div>`;
    }

    // ================================================================
    // [6] Drucken – 1 A4-Seite, Label oben links
    // ================================================================
    function printLabelFromDom() {
        const el = document.getElementById('labelInner');
        if (!el) return alert('Kein Label gefunden.');

        const html = el.innerHTML;

        const w = window.open('', '_blank');
        w.document.write(`
<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
@page { size:A4 portrait; margin:0; }
body { margin:0; }
.label-wrap {
    margin-top:10mm;
    margin-left:10mm;
}
.label {
    width:100mm;
    height:150mm;
    border:1px solid #000;
    box-sizing:border-box;
}
</style></head><body>
<div class="label-wrap">
  <div class="label">${html}</div>
</div>
</body></html>`);
        w.document.close();

        w.onload = function () {
            w.focus();
            w.print();
        };
    }

    // ================================================================
    // [7] Screenshot → Zwischenablage (html2canvas)
    // ================================================================
    function copyImage() {
        const el = document.getElementById('labelInner');
        if (!el) return alert('Kein Label gefunden.');
        const s = document.createElement('script');
        s.src='https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js';
        s.onload=()=>{
            window.html2canvas(el,{scale:2,useCORS:true}).then(c=>{
                c.toBlob(b=>{
                    navigator.clipboard.write([new ClipboardItem({'image/png': b})]);
                });
            });
        };
        document.head.appendChild(s);
    }

    // ================================================================
    // [8] Overlay öffnen
    // ================================================================
    function openOverlay(data) {
        let ov = document.getElementById('labelOv');
        if (!ov) {
            ov = document.createElement('div');
            ov.id='labelOv';
            Object.assign(ov.style,{
                position:'fixed',inset:'0',background:'rgba(0,0,0,0.75)',
                display:'none',alignItems:'center',justifyContent:'center',
                zIndex:'99999'
            });
            ov.addEventListener('click',e=>{ if(e.target===ov) ov.style.display='none'; });
            document.body.appendChild(ov);
        }
        ov.innerHTML = `
<div style="background:#fff;width:420px;height:650px;padding:6px;display:flex;flex-direction:column;border:1px solid #000;">
  <div style="margin-bottom:6px;display:flex;gap:6px;font-size:11px;">
    <button id="btnPrint">Drucken</button>
    <button id="btnCopy">In Zwischenablage</button>
    <button id="btnClose">Abbrechen</button>
  </div>
  <div style="flex:1;display:flex;align-items:flex-start;justify-content:center;overflow:hidden;">
    <div id="labelInner" style="width:100mm;height:150mm;border:1px solid #000;box-sizing:border-box;">
      ${labelHtml(data)}
    </div>
  </div>
</div>`;
        ov.style.display='flex';

        document.getElementById('btnClose').onclick=()=>ov.style.display='none';
        document.getElementById('btnPrint').onclick=()=>printLabelFromDom();
        document.getElementById('btnCopy').onclick=()=>copyImage();
    }

    // ================================================================
    // [9] Schriftzug „Sendungssuche“ anklickbar
    // ================================================================
    function hook() {
        const h = [...document.querySelectorAll('h1,h2,h3,h4')]
            .find(e=>txt(e)==='Sendungssuche');
        if (h && !h.dataset.pk) {
            h.dataset.pk='1';
            h.style.cursor='pointer';
            h.onclick = handleClick;
        }
    }

    // ================================================================
    // [10] Klick → Daten sammeln + Overlay
    // ================================================================
    function handleClick() {
        const psn = getParcelNumber();
        if (!psn) return alert('Keine Paketscheinnummer.');
        const plz = getPlz();
        const svc = getServiceCode().replace(/\D/g,'');
        const date = printDate();

        const emp = address('Zustelladresse',['Auftraggeber']);
        const abs = address('Auftraggeber',['Paketscheinnummer']);

        const human = '00' + (plz||'') + psn + (svc||'') + '276';

        const data = {
            emp: emp,
            abs: abs,
            date: date,
            qr: asgQrUrl(),
            bc: barcodeUrl(human),
            psn: psn,
            service: (svc ? svc+'-DE-'+plz : '')
        };
        openOverlay(data);
    }

    // ================================================================
    // [11] Init
    // ================================================================
    function init() {
        hook();
        new MutationObserver(hook).observe(document.body,{childList:true,subtree:true});
    }
    document.readyState==='loading'
        ? document.addEventListener('DOMContentLoaded',init)
        : init();

})();
