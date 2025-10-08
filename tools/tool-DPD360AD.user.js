// ==UserScript==
// @name         DPD360 – Zustelldaten-Abfrage (8 Wochen) + Trackinglinks (robust)
// @namespace    dpd360.tools
// @version      1.6.1
// @description  Button robust einhängen, Hausnummern aus API (deliveryAddress.houseNo) sauber sortieren, Trackinglinks anzeigen.
// @match        https://dpd360.dpd.de/ops/notification_new.aspx*
// @grant        GM.xmlHttpRequest
// @grant        GM_addStyle
// @connect      depotportal.dpd.com
// ==/UserScript==

(function () {
  'use strict';

  const API_BASE = 'https://depotportal.dpd.com/rest/ShipmentInfoService/V1_0/ShipmentInfoService_1_0/ShipmentInfo';
  const TRACK_URL = n => `https://depotportal.dpd.com/dp/de_DE/tracking/parcels/${encodeURIComponent(n)}`;

  GM_addStyle(`
    .zustelldaten-btn{display:inline-block;margin-top:8px;padding:8px 12px;border:0;border-radius:6px;background:#a60019;color:#fff;font-weight:600;cursor:pointer}
    .zustelldaten-btn:hover{filter:brightness(.95)}
    .zd-modal-backdrop{position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99998}
    .zd-modal{position:fixed;top:6vh;left:50%;transform:translateX(-50%);width:min(900px,92vw);max-height:88vh;overflow:auto;z-index:99999;background:#fff;border-radius:10px;box-shadow:0 12px 40px rgba(0,0,0,.35);padding:18px;font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif}
    .zd-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
    .zd-title{font-size:18px;font-weight:700}
    .zd-close{border:0;background:#eee;border-radius:8px;padding:6px 10px;cursor:pointer}
    .zd-list{border:1px solid #eee;border-radius:8px;overflow:hidden}
    .zd-row{display:grid;grid-template-columns:90px 1fr 180px 200px;gap:10px;padding:10px 12px;border-top:1px solid #eee}
    .zd-row:first-child{border-top:0}
    .zd-row:nth-child(odd){background:#fafafa}
    .zd-head{font-weight:700;background:#f3f3f3}
    .zd-muted{color:#666}
    .zd-badge{display:inline-block;padding:2px 6px;border-radius:6px;background:#f0f0f0;font-size:12px}
    .zd-link{color:#a60019;text-decoration:none;font-weight:600}
    .zd-link:hover{text-decoration:underline}
    .zd-spinner{display:inline-block;width:16px;height:16px;border:2px solid #ddd;border-top-color:#a60019;border-radius:50%;animation:spin .8s linear infinite;vertical-align:-3px;margin-right:6px}
    @keyframes spin{to{transform:rotate(360deg)}}
  `);

  const $ = (s, r=document) => r.querySelector(s);

  // ---------- Modal ----------
  function showModal(html,title='Zustelldaten'){
    const back=document.createElement('div');back.className='zd-modal-backdrop';
    const box=document.createElement('div');box.className='zd-modal';
    box.innerHTML=`<div class="zd-header"><div class="zd-title">${title}</div><button class="zd-close">Schließen</button></div>${html}`;
    back.addEventListener('click',e=>{if(e.target===back){back.remove();box.remove();}});
    box.querySelector('.zd-close').onclick=()=>{back.remove();box.remove();};
    document.body.append(back,box);
  }
  function setModalBody(html){
    const box=$('.zd-modal'); if(!box){ showModal(html); return; }
    box.innerHTML=`<div class="zd-header"><div class="zd-title">Zustelldaten</div><button class="zd-close">Schließen</button></div>${html}`;
    box.querySelector('.zd-close').onclick=()=>{document.querySelector('.zd-modal-backdrop')?.remove();box.remove();};
  }
  const loading = () => showModal(`<div><span class="zd-spinner"></span>Abfrage läuft …</div>`);

  // ---------- DPD360 helpers ----------
  function getParcelNo(){
    const head=$('#ContentPlaceHolder1_labHeadline');
    return head?.textContent.trim() || $('#ContentPlaceHolder1_txtParcelNo')?.value?.trim() || '';
  }

  // holt den sichtbaren Lieferadress-Block (robust)
  function getAddressBlock(){
    // 1) Standard-Label (häufig)
    const lbl = $('#ContentPlaceHolder1_labAvisAddress');
    if (lbl && lbl.innerText.trim()) return lbl.innerText.trim();

    // 2) Fallback: Box mit Überschrift „LIEFERADRESSE“
    // wir suchen ein Element, dessen erste sichtbare Zeile exakt so beginnt
    const all = Array.from(document.querySelectorAll('div,section,article'));
    for (const el of all){
      const t = el.innerText?.trim();
      if (!t) continue;
      const first = t.split(/\r?\n/)[0]?.trim();
      if (/^LIEFERADRESSE$/i.test(first)) {
        // restliche Zeilen ohne die Überschrift zurückgeben
        return t.split(/\r?\n/).slice(1).join('\n').trim();
      }
    }
    return '';
  }

  function parseZipAndStreet(blockText){
  if (!blockText) return null;

  // Zeilen vorbereiten
  const lines = blockText
    .split(/\r?\n/)
    .map(s => s.replace(/^"+|"+$/g,'').trim())   // evtl. Anführungszeichen aus DevTools
    .filter(Boolean);

  // Bereich bis vor "Tel."
  const telIdx = lines.findIndex(l => /^tel\b|^tel\.:?/i.test(l));
  const pre = telIdx > -1 ? lines.slice(0, telIdx) : lines.slice(); // falls kein "Tel." gefunden

  // Straße nach deiner Regel wählen
  let streetLine = '';
  if (pre.length >= 4) {
    streetLine = pre[2];        // 4 Zeilen vor Tel. -> Zeile 3 (Index 2)
  } else if (pre.length >= 3) {
    streetLine = pre[1];        // 3 Zeilen vor Tel. -> Zeile 2 (Index 1)
  } else {
    // Fallback: erste Zeile, die nicht PLZ/Ort ist
    streetLine = pre.find(l => !/\b\d{5}\b/.test(l)) || pre[0] || '';
  }

  // PLZ aus irgendeiner der "pre"-Zeilen
  const zipMatchLine = pre.find(l => /\b\d{5}\b/.test(l)) || '';
  const zipMatch = zipMatchLine.match(/\b(\d{5})\b/);

  // reiner Straßenname (Hausnr., Zusätze abtrennen)
  const streetOnly = (streetLine || '').replace(/[,]*\s*\d+[A-Za-z\-\/]*.*$/, '').trim();

  return zipMatch ? { zip: zipMatch[1], streetOnly } : null;
}


  function last8WeeksRange(){
    const to=new Date(), from=new Date(); from.setDate(to.getDate()-56);
    const pad=n=>String(n).padStart(2,'0');
    const fmt=(d,end)=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${end?'23:59:59':'00:00:00'}`;
    return { from:fmt(from,false), to:fmt(to,true) };
  }

  // ---------- API ----------
  function buildUrl({zip,street,from,to}){
    const p=new URLSearchParams({
      messageLanguage:'de_DE',
      addressType:'DELIVERYADDRESS',
      zipCode:zip,
      country:'DE',
      limit:'100',
      consignmentDateFrom:from,
      consignmentDateTo:to,
      street:`%${street}%`
    });
    return `${API_BASE}?${p.toString()}`;
  }

  function httpGet(url,cb){
    GM.xmlHttpRequest({
      method:'GET', url, responseType:'json',
      headers:{'Accept':'application/json','X-Requested-With':'XMLHttpRequest'},
      onload:r=>{
        let d=r.response; try{ if(!d) d=JSON.parse(r.responseText); }catch{}
        cb(d);
      },
      onerror:()=>cb(null)
    });
  }

  // ---------- Normalisieren & Sortieren ----------
  const hnKey = s=>{
    const m=String(s||'').match(/(\d+)([A-Za-z]*)/);
    return { num: m?parseInt(m[1]):99999, suf: m?m[2].toLowerCase():'' };
  };
  const sortByHouseNo = (a,b)=>{
    const A=hnKey(a.houseNo), B=hnKey(b.houseNo);
    if (A.num!==B.num) return A.num-B.num;
    return A.suf.localeCompare(B.suf,'de');
  };

  function normalizeApi(data,fallbackStreet,zip){
    if(!data?.ShipmentInfoResponse?.ShipmentInfo) return [];
    return data.ShipmentInfoResponse.ShipmentInfo.map(x=>{
      const d=x.deliveryAddress||{};
      const when=(x.consignmentDate||'')+(x.consignmentTime?` ${x.consignmentTime}`:'');
      const parcel=(Array.isArray(x.parcels)&&x.parcels[0]?.parcelLabelNumber)||'';
      return {
        houseNo: d.houseNo||'',
        street: d.street||fallbackStreet,
        zipcode: d.zipCode||zip,
        city: d.city||'',
        status:`Service ${x.serviceCode||''}`,
        consNo: parcel,
        when
      };
    });
  }

  // ---------- UI Aktion ----------
  function onFetch(){
    const addr = getAddressBlock();
    const parsed = parseZipAndStreet(addr);
    if(!parsed){ showModal('<div>Adresse konnte nicht erkannt werden.</div>'); return; }

    const parcel = getParcelNo();
    const { from, to } = last8WeeksRange();
    const url = buildUrl({ zip: parsed.zip, street: parsed.streetOnly, from, to });

    loading();
    httpGet(url, data=>{
      const rows = normalizeApi(data, parsed.streetOnly, parsed.zip).sort(sortByHouseNo);
      if (!rows.length) { setModalBody(`<div>Keine Treffer für <b>${parsed.streetOnly}</b> in <b>${parsed.zip}</b> in den letzten 8 Wochen.</div>`); return; }

      const html = rows.map(r=>{
        const link = r.consNo ? `<a class="zd-link" target="_blank" href="${TRACK_URL(r.consNo)}">${r.consNo}</a>` : '-';
        return `<div class="zd-row">
          <div>${r.houseNo?`<span class="zd-badge">${r.houseNo}</span>`:'-'}</div>
          <div>${r.street}<br><span class="zd-muted">${r.zipcode} ${r.city}</span></div>
          <div>${link}</div>
          <div>${r.when?new Date(r.when).toLocaleString():'-'}<br><span class="zd-muted">${r.status}</span></div>
        </div>`;
      }).join('');

      setModalBody(`
        <div class="zd-hint">
          Treffer für <b>${parsed.streetOnly}</b> in <b>${parsed.zip}</b> – letzte 8 Wochen.
          ${parcel?`(aktuelle Paketnr.: <span class="zd-badge">${parcel}</span>)`:''}
        </div>
        <div class="zd-list">
          <div class="zd-row zd-head">
            <div>Hausnr.</div><div>Adresse</div><div>Sendung</div><div>Zeit/Status</div>
          </div>${html}
        </div>
      `);
    });
  }

  // ---------- Button robust einfügen ----------
  function placeButton(){
    if (document.querySelector('.zustelldaten-btn')) return;

    const btn = document.createElement('button');
    btn.className = 'zustelldaten-btn';
    btn.textContent = 'Zustelldaten abfragen';
    btn.addEventListener('click', onFetch);

    // 1) Bevorzugt direkt am bekannten Label
    const lbl = $('#ContentPlaceHolder1_labAvisAddress');
    if (lbl) { lbl.insertAdjacentElement('afterend', btn); return true; }

    // 2) Fallback: Box mit Überschrift „LIEFERADRESSE“
    const candidates = Array.from(document.querySelectorAll('div,section,article'));
    for (const el of candidates){
      const t = el.innerText?.trim();
      if (!t) continue;
      const first = t.split(/\r?\n/)[0]?.trim();
      if (/^LIEFERADRESSE$/i.test(first)) {
        el.appendChild(btn);
        return true;
      }
    }

    // 3) Notfall: rechts unten fix anheften
    btn.style.position='fixed'; btn.style.bottom='18px'; btn.style.right='18px';
    document.body.appendChild(btn);
    return true;
  }

  function init(){
    placeButton();
    // Beobachten, falls die Seite per Ajax neu rendert
    const obs = new MutationObserver(()=>placeButton());
    obs.observe(document.body, {childList:true, subtree:true});
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
