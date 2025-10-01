// ==UserScript==
// @name         OpenPricer: Abholdepot hinzufügen
// @namespace    dpd-openpricer
// @version      2.2.0
// @description  Spalte "Abholdepot" in /app/rfq/list. Liest aus /app/rfq/<ID>/requirements (#agency). Mit Cache & Queue.
// @match        https://dpdde.openpricer.com/app/rfq/list*
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  const injected = `
    (function(){
      const COL_FIELD = "pickupDepot";
      const COL_TITLE = "Abholdepot";
      const AGENCY_SELECTOR = "#agency";         // <span id="agency">10127 ...</span>
      const CACHE_TTL_MS = 12 * 60 * 60 * 1000;  // 12h
      const MAX_PARALLEL = 3;
      const DEBUG = false;
      function log(...a){ if (DEBUG) console.log("[Abholdepot]", ...a); }

      let active = 0, queue = [];
      function enqueue(fn){ queue.push(fn); pump(); }
      function pump(){ while(active<MAX_PARALLEL && queue.length){ active++; queue.shift()().finally(()=>{active--; pump();}); } }

      function cacheKey(id){ return \`op:agency:rfq:\${id}\`; }
      function cacheGet(id){ try{ const r=sessionStorage.getItem(cacheKey(id)); if(!r) return null; const {v,t}=JSON.parse(r); return (Date.now()-t> CACHE_TTL_MS)?null:v; }catch{ return null; } }
      function cacheSet(id,val){ try{ sessionStorage.setItem(cacheKey(id), JSON.stringify({v:val,t:Date.now()})); }catch{} }

      async function fetchAgency(id){
        // 1) OPS/requirements (enthält Abholung)
        // 2) Standardseite
        // 3) SimpleQuote-Fallback
        const urls = [
          \`/app/rfq/\${encodeURIComponent(id)}/requirements\`,
          \`/app/rfq/\${encodeURIComponent(id)}\`,
          \`/app/rfq/showQuoteSimpleQuote?rfqID=\${encodeURIComponent(id)}\`
        ];
        for (const url of urls){
          try{
            const res = await fetch(url, { credentials: "same-origin" });
            if(!res.ok) continue;
            const html = await res.text();
            const doc  = new DOMParser().parseFromString(html, "text/html");
            const el   = doc.querySelector(AGENCY_SELECTOR);
            const txt  = (el ? el.textContent.trim() : "");
            if (txt){ cacheSet(id, txt); return txt; }
          }catch(e){ log("fetch fail", id, url, e); }
        }
        return "-";
      }

      function ensureColumn(go){
        if (go.columnApi.getColumn(COL_FIELD)) return;
        const defs = go.columnApi.getAllGridColumns().map(c => c.getColDef());
        defs.push({
          colId: COL_FIELD,
          field: COL_FIELD,            // wichtig: echtes Feld, kein valueGetter
          headerName: COL_TITLE,
          cellClass: ' ag-left-aligned-cell ',
          sortable: true,
          filter: true,
          autoHeight: true,
          width: 180
        });
        go.api.setColumnDefs(defs);
      }

      function processRows(go){
        go.api.forEachNodeAfterFilterAndSort(node=>{
          const data = node.data || {};
          const id = data.id;
          if (!id) return;

          if (data.__loadingAgency) return;
          if (data[COL_FIELD] && data[COL_FIELD] !== "lädt…") return;

          const cached = cacheGet(id);
          if (cached){
            data[COL_FIELD] = cached;
            go.api.refreshCells({ rowNodes:[node], columns:[COL_FIELD], force:true });
            return;
          }

          data[COL_FIELD] = "lädt…";
          go.api.refreshCells({ rowNodes:[node], columns:[COL_FIELD], force:true });

          data.__loadingAgency = true;
          enqueue(async ()=>{
            try{
              const val = await fetchAgency(id);
              data[COL_FIELD] = val || "-";
            }catch(e){
              data[COL_FIELD] = "Fehler";
              log("load error", id, e);
            }finally{
              data.__loadingAgency = false;
              go.api.refreshCells({ rowNodes:[node], columns:[COL_FIELD], force:true });
            }
          });
        });
      }

      function attach(){
        const go = gridOptionsAgGridrfqs;
        ensureColumn(go);
        processRows(go);
        go.api.addEventListener('firstDataRendered', ()=>{ ensureColumn(go); processRows(go); });
        go.api.addEventListener('modelUpdated',       ()=>{ ensureColumn(go); processRows(go); });
      }

      (function wait(){
        if (typeof agGrid!=="undefined" && typeof gridOptionsAgGridrfqs!=="undefined" && gridOptionsAgGridrfqs.api){
          attach();
        } else {
          setTimeout(wait, 150);
        }
      })();
    })();
  `;
  const s = document.createElement('script');
  s.textContent = injected;
  document.documentElement.appendChild(s);
  s.remove();
})();
