// ==UserScript==
// @name         Company Bootstrap
// @namespace    https://github.com/<deinUser>/company-userscripts
// @version      1.0.0
// @description  Lädt alle Tools aus dem Manifest
// @match        *://*/*
// @grant        GM_xmlhttpRequest
// @connect      raw.githubusercontent.com
// @updateURL    https://raw.githubusercontent.com/<deinUser>/company-userscripts/main/bootstrap.user.js
// @downloadURL  https://raw.githubusercontent.com/<deinUser>/company-userscripts/main/bootstrap.user.js
// ==/UserScript==
(async function () {
  const manifestUrl = "https://raw.githubusercontent.com/<deinUser>/company-userscripts/main/manifest.json";
  function httpGet(url) {
    return new Promise((res, rej) => {
      GM_xmlhttpRequest({
        method: "GET",
        url,
        onload: r => (r.status >= 200 && r.status < 300) ? res(r.responseText) : rej(r),
        onerror: rej
      });
    });
  }
  try {
    const manifest = JSON.parse(await httpGet(manifestUrl));
    for (const t of manifest.tools) {
      const code = await httpGet(t.url);
      (0, eval)(code);
      console.log(`[Bootstrap] geladen: ${t.name} v${t.version}`);
    }
  } catch (e) { console.error("Bootstrap Fehler", e); }
})();
