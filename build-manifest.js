// build-manifest.js
// Liest Header aus tools/*.user.js und erzeugt manifest.json
// Unterstützt mehrere @match/@include & mehrere @grant

const fs = require('fs');
const path = require('path');

const REPO_USER = 'toni2123a';
const REPO_NAME = 'company-userscripts';
const TOOLS_DIR = path.join(__dirname, 'tools');
const RAW_PREFIX = `https://raw.githubusercontent.com/${REPO_USER}/${REPO_NAME}/main/tools/`;

function parseHeader(str) {
  const meta = { grants: [], matches: [] };
  const block = (str.match(/\/\/\s*==UserScript==([\s\S]*?)\/\/\s*==\/UserScript==/) || [])[1] || '';
  block.split(/\r?\n/).forEach(line => {
    const m = line.match(/^\s*\/\/\s*@(\w+)\s+(.+)\s*$/);
    if (!m) return;
    const key = m[1].toLowerCase();
    const val = m[2].trim();
    if (key === 'grant') {
      if (!meta.grants.includes(val)) meta.grants.push(val);
    } else if (key === 'match' || key === 'include') {
      meta.matches.push(val);
    } else {
      meta[key] = val;
    }
  });
  if (!meta.matches.length) meta.matches = ['*://*/*'];
  if (!meta.grants.length) meta.grants = ['none'];
  return meta;
}

function main() {
  if (!fs.existsSync(TOOLS_DIR)) fs.mkdirSync(TOOLS_DIR, { recursive: true });
  const files = fs.readdirSync(TOOLS_DIR).filter(f => f.endsWith('.user.js'));
  const tools = files.map(f => {
    const p = path.join(TOOLS_DIR, f);
    const txt = fs.readFileSync(p, 'utf8');
    const meta = parseHeader(txt);
    const id = (path.basename(f, '.user.js') || (meta.name || 'tool')).toLowerCase().replace(/\s+/g,'-');
    return {
      id,
      name: meta.name || id,
      version: meta.version || '0.0.0',
      url: RAW_PREFIX + encodeURIComponent(f),
      match: meta.matches,         // Array
      grants: meta.grants,         // Array
      description: meta.description || ''
    };
  }).sort((a,b)=> a.name.localeCompare(b.name,'de'));

  const out = { tools };
  fs.writeFileSync(path.join(__dirname,'manifest.json'), JSON.stringify(out, null, 2));
  console.log('manifest.json geschrieben mit', tools.length, 'Tools');
}

main();
