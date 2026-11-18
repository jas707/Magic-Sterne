/**
 * Sucht Produktbilder für die Sterne in den Lieferanten-Shops und legt sie als <Nr>.jpg ab.
 *
 *   node bilder-holen.js <Ausgabeordner> "<Magic.xlsx>" "<Mea.xlsx>" ["<Sterne 2025.xlsx>"]
 *
 * - Magic-Sterne: guru-shop.de, Suche nach Modellname (und alter Lieferanten-Nr. aus 2025, falls angegeben)
 * - Blumen-Sterne: mea-living.de, Zuordnung über die Nummer in der Produkt-Adresse („…-stella-029“)
 * Ergebnis: Bilder (max. 600 px, JPEG) + uebersicht.html zum Prüfen + ergebnis.json.
 * Die Artikelliste kommt aus derselben Auswertung wie der Import (gas/Import.gs).
 */
const fs = require('fs'), path = require('path'), vm = require('vm'), X = require('xlsx'), sharp = require('sharp');

const [OUT, ...FILES] = process.argv.slice(2);
if (!OUT || FILES.length < 2) { console.error('Aufruf: node bilder-holen.js <Ausgabeordner> <Magic.xlsx> <Mea.xlsx> [Sterne 2025.xlsx]'); process.exit(1); }
fs.mkdirSync(OUT, { recursive: true });

/* ---------- Artikelliste wie beim Import ---------- */
const ctx = { console, Utilities: { formatDate: () => '' } };
vm.createContext(ctx);
['Data.gs', 'Setup.gs', 'Import.gs'].forEach(f => vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'gas', f), 'utf8'), ctx));
const quelle = file => {
  const wb = X.readFile(file);
  return { datei: path.basename(file, '.xlsx'), blaetter: wb.SheetNames.map(n => {
    const rows = X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
    return { name: n, werte: rows, anzeige: rows.map(r => r.map(String)) };
  }) };
};
const listen = FILES.filter(f => !/2025/.test(path.basename(f)));
const plan = ctx.planImport_(listen.map(quelle), { lager: 'Lager', staende: ['Fulda'] }, { head: ['Nr'], rows: [] });
const sterne = plan.artikel.filter(a => !a.zub);

// alte Lieferanten-Nr. (Guru) aus „Sterne 2025“ → Kalkulation: eigene Nr. → Art-Nr.
const alt = {};
const f2025 = FILES.find(f => /2025/.test(path.basename(f)));
if (f2025) {
  const k = X.utils.sheet_to_json(X.readFile(f2025).Sheets['Kalkulation'], { header: 1, defval: '' });
  k.slice(2).forEach(r => { if (r[2] && String(r[2]) !== '0' && r[1]) alt[ctx.nrKey_(r[2])] = String(r[1]); });
}

/* ---------- Netz ---------- */
const sleep = ms => new Promise(r => setTimeout(r, ms));
const UA = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Sternkasse-Bildabgleich' };
async function get(url, bin, ok410) {
  await sleep(400); // höflich zu den Shops
  const r = await fetch(url, { headers: UA });
  if (!r.ok && !(ok410 && r.status === 410)) throw new Error(r.status + ' ' + url); // Guru: 410 = keine Treffer
  const buf = Buffer.from(await r.arrayBuffer());
  if (bin) return buf;
  const latin = /iso-8859-1|latin1|windows-1252/i.test(r.headers.get('content-type') || '') || /guru-shop/.test(url);
  return new TextDecoder(latin ? 'windows-1252' : 'utf-8').decode(buf);
}
const decode = s => s.replace(/&amp;/g, '&').replace(/&uuml;/g, 'ü').replace(/&ouml;/g, 'ö').replace(/&auml;/g, 'ä').replace(/&szlig;/g, 'ß').replace(/&quot;/g, '"');
const og = (html, p) => ((new RegExp('property="og:' + p + '" content="([^"]+)').exec(html) || new RegExp('content="([^"]+)"[^>]*property="og:' + p + '"').exec(html) || [])[1]);

/* ---------- Guru-Shop ---------- */
const FARB = { red: 'rot', blue: 'blau', black: 'schwarz', white: 'weiß', violett: 'lila', mixed: 'mix' };
const norm = s => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/ß/g, 'ss');
const toks = s => norm(s).split(/[^a-z0-9]+/).filter(Boolean).map(t => norm(FARB[t] || t));
async function guruSuche(kw) {
  const html = await get('https://guru-shop.de/advanced_search_result.php?keywords=' + encodeURIComponent(kw), false, true);
  const out = new Map();
  // Titel aus der Adresse: …/Faltbarer-Advents-Leucht-Papierstern-Weihnachtsstern-60-cm-Bilal-rot.html → „… 60 cm - Bilal rot“
  for (const m of html.matchAll(/href="(https:\/\/guru-shop\.de\/products\/de\/(?:[^"]*\/)?([^"\/]+)\.html)"/g)) {
    const slug = decode(m[2]).replace(/\d+$/, '');
    const t = slug.replace(/-(\d{2})-cm-/, ' $1 cm - ').replace(/-/g, ' ');
    if (/Papierstern|Stern/i.test(t) && !/Komplettset|^Set\b|\bSet \d/i.test(t)) out.set(m[1], t);
  }
  return [...out].map(([url, titel]) => ({ url, titel }));
}
const wantToks = a => toks(a.name.replace(/\d{2}\s?cm/ig, '')).filter(t => !/^(mit|und|glitter|gold|perlenbestickt|strahlen|small|neu|\d+)$/.test(t));
function score(a, c) {
  const tail = c.titel.split(' - ').slice(1).join(' ') || c.titel.replace(/^.*Weihnachtsstern\s*/i, '');
  const want = wantToks(a);
  const have = new Set(toks(tail)), haveAll = toks(c.titel);
  const model = want[0];
  let s = 0;
  if (!have.has(model) && !toks(tail).some(t => t.startsWith(model.slice(0, 5)))) return -99;
  s += 10;
  want.slice(1).forEach(t => { s += have.has(t) ? 3 : -1; });
  toks(tail).forEach(t => { if (!want.includes(t) && !/^(\d+|cm|led|neu)$/.test(t)) s -= 2; }); // Wörter im Titel, die nicht bei uns stehen
  const gr = (/(\d{2})\s?cm/.exec(c.titel) || [])[1], ours = (a.groesse || '').slice(0, 2);
  if (gr && ours) s += gr === ours ? 4 : -6;
  if (/small/i.test(a.name) !== /small/i.test(c.titel)) s -= 2;
  if (haveAll.includes('led')) s -= 3;
  return s;
}
async function guru(a) {
  const model = wantToks(a)[0];
  const suche = [model === 'siddharta' ? 'Siddhartha' : model];
  let cand = [];
  for (const kw of suche) cand.push(...await guruSuche(kw));
  const altNr = alt[ctx.nrKey_(a.nr)];
  let perNr = [];
  if (altNr) perNr = await guruSuche(altNr);
  const seen = new Set(); cand = [...perNr, ...cand].filter(c => !seen.has(c.url) && seen.add(c.url));
  cand.forEach(c => { c.score = score(a, c) + (perNr.some(p => p.url === c.url) ? 8 : 0); });
  cand = cand.filter(c => c.score > -50).sort((x, y) => y.score - x.score);
  if (!cand.length) return { status: 'nicht gefunden', alternativen: [] };
  const best = cand[0];
  const page = await get(best.url);
  const bild = og(page, 'image'), artNr = (/Art\.Nr\.:\s*(?:<\/label>\s*)?(\d+)/.exec(page) || [])[1];
  const nrPasst = altNr && (artNr === altNr || (bild || '').includes(altNr));
  const eindeutig = cand.length === 1 || best.score - cand[1].score >= 3;
  return {
    status: nrPasst ? 'sicher (Lieferanten-Nr.)' : eindeutig && best.score >= 14 ? 'wahrscheinlich' : 'bitte prüfen',
    shop: 'guru-shop.de', titel: best.titel, url: best.url, bild, artNr,
    alternativen: cand.slice(1, 4).map(c => ({ titel: c.titel, url: c.url, score: c.score }))
  };
}

/* ---------- mea-living ---------- */
let meaMap = null;
async function meaUrls() {
  if (meaMap) return meaMap;
  const zlib = require('zlib');
  const xml = zlib.gunzipSync(await get('https://mea-living.de/export/sitemap_0.xml.gz', true)).toString('utf8');
  meaMap = {};
  for (const m of xml.matchAll(/<loc>([^<]+stella-(\d{3}))<\/loc>/gi)) meaMap[m[2]] = m[1];
  return meaMap;
}
// Ersatz, wenn die eigene Nummer im Shop fehlt: gleiche Serie/Farbe in anderer Größe bzw. ähnliche Farbe
const MEA_ERSATZ = { '047': ['048', 'gleicher Stern (mint-weiß) in 40 cm'], '066': ['073', 'Dalia pink 40 cm – Farbe „rosa“ bitte prüfen'], '067': ['073', 'Dalia pink 40 cm – Farbe „rosa“ bitte prüfen'] };
async function mea(a) {
  const map = await meaUrls();
  const nr = a.nr.replace(/^[A-Z]+/, '').padStart(3, '0');
  let url = map[nr], status = 'sicher (Nummer)', notiz = '';
  if (!url && MEA_ERSATZ[nr]) {
    const [e, n] = MEA_ERSATZ[nr];
    url = map[e] || (e === '073' ? 'https://mea-living.de/Stella-Dalia-Stern-aus-Pergamentpapier-pink-40-cm' : null);
    status = 'Ersatzbild – bitte prüfen'; notiz = n;
  }
  if (!url) return { status: 'nicht gefunden', alternativen: [] };
  const page = await get(url);
  return { status, notiz, shop: 'mea-living.de', titel: decode(og(page, 'title') || ''), url, bild: og(page, 'image'), alternativen: [] };
}

/* ---------- Ablauf ---------- */
(async () => {
  const res = [];
  for (const a of sterne) {
    process.stdout.write(a.nr.padEnd(5) + a.name.slice(0, 40).padEnd(42));
    let r;
    try { r = a.kategorie === 'Blumen-Sterne' ? await mea(a) : await guru(a); } catch (e) { r = { status: 'Fehler: ' + e.message, alternativen: [] }; }
    if (r.bild) {
      try {
        const img = await get(r.bild, true);
        const unsicher = !/^(sicher|wahrscheinlich)/.test(r.status);
        r.datei = (unsicher ? 'pruefen/' : '') + a.nr + '.jpg';
        fs.mkdirSync(path.join(OUT, 'pruefen'), { recursive: true });
        // quadratisch (die App zeigt Bilder randlos im Quadrat): weißen Rand weg, Stern mittig mit etwas Luft
        const weiss = { r: 255, g: 255, b: 255, alpha: 1 };
        const ohneRand = await sharp(img).flatten({ background: '#ffffff' }).trim({ threshold: 12 }).toBuffer();
        await sharp(ohneRand).resize(560, 560, { fit: 'contain', background: weiss }).extend({ top: 20, bottom: 20, left: 20, right: 20, background: weiss })
          .jpeg({ quality: 82, mozjpeg: true }).toFile(path.join(OUT, r.datei));
        r.kb = Math.round(fs.statSync(path.join(OUT, r.datei)).size / 1024);
      } catch (e) { r.status = 'Bild-Fehler: ' + e.message; }
    }
    console.log(r.status + (r.titel ? '  ← ' + r.titel.replace(/^.* - /, '') : ''));
    res.push({ nr: a.nr, name: a.name, kategorie: a.kategorie, groesse: a.groesse, altNr: alt[ctx.nrKey_(a.nr)] || '', ...r });
  }
  fs.writeFileSync(path.join(OUT, 'ergebnis.json'), JSON.stringify(res, null, 2));
  fs.writeFileSync(path.join(OUT, 'uebersicht.html'), uebersicht(res));
  const z = s => res.filter(r => r.status.startsWith(s)).length;
  console.log(`\nSicher: ${z('sicher')}, wahrscheinlich: ${z('wahrscheinlich')}, prüfen: ${z('bitte') + z('Ersatz')}, nicht gefunden/Fehler: ${res.length - z('sicher') - z('wahrscheinlich') - z('bitte') - z('Ersatz')}`);
})();

function uebersicht(res) {
  const e = s => String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const cls = s => /^sicher/.test(s) ? 'ok' : /^wahrsch/.test(s) ? 'mid' : 'warn';
  return `<!doctype html><meta charset="utf-8"><title>Sternbilder prüfen</title>
<style>body{font:14px system-ui,sans-serif;margin:16px;background:#f4f4f1;color:#1b2320}h1{font-size:20px}
.g{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:12px}.c{background:#fff;border-radius:10px;padding:10px;box-shadow:0 0 0 1px #ddd}
.c img{width:100%;aspect-ratio:1;object-fit:contain;background:#eee;border-radius:6px}.n{font-weight:700}.s{font-size:12px;font-weight:600;padding:2px 6px;border-radius:4px;display:inline-block;margin:4px 0}
.ok{background:#d6efdf;color:#1f5a45}.mid{background:#fdf0d2;color:#7a4c06}.warn{background:#f8d7d3;color:#8c1d18}.t{font-size:12px;color:#5f6a65}a{color:#1f5a45}
.miss{aspect-ratio:1;display:grid;place-items:center;background:#eee;border-radius:6px;color:#999}</style>
<h1>Sternbilder prüfen (${res.length} Sterne)</h1><p>Im Hauptordner liegen nur sichere und wahrscheinliche Treffer. Unsichere Treffer (andere Farbe, Ersatzbild) liegen im Unterordner <b>pruefen</b> – passt eins, die Datei in den Hauptordner verschieben. Falsche Bilder löschen oder durch ein eigenes Foto mit gleichem Dateinamen ersetzen.</p><div class="g">` +
  res.map(r => `<div class="c">${r.datei ? `<img src="${e(r.datei)}" alt="">` : '<div class="miss">kein Bild</div>'}
  <div class="n">${e(r.nr)} · ${e(r.name)}</div><div class="t">${e(r.kategorie)}${r.altNr ? ' · alte Lief.-Nr. ' + e(r.altNr) : ''}</div>
  <span class="s ${cls(r.status)}">${e(r.status)}</span>${r.notiz ? `<div class="t">${e(r.notiz)}</div>` : ''}
  ${r.url ? `<div class="t">Shop: <a href="${e(r.url)}" target="_blank">${e(r.titel)}</a>${r.artNr ? ' (Art.-Nr. ' + e(r.artNr) + ')' : ''}</div>` : ''}
  ${r.alternativen.length ? `<div class="t">Alternativen: ${r.alternativen.map(x => `<a href="${e(x.url)}" target="_blank">${e(x.titel.replace(/^.* - /, ''))}</a>`).join(', ')}</div>` : ''}</div>`).join('') + '</div>';
}
