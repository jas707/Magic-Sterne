/**
 * Prüft den Import (gas/Import.gs) lokal gegen echte Excel-Dateien – ohne Google.
 *
 *   cd tools && npm install && node test-import.js "<Datei1.xlsx>" "<Datei2.xlsx>" …
 *
 * Zeigt, was der Import im Sheet anlegen würde (Artikel, Bestände, Hinweise, Fehler),
 * und führt danach Selbsttests mit konstruierten Sonderfällen aus.
 */
const fs = require('fs'), path = require('path'), vm = require('vm'), X = require('xlsx');

const GAS = path.join(__dirname, '..', 'gas');
const ctx = { console, Utilities: { formatDate: () => '20261118-000000' } };
vm.createContext(ctx);
['Data.gs', 'Setup.gs', 'Import.gs'].forEach(f => vm.runInContext(fs.readFileSync(path.join(GAS, f), 'utf8'), ctx, { filename: f }));

const ORTE = [{ name: 'Lager', typ: 'Lager', aktiv: true }, { name: 'Fulda', typ: 'Stand', aktiv: true }];
const orte = { lager: 'Lager', staende: ['Fulda'] };

function quelleAusXlsx(file) {
  const wb = X.readFile(file);
  return {
    datei: path.basename(file).replace(/\.xlsx$/i, ''),
    blaetter: wb.SheetNames.map(n => {
      const ws = wb.Sheets[n], rg = X.utils.decode_range(ws['!ref']), werte = [], anzeige = [];
      for (let r = 0; r <= rg.e.r; r++) {
        const w = [], a = [];
        for (let c = 0; c <= rg.e.c; c++) {
          const cell = ws[X.utils.encode_cell({ r, c })];
          w.push(cell ? cell.v : '');
          // Google zeigt Zahlen im Format „Allgemein“ ohne Tausender-/Währungszeichen; für Nr. genügt String(v)
          a.push(cell ? (cell.t === 'n' ? String(cell.v) : String(cell.v)) : '');
        }
        werte.push(w); anzeige.push(a);
      }
      return { name: n, werte, anzeige };
    })
  };
}

/** Simuliertes Artikelblatt, damit applyImport_ wirklich läuft. */
function fakeSheet(head) {
  let rows = [head.slice()];
  const sh = {
    getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
    getLastRow: () => rows.length,
    getRange: (r, c, nr, nc) => ({
      clearContent() { for (let i = r - 1; i < r - 1 + nr; i++) for (let j = c - 1; j < c - 1 + nc; j++) if (rows[i]) rows[i][j] = ''; return this; },
      setNumberFormat() { return this; },
      setValues(v) { v.forEach((row, i) => { rows[r - 1 + i] = rows[r - 1 + i] || head.map(() => ''); row.forEach((x, j) => { rows[r - 1 + i][c - 1 + j] = x; }); }); return this; }
    }),
    rows: () => rows
  };
  return sh;
}
function runApply(plan, sheet) {
  const booked = [];
  ctx.ss_ = () => ({ getSheetByName: () => sheet });
  ctx.getOrte_ = () => ORTE;
  ctx.book_ = ops => { booked.push(...ops); return ops.map(o => ({ id: o.id, ok: true })); };
  ctx.rebuildBestand_ = () => {};
  const msg = ctx.applyImport_(plan, 'Test');
  return { msg, booked };
}
const HEAD = ['Nr', 'Name', 'Kurzname', 'Art', 'Preis', 'Zubehör', 'Aktiv', 'Bild', 'Lieferanten-Nr', 'EK', 'Filter: Kategorie', 'Filter: Größe', 'Filter: Farbe', 'Hinweis', 'Min: Lager', 'Min: Fulda'];

/* ---------- 1. Echte Dateien ---------- */
const files = process.argv.slice(2);
if (files.length) {
  const quellen = files.map(quelleAusXlsx);
  const plan = ctx.planImport_(quellen, orte, { head: HEAD, rows: [] });
  const show = (t, l) => { if (l.length) { console.log('\n## ' + t); l.forEach(x => console.log('  ' + x)); } };
  show('FEHLER', plan.fehler); show('Zusammenfassung', plan.zusammenfassung); show('Hinweise', plan.warnungen); show('Erkannte Struktur', plan.notizen);
  if (!plan.fehler.length) {
    const sh = fakeSheet(HEAD), { msg, booked } = runApply(plan, sh);
    console.log('\n## Artikelblatt nach Import');
    const H = Object.fromEntries(HEAD.map((h, i) => [h, i]));
    sh.rows().slice(1).forEach(r => console.log('  ' + [r[H.Nr], r[H.Art], r[H['Filter: Kategorie']] || '-', r[H.Kurzname], r[H.Preis], r[H['Filter: Größe']] || '-', r[H['Filter: Farbe']] || '-', r[H.Hinweis] ? '[' + r[H.Hinweis] + ']' : '', 'Zub: ' + (r[H['Zubehör']] || '-')].join(' | ')));
    console.log('\n## Buchungen');
    booked.forEach(b => console.log('  ' + b.id + ' ' + b.ort + ': ' + b.lines.length + ' Zeilen, ' + b.lines.reduce((s, l) => s + l.menge, 0) + ' Stück'));
    show('Meldung', msg);
  }
}

/* ---------- 2. Selbsttests ---------- */
let fail = 0;
const ok = (cond, name) => { console.log((cond ? '  ok   ' : '  FAIL ') + name); if (!cond) fail++; };
const mk = (datei, kopf2, rows) => ({ datei, blaetter: [{ name: 'Tabelle1', werte: [['Nr.', 'Bezeichnung', 'Lager', 'Fulda', 'VK', 'Nr.'], kopf2, ...rows], anzeige: [['Nr.', 'Bezeichnung', 'Lager', 'Fulda', 'VK', 'Nr.'], kopf2, ...rows].map(r => r.map(String)) }] });
console.log('\n## Selbsttests');

let sp = ctx.splitHinweis_('Priamos 7, schwarz, 60 cm,    Restbestand');
ok(sp.name === 'Priamos 7, schwarz, 60 cm' && sp.hinweis === 'Restbestand', 'Hinweis „Restbestand“ abgetrennt');
sp = ctx.splitHinweis_('Siddharta rot,                               Keine Nachlieferung');
ok(sp.name === 'Siddharta rot' && sp.hinweis === 'Keine Nachlieferung', 'Hinweis „Keine Nachlieferung“ abgetrennt');
ok(ctx.splitHinweis_('Artemis beige 60 cm                  Rest').hinweis === 'Restbestand', '„Rest“ → Restbestand');
ok(ctx.splitHinweis_('Forest grün 60 cm').hinweis === '', 'Name ohne Hinweis bleibt');
ok(ctx.nrKey_('009') === ctx.nrKey_('9') && ctx.nrKey_('ZB  01') === 'ZB01', 'Nummern-Schlüssel');

// Kollision zwischen Kategorien (9 vs 009)
let p = ctx.planImport_([
  mk('Magic Sterne', ['', '', '', 'Markt', '', ''], [[9, 'Demian 60 cm', 5, 2, 14.9, 9]]),
  mk('Mea Liste', ['', 'Blumen-Sterne', 'Lager', 'Fulda', '', ''], [['009', 'Dalia 60 cm', 3, 1, 45, '009']])
], orte, { head: HEAD, rows: [] });
ok(!p.fehler.length && p.artikel.map(a => a.nr).sort().join() === 'B009,M9', 'Kategorie-Kollision → Präfix B/M  (' + p.artikel.map(a => a.nr) + ')');

// Dublette in derselben Datei → Fehler, genau einmal gemeldet
p = ctx.planImport_([mk('Magic', ['', '', '', '', '', ''], [[9, 'A 60 cm', 1, 1, 10, 9], ['09', 'B 60 cm', 1, 1, 10, '09']])], orte, { head: HEAD, rows: [] });
ok(p.fehler.length === 1 && /doppelt in derselben Datei/.test(p.fehler[0]), 'Dublette in einer Datei → 1 Fehler');

// Zubehör in beiden Dateien: gleich → Hinweis, verschieden → Fehler
p = ctx.planImport_([
  mk('Magic', ['', '', '', '', '', ''], [['ZB 01', 'LED- E27- 12 Watt', 1, 1, 4.5, 'ZB 01']]),
  mk('Blumen', ['', '', '', '', '', ''], [['ZB01', 'LED  E27 12 Watt', 1, 1, 4.5, 'ZB01']])
], orte, { head: HEAD, rows: [] });
ok(!p.fehler.length && p.artikel.length === 1, 'gleiches Zubehör in zwei Dateien → einmal angelegt');
p = ctx.planImport_([
  mk('Magic', ['', '', '', '', '', ''], [['ZB 01', 'LED E27', 1, 1, 4.5, '']]),
  mk('Blumen', ['', '', '', '', '', ''], [['ZB01', 'Kabel E14', 1, 1, 7, '']])
], orte, { head: HEAD, rows: [] });
ok(p.fehler.length === 1, 'verschiedenes Zubehör mit gleicher Nr. → Fehler');

// Kategorie nicht erkennbar
p = ctx.planImport_([mk('Liste', ['', '', '', '', '', ''], [[1, 'X 60 cm', 1, 1, 10, 1]])], orte, { head: HEAD, rows: [] });
ok(p.fehler.some(f => /Kategorie nicht erkannt/.test(f)), 'unbekannte Kategorie → Fehler');

// Re-Import: vorhandener Artikel „09“ wird von „9“ getroffen, behält Nr., Zubehör und Min; fehlender wird deaktiviert
const vorh = { head: HEAD, rows: [
  ['09', 'Demian alt', 'Demian', 'Stern', 13.9, 'Z4', '', 'bild-id', '', '', 'Magic-Sterne', '60 cm', 'Natur', '', 7, 4],
  ['17', 'Menor natur', 'Menor natur', 'Stern', 14.9, '', '', '', '', '', 'Magic-Sterne', '60 cm', 'Natur', '', 5, 3]
] };
p = ctx.planImport_([mk('Magic', ['', '', '', '', '', ''], [[9, 'Demian natur 60 cm', 5, 2, 14.9, 9]])], orte, vorh);
ok(p.artikel[0].nr === '09' && p.deaktivieren.join() === '17', 'Re-Import: 9 trifft 09, 17 wird ausgeblendet');
const sh = fakeSheet(HEAD); sh.getRange(2, 1, 2, HEAD.length).setValues(vorh.rows.map(r => r.slice()));
runApply(p, sh);
const r09 = sh.rows().find(r => r[0] === '09'), r17 = sh.rows().find(r => r[0] === '17');
ok(r09 && r09[4] === 14.9 && r09[5] === 'Z4' && r09[7] === 'bild-id' && r09[14] === 7, 'Re-Import: Preis neu, Zubehör/Bild/Min bleiben');
ok(r17 && r17[6] === 'nein', 'Re-Import: fehlender Artikel Aktiv = nein');

// Zubehör-Vorschlag: E14 bei allen, E27 nur bei Blumen-Sternen (zuerst), Ständer ab 60 cm
const Z = [{ nr: 'ZB01', name: 'LED- E27- 12 Watt' }, { nr: 'ZB02', name: 'LED- E14- 8 Watt' }, { nr: 'ZB03', name: 'Kabel  E27 schwarz' },
  { nr: 'ZB04', name: 'Ständer' }, { nr: 'ZB05', name: 'Kabel  E14  weiß' }, { nr: 'ZB06', name: 'LED- E14- 4 Watt dimmbar' }];
ok(ctx.zubehoerFuer_('Demian natur 60 cm', 'Magic-Sterne', Z).join() === 'ZB02,ZB04,ZB05,ZB06', 'Magic 60 cm: nur E14 + Ständer');
ok(ctx.zubehoerFuer_('Phoenix natur, 40 cm', 'Magic-Sterne', Z).join() === 'ZB02,ZB05,ZB06', 'Magic 40 cm: nur E14');
ok(ctx.zubehoerFuer_('Dalia weiß 60 cm', 'Blumen-Sterne', Z).join() === 'ZB01,ZB03,ZB02,ZB04,ZB05,ZB06', 'Blumen 60 cm: E27 zuerst, dann E14 + Ständer');

console.log(fail ?`\n${fail} Test(s) fehlgeschlagen.` : '\nAlle Selbsttests bestanden.');
process.exit(fail ? 1 : 0);
