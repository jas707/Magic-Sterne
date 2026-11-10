/**
 * Sternkasse – Datenhaltung im Google Sheet.
 *
 * Grundsatz: Das Blatt „Bewegungen“ ist die einzige Quelle der Wahrheit.
 * Jeder Bestand ergibt sich aus der Summe der Mengen je Artikel und Ort
 * (ohne Zeilen der Art „Bestellung“, die nur vormerken). Es wird nie ein
 * Bestand überschrieben, nur Bewegungen angehängt.
 */

var SH = { ART: 'Artikel', ORTE: 'Orte', BEW: 'Bewegungen', BEST: 'Bestand', INFO: 'Anleitung' };

var ART_HEAD = ['Nr', 'Name', 'Kurzname', 'Art', 'Preis', 'Zubehör', 'Aktiv', 'Bild',
  'Lieferanten-Nr', 'EK', 'Filter: Größe', 'Filter: Farbe'];
var ORTE_HEAD = ['Ort', 'Typ', 'Aktiv'];
var BEW_HEAD = ['Zeitpunkt', 'Vorgang', 'Art', 'Nr', 'Ort', 'Menge', 'Gerät', 'Notiz', 'Erfasst'];

/** Arten von Bewegungen und wer sie buchen darf. */
var ARTEN = {
  'Verkauf':        { rolle: 'verkauf' },
  'Rückgabe':       { rolle: 'verkauf' },
  'Storno':         { rolle: 'verkauf' },
  'Lieferung':      { rolle: 'inhaber' },  // Lager → Stand
  'Rücklieferung':  { rolle: 'inhaber' },  // Stand → Lager
  'Wareneingang':   { rolle: 'inhaber' },  // Lieferant → Lager
  'Bestellung':     { rolle: 'inhaber' },  // nur vorgemerkt, kein Bestand
  'Retoure':        { rolle: 'inhaber' },  // Lager → Lieferant
  'Korrektur':      { rolle: 'inhaber' },
  'Inventur':       { rolle: 'inhaber' }
};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

/** Liest ein Blatt als Tabelle: Kopfzeile + Zeilen, Spaltenzugriff über den Namen. */
function readTable_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Blatt „' + name + '“ fehlt. Bitte Menü Magic-Sterne → Einrichten ausführen.');
  var values = sh.getDataRange().getValues();
  var head = values.shift().map(function (h) { return String(h).trim(); });
  var idx = {};
  head.forEach(function (h, i) { idx[h] = i; });
  var rows = values.filter(function (r) { return r.some(function (v) { return v !== '' && v !== null; }); });
  return { sheet: sh, head: head, idx: idx, rows: rows };
}

function col_(t, name) {
  var i = t.idx[name];
  return function (row) { return i === undefined ? '' : row[i]; };
}

function nr_(v) { return String(v === null || v === undefined ? '' : v).trim(); }

function truthy_(v) {
  if (v === '' || v === null || v === undefined) return true; // leer = aktiv
  if (v === true) return true;
  return /^(ja|j|x|1|wahr|true)$/i.test(String(v).trim());
}

function getOrte_() {
  var t = readTable_(SH.ORTE);
  var ort = col_(t, 'Ort'), typ = col_(t, 'Typ'), akt = col_(t, 'Aktiv');
  return t.rows.map(function (r) {
    return { name: nr_(ort(r)), typ: nr_(typ(r)) || 'Stand', aktiv: truthy_(akt(r)) };
  }).filter(function (o) { return o.name; });
}

function lagerName_() {
  var l = getOrte_().filter(function (o) { return /^lager$/i.test(o.typ); })[0];
  if (!l) throw new Error('Im Blatt „Orte“ fehlt ein Ort mit Typ „Lager“.');
  return l.name;
}

function getArtikel_() {
  var t = readTable_(SH.ART);
  var c = {};
  t.head.forEach(function (h) { c[h] = col_(t, h); });
  var minCols = t.head.filter(function (h) { return /^Min:\s*/.test(h); });
  var filterCols = t.head.filter(function (h) { return /^Filter:\s*/.test(h); });
  return t.rows.map(function (r) {
    var preis = c['Preis'](r);
    var a = {
      nr: nr_(c['Nr'](r)),
      name: nr_(c['Name'](r)),
      kurz: nr_(c['Kurzname'](r)) || nr_(c['Name'](r)),
      art: /zub/i.test(nr_(c['Art'](r))) ? 'Zubehör' : 'Stern',
      preis: (preis === '' || preis === null || isNaN(Number(preis))) ? null : Math.round(Number(preis) * 100) / 100,
      zubehoer: nr_(c['Zubehör'](r)).split(/[,;\s]+/).filter(String),
      aktiv: truthy_(c['Aktiv'](r)),
      bild: nr_(c['Bild'](r)),
      min: {},
      filter: {}
    };
    minCols.forEach(function (h) {
      var v = Number(c[h](r));
      a.min[h.replace(/^Min:\s*/, '')] = isNaN(v) ? 0 : v;
    });
    filterCols.forEach(function (h) {
      var v = nr_(c[h](r));
      if (v) a.filter[h.replace(/^Filter:\s*/, '')] = v.split(/\s*[,\/]\s*/).filter(String);
    });
    return a;
  }).filter(function (a) { return a.nr; });
}

/** Bestand je Artikel und Ort sowie offene Bestellmengen (Bestellung − Wareneingang). */
function computeStock_() {
  var t = readTable_(SH.BEW);
  var art = col_(t, 'Art'), nr = col_(t, 'Nr'), ort = col_(t, 'Ort'), menge = col_(t, 'Menge');
  var stock = {}, bestellt = {}, eingang = {};
  t.rows.forEach(function (r) {
    var a = nr_(art(r)), n = nr_(nr(r)), o = nr_(ort(r)), m = Number(menge(r)) || 0;
    if (!n) return;
    if (a === 'Bestellung') { bestellt[n] = (bestellt[n] || 0) + m; return; }
    if (a === 'Wareneingang') eingang[n] = (eingang[n] || 0) + m;
    stock[n] = stock[n] || {};
    stock[n][o] = (stock[n][o] || 0) + m;
  });
  var offen = {};
  Object.keys(bestellt).forEach(function (n) {
    var rest = bestellt[n] - (eingang[n] || 0);
    if (rest > 0) offen[n] = rest;
  });
  return { stock: stock, offen: offen, rows: t.rows.length };
}

/** Alle bereits gebuchten Vorgangs-IDs (für doppelte Sendungen nach Offline-Phasen). */
function knownVorgaenge_() {
  var t = readTable_(SH.BEW);
  var v = col_(t, 'Vorgang'), seen = {};
  t.rows.forEach(function (r) { seen[nr_(v(r))] = true; });
  return seen;
}

/**
 * Wandelt einen Vorgang in Bewegungszeilen um.
 * op = { id, art, ort, von, nach, zeit, geraet, notiz, ref, lines:[{nr, menge}] }
 */
function rowsFor_(op, ctx) {
  var lager = ctx.lager;
  var zeit = op.zeit ? new Date(op.zeit) : new Date();
  if (isNaN(zeit.getTime())) zeit = new Date();
  var now = new Date();
  var out = [];
  function push(art, nr, ort, menge, notiz) {
    if (!menge) return;
    out.push([zeit, op.id, art, nr, ort, menge, op.geraet || '', notiz || op.notiz || '', now]);
  }
  function needOrt(o) {
    if (!o || !ctx.orte[o]) throw new Error('Unbekannter Ort „' + o + '“.');
    return o;
  }
  function needArt(n) {
    if (!ctx.artikel[n]) throw new Error('Unbekannte Artikelnummer „' + n + '“.');
    return n;
  }
  var lines = (op.lines || []).map(function (l) { return { nr: nr_(l.nr), menge: Math.round(Number(l.menge) || 0) }; })
    .filter(function (l) { return l.nr && l.menge; });

  switch (op.art) {
    case 'Verkauf':
      needOrt(op.ort);
      lines.forEach(function (l) { push('Verkauf', needArt(l.nr), op.ort, -Math.abs(l.menge)); });
      break;
    case 'Rückgabe':
      needOrt(op.ort);
      lines.forEach(function (l) { push('Rückgabe', needArt(l.nr), op.ort, Math.abs(l.menge)); });
      break;
    case 'Lieferung':
    case 'Rücklieferung':
      var von = needOrt(op.von || (op.art === 'Lieferung' ? lager : op.ort));
      var nach = needOrt(op.nach || (op.art === 'Lieferung' ? op.ort : lager));
      if (von === nach) throw new Error('Von- und Nach-Ort sind gleich.');
      lines.forEach(function (l) {
        push(op.art, needArt(l.nr), von, -Math.abs(l.menge), 'nach ' + nach);
        push(op.art, l.nr, nach, Math.abs(l.menge), 'von ' + von);
      });
      break;
    case 'Wareneingang':
    case 'Bestellung':
      lines.forEach(function (l) { push(op.art, needArt(l.nr), lager, Math.abs(l.menge)); });
      break;
    case 'Retoure':
      lines.forEach(function (l) { push('Retoure', needArt(l.nr), lager, -Math.abs(l.menge)); });
      break;
    case 'Korrektur':
      needOrt(op.ort);
      lines.forEach(function (l) { push('Korrektur', needArt(l.nr), op.ort, l.menge); });
      break;
    case 'Inventur': // Mengen = gezählter Bestand; gebucht wird die Differenz
      needOrt(op.ort);
      lines.forEach(function (l) {
        var ist = ((ctx.stock[l.nr] || {})[op.ort]) || 0;
        push('Inventur', needArt(l.nr), op.ort, Math.abs(l.menge) - ist, 'gezählt ' + Math.abs(l.menge));
      });
      // auch Nullzählungen zulassen: lines mit menge 0 wurden oben gefiltert
      (op.lines || []).forEach(function (l) {
        if (Number(l.menge) === 0 && ctx.artikel[nr_(l.nr)]) {
          var ist0 = ((ctx.stock[nr_(l.nr)] || {})[op.ort]) || 0;
          push('Inventur', nr_(l.nr), op.ort, -ist0, 'gezählt 0');
        }
      });
      break;
    case 'Storno':
      var ref = nr_(op.ref);
      if (!ref) throw new Error('Storno ohne Bezug.');
      if (ctx.storniert[ref]) return []; // schon storniert → nichts tun
      var orig = ctx.byVorgang[ref];
      if (!orig || !orig.length) throw new Error('Vorgang „' + ref + '“ nicht gefunden.');
      if (ctx.rolle !== 'inhaber' && orig.some(function (r) { return r.art !== 'Verkauf' && r.art !== 'Rückgabe'; })) {
        throw new Error('Diesen Vorgang darf nur der Inhaber stornieren.');
      }
      orig.forEach(function (r) { push('Storno', r.nr, r.ort, -r.menge, 'Storno von ' + ref); });
      break;
    default:
      throw new Error('Unbekannte Buchungsart „' + op.art + '“.');
  }
  return out;
}

/** Bucht mehrere Vorgänge; bereits bekannte Vorgangs-IDs werden übersprungen. */
function book_(ops, rolle) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var t = readTable_(SH.BEW);
    var cV = col_(t, 'Vorgang'), cA = col_(t, 'Art'), cN = col_(t, 'Nr'), cO = col_(t, 'Ort'), cM = col_(t, 'Menge'), cNo = col_(t, 'Notiz');
    var known = {}, byVorgang = {}, storniert = {};
    t.rows.forEach(function (r) {
      var v = nr_(cV(r));
      known[v] = true;
      (byVorgang[v] = byVorgang[v] || []).push({ art: nr_(cA(r)), nr: nr_(cN(r)), ort: nr_(cO(r)), menge: Number(cM(r)) || 0 });
      var m = /^Storno von (.+)$/.exec(nr_(cNo(r)));
      if (m) storniert[m[1]] = true;
    });
    var artikel = {};
    getArtikel_().forEach(function (a) { artikel[a.nr] = a; });
    var orte = {};
    getOrte_().forEach(function (o) { orte[o.name] = o; });
    var ctx = { lager: lagerName_(), artikel: artikel, orte: orte, byVorgang: byVorgang, storniert: storniert,
      stock: computeStock_().stock, rolle: rolle };

    var results = [], newRows = [];
    ops.forEach(function (op) {
      var id = nr_(op.id);
      if (!id) { results.push({ id: id, ok: false, fehler: 'Vorgang ohne ID.' }); return; }
      if (known[id]) { results.push({ id: id, ok: true, doppelt: true }); return; }
      var def = ARTEN[op.art];
      if (!def) { results.push({ id: id, ok: false, fehler: 'Unbekannte Buchungsart.' }); return; }
      if (def.rolle === 'inhaber' && rolle !== 'inhaber') {
        results.push({ id: id, ok: false, fehler: 'Nur der Inhaber darf „' + op.art + '“ buchen.' }); return;
      }
      try {
        var rows = rowsFor_(op, ctx);
        rows.forEach(function (r) { newRows.push(r); });
        known[id] = true;
        // im selben Paket gebuchte Vorgänge sind für ein folgendes Storno sofort auffindbar
        ctx.byVorgang[id] = rows.map(function (r) { return { art: r[2], nr: r[3], ort: r[4], menge: r[5] }; });
        if (op.art === 'Storno') ctx.storniert[nr_(op.ref)] = true;
        results.push({ id: id, ok: true, zeilen: rows.length });
      } catch (e) {
        results.push({ id: id, ok: false, fehler: e.message });
      }
    });
    if (newRows.length) {
      var sh = t.sheet, start = sh.getLastRow() + 1;
      var rng = sh.getRange(start, 1, newRows.length, BEW_HEAD.length);
      sh.getRange(start, 4, newRows.length, 2).setNumberFormat('@'); // Nr und Ort als Text („09“ bleibt „09“)
      rng.setValues(newRows.map(function (r) { return r.map(function (v, i) { return (i === 3 || i === 4) ? String(v) : v; }); }));
    }
    return results;
  } finally {
    lock.releaseLock();
  }
}
