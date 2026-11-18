/**
 * Sternkasse – Sicherung, Zurücksetzen und Import aus Bestandslisten (ab Saison 2026).
 *
 * Erwartet je Datei ein Blatt mit einer Kopfzeile, in der eine Spalte „Nr.“ heißt.
 * Daneben werden erkannt (Kopfzeile + Zeile darunter, Groß/klein egal):
 *   Bezeichnung (sonst die Spalte direkt rechts von „Nr.“),
 *   „VK“/„Preis“, „Lager“ und je Marktstand eine Spalte mit dessen Namen (z. B. „Fulda“).
 * Alles rechts einer zweiten „Nr.“-Spalte (Strichliste) wird ignoriert.
 * Die Produktkategorie (Magic-Sterne / Blumen-Sterne) kommt aus den Kopfzeilen,
 * sonst aus dem Dateinamen. Nummern, die mit „Z“ beginnen, sind Zubehör.
 *
 * Die Auswertung (parseQuelle_, planImport_) ist bewusst rein, also ohne Zugriff auf
 * Google-Dienste, damit sie außerhalb von Apps Script getestet werden kann.
 */

var KATEGORIEN_ = [
  { name: 'Blumen-Sterne', re: /blumen/i, praefix: 'B' },
  { name: 'Magic-Sterne', re: /magic/i, praefix: 'M' }
];
var SICHERUNG_MAX_ALTER_MS = 60 * 60 * 1000;

/* ---------- Menü ---------- */

function menuBackup() {
  var f = backup_();
  SpreadsheetApp.getUi().alert('Sicherung erstellt:\n\n' + f.getName() + '\n\nSie liegt im selben Drive-Ordner wie dieses Sheet.');
}

function menuReset() {
  var ui = SpreadsheetApp.getUi();
  var last = Number(PropertiesService.getScriptProperties().getProperty('LETZTE_SICHERUNG') || 0);
  if (Date.now() - last > SICHERUNG_MAX_ALTER_MS) {
    ui.alert('Bitte zuerst „Sicherung erstellen“ ausführen.\n\nDie letzte Sicherung fehlt oder ist älter als eine Stunde.');
    return;
  }
  var nA = Math.max(0, ss_().getSheetByName(SH.ART).getLastRow() - 1);
  var nB = Math.max(0, ss_().getSheetByName(SH.BEW).getLastRow() - 1);
  if (ui.alert('Testdaten entfernen',
    'Gelöscht werden ' + nA + ' Zeilen im Blatt „Artikel“ und ' + nB + ' Zeilen im Blatt „Bewegungen“.\n\n' +
    'Erhalten bleiben: Orte, PINs, Bilderordner, Kopfzeilen und eigene Filter-Spalten.\n\nFortfahren?',
    ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  reset_();
  ui.alert('Testdaten entfernt. Als Nächstes: Menü Magic-Sterne → Import aus Bestandslisten …');
}

function menuImport() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Import aus Bestandslisten',
    'Jede Excel-Datei in Google Drive öffnen und über „Datei → Als Google Tabellen speichern“ umwandeln.\n' +
    'Dann hier die Links zu diesen Google-Tabellen einfügen (mehrere mit Leerzeichen getrennt):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var urls = r.getResponseText().split(/\s+/).filter(function (u) { return /^https?:\/\//.test(u); });
  if (!urls.length) { ui.alert('Kein Link erkannt.'); return; }

  setup_();
  var quellen = urls.map(readQuelle_);
  var plan = planImport_(quellen, ortInfo_(), getArtikelZeilen_());
  writeProtokoll_(plan, false);
  if (plan.fehler.length) {
    ui.alert('Import abgebrochen – bitte die Quelldateien korrigieren.\n\n' + plan.fehler.slice(0, 15).join('\n') +
      '\n\nAlle Meldungen stehen im Blatt „' + SH.PROT + '“.');
    return;
  }
  if (ui.alert('Import prüfen', plan.zusammenfassung.join('\n') +
    (plan.warnungen.length ? '\n\nHinweise (' + plan.warnungen.length + ', vollständig im Blatt „' + SH.PROT + '“):\n' +
      plan.warnungen.slice(0, 12).join('\n') + (plan.warnungen.length > 12 ? '\n…' : '') : '') +
    '\n\nImport jetzt ausführen?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var msg = applyImport_(plan, quellen.map(function (q) { return q.datei; }).join(', '));
  writeProtokoll_(plan, true, msg);
  ui.alert(msg.join('\n'));
}

/* ---------- Sicherung und Zurücksetzen ---------- */

function backup_() {
  var ss = ss_(), file = DriveApp.getFileById(ss.getId());
  var parents = file.getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var copy = file.makeCopy(ss.getName() + ' – Sicherung ' + Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd HH.mm'), folder);
  PropertiesService.getScriptProperties().setProperty('LETZTE_SICHERUNG', String(Date.now()));
  return copy;
}

function reset_() {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    [SH.ART, SH.BEW].forEach(function (name) {
      var sh = ss_().getSheetByName(name);
      if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
    });
  } finally {
    lock.releaseLock();
  }
  rebuildBestand_();
}

/* ---------- Lesen (Google-Dienste) ---------- */

function readQuelle_(url) {
  var src = SpreadsheetApp.openByUrl(url);
  return {
    datei: src.getName(),
    blaetter: src.getSheets().map(function (sh) {
      var rng = sh.getDataRange();
      return { name: sh.getName(), werte: rng.getValues(), anzeige: rng.getDisplayValues() };
    })
  };
}

function ortInfo_() {
  return { lager: lagerName_(), staende: getOrte_().filter(function (o) { return !/^lager$/i.test(o.typ); }).map(function (o) { return o.name; }) };
}

/** Artikelblatt als Kopf + Zeilen (Rohwerte), für den Abgleich. */
function getArtikelZeilen_() {
  var sh = ss_().getSheetByName(SH.ART);
  var all = sh.getDataRange().getValues();
  return { head: all[0].map(String), rows: all.slice(1).filter(function (r) { return nr_(r[0]); }) };
}

/* ---------- Auswertung (rein) ---------- */

function normNr_(v) { return String(v === null || v === undefined ? '' : v).replace(/\s+/g, '').toUpperCase(); }
/** Vergleichsschlüssel: führende Nullen zählen nicht („9“ = „09“ = „009“). */
function nrKey_(nr) { return normNr_(nr).replace(/^0+(?=[0-9A-Z])/, ''); }

var HINWEIS_RE_ = /[\s,]*\b(keine\s+nachlieferung|restbestand|rest)\s*[,.]?\s*$/i;

/** Trennt Hinweise wie „Keine Nachlieferung“ oder „Restbestand“ vom Namen. */
function splitHinweis_(name) {
  var s = String(name || '').replace(/\s+/g, ' ').trim(), h = [], m;
  while ((m = HINWEIS_RE_.exec(s))) {
    var t = m[1].replace(/\s+/g, ' ').toLowerCase();
    h.unshift(t === 'keine nachlieferung' ? 'Keine Nachlieferung' : 'Restbestand');
    s = s.slice(0, m.index);
  }
  s = s.replace(/[\s,]+$/, '').replace(/\s+,/g, ',').trim();
  return { name: s, hinweis: h.join(', ') };
}

function zubKurz_(name) { return name.replace(/\s*-\s+/g, ' ').replace(/\s+/g, ' ').trim(); }

function kategorieAus_(text) {
  for (var i = 0; i < KATEGORIEN_.length; i++) if (KATEGORIEN_[i].re.test(text)) return KATEGORIEN_[i].name;
  return '';
}

/**
 * Wertet eine Quelldatei aus.
 * q = { datei, blaetter: [{ name, werte: [[…]], anzeige: [[…]] }] }
 * orte = { lager: 'Lager', staende: ['Fulda', …] }
 */
function parseQuelle_(q, orte) {
  var out = { datei: q.datei, kategorie: '', artikel: [], fehler: [], warnungen: [], notizen: [] };
  var gefunden = 0;
  q.blaetter.forEach(function (b) {
    var W = b.werte, A = b.anzeige, wo = '„' + q.datei + '“ / ' + b.name.trim();
    var hr = -1, cNr = -1;
    for (var r = 0; r < Math.min(W.length, 10) && hr < 0; r++) {
      for (var c = 0; c < A[r].length; c++) if (/^nr\.?$/i.test(String(A[r][c]).trim())) { hr = r; cNr = c; break; }
    }
    if (hr < 0) return; // Blatt ohne Artikelliste
    gefunden++;
    var head = A[hr].map(function (v) { return String(v).trim(); });
    var limit = head.length;
    for (var c2 = cNr + 1; c2 < head.length; c2++) if (/^nr\.?$/i.test(head[c2])) { limit = c2; break; }
    var sub = hr + 1 < A.length && !String(A[hr + 1][cNr]).trim() ? A[hr + 1] : null;
    var start = hr + (sub ? 2 : 1);
    var txt = function (c) { return (head[c] + ' ' + (sub ? String(sub[c]).trim() : '')).trim(); };

    // Kategorie: Kopfbereich zuerst, dann Blatt- und Dateiname
    var kopf = A.slice(0, start).map(function (row) { return row.slice(0, limit).join(' '); }).join(' ');
    var kat = kategorieAus_(kopf) || kategorieAus_(b.name) || kategorieAus_(q.datei);
    if (!kat) out.fehler.push(wo + ': Kategorie nicht erkannt. Kopfzeile oder Dateiname muss „Magic“ oder „Blumen“ enthalten.');
    if (out.kategorie && kat && kat !== out.kategorie) out.fehler.push(wo + ': gemischte Kategorien in einer Datei (' + out.kategorie + ' / ' + kat + ').');
    out.kategorie = out.kategorie || kat;

    var cName = -1, cPreis = -1, cBest = {};
    for (var c3 = cNr + 1; c3 < limit; c3++) if (/bezeichnung|artikel|name/i.test(txt(c3))) { cName = c3; break; }
    if (cName < 0) cName = cNr + 1;
    for (var c4 = cNr + 1; c4 < limit; c4++) {
      if (c4 === cName) continue;
      var t = txt(c4);
      if (cPreis < 0 && /\bvk\b|vk-|preis/i.test(t)) { cPreis = c4; continue; }
      if (/lager/i.test(t)) { if (cBest[orte.lager] === undefined) cBest[orte.lager] = c4; continue; }
      orte.staende.forEach(function (o) {
        if (cBest[o] === undefined && t.toLowerCase().indexOf(o.toLowerCase()) >= 0) cBest[o] = c4;
      });
    }
    if (cPreis < 0) out.fehler.push(wo + ': keine Preisspalte („VK“ oder „Preis“) gefunden.');
    if (!Object.keys(cBest).length) out.fehler.push(wo + ': keine Bestandsspalte („Lager“ oder Name eines Marktstands) gefunden.');
    out.notizen.push(wo + ': Kopfzeile ' + (hr + 1) + ', Daten ab Zeile ' + (start + 1) + '. Spalten: Nr = ' + colLetter_(cNr + 1) +
      ', Name = ' + colLetter_(cName + 1) + (cPreis >= 0 ? ', Preis = ' + colLetter_(cPreis + 1) : '') +
      Object.keys(cBest).map(function (o) { return ', Bestand ' + o + ' = ' + colLetter_(cBest[o] + 1); }).join('') + '. Kategorie: ' + (kat || '?') + '.');

    for (var i = start; i < W.length; i++) {
      var zeile = wo + ' Zeile ' + (i + 1);
      var nr = normNr_(A[i][cNr]);
      var roh = String(W[i][cName] === null || W[i][cName] === undefined ? '' : W[i][cName]).trim();
      if (!nr) {
        var rest = A[i].slice(0, limit).join(' ').replace(/\s+/g, ' ').trim();
        if (rest) out.notizen.push(zeile + ': Zeile ohne Nr. übersprungen („' + rest + '“).');
        continue;
      }
      if (!roh) { out.notizen.push(zeile + ': Nr. ' + nr + ' ohne Bezeichnung übersprungen.'); continue; }
      var zub = /^Z/.test(nr), sp = splitHinweis_(roh);
      var preis = cPreis >= 0 ? W[i][cPreis] : '';
      var a = {
        nr: nr, name: sp.name, hinweis: sp.hinweis, zub: zub, kategorie: zub ? '' : kat, zeile: zeile, quelle: q.datei,
        preis: typeof preis === 'number' && preis > 0 ? Math.round(preis * 100) / 100 : null,
        groesse: zub ? '' : ((/(\d{2})\s?cm/i.exec(sp.name) || [])[1] ? (/(\d{2})\s?cm/i.exec(sp.name))[1] + ' cm' : ''),
        bestand: {}
      };
      Object.keys(cBest).forEach(function (o) {
        var v = W[i][cBest[o]];
        if (v === '' || v === null || v === undefined) return;
        if (typeof v !== 'number' || isNaN(v) || Math.round(v) !== v) { out.warnungen.push(zeile + ' (' + nr + '): Bestand ' + o + ' „' + v + '“ ist keine ganze Zahl – nicht gebucht.'); return; }
        if (v < 0) out.warnungen.push(zeile + ' (' + nr + '): Bestand ' + o + ' ist negativ (' + v + ').');
        a.bestand[o] = v;
      });
      if (a.preis === null) out.warnungen.push(zeile + ' (' + nr + ' ' + a.name + '): kein Preis – in der Kasse gesperrt.');
      if (!zub && !a.groesse) out.warnungen.push(zeile + ' (' + nr + ' ' + a.name + '): keine Größe im Namen – Filter „Größe“ bleibt leer, kein Ständer vorgeschlagen.');
      out.artikel.push(a);
    }
  });
  if (!gefunden) out.fehler.push('„' + q.datei + '“: kein Blatt mit einer Spalte „Nr.“ gefunden.');
  return out;
}

/**
 * Führt die ausgewerteten Quellen zusammen, löst Nummern-Dubletten auf und
 * gleicht mit dem bestehenden Artikelblatt ab. Schreibt nichts.
 */
function planImport_(quellen, orte, bestand) {
  var plan = { artikel: [], fehler: [], warnungen: [], notizen: [], zusammenfassung: [], deaktivieren: [], kategorien: [], orte: orte };
  var parsed = quellen.map(function (q) { return parseQuelle_(q, orte); });
  parsed.forEach(function (p) {
    plan.fehler = plan.fehler.concat(p.fehler);
    plan.warnungen = plan.warnungen.concat(p.warnungen);
    plan.notizen = plan.notizen.concat(p.notizen);
    if (p.kategorie && p.artikel.some(function (a) { return !a.zub; })) {
      if (plan.kategorien.indexOf(p.kategorie) >= 0) plan.fehler.push('Kategorie ' + p.kategorie + ' kommt in mehreren Dateien vor („' + p.datei + '“).');
      else plan.kategorien.push(p.kategorie);
    }
  });

  // 1. Dubletten innerhalb einer Datei: Fehler (welche Zeile gilt, kann nur der Mensch entscheiden)
  parsed.forEach(function (p) {
    var seen = {};
    p.artikel.forEach(function (a) {
      var k = nrKey_(a.nr);
      if (seen[k]) plan.fehler.push('Nr. ' + a.nr + ' doppelt in derselben Datei: ' + seen[k].zeile + ' und ' + a.zeile + '.');
      else seen[k] = a;
    });
  });

  // 2. Über Dateien hinweg
  var byKey = {};
  parsed.forEach(function (p) {
    p.artikel.forEach(function (a) {
      var k = nrKey_(a.nr), alt = byKey[k];
      if (!alt) { byKey[k] = a; plan.artikel.push(a); return; }
      if (alt === a || alt.quelle === a.quelle) return; // gleiche Datei: schon in Schritt 1 gemeldet
      if (a.zub && alt.zub) {
        var gleich = zubKurz_(a.name).toLowerCase() === zubKurz_(alt.name).toLowerCase();
        if (!gleich) { plan.fehler.push('Zubehör-Nr. ' + a.nr + ' steht für zwei verschiedene Artikel: „' + alt.name + '“ (' + alt.zeile + ') und „' + a.name + '“ (' + a.zeile + ').'); return; }
        plan.warnungen.push('Zubehör ' + a.nr + ' steht in mehreren Dateien – verwendet wird ' + alt.zeile + ', ' + a.zeile + ' wird ignoriert.' +
          (JSON.stringify(a.bestand) !== JSON.stringify(alt.bestand) || a.preis !== alt.preis ? ' ACHTUNG: Bestand oder Preis weichen ab.' : ''));
        return;
      }
      if (a.kategorie && alt.kategorie && a.kategorie !== alt.kategorie) {
        // Verschiedene Kategorien, gleiche Nr.: beide bekommen das Kategorie-Präfix, damit Etikett und Kasse eindeutig bleiben
        [alt, a].forEach(function (x) {
          var pre = KATEGORIEN_.filter(function (kk) { return kk.name === x.kategorie; })[0].praefix;
          if (x.nr.indexOf(pre) !== 0 || !/^\d/.test(x.nr.slice(pre.length))) {
            plan.warnungen.push('Nr. ' + x.nr + ' gibt es bei Magic- und Blumen-Sternen – ' + x.kategorie + ' ' + x.name + ' wird als ' + pre + x.nr + ' angelegt (' + x.zeile + ').');
            x.nr = pre + x.nr;
          }
        });
        delete byKey[k];
        byKey[nrKey_(alt.nr)] = alt;
        if (byKey[nrKey_(a.nr)]) { plan.fehler.push('Nr. ' + a.nr + ' ist auch mit Präfix nicht eindeutig (' + a.zeile + ').'); return; }
        byKey[nrKey_(a.nr)] = a; plan.artikel.push(a);
        return;
      }
      plan.fehler.push('Nr. ' + a.nr + ' doppelt: ' + alt.zeile + ' und ' + a.zeile + ' (Stern und Zubehör oder gleiche Kategorie).');
    });
  });

  // 3. Abgleich mit dem Artikelblatt
  var H = {};
  bestand.head.forEach(function (h, i) { H[h] = i; });
  var vorhanden = {};
  bestand.rows.forEach(function (r) { vorhanden[nrKey_(r[H['Nr']])] = r; });
  var neu = 0, akt = 0;
  plan.artikel.forEach(function (a) {
    var r = vorhanden[nrKey_(a.nr)];
    if (r && normNr_(r[H['Nr']]) !== a.nr) {
      plan.warnungen.push('Nr. ' + a.nr + ' (' + a.zeile + ') entspricht dem vorhandenen Artikel ' + r[H['Nr']] + ' – dieser wird aktualisiert und behält seine Nr.');
      a.nr = String(r[H['Nr']]);
    }
    if (r) { akt++; a.vorhanden = true; } else neu++;
  });
  var imp = {};
  plan.artikel.forEach(function (a) { imp[nrKey_(a.nr)] = true; });
  var zubImport = plan.artikel.some(function (a) { return a.zub; });
  bestand.rows.forEach(function (r) {
    var nr = String(r[H['Nr']]), art = String(r[H['Art']]), kat = H['Filter: Kategorie'] !== undefined ? String(r[H['Filter: Kategorie']]) : '';
    if (imp[nrKey_(nr)] || /^nein$/i.test(String(r[H['Aktiv']]).trim())) return;
    var betroffen = /zub/i.test(art) ? zubImport : (kat ? plan.kategorien.indexOf(kat) >= 0 : plan.kategorien.length > 0);
    if (betroffen) plan.deaktivieren.push(nr);
  });
  if (plan.deaktivieren.length) plan.warnungen.push('Nicht mehr in den Listen, wird ausgeblendet (Aktiv = nein): ' + plan.deaktivieren.join(', ') + '.');

  var stern = plan.artikel.filter(function (a) { return !a.zub; });
  plan.zusammenfassung.push('Artikel: ' + plan.artikel.length + ' (neu ' + neu + ', aktualisiert ' + akt + ').');
  plan.kategorien.forEach(function (k) {
    plan.zusammenfassung.push('  ' + k + ': ' + stern.filter(function (a) { return a.kategorie === k; }).length + ' Sterne');
  });
  plan.zusammenfassung.push('  Zubehör: ' + plan.artikel.filter(function (a) { return a.zub; }).length);
  [orte.lager].concat(orte.staende).forEach(function (o) {
    var n = plan.artikel.filter(function (a) { return a.bestand[o] !== undefined; });
    if (n.length) plan.zusammenfassung.push('Bestand ' + o + ': ' + n.length + ' Artikel, ' + n.reduce(function (s, a) { return s + a.bestand[o]; }, 0) + ' Stück (als Inventur).');
  });
  return plan;
}

/**
 * Zubehör-Vorschlag für einen Stern:
 *   E14 (Kabel, LED) bei allen Sternen,
 *   E27 nur bei Blumen-Sternen (Mea) – dort als Empfehlung zuerst,
 *   Ständer nur ab 60 cm, sonstiges Zubehör immer.
 * zub = [{ nr, name }]
 */
function zubehoerFuer_(name, kategorie, zub) {
  var size = Number((/(\d{2})\s?cm/i.exec(name) || [])[1] || 0);
  var blumen = kategorie === 'Blumen-Sterne';
  var e27 = [], rest = [];
  zub.forEach(function (z) {
    if (/E\s?27/i.test(z.name)) { if (blumen) e27.push(z.nr); return; }
    if (/Ständer/i.test(z.name) && size < 60) return;
    rest.push(z.nr);
  });
  return e27.concat(rest);
}

/* ---------- Schreiben ---------- */

function applyImport_(plan, quelleName) {
  var sh = ss_().getSheetByName(SH.ART);
  var all = sh.getDataRange().getValues();
  var head = all[0].map(String), H = {};
  head.forEach(function (h, i) { H[h] = i; });
  var data = all.slice(1).filter(function (row) { return nr_(row[H['Nr']]); });
  var byKey = {};
  data.forEach(function (row) { byKey[nrKey_(row[H['Nr']])] = row; });
  var orte = getOrte_();
  var set = function (row, h, v) { if (H[h] !== undefined) row[H[h]] = v; };
  var leer = function (row, h) { return H[h] !== undefined && (row[H[h]] === '' || row[H[h]] === null); };

  plan.artikel.forEach(function (a) {
    var row = byKey[nrKey_(a.nr)];
    if (!row) { row = head.map(function () { return ''; }); data.push(row); byKey[nrKey_(a.nr)] = row; set(row, 'Nr', a.nr); }
    set(row, 'Name', a.name);
    if (leer(row, 'Kurzname')) set(row, 'Kurzname', a.zub ? zubKurz_(a.name) : kurzname_(a.name, false));
    set(row, 'Art', a.zub ? 'Zubehör' : 'Stern');
    set(row, 'Preis', a.preis === null ? '' : a.preis);
    set(row, 'Hinweis', a.hinweis);
    set(row, 'Filter: Kategorie', a.kategorie);
    if (/^nein$/i.test(String(row[H['Aktiv']]).trim())) set(row, 'Aktiv', '');
    if (!a.zub) {
      if (leer(row, 'Filter: Größe') && a.groesse) set(row, 'Filter: Größe', a.groesse);
      if (leer(row, 'Filter: Farbe')) set(row, 'Filter: Farbe', farben_(kurzname_(a.name, false)) || 'Sonstige');
    }
    orte.forEach(function (o) {
      var h = 'Min: ' + o.name;
      if (leer(row, h)) set(row, h, /^lager$/i.test(o.typ) ? (a.zub ? 20 : 5) : (a.zub ? 10 : 3));
    });
  });
  plan.deaktivieren.forEach(function (nr) { var row = byKey[nrKey_(nr)]; if (row) set(row, 'Aktiv', 'nein'); });

  // Zubehör-Vorbelegung für Sterne ohne Eintrag (siehe zubehoerFuer_)
  var zubRows = data.filter(function (row) { return row[H['Art']] === 'Zubehör' && !/^nein$/i.test(String(row[H['Aktiv']]).trim()); })
    .map(function (z) { return { nr: String(z[H['Nr']]), name: String(z[H['Name']]) }; });
  data.forEach(function (row) {
    if (row[H['Art']] !== 'Stern' || row[H['Zubehör']]) return;
    row[H['Zubehör']] = zubehoerFuer_(String(row[H['Name']]), String(row[H['Filter: Kategorie']] || ''), zubRows).join(', ');
  });

  data.sort(function (a, b) {
    return (a[H['Art']] === b[H['Art']] ? 0 : a[H['Art']] === 'Stern' ? -1 : 1) ||
      String(a[H['Filter: Kategorie']] || '').localeCompare(String(b[H['Filter: Kategorie']] || '')) ||
      String(a[H['Nr']]).localeCompare(String(b[H['Nr']]), 'de', { numeric: true });
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, head.length).clearContent();
  if (data.length) {
    sh.getRange(2, H['Nr'] + 1, data.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, data.length, head.length).setValues(data.map(function (row) { row[H['Nr']] = String(row[H['Nr']]); return row; }));
  }

  // Bestände als Inventur je Ort (gebucht wird die Differenz zum aktuellen Bestand)
  var msg = ['Import fertig.', ''].concat(plan.zusammenfassung), stamp = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyyMMdd-HHmmss');
  [plan.orte.lager].concat(plan.orte.staende).forEach(function (o) {
    var lines = plan.artikel.filter(function (a) { return a.bestand[o] !== undefined; }).map(function (a) { return { nr: a.nr, menge: a.bestand[o] }; });
    if (!lines.length) return;
    var res = book_([{ id: 'IMPORT-' + stamp + '-' + o, art: 'Inventur', ort: o, geraet: 'Import', notiz: 'Bestand aus ' + quelleName, lines: lines }], 'inhaber');
    msg.push(res[0].ok ? 'Bestand ' + o + ' gebucht.' : 'Bestand ' + o + ' NICHT gebucht: ' + res[0].fehler);
  });
  rebuildBestand_();
  msg.push('', 'Bitte im Blatt „Artikel“ prüfen: Zubehör je Stern und Mindestbestände sind vorbelegt.');
  return msg;
}

function writeProtokoll_(plan, ausgefuehrt, msg) {
  var sh = ss_().getSheetByName(SH.PROT) || ss_().insertSheet(SH.PROT);
  sh.clear();
  var rows = [['Importprotokoll ' + Utilities.formatDate(new Date(), 'Europe/Berlin', 'dd.MM.yyyy HH:mm') + (ausgefuehrt ? ' – ausgeführt' : ' – nur geprüft')], ['']];
  var add = function (titel, list) { if (list.length) { rows.push([titel]); list.forEach(function (l) { rows.push(['   ' + l]); }); rows.push(['']); } };
  add('FEHLER (Import nicht möglich)', plan.fehler);
  add('Zusammenfassung', plan.zusammenfassung);
  add('Hinweise', plan.warnungen);
  add('Erkannte Struktur', plan.notizen);
  if (msg) add('Ergebnis', msg);
  sh.getRange(1, 1, rows.length, 1).setValues(rows);
  sh.getRange('A1').setFontWeight('bold').setFontSize(13);
  sh.setColumnWidth(1, 1000);
}
