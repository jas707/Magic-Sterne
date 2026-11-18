/**
 * Sternkasse – Einrichtung, Menü, Bestandsübersicht. Import siehe Import.gs.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Magic-Sterne')
    .addItem('Einrichten / reparieren', 'menuSetup')
    .addItem('PINs festlegen …', 'menuPins')
    .addSeparator()
    .addItem('Sicherung erstellen', 'menuBackup')
    .addItem('Testdaten entfernen …', 'menuReset')
    .addItem('Import aus Bestandslisten …', 'menuImport')
    .addSeparator()
    .addItem('Bestandsübersicht neu aufbauen', 'menuBestand')
    .addItem('Zubehör-Vorschläge neu setzen …', 'menuZubehoer')
    .addItem('Bilder aus Ordner zuordnen', 'menuBilder')
    .addToUi();
}

function menuSetup() {
  setup_();
  SpreadsheetApp.getUi().alert('Magic-Sterne ist eingerichtet.\n\nNächste Schritte: PINs festlegen, dann Import aus Bestandslisten.');
}

function menuPins() {
  var ui = SpreadsheetApp.getUi();
  var a = ui.prompt('PIN für Verkäufer', 'Mindestens 4 Ziffern. Öffnet die Kasse.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  var b = ui.prompt('PIN für Inhaber', 'Mindestens 4 Ziffern, anders als die Verkäufer-PIN. Öffnet alles.', ui.ButtonSet.OK_CANCEL);
  if (b.getSelectedButton() !== ui.Button.OK) return;
  var pv = a.getResponseText().trim(), pi = b.getResponseText().trim();
  if (!/^\d{4,}$/.test(pv) || !/^\d{4,}$/.test(pi)) { ui.alert('Bitte nur Ziffern, mindestens 4.'); return; }
  if (pv === pi) { ui.alert('Die beiden PINs müssen sich unterscheiden.'); return; }
  PropertiesService.getScriptProperties().setProperties({ PIN_VERKAUF: pv, PIN_INHABER: pi });
  ui.alert('PINs gespeichert.');
}

function menuBestand() { rebuildBestand_(); }

function menuBilder() {
  syncImagesFromFolder_();
  SpreadsheetApp.getUi().alert('Bilder zugeordnet. Ordner: „Magic-Sterne Bilder“ – Dateien heißen wie die Artikel-Nr., z. B. 16B.jpg.');
}

/* ---------- Einrichtung ---------- */

function setup_() {
  var ss = ss_();
  ensureSheet_(SH.ART, ART_HEAD, function (sh) {
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('I:I').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('#,##0.00 €');
    sh.getRange('J:J').setNumberFormat('#,##0.00 €');
    sh.getRange('D2:D').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Stern', 'Zubehör']).build());
  });
  ensureSheet_(SH.ORTE, ORTE_HEAD, function (sh) {
    sh.getRange(2, 1, 2, 3).setValues([['Lager', 'Lager', 'ja'], ['Fulda', 'Stand', 'ja']]);
  });
  ensureSheet_(SH.BEW, BEW_HEAD, function (sh) {
    sh.getRange('D:E').setNumberFormat('@');
    sh.getRange('A:A').setNumberFormat('dd.MM.yyyy HH:mm:ss');
    sh.getRange('I:I').setNumberFormat('dd.MM.yyyy HH:mm:ss');
    sh.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Object.keys(ARTEN)).build());
  });
  ensureArtColumns_();
  ensureMinColumns_();
  ensureInfo_();
  rebuildBestand_();
  var def = ss.getSheetByName('Tabellenblatt1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  ss.setActiveSheet(ss.getSheetByName(SH.BEST));
}

function ensureSheet_(name, head, init) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (init) init(sh);
  return sh;
}

/**
 * Ergänzt im Artikelblatt fehlende Standardspalten. „Filter: Kategorie“ kommt vor
 * die übrigen Filter, damit es in der Kasse als erster Filter erscheint.
 */
function ensureArtColumns_() {
  var sh = ss_().getSheetByName(SH.ART);
  ART_HEAD.forEach(function (h) {
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    if (head.indexOf(h) >= 0) return;
    var vor = h === 'Filter: Kategorie' ? head.indexOf('Filter: Größe') : -1;
    var col;
    if (vor >= 0) { sh.insertColumnBefore(vor + 1); col = vor + 1; }
    else col = head.length + 1;
    sh.getRange(1, col).setValue(h).setFontWeight('bold');
  });
}

/** Für jeden Ort eine Spalte „Min: <Ort>“ im Artikelblatt. */
function ensureMinColumns_() {
  var sh = ss_().getSheetByName(SH.ART);
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  getOrte_().forEach(function (o) {
    var h = 'Min: ' + o.name;
    if (head.indexOf(h) < 0) {
      sh.getRange(1, head.length + 1).setValue(h).setFontWeight('bold');
      head.push(h);
    }
  });
}

function ensureInfo_() {
  var sh = ss_().getSheetByName(SH.INFO);
  if (sh) return;
  sh = ss_().insertSheet(SH.INFO);
  var txt = [
    ['Magic-Sterne – so funktioniert dieses Sheet'],
    [''],
    ['Artikel: Stammdaten. Preise, Mindestbestände (Spalten „Min: <Ort>“), Zubehör je Stern (Artikel-Nr. mit Komma getrennt) und Filter hier pflegen.'],
    ['   „Filter: Kategorie“ = Magic-Sterne oder Blumen-Sterne (kommt aus dem Import). „Hinweis“ = Vermerke aus der Liste, z. B. „Keine Nachlieferung“.'],
    ['   Spalten, deren Überschrift mit „Filter:“ beginnt, erscheinen in der Kasse als Filter. Neue Filter-Spalten dürfen jederzeit angelegt werden.'],
    ['   Aktiv = leer oder „ja“ zeigt den Artikel, „nein“ blendet ihn aus.'],
    ['Orte: Lager (genau einer mit Typ „Lager“) und Marktstände (Typ „Stand“). Neuer Stand = neue Zeile, danach Menü → Einrichten.'],
    ['Bewegungen: das Buch aller Bestandsänderungen. Wird von der App geschrieben. Nicht von Hand ändern, keine Spalten umsortieren.'],
    ['   Fehler werden durch neue Buchungen korrigiert (Storno, Korrektur, Inventur), nicht durch Löschen.'],
    ['Bestand: wird aus den Bewegungen berechnet. Rot = unter Mindestbestand. Hinweis „bestellen“ (Lager) bzw. „nachliefern“ (Stand).'],
    ['   Nach neuen Artikeln: Menü Magic-Sterne → Bestandsübersicht neu aufbauen.'],
    ['Bilder: Ordner „Magic-Sterne Bilder“ neben diesem Sheet. Dateiname = Artikel-Nr. (z. B. 16B.jpg), dann Menü → Bilder aus Ordner zuordnen.'],
    ['Neue Saison: Menü → Sicherung erstellen, → Testdaten entfernen (nur wenn wirklich alles neu soll), → Import aus Bestandslisten. Ergebnis im Blatt „Importprotokoll“.']
  ];
  sh.getRange(1, 1, txt.length, 1).setValues(txt);
  sh.getRange('A1').setFontWeight('bold').setFontSize(14);
  sh.setColumnWidth(1, 900);
}

/* ---------- Bestandsübersicht ---------- */

function colLetter_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

/**
 * Ermittelt, welches Argument-Trennzeichen das Sheet versteht („,“ oder „;“,
 * je nach Gebietsschema), indem eine Testformel geschrieben und ausgewertet wird.
 */
function formulaSep_(sh) {
  var cell = sh.getRange(1, sh.getMaxColumns());
  var probe = '=IF(SUMIFS(A1:A1,A1:A1,"<>x")>=0,TEXTJOIN("-",TRUE,"a","b"),"")';
  var seps = [',', ';'];
  for (var i = 0; i < seps.length; i++) {
    cell.setFormula(localize_(probe, seps[i]));
    SpreadsheetApp.flush();
    var ok = cell.getDisplayValue() === 'a-b';
    cell.clearContent();
    if (ok) return seps[i];
  }
  throw new Error('Formeln werden in diesem Sheet nicht verstanden (Gebietsschema ' + ss_().getSpreadsheetLocale() + ').');
}

/** Ersetzt Kommas außerhalb von Anführungszeichen durch das Trennzeichen des Sheets. */
function localize_(f, sep) {
  if (sep === ',') return f;
  var out = '', inStr = false;
  for (var i = 0; i < f.length; i++) {
    var ch = f[i];
    if (ch === '"') inStr = !inStr;
    out += (!inStr && ch === ',') ? sep : ch;
  }
  return out;
}

function rebuildBestand_() {
  var ss = ss_(), sh = ss.getSheetByName(SH.BEST);
  if (!sh) sh = ss.insertSheet(SH.BEST, 0);
  sh.clear();
  var sep = formulaSep_(sh);
  sh.clearConditionalFormatRules();
  var orte = getOrte_().filter(function (o) { return o.aktiv; });
  var lager = lagerName_();
  var artHead = ss.getSheetByName(SH.ART).getRange(1, 1, 1, ss.getSheetByName(SH.ART).getLastColumn()).getValues()[0].map(String);
  var head = ['Nr', 'Artikel', 'Art'];
  orte.forEach(function (o) { head.push(o.name, 'Min ' + o.name); });
  head.push('Bestellt, offen', 'Hinweis');
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  sh.setFrozenRows(1);

  var arts = getArtikel_().filter(function (a) { return a.aktiv; });
  if (!arts.length) { sh.getRange(2, 1).setValue('Noch keine Artikel – Menü Magic-Sterne → Import aus Excel-Datei.'); return; }
  arts.sort(function (a, b) { return (a.art === b.art ? 0 : a.art === 'Stern' ? -1 : 1) || a.nr.localeCompare(b.nr, 'de', { numeric: true }); });

  var rows = arts.map(function (a, i) {
    var r = i + 2, row = [a.nr, a.kurz, a.art], hints = [];
    orte.forEach(function (o, j) {
      var stockCol = 4 + j * 2, minCol = stockCol + 1;
      var minArtCol = artHead.indexOf('Min: ' + o.name) + 1;
      row.push('=SUMIFS(Bewegungen!$F:$F,Bewegungen!$D:$D,$A' + r + ',Bewegungen!$E:$E,"' + o.name + '",Bewegungen!$C:$C,"<>Bestellung")');
      row.push(minArtCol > 0 ? '=IFERROR(INDEX(Artikel!$' + colLetter_(minArtCol) + ':$' + colLetter_(minArtCol) + ',MATCH($A' + r + ',Artikel!$A:$A,0)),0)' : '=0');
      var s = colLetter_(stockCol) + r, m = colLetter_(minCol) + r;
      hints.push('IF(' + s + '<' + m + ',"' + (o.name === lager ? 'bestellen' : 'nachliefern ' + o.name) + '","")');
    });
    row.push('=MAX(0,SUMIFS(Bewegungen!$F:$F,Bewegungen!$D:$D,$A' + r + ',Bewegungen!$C:$C,"Bestellung")-SUMIFS(Bewegungen!$F:$F,Bewegungen!$D:$D,$A' + r + ',Bewegungen!$C:$C,"Wareneingang"))');
    row.push('=TEXTJOIN(", ",TRUE,' + hints.join(',') + ')');
    return row;
  });
  // Texte als Werte, Formeln über setFormulas: nur so übersetzt Google die
  // Formeln (englische Schreibweise mit Kommas) in das Gebietsschema des Sheets.
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 3).setValues(rows.map(function (row) { return row.slice(0, 3); }));
  sh.getRange(2, 4, rows.length, head.length - 3).setFormulas(rows.map(function (row) {
    return row.slice(3).map(function (f) { return localize_(f, sep); });
  }));

  var rules = [];
  orte.forEach(function (o, j) {
    var c = colLetter_(4 + j * 2), m = colLetter_(5 + j * 2);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(localize_('=AND($A2<>"",' + c + '2<' + m + '2)', sep))
      .setBackground('#f4c7c3').setFontColor('#8c1d18')
      .setRanges([sh.getRange(c + '2:' + c + (rows.length + 1))]).build());
  });
  var hc = colLetter_(head.length);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty().setFontColor('#8c1d18').setBold(true)
    .setRanges([sh.getRange(hc + '2:' + hc + (rows.length + 1))]).build());
  sh.setConditionalFormatRules(rules);
  sh.autoResizeColumns(1, head.length);
}

/* ---------- Import ---------- */

var COLORS_ = [['bordeaux', 'Bordeaux'], ['türkis', 'Türkis'], ['violett', 'Lila'], ['lila', 'Lila'], ['pink', 'Pink'],
  ['orange', 'Orange'], ['gelb', 'Gelb'], ['rot', 'Rot'], ['red', 'Rot'], ['blau', 'Blau'], ['blue', 'Blau'], ['grün', 'Grün'],
  ['schwarz', 'Schwarz'], ['black', 'Schwarz'], ['silber', 'Silber'], ['braun', 'Braun'], ['beige', 'Natur'],
  ['naturweiß', 'Weiß'], ['naturwiß', 'Weiß'], ['natur', 'Natur'], ['creme', 'Weiß'], ['weiß', 'Weiß'],
  ['rosa', 'Rosa'], ['mint', 'Mint'], ['grau', 'Grau'], ['anthrazit', 'Anthrazit'], ['gold', 'Gold']];

function kurzname_(n, zub) {
  if (zub) {
    if (/Ständer/.test(n)) return 'Ständer Edelstahl';
    if (/Led|Leuchtmittel/i.test(n)) return 'LED-Kerze E14';
    if (/Kabel/i.test(n)) return 'Kabel ' + ((/^([\d,]+)\s?m/.exec(n) || [, ''])[1]) + ' m ' + (/schwarz/i.test(n) ? 'schwarz' : 'weiß');
    return n.split(',')[0];
  }
  var s = n.replace(/^\d+\s?cm\s*-\s*/i, '').replace(/[\s-]*\d+\s?cm\b.*$/i, '');
  var parts = s.split(',').map(function (p) { return p.trim(); }).filter(String);
  var keep = [parts[0]];
  parts.slice(1).forEach(function (p) {
    if (!/papier|bestickt|pailletten|pailetten|hochwertig|glänzt|zaubert|^\s*mit\b|^\s*neu\s*$/i.test(p) && p.length < 16) keep.push(p);
  });
  return keep.join(' ').replace(/\s+neu$/i, '').replace(/\s+/g, ' ').trim();
}

function farben_(s) {
  var l = s.toLowerCase(), out = [];
  COLORS_.forEach(function (c) {
    var re = new RegExp('(^|[^a-zäöüß])' + c[0] + '(?![a-zäöüß])');
    if (re.test(l) && out.indexOf(c[1]) < 0) out.push(c[1]);
  });
  return out.join(', ');
}
