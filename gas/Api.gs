/**
 * Sternkasse – Schnittstelle für die Web-App.
 *
 * Alle Aufrufe gehen als POST mit Content-Type text/plain und JSON im Rumpf
 * an die /exec-Adresse (so braucht der Browser keine CORS-Vorabfrage).
 * Jeder Aufruf trägt die PIN; daraus ergibt sich die Rolle.
 */

var API_VERSION = 1;

function doGet() {
  return json_({ ok: true, app: 'Magic-Sterne', api: API_VERSION });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, fehler: 'Ungültige Anfrage.' });
  }
  try {
    var rolle = checkPin_(req.pin);
    var a = req.a;
    var out;
    switch (a) {
      case 'ping':   out = { rolle: rolle, orte: getOrte_() }; break;
      case 'state':  out = state_(rolle); break;
      case 'book':   out = { ergebnisse: book_(req.ops || [], rolle) }; out.stand = state_(rolle); break;
      case 'moves':  needInhaber_(rolle); out = moves_(req.von, req.bis); break;
      case 'images': out = { bilder: imageIndex_() }; break;
      case 'image':  out = imageData_(req.nr); break;
      case 'setImage': needInhaber_(rolle); out = setImage_(req.nr, req.data); break;
      default: throw new Error('Unbekannte Aktion „' + a + '“.');
    }
    out.ok = true;
    out.api = API_VERSION;
    out.serverZeit = new Date().toISOString();
    return json_(out);
  } catch (err) {
    return json_({ ok: false, fehler: err.message, code: err.code || '' });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- PIN ---------- */

function checkPin_(pin) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) { var e = new Error('Zu viele falsche PINs. Bitte 10 Minuten warten.'); e.code = 'gesperrt'; throw e; }
  var p = PropertiesService.getScriptProperties();
  var pv = p.getProperty('PIN_VERKAUF'), pi = p.getProperty('PIN_INHABER');
  if (!pv || !pi) { var e2 = new Error('PINs sind noch nicht festgelegt (Menü Magic-Sterne → PINs festlegen).'); e2.code = 'keinePin'; throw e2; }
  pin = String(pin || '');
  if (pin && pin === pi) return 'inhaber';
  if (pin && pin === pv) return 'verkauf';
  cache.put('pinFails', String(fails + 1), 600);
  var e3 = new Error('PIN falsch.'); e3.code = 'pin'; throw e3;
}

function needInhaber_(rolle) {
  if (rolle !== 'inhaber') { var e = new Error('Nur für den Inhaber.'); e.code = 'rolle'; throw e; }
}

/* ---------- Lesen ---------- */

function state_(rolle) {
  var s = computeStock_();
  var artikel = getArtikel_().filter(function (a) { return a.aktiv; }).map(function (a) {
    var o = { nr: a.nr, name: a.name, kurz: a.kurz, art: a.art, preis: a.preis, zubehoer: a.zubehoer,
      min: a.min, filter: a.filter, bild: a.bild ? true : false };
    return o;
  });
  return {
    rolle: rolle,
    orte: getOrte_().filter(function (o) { return o.aktiv; }),
    artikel: artikel,
    bestand: s.stock,
    offenBestellt: rolle === 'inhaber' ? s.offen : {},
    renner: renner_(),
    bilderStand: PropertiesService.getScriptProperties().getProperty('BILDER_STAND') || '0'
  };
}

/** Verkaufte Stück je Artikel in dieser Saison (für „Renner zuerst“). */
function renner_() {
  var t = readTable_(SH.BEW);
  var art = col_(t, 'Art'), nr = col_(t, 'Nr'), menge = col_(t, 'Menge'), notiz = col_(t, 'Notiz'), vg = col_(t, 'Vorgang');
  var verkaufsVorgang = {}, out = {};
  t.rows.forEach(function (r) {
    var a = nr_(art(r));
    if (a === 'Verkauf' || a === 'Rückgabe') verkaufsVorgang[nr_(vg(r))] = true;
  });
  t.rows.forEach(function (r) {
    var a = nr_(art(r));
    if (a === 'Storno') {
      var m = /^Storno von (.+)$/.exec(nr_(notiz(r)));
      if (!m || !verkaufsVorgang[m[1]]) return; // nur Stornos von Verkäufen/Rückgaben zählen
    } else if (a !== 'Verkauf' && a !== 'Rückgabe') return;
    var n = nr_(nr(r));
    out[n] = (out[n] || 0) - (Number(menge(r)) || 0);
  });
  Object.keys(out).forEach(function (k) { if (out[k] <= 0) delete out[k]; });
  return out;
}

/** Bewegungen in einem Zeitraum (für Berichte). von/bis als ISO-Datum, bis inklusive. */
function moves_(von, bis) {
  var t = readTable_(SH.BEW);
  var c = {};
  BEW_HEAD.forEach(function (h) { c[h] = col_(t, h); });
  var v = von ? new Date(von) : new Date(0);
  var b = bis ? new Date(bis) : new Date(8640000000000000);
  return {
    bewegungen: t.rows.filter(function (r) {
      var z = c['Zeitpunkt'](r);
      z = z instanceof Date ? z : new Date(z);
      return z >= v && z <= b;
    }).map(function (r) {
      var z = c['Zeitpunkt'](r);
      return { zeit: (z instanceof Date ? z : new Date(z)).toISOString(), vorgang: nr_(c['Vorgang'](r)), art: nr_(c['Art'](r)),
        nr: nr_(c['Nr'](r)), ort: nr_(c['Ort'](r)), menge: Number(c['Menge'](r)) || 0, notiz: nr_(c['Notiz'](r)) };
    })
  };
}

/* ---------- Bilder ---------- */

function imageFolder_() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('BILDER_ORDNER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* neu anlegen */ } }
  var file = DriveApp.getFileById(ss_().getId());
  var parents = file.getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var f = parent.createFolder('Magic-Sterne Bilder');
  p.setProperty('BILDER_ORDNER', f.getId());
  return f;
}

/** Ordnet Bilddateien „<Nr>.jpg/png“ aus dem Bilderordner den Artikeln zu. */
function syncImagesFromFolder_() {
  var folder = imageFolder_(), files = folder.getFiles(), map = {};
  while (files.hasNext()) {
    var f = files.next(), m = /^(.+?)\.(jpe?g|png|webp)$/i.exec(f.getName());
    if (m && !f.isTrashed()) map[m[1].trim()] = f;
  }
  var t = readTable_(SH.ART), iNr = t.idx['Nr'], iBild = t.idx['Bild'], changed = false;
  if (iBild === undefined) return;
  var all = t.sheet.getDataRange().getValues();
  for (var r = 1; r < all.length; r++) {
    var n = nr_(all[r][iNr]), f2 = map[n];
    var want = f2 ? f2.getId() : '';
    if (nr_(all[r][iBild]) !== want && (f2 || !nr_(all[r][iBild]))) {
      t.sheet.getRange(r + 1, iBild + 1).setValue(want);
      changed = true;
    }
  }
  if (changed) bumpImages_();
}

function bumpImages_() {
  PropertiesService.getScriptProperties().setProperty('BILDER_STAND', String(Date.now()));
}

function imageIndex_() {
  var out = {};
  getArtikel_().forEach(function (a) {
    if (!a.bild) return;
    try { out[a.nr] = DriveApp.getFileById(a.bild).getLastUpdated().getTime(); } catch (e) { /* gelöscht */ }
  });
  return out;
}

function imageData_(nr) {
  var a = getArtikel_().filter(function (x) { return x.nr === nr_(nr); })[0];
  if (!a || !a.bild) return { nr: nr, data: null };
  var blob = DriveApp.getFileById(a.bild).getBlob();
  return { nr: a.nr, data: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
}

/** data = "data:image/jpeg;base64,…" (in der App bereits verkleinert). */
function setImage_(nr, data) {
  nr = nr_(nr);
  var m = /^data:(image\/(?:jpeg|png|webp));base64,(.+)$/.exec(String(data || ''));
  if (!m) throw new Error('Bild nicht erkannt.');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 400000) throw new Error('Bild zu groß (max. 400 KB).');
  var t = readTable_(SH.ART), iNr = t.idx['Nr'], iBild = t.idx['Bild'];
  var all = t.sheet.getDataRange().getValues(), row = -1;
  for (var r = 1; r < all.length; r++) if (nr_(all[r][iNr]) === nr) { row = r; break; }
  if (row < 0) throw new Error('Artikel „' + nr + '“ nicht gefunden.');
  var folder = imageFolder_(), ext = m[1].split('/')[1].replace('jpeg', 'jpg');
  var old = folder.getFilesByName(nr + '.' + ext);
  while (old.hasNext()) old.next().setTrashed(true);
  var file = folder.createFile(Utilities.newBlob(bytes, m[1], nr + '.' + ext));
  t.sheet.getRange(row + 1, iBild + 1).setValue(file.getId());
  bumpImages_();
  return { nr: nr, bild: true };
}
