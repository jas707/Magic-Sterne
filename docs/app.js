'use strict';
/* Magic-Sterne – Web-App (Kasse für den Stand, Verwaltung für den Inhaber).
   Daten kommen von der Apps-Script-Schnittstelle am Google Sheet. Buchungen werden
   zuerst lokal gepuffert (Warteschlange) und gesendet, sobald Netz da ist. */

const APP_VERSION = '1.0.0';

/* ---------- Speicher ---------- */
const LS = {
  get(k, d) { try { const v = localStorage.getItem('ms.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem('ms.' + k, JSON.stringify(v)); } catch (e) { /* voll oder gesperrt */ } },
  del(k) { try { localStorage.removeItem('ms.' + k); } catch (e) { /* egal */ } }
};

let cfg = LS.get('cfg', null);         // {api, pin, rolle, ort, geraet}
let st = LS.get('state', null);        // letzter Stand vom Server
let queue = LS.get('queue', []);       // noch nicht gesendete Vorgänge
let failed = LS.get('failed', []);     // vom Server abgelehnte Vorgänge
let lastSale = LS.get('lastSale', null);
const imgs = {};                        // Nr -> data-URL

const ui = {
  view: 'kasse', tab: 'Stern', q: '', filters: {}, pop: null, sort: 'renner', mode: 'Verkauf',
  cart: new Map(), undo: [], lastStar: null, open: false,
  online: navigator.onLine, syncing: false, syncErr: null, lastSync: LS.get('lastSync', null),
  book: { art: 'Lieferung', ort: null, qty: {}, q: '' },
  over: { nur: false, q: '' },
  report: { range: 'heute', von: '', bis: '', moves: null, loading: false, err: null },
  setup: { api: '', step: 'pin', orte: [], rolle: null, busy: false, err: null },
  reloadPending: false
};

const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const eur = v => (v || 0).toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));

/* ---------- Schnittstelle ---------- */
async function api(a, extra = {}, c = cfg) {
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), 25000);
  try {
    const r = await fetch(c.api, {
      method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ a, pin: c.pin }, extra)), signal: ctl.signal, redirect: 'follow'
    });
    const j = await r.json();
    if (!j.ok) { const e = new Error(j.fehler || 'Unbekannter Fehler'); e.code = j.code; throw e; }
    return j;
  } catch (e) {
    if (e.name === 'AbortError') { const e2 = new Error('Keine Antwort vom Server (Zeitüberschreitung).'); e2.code = 'netz'; throw e2; }
    if (!e.code) e.code = 'netz';
    throw e;
  } finally { clearTimeout(t); }
}

function setState(s) {
  if (!s || !s.artikel) return;
  st = { rolle: s.rolle, orte: s.orte, artikel: s.artikel, bestand: s.bestand, offenBestellt: s.offenBestellt || {},
    renner: s.renner || {}, bilderStand: s.bilderStand, serverZeit: s.serverZeit };
  LS.set('state', st);
  if (cfg && s.rolle && cfg.rolle !== s.rolle) { cfg.rolle = s.rolle; LS.set('cfg', cfg); }
}

/* ---------- Stammdaten-Helfer ---------- */
const arts = () => (st && st.artikel) || [];
const byNr = nr => arts().find(a => a.nr === nr);
const lager = () => ((st && st.orte) || []).find(o => /^lager$/i.test(o.typ));
const lagerName = () => (lager() || { name: 'Lager' }).name;
const staende = () => ((st && st.orte) || []).filter(o => !/^lager$/i.test(o.typ));
const isInhaber = () => cfg && cfg.rolle === 'inhaber';

/** Bestand = Serverstand + Wirkung der noch nicht gesendeten Vorgänge. */
function stockAt(nr, ort) {
  let s = (((st && st.bestand) || {})[nr] || {})[ort] || 0;
  for (const op of queue) for (const e of (op._eff || [])) if (e.nr === nr && e.ort === ort) s += e.d;
  return s;
}
function minAt(a, ort) { return (a.min && a.min[ort]) || 0; }

/** Lokale Bestandswirkung eines Vorgangs (spiegelt die Buchungslogik im Script). */
function effFor(op) {
  const L = lagerName(), out = [];
  const add = (nr, ort, d) => { if (d) out.push({ nr, ort, d }); };
  for (const l of op.lines || []) {
    const q = Math.abs(l.menge);
    switch (op.art) {
      case 'Verkauf': add(l.nr, op.ort, -q); break;
      case 'Rückgabe': add(l.nr, op.ort, q); break;
      case 'Lieferung': add(l.nr, L, -q); add(l.nr, op.ort, q); break;
      case 'Rücklieferung': add(l.nr, op.ort, -q); add(l.nr, L, q); break;
      case 'Wareneingang': add(l.nr, L, q); break;
      case 'Retoure': add(l.nr, L, -q); break;
      case 'Korrektur': add(l.nr, op.ort, l.menge); break;
      case 'Inventur': add(l.nr, op.ort, q - stockAt(l.nr, op.ort)); break;
      default: break;
    }
  }
  return out;
}

function enqueue(op) {
  op.id = op.id || uid();
  op.zeit = new Date().toISOString();
  op.geraet = cfg.geraet || '';
  if (!op._eff) op._eff = effFor(op);
  queue.push(op);
  LS.set('queue', queue);
  sync();
  return op;
}

/* ---------- Abgleich ---------- */
let syncTimer = null;
async function sync() {
  if (!cfg) return;
  if (ui.syncing) { ui.syncAgain = true; return; }
  ui.syncing = true; renderStatus();
  try {
    if (queue.length) {
      const batch = queue.slice(0, 40);
      const ops = batch.map(o => { const c = Object.assign({}, o); delete c._eff; delete c._label; return c; });
      ui.inflight = new Set(batch.map(o => o.id));
      let j;
      try { j = await api('book', { ops }); } finally { ui.inflight = null; }
      const res = new Map(j.ergebnisse.map(r => [r.id, r]));
      queue = queue.filter(op => {
        const r = res.get(op.id);
        if (!r) return true;
        if (!r.ok) failed.push(Object.assign({}, op, { fehler: r.fehler }));
        return false;
      });
      LS.set('queue', queue); LS.set('failed', failed);
      setState(j.stand);
      if (queue.length) ui.syncAgain = true;
    } else {
      setState(await api('state'));
    }
    ui.syncErr = null; ui.lastSync = Date.now(); LS.set('lastSync', ui.lastSync);
    loadImages();
  } catch (e) {
    ui.syncErr = e;
    if (e.code === 'pin' || e.code === 'keinePin') toast('Anmeldung ungültig: ' + e.message + ' Bitte neu einrichten (⚙).', true);
  } finally {
    ui.syncing = false;
    if (ui.syncAgain) { ui.syncAgain = false; if (!ui.syncErr) setTimeout(sync, 200); }
    refresh();
  }
}
function scheduleSync() {
  clearInterval(syncTimer);
  syncTimer = setInterval(() => {
    if (document.visibilityState !== 'visible') return;
    if (queue.length || !ui.lastSync || Date.now() - ui.lastSync > 60000) sync();
  }, 15000);
}

/* ---------- Bilder ---------- */
async function initImages() {
  try {
    const cache = await caches.open('ms-bilder');
    for (const req of await cache.keys()) {
      const nr = decodeURIComponent(req.url.split('/bild/')[1] || '');
      if (nr) imgs[nr] = await (await cache.match(req)).text();
    }
  } catch (e) { /* kein Cache verfügbar */ }
}
async function loadImages() {
  if (!st || ui.imgBusy) return;
  const have = LS.get('imgIdx', {});
  if (have._stand === st.bilderStand) return;
  ui.imgBusy = true;
  try {
    const idx = (await api('images')).bilder || {};
    let cache = null; try { cache = await caches.open('ms-bilder'); } catch (e) { /* ohne Cache */ }
    for (const nr of Object.keys(idx)) {
      if (have[nr] === idx[nr] && imgs[nr]) continue;
      const r = await api('image', { nr });
      if (r.data) { imgs[nr] = r.data; if (cache) await cache.put('bild/' + encodeURIComponent(nr), new Response(r.data)); }
    }
    for (const nr of Object.keys(imgs)) if (!(nr in idx)) { delete imgs[nr]; if (cache) await cache.delete('bild/' + encodeURIComponent(nr)); }
    idx._stand = st.bilderStand; LS.set('imgIdx', idx);
    refresh();
  } catch (e) { /* nächster Versuch beim nächsten Abgleich */ } finally { ui.imgBusy = false; }
}
async function resizeImage(file, max = 480) {
  const bmp = await createImageBitmap(file);
  const s = Math.min(1, max / Math.max(bmp.width, bmp.height));
  const c = document.createElement('canvas');
  c.width = Math.round(bmp.width * s); c.height = Math.round(bmp.height * s);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return c.toDataURL('image/jpeg', 0.82);
}

/* ---------- Platzhalterbilder ---------- */
const FARBE = { Rot: '#c62a22', Bordeaux: '#7d1d33', Natur: '#e0cb98', Weiß: '#f4f0e8', Lila: '#7b4bb0', Blau: '#2f62b8', Türkis: '#1f9e9a',
  Schwarz: '#2a2a30', Orange: '#e6762a', Gelb: '#e3b320', Pink: '#dd4c86', Grün: '#2f8a4c', Silber: '#b4bac2', Braun: '#7b5133', Sonstige: '#a99a7c' };
function colorFor(label) {
  if (FARBE[label]) return FARBE[label];
  let h = 0; for (const ch of String(label)) h = (h * 31 + ch.charCodeAt(0)) % 360;
  return `hsl(${h} 55% 50%)`;
}
const hex2 = h => /^#/.test(h) ? [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16)) : [160, 150, 130];
const mix = (h, w, t) => { const a = hex2(h), b = hex2(w); return 'rgb(' + a.map((v, i) => Math.round(v + (b[i] - v) * t)).join(',') + ')'; };
function starSVG(a) {
  const cols = ((a.filter && a.filter['Farbe']) || ['Sonstige']).map(colorFor);
  const n = +((a.kurz.match(/\b([5-9])\b/) || [])[1]) || 7, R = 46, r = 18, parts = [];
  for (let i = 0; i < n; i++) {
    const c = cols[i % cols.length], t = i * 2 * Math.PI / n - Math.PI / 2, d = Math.PI / n;
    const P = (ang, rad) => (Math.cos(ang) * rad).toFixed(1) + ',' + (Math.sin(ang) * rad).toFixed(1);
    parts.push(`<polygon points="0,0 ${P(t - d, r)} ${P(t, R)}" fill="${mix(c, '#fff3d6', .32)}"/><polygon points="0,0 ${P(t, R)} ${P(t + d, r)}" fill="${mix(c, '#000000', .12)}"/>`);
  }
  return `<svg viewBox="-50 -50 100 100" aria-hidden="true"><circle r="44" fill="${mix(cols[0], '#ffd27a', .4)}" opacity=".16"/>${parts.join('')}<circle r="9" fill="#fff4cf" opacity=".85"/><circle r="4.5" fill="#fffbea"/></svg>`;
}
function zubSVG(a) {
  const n = a.kurz + ' ' + a.name;
  if (/Ständer/i.test(n)) return `<svg viewBox="0 0 40 40" aria-hidden="true"><path d="M20 6v24M12 34h16M20 30l-7 4M20 30l7 4" stroke="#c9ced4" stroke-width="2.6" stroke-linecap="round" fill="none"/><circle cx="20" cy="7" r="3" fill="#e8b84a"/></svg>`;
  if (/LED|Leuchtmittel|Birne/i.test(n)) return `<svg viewBox="0 0 40 40" aria-hidden="true"><path d="M20 5c-6 6-8 11-8 16a8 8 0 0 0 16 0c0-5-2-10-8-16z" fill="#ffe7a3"/><rect x="15" y="29" width="10" height="7" rx="1.5" fill="#b9bec4"/></svg>`;
  const c = /schwarz/i.test(n) ? '#2a2a30' : '#eef0f2';
  return `<svg viewBox="0 0 40 40" aria-hidden="true"><path d="M6 30c8 0 8-18 16-18s6 12 12 12" stroke="${c}" stroke-width="3" fill="none" stroke-linecap="round"/><rect x="29" y="18" width="8" height="10" rx="2" fill="${c}" stroke="#8a9096"/><rect x="3" y="27" width="7" height="6" rx="1.5" fill="#8a9096"/></svg>`;
}
const artImg = a => imgs[a.nr] ? `<img src="${imgs[a.nr]}" alt="">` : (a.art === 'Zubehör' ? zubSVG(a) : starSVG(a));

/* ---------- Grundgerüst ---------- */
function render() {
  const root = $('root');
  if (!cfg) { renderSetup(); return; }
  if (!st) {
    root.innerHTML = `<div class="setup"><form onsubmit="return false"><h1>Magic-Sterne</h1><p class="note">Lade Artikel und Bestände …</p>${ui.syncErr ? `<div class="err">${esc(ui.syncErr.message)}</div><button class="btn primary" data-act="retry">Erneut versuchen</button> <button class="btn" data-act="logout">Neu einrichten</button>` : ''}</form></div>`;
    return;
  }
  if (!isInhaber() && ui.view !== 'kasse') ui.view = 'kasse';
  const tabs = isInhaber() ? [['kasse', 'Kasse'], ['over', 'Übersicht'], ['book', 'Buchen'], ['report', 'Bericht']] : [];
  root.innerHTML = `<div class="shell">
    <header class="nav">
      <div class="brand">Magic-Sterne <small>${esc(isInhaber() ? 'Inhaber' : (cfg.ort || ''))}</small></div>
      ${tabs.length ? `<div class="seg" role="group" aria-label="Bereich">${tabs.map(([k, l]) => `<button data-view="${k}" aria-pressed="${ui.view === k}">${l}</button>`).join('')}</div>` : ''}
      <span id="status"></span>
      <button class="iconbtn" data-act="settings" aria-label="Einstellungen">⚙</button>
    </header>
    <main class="main" id="main"></main></div>`;
  renderView();
  renderStatus();
}
function renderView() {
  const m = $('main'); if (!m) return;
  if (ui.view === 'kasse') renderKasse(m);
  else if (ui.view === 'over') renderOver(m);
  else if (ui.view === 'book') renderBook(m);
  else if (ui.view === 'report') renderReport(m);
}
/** Aktualisiert nach einem Abgleich nur, was nicht gerade bearbeitet wird. */
function refresh() {
  if (!cfg) return;
  if (!$('main')) { render(); return; }
  renderStatus();
  if (ui.view === 'kasse') { renderGrid(); renderZbar(); }
  else if (ui.view === 'over') renderView();
  else if (ui.view === 'report') renderView();
  else if (ui.view === 'book' && !document.activeElement?.classList.contains('qin')) renderView();
}
function renderStatus() {
  const el = $('status'); if (!el) return;
  const n = queue.length;
  let cls = '', txt;
  if (ui.syncing) txt = 'Abgleich …';
  else if (ui.syncErr && ui.syncErr.code !== 'netz') { cls = 'err'; txt = 'Fehler'; }
  else if (!ui.online || ui.syncErr) { cls = 'off'; txt = n ? `offline · ${n} offen` : 'offline'; }
  else txt = n ? `${n} offen` : 'aktuell';
  if (failed.length) { cls = 'err'; txt += ` · ${failed.length} abgelehnt`; }
  el.innerHTML = `<button class="sync ${cls}" data-act="settings" title="Abgleich"><i></i>${txt}</button>`;
}

/* ---------- Kasse ---------- */
function renderKasse(m) {
  if (!cfg.ort || !staende().some(o => o.name === cfg.ort)) cfg.ort = (staende()[0] || {}).name;
  if (!cfg.ort) { m.innerHTML = `<div class="page"><div class="wrap"><h2>Kein Marktstand angelegt</h2><p class="note">Im Sheet im Blatt „Orte“ einen Ort mit Typ „Stand“ anlegen.</p></div></div>`; return; }
  m.innerHTML = `<div class="kasse" id="kasse">
    <section class="shop" aria-label="Artikel">
      <div class="top">
        <div class="seg" role="group" aria-label="Artikelart">
          <button data-tab="Stern" aria-pressed="${ui.tab === 'Stern'}">Sterne</button>
          <button data-tab="Zubehör" aria-pressed="${ui.tab === 'Zubehör'}">Zubehör</button>
        </div>
        <label class="search"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
          <input id="q" type="search" placeholder="Name oder Nr. suchen" autocomplete="off" value="${esc(ui.q)}"></label>
        ${isInhaber() && staende().length > 1 ? `<select id="kOrt" class="chip">${staende().map(o => `<option ${o.name === cfg.ort ? 'selected' : ''}>${esc(o.name)}</option>`).join('')}</select>` : ''}
      </div>
      <div class="filters" id="filters"></div>
      <div class="zbar" id="zbar"></div>
      <div class="grid" id="grid"></div>
    </section>
    <aside class="cart" id="cart" aria-label="Aktueller Verkauf">
      <div class="chead">
        <h2 id="ctitle">Aktueller Verkauf</h2>
        <div class="seg mode" role="group" aria-label="Modus">
          <button data-m="Verkauf">Verkauf</button><button data-m="Rückgabe">Rückgabe</button>
        </div>
        <button class="ctoggle" data-act="ctoggle" id="ctoggle">Liste ▲</button>
      </div>
      <div class="lines" id="lines"></div>
      <div class="sum">
        <div class="total"><span class="tl" id="tlabel">Summe</span><span class="tv" id="total">0,00 €</span></div>
        <div class="acts">
          <button class="btn" data-act="undo" id="undo">↶ Rückgängig</button>
          <button class="btn" data-act="clear" id="clear">Leeren</button>
          <button class="go" data-act="go" id="go">Fertigstellen</button>
        </div>
      </div>
    </aside>
    <footer class="foot"><div class="last" id="last"></div></footer>
  </div>`;
  $('q').addEventListener('input', e => { ui.q = e.target.value; renderGrid(); });
  const ko = $('kOrt'); if (ko) ko.addEventListener('change', e => { cfg.ort = e.target.value; LS.set('cfg', cfg); renderGrid(); });
  renderFilters(); renderZbar(); renderGrid(); renderCart(); renderLast();
}
function filterKeys() {
  const keys = new Map();
  arts().filter(a => a.art === ui.tab).forEach(a => Object.entries(a.filter || {}).forEach(([k, vs]) => {
    if (!keys.has(k)) keys.set(k, new Map());
    vs.forEach(v => keys.get(k).set(v, (keys.get(k).get(v) || 0) + 1));
  }));
  return keys;
}
function renderFilters() {
  const f = $('filters'); if (!f) return;
  const keys = filterKeys();
  let row = '', pop = '';
  for (const [k, vals] of keys) {
    const sel = ui.filters[k] || new Set();
    const list = [...vals.entries()].sort((a, b) => b[1] - a[1]).map(([v]) => v);
    const chip = v => `<button class="chip" data-f="${esc(k)}" data-v="${esc(v)}" aria-pressed="${sel.has(v)}">${k === 'Farbe' ? `<span class="dot" style="background:${colorFor(v)}"></span>` : ''}${esc(v)}</button>`;
    if (list.length <= 3) row += list.map(chip).join('');
    else {
      row += `<button class="chip more" data-pop="${esc(k)}" aria-expanded="${ui.pop === k}">${esc(k)}${sel.size ? ` (${sel.size})` : ''} ▾</button>`;
      row += [...sel].map(v => `<button class="chip" data-f="${esc(k)}" data-v="${esc(v)}" aria-pressed="true">${k === 'Farbe' ? `<span class="dot" style="background:${colorFor(v)}"></span>` : ''}${esc(v)}<span class="x">✕</span></button>`).join('');
      if (ui.pop === k) pop = `<div class="pop" id="pop">${list.map(chip).join('')}${sel.size ? `<button class="chip" data-fclear="${esc(k)}">Alle</button>` : ''}</div>`;
    }
  }
  f.innerHTML = `<div class="frow">${row}<div class="sort"><button data-sort="renner" aria-pressed="${ui.sort === 'renner'}">Renner zuerst</button><button data-sort="nr" aria-pressed="${ui.sort === 'nr'}">Nr.</button></div></div>${pop}`;
}
function visible() {
  const q = ui.q.trim().toLowerCase();
  let list = arts().filter(a => a.art === ui.tab);
  if (q) list = list.filter(a => a.nr.toLowerCase().startsWith(q) || a.kurz.toLowerCase().includes(q) || a.name.toLowerCase().includes(q));
  else for (const [k, sel] of Object.entries(ui.filters)) if (sel.size) list = list.filter(a => (a.filter[k] || []).some(v => sel.has(v)));
  const ren = (st && st.renner) || {};
  if (ui.sort === 'renner') list = [...list].sort((a, b) => (ren[b.nr] || 0) - (ren[a.nr] || 0) || a.nr.localeCompare(b.nr, 'de', { numeric: true }));
  else list = [...list].sort((a, b) => a.nr.localeCompare(b.nr, 'de', { numeric: true }));
  return list;
}
function stockLabel(a) {
  const s = stockAt(a.nr, cfg.ort), min = minAt(a, cfg.ort);
  if (s <= 0) return `<span class="stock out">${s < 0 ? s : 'leer'}</span>`;
  return `<span class="stock${s < min ? ' low' : ''}">${s < min ? '⚠ ' : ''}${s} da</span>`;
}
function renderGrid() {
  const g = $('grid'); if (!g) return;
  const list = visible();
  g.innerHTML = list.length ? list.map(a => {
    const q = ui.cart.get(a.nr) || 0, noPrice = a.preis == null;
    const sub = (a.filter && a.filter['Größe'] || []).join(' / ');
    return `<button class="tile${q ? ' inCart' : ''}" data-add="${esc(a.nr)}" ${noPrice ? 'disabled' : ''} aria-label="${esc(a.kurz)}, ${noPrice ? 'kein Preis' : eur(a.preis)}">
      <div class="art">${artImg(a)}</div>
      <span class="tnr">${esc(a.nr)}</span>${q ? `<span class="badge">${q}</span>` : ''}
      <div class="meta"><div class="tname">${esc(a.kurz)}</div>
      <div class="trow"><span class="num">${noPrice ? 'Preis fehlt' : eur(a.preis)}</span>${stockLabel(a)}</div>
      ${sub ? `<div class="trow"><span class="stock">${esc(sub)}</span></div>` : ''}</div></button>`;
  }).join('') : `<div class="empty">Kein Artikel passt zu Suche oder Filter.</div>`;
}
function renderZbar() {
  const z = $('zbar'); if (!z) return;
  const a = ui.lastStar && byNr(ui.lastStar);
  if (!a) { z.innerHTML = `<span class="zt">Stern antippen – passendes Zubehör erscheint hier.</span>`; return; }
  const zs = (a.zubehoer || []).map(byNr).filter(x => x && x.preis != null);
  z.innerHTML = `<span class="zt">Zu <b>${esc(a.kurz)}</b>:</span>` + (zs.length ? zs.map(x => {
    const q = ui.cart.get(x.nr) || 0;
    return `<button class="zbtn" data-add="${esc(x.nr)}">${imgs[x.nr] ? `<img src="${imgs[x.nr]}" alt="">` : zubSVG(x)}+ ${esc(x.kurz)} <span class="num">${eur(x.preis)}</span>${q ? `<span class="zq">${q}</span>` : ''}</button>`;
  }).join('') : `<span class="zt">kein Zubehör hinterlegt</span>`);
}
const cartTotal = () => { let t = 0; ui.cart.forEach((q, nr) => { const a = byNr(nr); t += q * ((a && a.preis) || 0); }); return t; };
function renderCart() {
  const c = $('cart'); if (!c) return;
  const r = ui.mode === 'Rückgabe';
  c.classList.toggle('rueck', r); c.classList.toggle('open', ui.open);
  $('ctitle').textContent = r ? 'Rückgabe' : 'Aktueller Verkauf';
  c.querySelectorAll('.mode button').forEach(b => b.setAttribute('aria-pressed', b.dataset.m === ui.mode));
  const lines = [...ui.cart.entries()];
  $('lines').innerHTML = lines.length ? lines.map(([nr, q]) => {
    const a = byNr(nr) || { nr, kurz: nr, preis: 0 };
    return `<div class="line"><div class="ln"><span>${esc(a.nr)}</span>${esc(a.kurz)}</div><div class="lt">${eur(q * (a.preis || 0))}</div>
      <div class="qty"><button data-dec="${esc(nr)}" aria-label="eins weniger">−</button><button class="q" data-pad="${esc(nr)}" aria-label="Menge eingeben">${q}</button><button data-inc="${esc(nr)}" aria-label="eins mehr">+</button></div>
      <div class="up num">je ${eur(a.preis)}</div></div>`;
  }).join('') : `<div class="cempty">${r ? '<b>Rückgabe:</b> Artikel antippen, die der Kunde zurückbringt. Sie gehen zurück in den Standbestand.' : '<b>Kachel antippen = 1 Stück.</b><br>Mehrmals tippen für mehr, oder im Warenkorb auf die Zahl tippen und die Menge eingeben.'}</div>`;
  $('tlabel').textContent = r ? 'Erstattung' : 'Summe';
  $('total').textContent = eur(cartTotal());
  const n = [...ui.cart.values()].reduce((s, q) => s + q, 0);
  $('ctoggle').textContent = (ui.open ? 'Liste ▼' : 'Liste ▲') + (n ? ` · ${n} Stk.` : '');
  $('undo').disabled = !ui.undo.length; $('clear').disabled = !ui.cart.size; $('go').disabled = !ui.cart.size;
  $('go').textContent = r ? 'Rückgabe buchen' : 'Fertigstellen';
}
function renderLast() {
  const el = $('last'); if (!el) return;
  const l = lastSale;
  if (!l || l.ort !== cfg.ort) { el.textContent = 'Noch kein Verkauf abgeschlossen.'; return; }
  el.innerHTML = `${l.mode === 'Rückgabe' ? 'Letzte Rückgabe' : 'Letzter Verkauf'}: <span class="num">${eur(l.total)}</span> <span>(${l.time}, ${l.n} Stk.)</span> <button class="link" data-act="revert">zurücknehmen</button>`;
}
function setQty(nr, q, record = true) {
  const before = ui.cart.get(nr) || 0; q = Math.max(0, Math.min(999, q));
  if (q === before) return;
  if (record) ui.undo.push({ nr, before });
  if (q === 0) ui.cart.delete(nr); else ui.cart.set(nr, q);
  renderGrid(); renderZbar(); renderCart();
}
function addToCart(nr) {
  const a = byNr(nr); if (!a || a.preis == null) return;
  if (a.art === 'Stern') ui.lastStar = nr;
  setQty(nr, (ui.cart.get(nr) || 0) + 1);
}
function finishSale() {
  const lines = [...ui.cart.entries()].map(([nr, menge]) => ({ nr, menge }));
  if (!lines.length) return;
  const total = cartTotal();
  const op = enqueue({ art: ui.mode, ort: cfg.ort, lines });
  lastSale = { id: op.id, mode: ui.mode, ort: cfg.ort, lines, total, n: lines.reduce((s, l) => s + l.menge, 0),
    time: new Date().toLocaleTimeString('de-DE', { hour: '2-digit', minute: '2-digit' }), eff: op._eff };
  LS.set('lastSale', lastSale);
  toast(ui.mode === 'Rückgabe' ? 'Rückgabe gebucht · Bestand erhöht' : 'Verkauf gebucht · Bestand aktualisiert');
  ui.cart.clear(); ui.undo = []; ui.lastStar = null; ui.mode = 'Verkauf'; ui.open = false;
  renderGrid(); renderZbar(); renderCart(); renderLast(); renderStatus();
  if (ui.reloadPending) location.reload();
}
function revertSale() {
  const l = lastSale; if (!l) return;
  const i = queue.findIndex(o => o.id === l.id);
  // Noch nicht unterwegs → einfach aus der Warteschlange nehmen; sonst Storno buchen.
  if (i >= 0 && !(ui.inflight && ui.inflight.has(l.id))) { queue.splice(i, 1); LS.set('queue', queue); }
  else enqueue({ art: 'Storno', ref: l.id, ort: l.ort, _eff: (l.eff || []).map(e => ({ nr: e.nr, ort: e.ort, d: -e.d })) });
  lastSale = null; LS.del('lastSale');
  toast(l.mode === 'Rückgabe' ? 'Rückgabe zurückgenommen' : 'Verkauf zurückgenommen · Bestand wiederhergestellt');
  renderGrid(); renderLast(); renderStatus();
}

/* ---------- Übersicht (Inhaber) ---------- */
function needs() {
  const L = lagerName(), out = { bestellen: [], nachliefern: [] };
  arts().forEach(a => {
    if (stockAt(a.nr, L) < minAt(a, L)) out.bestellen.push(a);
    staende().forEach(o => { if (stockAt(a.nr, o.name) < minAt(a, o.name)) out.nachliefern.push({ a, ort: o.name }); });
  });
  return out;
}
function renderOver(m) {
  const orte = [lager(), ...staende()].filter(Boolean), nd = needs(), L = lagerName();
  const offen = (st && st.offenBestellt) || {};
  const q = ui.over.q.trim().toLowerCase();
  let list = [...arts()].sort((a, b) => (a.art === b.art ? 0 : a.art === 'Stern' ? -1 : 1) || a.nr.localeCompare(b.nr, 'de', { numeric: true }));
  if (q) list = list.filter(a => a.nr.toLowerCase().startsWith(q) || a.kurz.toLowerCase().includes(q));
  if (ui.over.nur) list = list.filter(a => orte.some(o => stockAt(a.nr, o.name) < minAt(a, o.name)));
  m.innerHTML = `<div class="page"><div class="wrap">
    ${failedHtml()}
    <div class="cards">
      <div class="card ${nd.bestellen.length ? 'alert' : ''}"><div class="k">Bestellen (Lager)</div><div class="v">${nd.bestellen.length}</div></div>
      <div class="card ${nd.nachliefern.length ? 'alert' : ''}"><div class="k">Nachliefern (Stand)</div><div class="v">${nd.nachliefern.length}</div></div>
      <div class="card"><div class="k">Bestellt, offen</div><div class="v">${Object.keys(offen).length}</div></div>
      <div class="card"><div class="k">Nicht gesendet</div><div class="v">${queue.length}</div></div>
    </div>
    <div class="toolbar">
      <label class="search"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg><input id="oq" type="search" placeholder="Name oder Nr." value="${esc(ui.over.q)}"></label>
      <button class="chip" data-act="overNur" aria-pressed="${ui.over.nur}">Nur Handlungsbedarf</button>
      <button class="btn" data-act="syncNow">Aktualisieren</button>
    </div>
    <div class="tablewrap"><table>
      <thead><tr><th>Nr</th><th></th><th>Artikel</th>${orte.map(o => `<th class="n">${esc(o.name)}</th>`).join('')}<th class="n">Bestellt</th><th>Hinweis</th><th>Foto</th></tr></thead>
      <tbody>${list.map(a => {
        const hints = [];
        const cells = orte.map(o => {
          const s = stockAt(a.nr, o.name), mi = minAt(a, o.name), bad = s < mi;
          if (bad) hints.push(o.name === L ? 'bestellen' : 'nachliefern ' + o.name);
          return `<td class="n ${bad ? 'bad' : ''}" title="Mindestbestand ${mi}">${s}<span class="note"> / ${mi}</span></td>`;
        }).join('');
        return `<tr><td class="num">${esc(a.nr)}</td><td>${imgs[a.nr] ? `<img class="thumb" src="${imgs[a.nr]}" alt="">` : ''}</td><td>${esc(a.kurz)}${a.preis == null ? ' <span class="hint">· Preis fehlt</span>' : ''}</td>${cells}
          <td class="n">${offen[a.nr] || ''}</td><td><span class="hint">${hints.join(', ')}</span></td>
          <td><label class="photo">📷<input type="file" accept="image/*" capture="environment" data-photo="${esc(a.nr)}" hidden></label></td></tr>`;
      }).join('')}</tbody></table></div>
    <p class="note">Zahl / Mindestbestand. Rot = unter Mindestbestand. Preise, Mindestbestände, Zubehör und Filter werden im Google Sheet (Blatt „Artikel“) gepflegt.</p>
  </div></div>`;
  $('oq').addEventListener('input', e => { ui.over.q = e.target.value; const pos = e.target.selectionStart; renderView(); const el = $('oq'); el.focus(); el.setSelectionRange(pos, pos); });
}
function failedHtml() {
  if (!failed.length) return '';
  return `<div class="err">${failed.length} Buchung(en) wurden vom Server abgelehnt:<ul>${failed.map((f, i) => `<li>${esc(f.art)} ${esc(new Date(f.zeit).toLocaleString('de-DE'))}: ${esc(f.fehler)} <button class="link" data-dropfail="${i}">verwerfen</button></li>`).join('')}</ul></div>`;
}

/* ---------- Buchen (Inhaber) ---------- */
const BOOK = {
  'Lieferung': { label: 'Lieferung an Stand', ort: 'stand', hint: 'Bucht vom Lager an den Stand. Vorschlag: Artikel unter Mindestbestand am Stand auf das Doppelte des Mindestbestands auffüllen (soweit im Lager vorhanden).' },
  'Rücklieferung': { label: 'Rücklieferung vom Stand', ort: 'stand', hint: 'Bucht vom Stand zurück ins Lager, z. B. am Saisonende. Vorschlag: alles, was am Stand liegt.' },
  'Wareneingang': { label: 'Wareneingang Lager', ort: null, hint: 'Ware vom Lieferanten ist im Lager angekommen. Vorschlag: offene Bestellmengen.' },
  'Bestellung': { label: 'Bestellung beim Lieferanten', ort: null, hint: 'Merkt bestellte Mengen vor (ändert keinen Bestand). Vorschlag: Lager unter Mindestbestand auf das Doppelte auffüllen, abzüglich offener Bestellungen.' },
  'Retoure': { label: 'Retoure an Lieferant', ort: null, hint: 'Ware geht aus dem Lager an den Lieferanten zurück.' },
  'Inventur': { label: 'Inventur (Zählung)', ort: 'alle', hint: 'Gezählte Mengen eintragen. Gebucht wird die Differenz zum aktuellen Bestand. Nur Artikel mit eingetragener Zahl werden geändert (0 eintragen = leer).' },
  'Korrektur': { label: 'Korrektur ±', ort: 'alle', hint: 'Bestand um die eingetragene Menge ändern, z. B. −1 für Bruch. Minuszeichen erlaubt.' }
};
function bookOrt() {
  const def = BOOK[ui.book.art];
  const list = def.ort === 'stand' ? staende() : def.ort === 'alle' ? [lager(), ...staende()].filter(Boolean) : [lager()].filter(Boolean);
  if (!list.some(o => o && o.name === ui.book.ort)) ui.book.ort = (list[0] || {}).name;
  return list;
}
function suggestion(a) {
  const L = lagerName(), o = ui.book.ort, offen = ((st && st.offenBestellt) || {})[a.nr] || 0;
  switch (ui.book.art) {
    case 'Lieferung': { const s = stockAt(a.nr, o), mi = minAt(a, o); return s < mi ? Math.max(0, Math.min(2 * mi - s, stockAt(a.nr, L))) : 0; }
    case 'Rücklieferung': return Math.max(0, stockAt(a.nr, o));
    case 'Wareneingang': return offen;
    case 'Bestellung': { const s = stockAt(a.nr, L), mi = minAt(a, L); return s < mi ? Math.max(0, 2 * mi - s - offen) : 0; }
    default: return 0;
  }
}
function renderBook(m) {
  const def = BOOK[ui.book.art], orte = bookOrt(), L = lagerName(), o = ui.book.ort;
  const q = ui.book.q.trim().toLowerCase();
  let list = [...arts()].sort((a, b) => (a.art === b.art ? 0 : a.art === 'Stern' ? -1 : 1) || a.nr.localeCompare(b.nr, 'de', { numeric: true }));
  if (q) list = list.filter(a => a.nr.toLowerCase().startsWith(q) || a.kurz.toLowerCase().includes(q));
  const showStand = o && o !== L;
  const set = Object.entries(ui.book.qty).filter(([, v]) => v !== '' && v != null && !isNaN(v) && (ui.book.art === 'Inventur' || Number(v) !== 0));
  m.innerHTML = `<div class="page"><div class="wrap">
    ${failedHtml()}
    <h2>Buchen</h2>
    <div class="toolbar">
      <select id="bArt">${Object.entries(BOOK).map(([k, d]) => `<option value="${k}" ${k === ui.book.art ? 'selected' : ''}>${d.label}</option>`).join('')}</select>
      ${def.ort ? `<select id="bOrt">${orte.map(x => `<option ${x.name === o ? 'selected' : ''}>${esc(x.name)}</option>`).join('')}</select>` : ''}
      <label class="search"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg><input id="bq" type="search" placeholder="Name oder Nr." value="${esc(ui.book.q)}"></label>
      ${['Lieferung', 'Rücklieferung', 'Wareneingang', 'Bestellung'].includes(ui.book.art) ? `<button class="btn" data-act="suggest">Vorschlag übernehmen</button>` : ''}
      ${ui.book.art === 'Inventur' ? `<button class="btn" data-act="suggestInv">Aktuellen Bestand eintragen</button>` : ''}
    </div>
    <p class="note">${def.hint}</p>
    <div class="tablewrap"><table>
      <thead><tr><th>Nr</th><th>Artikel</th><th class="n">${esc(L)}</th>${showStand ? `<th class="n">${esc(o)}</th>` : ''}<th class="n">Min</th><th class="n">Vorschlag</th><th class="n">Menge</th></tr></thead>
      <tbody>${list.map(a => {
        const sg = suggestion(a), v = ui.book.qty[a.nr];
        const mo = showStand ? o : L;
        return `<tr><td class="num">${esc(a.nr)}</td><td>${esc(a.kurz)}</td><td class="n ${stockAt(a.nr, L) < minAt(a, L) ? 'bad' : ''}">${stockAt(a.nr, L)}</td>
          ${showStand ? `<td class="n ${stockAt(a.nr, o) < minAt(a, o) ? 'bad' : ''}">${stockAt(a.nr, o)}</td>` : ''}
          <td class="n">${minAt(a, mo)}</td><td class="n">${sg || ''}</td>
          <td class="n"><input class="qin${v !== undefined && v !== '' ? ' set' : ''}" inputmode="${ui.book.art === 'Korrektur' ? 'text' : 'numeric'}" data-qin="${esc(a.nr)}" value="${v === undefined ? '' : esc(v)}" aria-label="Menge ${esc(a.kurz)}"></td></tr>`;
      }).join('')}</tbody></table></div>
    <div class="stick"><span id="bSum">${set.length} Artikel, ${set.reduce((s, [, v]) => s + Math.abs(Number(v)), 0)} Stück</span>
      <button class="btn" data-act="bookClear">Leeren</button>
      <button class="btn primary" data-act="bookGo" ${set.length ? '' : 'disabled'}>${esc(def.label)} buchen</button></div>
  </div></div>`;
  $('bArt').addEventListener('change', e => { ui.book.art = e.target.value; ui.book.qty = {}; renderView(); });
  const bo = $('bOrt'); if (bo) bo.addEventListener('change', e => { ui.book.ort = e.target.value; ui.book.qty = {}; renderView(); });
  $('bq').addEventListener('input', e => { ui.book.q = e.target.value; const pos = e.target.selectionStart; renderView(); const el = $('bq'); el.focus(); el.setSelectionRange(pos, pos); });
  m.querySelectorAll('.qin').forEach(inp => inp.addEventListener('input', e => {
    const raw = e.target.value.replace(',', '.').trim();
    if (raw === '' ) delete ui.book.qty[e.target.dataset.qin];
    else if (/^-?\d+$/.test(raw) && (ui.book.art === 'Korrektur' || Number(raw) >= 0)) ui.book.qty[e.target.dataset.qin] = Number(raw);
    e.target.classList.toggle('set', e.target.dataset.qin in ui.book.qty);
    const s = Object.entries(ui.book.qty).filter(([, v]) => ui.book.art === 'Inventur' || Number(v) !== 0);
    $('bSum').textContent = `${s.length} Artikel, ${s.reduce((t, [, v]) => t + Math.abs(Number(v)), 0)} Stück`;
    m.querySelector('[data-act="bookGo"]').disabled = !s.length;
  }));
}
function bookGo() {
  const def = BOOK[ui.book.art];
  const lines = Object.entries(ui.book.qty).filter(([, v]) => ui.book.art === 'Inventur' || Number(v) !== 0).map(([nr, menge]) => ({ nr, menge: Number(menge) }));
  if (!lines.length) return;
  enqueue({ art: ui.book.art, ort: def.ort ? ui.book.ort : lagerName(), lines });
  toast(def.label + ' gebucht: ' + lines.length + ' Artikel');
  ui.book.qty = {}; renderView(); renderStatus();
}

/* ---------- Bericht (Inhaber) ---------- */
function dayKey(d) { return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); }
function rangeBounds() {
  const r = ui.report, now = new Date(), start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const day = 86400000;
  switch (r.range) {
    case 'heute': return [start, new Date(+start + day)];
    case 'gestern': return [new Date(+start - day), start];
    case '7': return [new Date(+start - 6 * day), new Date(+start + day)];
    case 'frei': return [r.von ? new Date(r.von + 'T00:00') : new Date(0), r.bis ? new Date(+new Date(r.bis + 'T00:00') + day) : new Date(+start + day)];
    default: return [new Date(0), new Date(+start + day)];
  }
}
async function loadMoves() {
  ui.report.loading = true; ui.report.err = null; renderView();
  try { ui.report.moves = (await api('moves')).bewegungen; } catch (e) { ui.report.err = e.message; }
  ui.report.loading = false; renderView();
}
function renderReport(m) {
  const r = ui.report;
  if (!r.moves && !r.loading && !r.err) { loadMoves(); return; }
  const [von, bis] = rangeBounds();
  const moves = r.moves || [];
  const saleVg = new Set(moves.filter(x => x.art === 'Verkauf' || x.art === 'Rückgabe').map(x => x.vorgang));
  const rows = moves.filter(x => {
    const z = new Date(x.zeit); if (z < von || z >= bis) return false;
    if (x.art === 'Verkauf' || x.art === 'Rückgabe') return true;
    const mm = /^Storno von (.+)$/.exec(x.notiz || '');
    return x.art === 'Storno' && mm && saleVg.has(mm[1]);
  });
  const per = {}, perDay = {}, perOrt = {};
  rows.forEach(x => {
    const n = -x.menge, a = byNr(x.nr), p = (a && a.preis) || 0, dk = dayKey(new Date(x.zeit));
    per[x.nr] = per[x.nr] || { n: 0, u: 0 }; per[x.nr].n += n; per[x.nr].u += n * p;
    perDay[dk] = perDay[dk] || { n: 0, u: 0 }; perDay[dk].n += n; perDay[dk].u += n * p;
    perOrt[x.ort] = perOrt[x.ort] || { n: 0, u: 0 }; perOrt[x.ort].n += n; perOrt[x.ort].u += n * p;
  });
  const sumN = Object.values(per).reduce((s, v) => s + v.n, 0), sumU = Object.values(per).reduce((s, v) => s + v.u, 0);
  const verk = new Set(rows.filter(x => x.art === 'Verkauf').map(x => x.vorgang)).size;
  const top = Object.entries(per).filter(([, v]) => v.n).sort((a, b) => b[1].n - a[1].n);
  const days = Object.entries(perDay).sort((a, b) => a[0] < b[0] ? 1 : -1);
  const nd = needs();
  const ranges = [['heute', 'Heute'], ['gestern', 'Gestern'], ['7', '7 Tage'], ['saison', 'Saison'], ['frei', 'Zeitraum']];
  m.innerHTML = `<div class="page"><div class="wrap">
    <h2>Bericht</h2>
    <div class="toolbar">
      <div class="seg">${ranges.map(([k, l]) => `<button data-range="${k}" aria-pressed="${r.range === k}">${l}</button>`).join('')}</div>
      ${r.range === 'frei' ? `<input type="date" id="rVon" value="${r.von}"> bis <input type="date" id="rBis" value="${r.bis}">` : ''}
      <button class="btn" data-act="reloadMoves">${r.loading ? 'Lädt …' : 'Aktualisieren'}</button>
    </div>
    ${r.err ? `<div class="err">${esc(r.err)}</div>` : ''}
    ${queue.length ? `<p class="note">${queue.length} Buchung(en) sind noch nicht beim Server und fehlen hier.</p>` : ''}
    <div class="cards">
      <div class="card"><div class="k">Stück verkauft</div><div class="v">${sumN}</div></div>
      <div class="card"><div class="k">Umsatz</div><div class="v">${eur(sumU)}</div></div>
      <div class="card"><div class="k">Verkäufe</div><div class="v">${verk}</div></div>
      ${Object.entries(perOrt).map(([o, v]) => `<div class="card"><div class="k">${esc(o)}</div><div class="v">${eur(v.u)}</div></div>`).join('')}
    </div>
    <h3>Verkaufte Artikel</h3>
    ${top.length ? `<div class="tablewrap"><table><thead><tr><th>Nr</th><th>Artikel</th><th class="n">Stück</th><th class="n">Umsatz</th></tr></thead><tbody>
      ${top.map(([nr, v]) => { const a = byNr(nr); return `<tr><td class="num">${esc(nr)}</td><td>${esc(a ? a.kurz : nr)}</td><td class="n">${v.n}</td><td class="n">${eur(v.u)}</td></tr>`; }).join('')}
    </tbody></table></div>` : `<p class="note">${r.loading ? 'Lädt …' : 'Im gewählten Zeitraum nichts verkauft.'}</p>`}
    ${days.length > 1 ? `<h3>Nach Tagen</h3><div class="tablewrap"><table><thead><tr><th>Tag</th><th class="n">Stück</th><th class="n">Umsatz</th></tr></thead><tbody>
      ${days.map(([d, v]) => `<tr><td>${new Date(d + 'T12:00').toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit', year: 'numeric' })}</td><td class="n">${v.n}</td><td class="n">${eur(v.u)}</td></tr>`).join('')}
    </tbody></table></div>` : ''}
    <h3>Unter Mindestbestand</h3>
    <p class="note">${nd.bestellen.length ? `<b>Bestellen:</b> ${nd.bestellen.map(a => esc(a.nr + ' ' + a.kurz)).join(', ')}` : 'Lager: alles über Mindestbestand.'}<br>
    ${nd.nachliefern.length ? `<b>Nachliefern:</b> ${nd.nachliefern.map(x => esc(x.a.nr + ' ' + x.a.kurz + (staende().length > 1 ? ' (' + x.ort + ')' : ''))).join(', ')}` : 'Stand: alles über Mindestbestand.'}</p>
    <p class="note">Umsatz = verkaufte Stück × aktueller Preis aus dem Sheet. Es werden keine Einzelbeträge gespeichert.</p>
  </div></div>`;
  const rv = $('rVon'), rb = $('rBis');
  if (rv) rv.addEventListener('change', e => { r.von = e.target.value; renderView(); });
  if (rb) rb.addEventListener('change', e => { r.bis = e.target.value; renderView(); });
}

/* ---------- Einrichtung ---------- */
function renderSetup() {
  const s = ui.setup;
  const root = $('root');
  if (s.step === 'ort') {
    root.innerHTML = `<div class="setup"><form id="fOrt"><h1>Welcher Stand?</h1><p class="note">Dieses Gerät bucht Verkäufe für:</p>
      <label>Marktstand<select id="sOrt">${s.orte.map(o => `<option>${esc(o.name)}</option>`).join('')}</select></label>
      <button class="btn primary">Fertig</button></form></div>`;
    $('fOrt').addEventListener('submit', e => { e.preventDefault(); finishSetup($('sOrt').value); });
    return;
  }
  root.innerHTML = `<div class="setup"><form id="fSetup" autocomplete="off">
    <h1>Magic-Sterne</h1>
    <p class="note">Gerät einrichten. Die Adresse kommt aus dem Einrichtungslink, die PIN vom Inhaber.</p>
    ${s.err ? `<div class="err">${esc(s.err)}</div>` : ''}
    <label>Adresse der Schnittstelle<input id="sApi" value="${esc(s.api)}" placeholder="https://script.google.com/macros/s/…/exec" ${s.api ? 'readonly' : ''}></label>
    <label>Name dieses Geräts<input id="sGeraet" value="${esc(s.geraet || 'Tablet')}" maxlength="30"></label>
    <label>PIN<input id="sPin" class="pinin" type="password" inputmode="numeric" pattern="[0-9]*" autocomplete="off" required></label>
    <button class="btn primary" ${s.busy ? 'disabled' : ''}>${s.busy ? 'Prüfe …' : 'Anmelden'}</button>
  </form></div>`;
  $('fSetup').addEventListener('submit', async e => {
    e.preventDefault();
    const c = { api: $('sApi').value.trim(), pin: $('sPin').value.trim(), geraet: $('sGeraet').value.trim() || 'Gerät' };
    s.api = c.api; s.geraet = c.geraet;
    if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(c.api)) { s.err = 'Die Adresse sieht nicht wie eine Apps-Script-Adresse aus (…/exec).'; renderSetup(); return; }
    s.busy = true; s.err = null; renderSetup();
    try {
      const j = await api('ping', {}, c);
      s.pending = Object.assign(c, { rolle: j.rolle });
      s.orte = (j.orte || []).filter(o => o.aktiv && !/^lager$/i.test(o.typ));
      s.busy = false;
      if (j.rolle === 'verkauf' && s.orte.length > 1) { s.step = 'ort'; renderSetup(); }
      else finishSetup((s.orte[0] || {}).name);
    } catch (err) {
      s.busy = false; s.err = err.code === 'netz' ? 'Keine Verbindung. Ist das Gerät online?' : err.message; renderSetup();
    }
  });
}
function finishSetup(ort) {
  cfg = Object.assign(ui.setup.pending, { ort });
  LS.set('cfg', cfg);
  ui.view = 'kasse';
  render(); sync();
}

/* ---------- Einstellungen ---------- */
function openSettings() {
  const L = $('layer');
  const last = ui.lastSync ? new Date(ui.lastSync).toLocaleString('de-DE') : 'noch nie';
  L.innerHTML = `<div class="scrim" id="scrim"><div class="pad" style="width:min(420px,100%)" role="dialog" aria-label="Einstellungen">
    <h3>Einstellungen</h3>
    <p class="note">Gerät: <b>${esc(cfg.geraet)}</b> · Rolle: <b>${isInhaber() ? 'Inhaber' : 'Verkauf'}</b><br>
      ${staende().length ? `Stand: <b>${esc(cfg.ort || '–')}</b><br>` : ''}
      Letzter Abgleich: ${last}<br>Nicht gesendet: ${queue.length} · Abgelehnt: ${failed.length}<br>App-Version ${APP_VERSION}</p>
    ${ui.syncErr ? `<div class="err">${esc(ui.syncErr.message)}</div>` : ''}
    ${!isInhaber() && staende().length > 1 ? `<label class="note">Stand wechseln <select id="setOrt">${staende().map(o => `<option ${o.name === cfg.ort ? 'selected' : ''}>${esc(o.name)}</option>`).join('')}</select></label>` : ''}
    <div class="acts" style="margin-top:12px">
      <button class="btn primary" data-act="syncNow">Jetzt abgleichen</button>
      <button class="btn" data-act="logout">${queue.length ? 'Abmelden (gesperrt)' : 'Abmelden'}</button>
      <button class="btn" data-act="closeLayer" style="grid-column:1/-1">Schließen</button>
    </div>
    ${queue.length ? `<p class="note">Abmelden geht erst, wenn alle Buchungen gesendet sind.</p>` : ''}
  </div></div>`;
  const so = $('setOrt'); if (so) so.addEventListener('change', e => { cfg.ort = e.target.value; LS.set('cfg', cfg); render(); });
}

/* ---------- Ziffernfeld, Toast, Bestätigung ---------- */
let toastT;
function toast(msg, bad) {
  clearTimeout(toastT);
  document.querySelector('.toast')?.remove();
  const t = document.createElement('div'); t.className = 'toast' + (bad ? ' bad' : ''); t.textContent = msg;
  document.body.append(t); toastT = setTimeout(() => t.remove(), bad ? 6000 : 2600);
}
function armed(btn, label, fn) { // zweiter Tipp bestätigt, verfällt nach 3 s
  if (btn.dataset.armed) { fn(); return; }
  btn.dataset.armed = '1'; const old = btn.textContent; btn.textContent = label; btn.classList.add('arm');
  setTimeout(() => { if (btn.isConnected) { delete btn.dataset.armed; btn.textContent = old; btn.classList.remove('arm'); } }, 3000);
}
function openPad(nr) {
  const a = byNr(nr); let v = '';
  const L = $('layer');
  L.innerHTML = `<div class="scrim" id="scrim"><div class="pad" role="dialog" aria-label="Menge eingeben"><h3>${esc(a ? a.kurz : nr)}</h3>
   <div class="pv" id="pv">${ui.cart.get(nr) || 0}</div><div class="keys">${[1, 2, 3, 4, 5, 6, 7, 8, 9].map(k => `<button data-k="${k}">${k}</button>`).join('')}
   <button data-k="del" aria-label="löschen">⌫</button><button data-k="0">0</button><button class="ok" data-k="ok">OK</button></div></div></div>`;
  L.querySelector('.pad').addEventListener('click', e => {
    const k = e.target.closest('[data-k]')?.dataset.k; if (!k) return;
    e.stopPropagation();
    if (k === 'ok') { if (v !== '') setQty(nr, +v); L.innerHTML = ''; return; }
    if (k === 'del') v = v.slice(0, -1); else if (v.length < 3) v = (v + k).replace(/^0+(?=\d)/, '');
    $('pv').textContent = v === '' ? '0' : v;
  });
}

/* ---------- Ereignisse ---------- */
document.addEventListener('click', e => {
  const t = e.target.closest('button,[data-dropfail]');
  if (e.target.id === 'scrim') { $('layer').innerHTML = ''; return; }
  if (ui.pop && !e.target.closest('#pop') && !t?.dataset.pop) { ui.pop = null; renderFilters(); }
  if (!t) return;
  const d = t.dataset;
  if (d.view) { ui.view = d.view; ui.pop = null; render(); return; }
  if (d.tab) { ui.tab = d.tab; ui.filters = {}; ui.pop = null; render(); return; }
  if (d.add) return addToCart(d.add);
  if (d.inc) return setQty(d.inc, (ui.cart.get(d.inc) || 0) + 1);
  if (d.dec) return setQty(d.dec, (ui.cart.get(d.dec) || 0) - 1);
  if (d.pad) return openPad(d.pad);
  if (d.pop) { ui.pop = ui.pop === d.pop ? null : d.pop; return renderFilters(); }
  if (d.f) { const s = ui.filters[d.f] = ui.filters[d.f] || new Set(); s.has(d.v) ? s.delete(d.v) : s.add(d.v); renderFilters(); return renderGrid(); }
  if (d.fclear) { ui.filters[d.fclear] = new Set(); ui.pop = null; renderFilters(); return renderGrid(); }
  if (d.sort) { ui.sort = d.sort; renderFilters(); return renderGrid(); }
  if (d.m) { ui.mode = d.m; return renderCart(); }
  if (d.range) { ui.report.range = d.range; return renderView(); }
  if (d.dropfail !== undefined) { failed.splice(+d.dropfail, 1); LS.set('failed', failed); renderView(); return renderStatus(); }
  switch (d.act) {
    case 'undo': { const u = ui.undo.pop(); if (u) { u.before ? ui.cart.set(u.nr, u.before) : ui.cart.delete(u.nr); } renderGrid(); renderZbar(); return renderCart(); }
    case 'clear': return armed(t, 'Wirklich leeren?', () => { ui.cart.clear(); ui.undo = []; renderGrid(); renderZbar(); renderCart(); });
    case 'go': return finishSale();
    case 'revert': return armed(t, 'Sicher? Nochmal tippen', revertSale);
    case 'ctoggle': ui.open = !ui.open; return renderCart();
    case 'settings': return openSettings();
    case 'closeLayer': $('layer').innerHTML = ''; return;
    case 'syncNow': $('layer').innerHTML = ''; ui.report.moves = null; return sync();
    case 'retry': return sync();
    case 'logout':
      if (queue.length) { toast('Erst abmelden, wenn alles gesendet ist.', true); return; }
      return armed(t, 'Wirklich abmelden?', () => { LS.del('cfg'); LS.del('state'); LS.del('lastSale'); cfg = null; st = null; $('layer').innerHTML = ''; ui.setup = { api: '', step: 'pin', orte: [] }; render(); });
    case 'overNur': ui.over.nur = !ui.over.nur; return renderView();
    case 'suggest': arts().forEach(a => { const s = suggestion(a); if (s) ui.book.qty[a.nr] = s; }); return renderView();
    case 'suggestInv': arts().forEach(a => { ui.book.qty[a.nr] = Math.max(0, stockAt(a.nr, ui.book.ort)); }); return renderView();
    case 'bookClear': ui.book.qty = {}; return renderView();
    case 'bookGo': return bookGo();
    case 'reloadMoves': ui.report.moves = null; return loadMoves();
  }
});
document.addEventListener('change', async e => {
  const nr = e.target.dataset && e.target.dataset.photo;
  if (!nr || !e.target.files[0]) return;
  if (!navigator.onLine) { toast('Fotos gehen nur mit Netz.', true); return; }
  toast('Foto wird hochgeladen …');
  try {
    const data = await resizeImage(e.target.files[0]);
    await api('setImage', { nr, data });
    imgs[nr] = data;
    try { const c = await caches.open('ms-bilder'); await c.put('bild/' + encodeURIComponent(nr), new Response(data)); } catch (_) { /* egal */ }
    toast('Foto gespeichert');
    sync();
  } catch (err) { toast('Foto nicht gespeichert: ' + err.message, true); }
});
window.addEventListener('online', () => { ui.online = true; sync(); });
window.addEventListener('offline', () => { ui.online = false; renderStatus(); });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible') { sync(); wake(); } });

/* Bildschirm an lassen, solange die Kasse offen ist */
let wakeLock = null;
async function wake() { try { if ('wakeLock' in navigator && !wakeLock) { wakeLock = await navigator.wakeLock.request('screen'); wakeLock.addEventListener('release', () => { wakeLock = null; }); } } catch (e) { /* nicht erlaubt */ } }

/* ---------- Start ---------- */
(async function start() {
  const m = /[#&]setup=([^&]+)/.exec(location.hash);
  if (m) {
    ui.setup.api = decodeURIComponent(m[1]);
    history.replaceState(null, '', location.pathname + location.search);
    if (cfg && cfg.api !== ui.setup.api && !queue.length) { cfg = null; LS.del('cfg'); LS.del('state'); st = null; }
  }
  await initImages();
  render();
  if (cfg) sync();
  scheduleSync();
  wake();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
    let first = !navigator.serviceWorker.controller;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (first) { first = false; return; }
      if (!ui.cart.size) location.reload(); else ui.reloadPending = true;
    });
  }
})();
