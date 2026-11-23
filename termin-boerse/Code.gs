// Lativ Fade Cut – Termin-Börse (nur Datum, Uhrzeit, Leistung – keine Namen, keine Nummern)
const PIN = "1234"; // ← HIER deine eigene geheime PIN eintragen
const MAX_OFFEN = 30;

function setup() { sheet_(); }

function sheet_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty("SHEET_ID");
  if (id) return SpreadsheetApp.openById(id).getSheets()[0];
  const ss = SpreadsheetApp.create("Lativ Fade Cut – Termin-Börse");
  p.setProperty("SHEET_ID", ss.getId());
  const sh = ss.getSheets()[0];
  sh.appendRow(["id", "erstellt", "datum", "uhrzeit", "leistung", "status"]);
  sh.setFrozenRows(1);
  return sh;
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function today_() { return Utilities.formatDate(new Date(), "Europe/Berlin", "yyyy-MM-dd"); }
function d_(v) { return v instanceof Date ? Utilities.formatDate(v, "Europe/Berlin", "yyyy-MM-dd") : String(v); }
function t_(v) { return v instanceof Date ? Utilities.formatDate(v, "Europe/Berlin", "HH:mm") : String(v); }
function clean_(s, n) { return String(s || "").replace(/[<>]/g, "").trim().slice(0, n); }

function rows_(sh) {
  const v = sh.getDataRange().getValues(), h = v.shift();
  return v.map(function (r, i) {
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    o.id = String(o.id); o.datum = d_(o.datum); o.uhrzeit = t_(o.uhrzeit);
    return o;
  });
}

// Vergangene Termine automatisch löschen
function cleanup_(sh) {
  const today = today_();
  rows_(sh).filter(function (r) { return r.datum < today; })
    .map(function (r) { return r._row; }).reverse()
    .forEach(function (row) { sh.deleteRow(row); });
}

function pub_(r) { return { id: r.id, datum: r.datum, uhrzeit: r.uhrzeit, leistung: r.leistung, status: r.status, vergeben: r.status === "vergeben" }; }

function doGet(e) {
  const sh = sheet_(), p = (e && e.parameter) || {}, today = today_();
  const all = rows_(sh).filter(function (r) { return r.datum >= today; })
    .sort(function (a, b) { return (a.datum + a.uhrzeit).localeCompare(b.datum + b.uhrzeit); });
  if (p.action === "admin") {
    if (String(p.pin) !== String(PIN)) return out_({ ok: false, error: "pin" });
    return out_({ ok: true, rows: all.map(pub_) });
  }
  return out_({ ok: true, offers: all.filter(function (r) { return r.status !== "vergeben"; }).map(pub_) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || "{}"), sh = sheet_(), today = today_();
    if (b.action === "create") {
      if (b.website) return out_({ ok: true }); // Spam-Schutz
      const d = clean_(b.datum, 10), t = clean_(b.uhrzeit, 5), l = clean_(b.leistung, 60);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^\d{2}:\d{2}$/.test(t) || !l) return out_({ ok: false, error: "daten" });
      if (d < today) return out_({ ok: false, error: "datum" });
      cleanup_(sh);
      const rows = rows_(sh);
      if (rows.some(function (r) { return r.datum === d && r.uhrzeit === t && r.status !== "vergeben"; })) return out_({ ok: false, error: "doppelt" });
      if (rows.filter(function (r) { return r.status !== "vergeben"; }).length >= MAX_OFFEN) return out_({ ok: false, error: "voll" });
      sh.appendRow([Utilities.getUuid().slice(0, 8), new Date(), "'" + d, "'" + t, l, "offen"]);
      return out_({ ok: true });
    }
    if (b.action === "mark" || b.action === "delete") {
      if (String(b.pin) !== String(PIN)) return out_({ ok: false, error: "pin" });
      const r = rows_(sh).filter(function (x) { return x.id === String(b.id); })[0];
      if (!r) return out_({ ok: false, error: "nicht gefunden" });
      if (b.action === "mark") sh.getRange(r._row, 6).setValue("vergeben");
      else sh.deleteRow(r._row);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: "aktion" });
  } finally {
    lock.releaseLock();
  }
}
