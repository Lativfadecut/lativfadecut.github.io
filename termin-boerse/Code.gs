// Lativ Fade Cut – Termin-Börse (Google Apps Script)
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
  sh.appendRow(["id", "erstellt", "vorname", "datum", "uhrzeit", "leistung", "telefon_abgeber", "status", "uebernommen_von", "telefon_uebernehmer", "uebernommen_am"]);
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
    o.datum = d_(o.datum); o.uhrzeit = t_(o.uhrzeit);
    return o;
  });
}

// Einträge, deren Termin älter als 14 Tage ist, werden gelöscht
function cleanup_(sh) {
  const limit = Utilities.formatDate(new Date(Date.now() - 14 * 864e5), "Europe/Berlin", "yyyy-MM-dd");
  rows_(sh).filter(function (r) { return r.datum < limit; })
    .map(function (r) { return r._row; }).reverse()
    .forEach(function (row) { sh.deleteRow(row); });
}

function doGet(e) {
  const sh = sheet_(), p = (e && e.parameter) || {}, today = today_();
  const all = rows_(sh);
  if (p.action === "admin") {
    if (String(p.pin) !== String(PIN)) return out_({ ok: false, error: "pin" });
    return out_({ ok: true, rows: all.reverse().map(function (r) {
      return { vorname: r.vorname, datum: r.datum, uhrzeit: r.uhrzeit, leistung: r.leistung, status: r.status,
               uebernommen_von: r.uebernommen_von, telefon_abgeber: String(r.telefon_abgeber), telefon_uebernehmer: String(r.telefon_uebernehmer) };
    }) });
  }
  return out_({ ok: true, offers: all.filter(function (r) { return r.datum >= today; }).map(function (r) {
    return { id: String(r.id), vorname: r.vorname, datum: r.datum, uhrzeit: r.uhrzeit, leistung: r.leistung, vergeben: r.status === "vergeben" };
  }) });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || "{}"), sh = sheet_(), today = today_();
    if (b.action === "create") {
      if (b.website) return out_({ ok: true }); // Spam-Schutz
      const v = clean_(b.vorname, 30), d = clean_(b.datum, 10), t = clean_(b.uhrzeit, 5), l = clean_(b.leistung, 60), tel = clean_(b.telefon, 25);
      if (!v || !/^\d{4}-\d{2}-\d{2}$/.test(d) || !/^\d{2}:\d{2}$/.test(t) || tel.replace(/\D/g, "").length < 6) return out_({ ok: false, error: "daten" });
      if (d < today) return out_({ ok: false, error: "datum" });
      cleanup_(sh);
      const offen = rows_(sh).filter(function (r) { return r.status !== "vergeben" && r.datum >= today; }).length;
      if (offen >= MAX_OFFEN) return out_({ ok: false, error: "voll" });
      sh.appendRow([Utilities.getUuid().slice(0, 8), new Date(), v, "'" + d, "'" + t, l, "'" + tel, "offen", "", "", ""]);
      return out_({ ok: true });
    }
    if (b.action === "take") {
      const n = clean_(b.name, 40), tel = clean_(b.telefon, 25);
      if (!n || tel.replace(/\D/g, "").length < 6) return out_({ ok: false, error: "daten" });
      const r = rows_(sh).filter(function (x) { return String(x.id) === String(b.id); })[0];
      if (!r) return out_({ ok: false, error: "nicht gefunden" });
      if (r.status === "vergeben") return out_({ ok: false, error: "vergeben" });
      sh.getRange(r._row, 8, 1, 4).setValues([["vergeben", n, "'" + tel, new Date()]]);
      return out_({ ok: true, vorname: r.vorname, telefon: String(r.telefon_abgeber) });
    }
    return out_({ ok: false, error: "aktion" });
  } finally {
    lock.releaseLock();
  }
}
