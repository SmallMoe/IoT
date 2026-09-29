/**
 * Gym Logger backend (Google Apps Script, container-bound to the Sheet).
 * Deploy: Execute as "Me", access "Anyone".
 *
 *   GET  /exec                 -> web page
 *   GET  /exec?action=bootstrap -> JSON for the ESP32 (exercise lists + last weights)
 *   POST /exec                 -> JSON {"action":"saveSession", ...} from the ESP32
 *
 * Tabs are created automatically on first use:
 *   Sets      one row per set: Timestamp | SessionID | Date | Type | BodyWeight | Exercise | Set | Weight | Reps | Volume | Source
 *   Templates one row per exercise: Type | Exercise | Order
 */
const SETS = 'Sets', TPL = 'Templates';
const TYPES = ['Push', 'Pull', 'Leg'];
const HEAD = ['Timestamp', 'SessionID', 'Date', 'Type', 'BodyWeight', 'Exercise', 'Set', 'Weight', 'Reps', 'Volume', 'Source'];
const DEFAULTS = {
  Push: ['Chest Press', 'Shoulder Press', 'Tricep Extension'],
  Pull: ['Lat Pulldown', 'Seated Row', 'Bicep Curl'],
  Leg:  ['Squat', 'Leg Press', 'Calf Raise']
};

// ---------------- entry points ----------------

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'bootstrap') return json_(bootstrap_());
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Load Log')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.action !== 'saveSession') throw new Error('Unknown action');
    return json_(saveSession_(p, p.source === 'ESP32' ? 'ESP32' : 'Web'));
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

// ---------------- called by the web page (google.script.run) ----------------

function getData() {
  const templates = {};
  TYPES.forEach(t => templates[t] = []);
  tab_(TPL).getDataRange().getValues().slice(1)
    .filter(r => templates[r[0]])
    .sort((a, b) => a[2] - b[2])
    .forEach(r => templates[r[0]].push(String(r[1])));
  return { templates: templates, rows: rows_() };
}

function saveSessionWeb(p) { return saveSession_(p, 'Web'); }

function saveTemplate(type, list) {
  if (TYPES.indexOf(type) < 0) throw new Error('Bad type');
  const sh = tab_(TPL);
  const keep = sh.getDataRange().getValues().slice(1).filter(r => r[0] !== type);
  list.forEach((n, i) => keep.push([type, clean_(n), i + 1]));
  sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 3).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, 3).setValues(keep);
}

// ---------------- core ----------------

/** Appends one row per set. Idempotent on sessionId, so retries never duplicate. */
function saveSession_(p, source) {
  if (TYPES.indexOf(p.type) < 0 || !Array.isArray(p.sets) || !p.sets.length || p.sets.length > 200) throw new Error('Bad payload');
  const sid = String(p.sessionId || '').slice(0, 40);
  if (!sid) throw new Error('Missing sessionId');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = tab_(SETS), last = sh.getLastRow();
    if (last > 1 && sh.getRange(2, 2, last - 1, 1).createTextFinder(sid).matchEntireCell(true).findNext()) {
      return { status: 'duplicate' };
    }
    const now = new Date(), tz = Session.getScriptTimeZone();
    const date = /^\d{4}-\d{2}-\d{2}$/.test(p.date) ? p.date : Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const bw = parseFloat(p.bodyWeight), count = {};
    const rows = p.sets.map(s => {
      const ex = clean_(s.exercise), w = Number(s.weight), r = Number(s.reps);
      if (!ex || !(w >= 0) || !(r > 0)) throw new Error('Bad set');
      count[ex] = (count[ex] || 0) + 1;
      return [now, sid, date, p.type, isNaN(bw) ? '' : bw, ex, count[ex], w, r, w * r, source];
    });
    sh.getRange(last + 1, 1, rows.length, HEAD.length).setValues(rows);
    return { status: 'ok', rows: rows.length };
  } finally {
    lock.releaseLock();
  }
}

/** Exercise lists plus the top set of the most recent session per exercise (prefill for the ESP32). */
function bootstrap_() {
  const d = getData(), latest = {}, out = { status: 'ok', types: {} };
  d.rows.slice().reverse().forEach(r => {           // newest first
    const l = latest[r[5]];
    if (!l) latest[r[5]] = { sid: r[1], w: r[7], r: r[8] };
    else if (l.sid === r[1] && r[7] > l.w) { l.w = r[7]; l.r = r[8]; }
  });
  TYPES.forEach(t => out.types[t] = d.templates[t].map(n => ({
    n: n, w: (latest[n] || {}).w || 0, r: (latest[n] || {}).r || 8
  })));
  return out;
}

// ---------------- helpers ----------------

function rows_() {
  const tz = Session.getScriptTimeZone();
  return tab_(SETS).getDataRange().getValues().slice(1).filter(r => r[1] !== '').map(r => [
    new Date(r[0]).getTime(), String(r[1]),
    r[2] instanceof Date ? Utilities.formatDate(r[2], tz, 'yyyy-MM-dd') : String(r[2]),
    r[3], r[4], r[5], r[6], r[7], r[8], r[10]
  ]);
}

/** Returns the tab, creating and seeding it if missing. */
function tab_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  if (name === SETS) {
    sh.appendRow(HEAD);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sh.getRange('C:C').setNumberFormat('@');          // keep Date as plain text (no timezone shifts)
  } else {
    sh.appendRow(['Type', 'Exercise', 'Order']);
    const seed = [];
    TYPES.forEach(t => DEFAULTS[t].forEach((e, i) => seed.push([t, e, i + 1])));
    sh.getRange(2, 1, seed.length, 3).setValues(seed);
  }
  sh.setFrozenRows(1);
  return sh;
}

/** Trims text and strips leading formula characters. */
function clean_(v) { return String(v || '').trim().slice(0, 60).replace(/^[=+\-@]+/, ''); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Run once from the editor to authorise the script and create both tabs. */
function setup() { tab_(SETS); tab_(TPL); }
