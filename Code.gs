/**
 * Code.gs — the spreadsheet backend for Tally.
 *
 * DEPLOY: Apps Script > paste Code.local.gs (made by `node make-local.js`,
 * it carries the real token and sheet id) > Deploy > New deployment > Web app
 * > execute as ME > access ANYONE WITH THE LINK. Paste the /exec url and the
 * token into the app under Settings > Data > Spreadsheet, press Test.
 *
 * It runs as YOU, so it opens the spreadsheet with your own permission and
 * the app never sees a Google credential. The token is the only thing
 * standing between the url and your sheet — make it long and random.
 *
 * DELIBERATELY GENERIC: the app sends tab names, headers and rows; this
 * writes them, and reads every tab back the same way. Adding a column to
 * Tally never requires redeploying this.
 */

var BACKEND_VERSION = 3;

/* ── configure ─────────────────────────────────────────────────────────── */
var SHEET_ID = 'PUT_YOUR_SPREADSHEET_ID_HERE';
var TOKEN    = 'PUT_A_LONG_RANDOM_STRING_HERE';
/* ──────────────────────────────────────────────────────────────────────── */

function doGet(e)  { return handle(e, null); }
function doPost(e) {
  var body = null;
  try { body = JSON.parse(e.postData.contents); } catch (err) { body = null; }
  return handle(e, body);
}

function handle(e, body) {
  var p = (e && e.parameter) || {};
  var token = (body && body.token) || p.token;
  if (token !== TOKEN) return json({ ok: false, error: 'bad token' });

  var action = (body && body.action) || p.action || 'ping';
  try {
    if (action === 'ping') return json({ ok: true, backend: BACKEND_VERSION });
    if (action === 'pull') return json({ ok: true, backend: BACKEND_VERSION, tabs: pull() });
    if (action === 'push') return json({ ok: true, backend: BACKEND_VERSION, wrote: push(body) });
    if (action === 'rates') return json({ ok: true, backend: BACKEND_VERSION, rows: fetchRates(body.from, body.to) });
    return json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function book() { return SpreadsheetApp.openById(SHEET_ID); }

/* Every tab, as { name, headers, rows }. */
function pull() {
  return book().getSheets().map(function (sh) {
    var v = sh.getDataRange().getValues();
    return { name: sh.getName(), headers: v.length ? v[0] : [], rows: v.slice(1) };
  });
}

/* Each tab is rewritten whole: cleared, then headers + rows. A tab the app
   does not send is left alone. */
function push(body) {
  var tabs = (body && body.tabs) || [];
  var ss = book(), wrote = {};
  tabs.forEach(function (t) {
    var sh = ss.getSheetByName(t.name) || ss.insertSheet(t.name);
    /* index: where the tab sits (1 = first). Config asks for 1. */
    if (t.index) { ss.setActiveSheet(sh); ss.moveActiveSheet(t.index); }
    sh.clearContents();
    var rows = [t.headers].concat(t.rows || []);
    if (rows.length && t.headers.length) sh.getRange(1, 1, rows.length, t.headers.length).setValues(rows);
    sh.setFrozenRows(1);
    wrote[t.name] = (t.rows || []).length;
  });
  return wrote;
}

/* ── exchange rates ─────────────────────────────────────────────────────────
   The second tab, "Rates": one row per month, the rate on the last trading
   day of that month, in USD per 1 EUR and USD per 1 AUD, from the ECB
   reference rates (api.frankfurter.dev). A month once written is never
   changed - the app's arithmetic must not move under it. On the 1st of each
   month a trigger writes the month that has just ended.                   */

var RATES_TAB = 'Rates';
var RATES_INDEX = 2;
var RATES_HEADERS = ['Month', 'USD per EUR', 'USD per AUD', 'Date'];
var SYMBOLS = ['EUR', 'AUD'];

/* From a daily series { 'YYYY-MM-DD': { EUR, AUD } } (units of currency per
   1 USD), the last day of each month, as USD per 1 unit. */
function pickMonthEnds(rates) {
  var out = {};
  Object.keys(rates || {}).sort().forEach(function (d) {
    var r = rates[d]; if (!r || !r.EUR || !r.AUD) return;
    out[d.slice(0, 7)] = { month: d.slice(0, 7), eur: round6(1 / r.EUR), aud: round6(1 / r.AUD), date: d };
  });
  return out;
}
function round6(x) { return Math.round(x * 1e6) / 1e6; }

/* Existing months win; only months not yet there are added. */
function mergeRates(existing, fresh) {
  var have = {};
  existing.forEach(function (r) { have[r.month] = r; });
  Object.keys(fresh).forEach(function (m) { if (!have[m]) have[m] = fresh[m]; });
  return Object.keys(have).sort().map(function (m) { return have[m]; });
}

/* Months up to the last one that has ENDED - the current month is not
   written until it is over. */
function lastCompleteMonth(now) {
  var d = now || new Date();
  var y = d.getFullYear(), m = d.getMonth();           // 0-based; current month excluded
  if (m === 0) { y -= 1; m = 12; }
  return y + '-' + (m < 10 ? '0' + m : '' + m);
}

function readRates() {
  var sh = book().getSheetByName(RATES_TAB);
  if (!sh) return [];
  return sh.getDataRange().getValues().slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    return { month: String(r[0]).slice(0, 7), eur: Number(r[1]), aud: Number(r[2]), date: String(r[3]).slice(0, 10) };
  });
}
function writeRates(rows) {
  var ss = book();
  var sh = ss.getSheetByName(RATES_TAB) || ss.insertSheet(RATES_TAB);
  ss.setActiveSheet(sh); ss.moveActiveSheet(RATES_INDEX);
  sh.clearContents();
  var values = [RATES_HEADERS].concat(rows.map(function (r) { return [r.month, r.eur, r.aud, r.date]; }));
  sh.getRange(1, 1, values.length, RATES_HEADERS.length).setValues(values);
  sh.setFrozenRows(1);
}

function seriesFor(start, end) {
  var url = 'https://api.frankfurter.dev/v1/' + start + '..' + end + '?base=USD&symbols=' + SYMBOLS.join(',');
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('rates service said ' + res.getResponseCode());
  return JSON.parse(res.getContentText()).rates || {};
}

/* Fill the Rates tab for the years from..to, one request per year, never
   touching a month already there. Returns every row now in the tab. Also
   makes sure the monthly trigger is installed. */
function fetchRates(from, to) {
  var y0 = Number(from) || 2013, y1 = Number(to) || new Date().getFullYear();
  var cap = lastCompleteMonth();
  var existing = readRates(), fresh = {};
  for (var y = y0; y <= y1; y++) {
    var end = y + '-12-31';
    if (String(y) > cap.slice(0, 4)) break;
    var got = pickMonthEnds(seriesFor(y + '-01-01', end));
    Object.keys(got).forEach(function (m) { if (m <= cap) fresh[m] = got[m]; });
  }
  var rows = mergeRates(existing, fresh);
  writeRates(rows);
  ensureTrigger();
  return rows;
}

/* Runs on the 1st of each month: the month that just ended, if missing. */
function monthEnd() {
  var m = lastCompleteMonth();
  var existing = readRates();
  if (existing.some(function (r) { return r.month === m; })) return;
  var got = pickMonthEnds(seriesFor(m + '-01', m + '-31'));
  if (got[m]) { var one = {}; one[m] = got[m]; writeRates(mergeRates(existing, one)); }
}
function ensureTrigger() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'monthEnd'; });
  if (!has) ScriptApp.newTrigger('monthEnd').timeBased().onMonthDay(1).atHour(6).create();
}
