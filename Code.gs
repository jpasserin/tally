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

var BACKEND_VERSION = 5;

/* ── configure ─────────────────────────────────────────────────────────── */
var SHEET_ID = 'PUT_YOUR_SPREADSHEET_ID_HERE';
var TOKEN    = 'PUT_A_LONG_RANDOM_STRING_HERE';
/* ──────────────────────────────────────────────────────────────────────── */

/* RUN THIS ONCE from the editor (pick "authorize" in the dropdown, press Run)
   whenever a new version needs a permission the old one did not have. A web
   app keeps the permissions it was first granted; only a run from the editor
   shows the consent screen. This touches everything the script uses. */
function authorize() {
  UrlFetchApp.fetch('https://api.frankfurter.dev/v1/latest?base=USD&symbols=EUR', { muteHttpExceptions: true });
  ScriptApp.getProjectTriggers();
  SpreadsheetApp.openById(SHEET_ID).getName();
  return 'ok';
}

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
    if (action === 'live') return json({ ok: true, backend: BACKEND_VERSION, rows: refreshLive() });
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

/* ── live figures ──────────────────────────────────────────────────────────
   The "Live" tab: what the outside world last said, written the moment it is
   fetched - every morning by a trigger, or when the app asks (action live).
   Fund prices come from Yahoo Finance's chart endpoint, one request per
   invested fund whose Config row carries a symbol in its Tags column (the
   fund's "index" in the app: OGIAX, JRLVX, 0P00005VUV.F...). I bond rates
   come from TreasuryDirect's two tables. The app reads this tab, turns each
   row into a Price mark or a rate announcement of its own, and pushes them
   into the record tabs; this tab is the landing strip, never the record.
   Rows: Kind (price | ibond), Account, Fund, Date, Price, Currency, Symbol,
   Name, Fixed, Fetched. A price row is one per account, fund and market day;
   an ibond row one per announcement month, Date = YYYY-MM, Price = the
   semiannual inflation rate, Fixed = the fixed rate for new bonds.        */

var LIVE_TAB = 'Live';
var LIVE_HEADERS = ['Kind', 'Account', 'Fund', 'Date', 'Price', 'Currency', 'Symbol', 'Name', 'Fixed', 'Fetched'];

function readLive() {
  var sh = book().getSheetByName(LIVE_TAB);
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var h = v[0].map(String);
  return v.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; h.forEach(function (k, i) { o[k.toLowerCase()] = r[i]; });
    o.date = o.date instanceof Date ? o.date.toISOString().slice(0, 10) : String(o.date || '');
    if (o.kind === 'ibond') o.date = o.date.slice(0, 7);
    return o;
  });
}
function writeLive(rows) {
  var ss = book();
  var sh = ss.getSheetByName(LIVE_TAB) || ss.insertSheet(LIVE_TAB);
  sh.clearContents();
  var values = [LIVE_HEADERS].concat(rows.map(function (o) { return [o.kind, o.account || '', o.fund || '', o.date, o.price, o.currency || '', o.symbol || '', o.name || '', o.fixed == null ? '' : o.fixed, o.fetched || '']; }));
  sh.getRange(1, 1, values.length, LIVE_HEADERS.length).setValues(values);
  sh.setFrozenRows(1);
}
var liveKey = function (o) { return o.kind === 'ibond' ? 'ibond|' + o.date : 'price|' + o.account + '|' + o.fund + '|' + o.date; };

/* The invested funds and their symbols, from Config's fund rows (Kind fund,
   Name, Type invested|cash, Tags = the symbol, Account). */
function fundSymbols() {
  var sh = book().getSheetByName('Config');
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var h = v[0].map(String);
  var ix = function (n) { return h.indexOf(n); };
  var out = [];
  v.slice(1).forEach(function (r) {
    if (String(r[ix('Kind')]) !== 'fund' || String(r[ix('Type')]) !== 'invested') return;
    var sym = String(r[ix('Tags')] || '').trim();
    if (sym) out.push({ account: String(r[ix('Account')]), fund: String(r[ix('Name')]), symbol: sym });
  });
  return out;
}

function quote(symbol) {
  var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(symbol) + '?range=5d&interval=1d';
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (res.getResponseCode() !== 200) return null;
  var j = JSON.parse(res.getContentText());
  var m = j && j.chart && j.chart.result && j.chart.result[0] && j.chart.result[0].meta;
  if (!m || !m.regularMarketPrice) return null;
  return { price: Math.round(m.regularMarketPrice * 10000) / 10000, currency: m.currency, name: m.longName || m.shortName || symbol,
           date: new Date(m.regularMarketTime * 1000).toISOString().slice(0, 10) };
}

var TD_URL = 'https://treasurydirect.gov/savings-bonds/i-bonds/i-bonds-interest-rates/';
var TD_MONTHS = { january: '01', february: '02', march: '03', april: '04', may: '05', june: '06', july: '07', august: '08', september: '09', october: '10', november: '11', december: '12' };
function tdText(s) { return s.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim(); }
function tdRows(table) {
  var rows = [], re = /<tr[\s\S]*?<\/tr>/g, m;
  while ((m = re.exec(table))) {
    var cells = [], ce = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g, c;
    while ((c = ce.exec(m[0]))) cells.push(tdText(c[1]));
    rows.push(cells);
  }
  return rows;
}
function tdMonth(s) { var m = /([A-Za-z]+)\.?\s+\d{1,2},\s+(\d{4})/.exec(s); return m && TD_MONTHS[m[1].toLowerCase()] ? m[2] + '-' + TD_MONTHS[m[1].toLowerCase()] : null; }
function tdPct(s) { var m = /(-?\d+(?:\.\d+)?)\s*%/.exec(s); return m ? Number(m[1]) : null; }
/* every announcement: [{ month, inflation, fixed }] */
function ibondRates() {
  var res = UrlFetchApp.fetch(TD_URL, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (res.getResponseCode() !== 200) return [];
  var html = res.getContentText();
  var tables = [], re = /<table[\s\S]*?<\/table>/g, m;
  while ((m = re.exec(html))) tables.push(tdRows(m[0]));
  var find = function (word) { for (var i = 0; i < tables.length; i++) { var r0 = tables[i][0]; if (r0 && r0[0] && r0[0].toLowerCase().indexOf(word) >= 0) return tables[i].slice(1); } return null; };
  var fixedRows = find('date the fixed rate was set'), inflRows = find('date the inflation rate was set');
  if (!fixedRows || !inflRows) return [];
  var fixed = {}, infl = {};
  fixedRows.forEach(function (r) { var mo = tdMonth(r[0]), v = tdPct(r[1] || ''); if (mo && v != null) fixed[mo] = v; });
  inflRows.forEach(function (r) { var mo = tdMonth(r[0]), v = tdPct(r[1] || ''); if (mo && v != null) infl[mo] = v; });
  return Object.keys(infl).sort().map(function (mo) { return { month: mo, inflation: infl[mo], fixed: fixed[mo] == null ? 0 : fixed[mo] }; });
}

/* Fetch everything, write the Live tab, return its rows. Also makes sure the
   daily trigger is installed. */
function refreshLive() {
  var now = new Date().toISOString().slice(0, 16).replace('T', ' ');
  var have = {}, rows = readLive();
  rows.forEach(function (o) { have[liveKey(o)] = o; });
  fundSymbols().forEach(function (f) {
    var q = quote(f.symbol); if (!q) return;
    var o = { kind: 'price', account: f.account, fund: f.fund, date: q.date, price: q.price, currency: q.currency, symbol: f.symbol, name: q.name, fixed: '', fetched: now };
    have[liveKey(o)] = o;
  });
  ibondRates().forEach(function (r) {
    var o = { kind: 'ibond', account: '', fund: '', date: r.month, price: r.inflation, currency: '', symbol: '', name: 'TreasuryDirect', fixed: r.fixed, fetched: now };
    var k = liveKey(o); if (!have[k] || have[k].price !== o.price || Number(have[k].fixed) !== o.fixed) have[k] = o;
  });
  var all = Object.keys(have).map(function (k) { return have[k]; }).sort(function (a, b) { return (a.kind + a.date + a.account + a.fund) < (b.kind + b.date + b.account + b.fund) ? -1 : 1; });
  writeLive(all);
  ensureLiveTrigger();
  return all;
}
/* Every morning. */
function dailyLive() { refreshLive(); }
function ensureLiveTrigger() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dailyLive'; });
  if (!has) ScriptApp.newTrigger('dailyLive').timeBased().everyDays(1).atHour(7).create();
}
