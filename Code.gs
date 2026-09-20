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

var BACKEND_VERSION = 2;

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
