/** @OnlyCurrentDoc */
/**
 * Carryon waitlist receiver.
 * Paste into the "Carryon waitlist" sheet: Extensions > Apps Script, replace everything, Save.
 * Then Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone". Copy the web app URL.
 */
const SHEET_NAME = 'Responses';
const FIELDS = ['session_id', 'started_at', 'updated_at', 'last_step', 'step_name', 'consent_variant',
  'price_variant', 'country', 'status', 'university', 'arrival', 'first_needs', 'experience',
  'home_score', 'home_score_band', 'pan', 'consent_given', 'share_report_early', 'price_choice',
  'email', 'heard_from', 'ref_in', 'ref_code', 'source'];
const NUMERIC = { last_step: true, price_variant: true };
const EMAIL_COL = FIELDS.indexOf('email') + 1;

// Text from the page is stored as plain text, never as a formula.
function clean_(key, value) {
  if (value === undefined || value === null) return '';
  if (NUMERIC[key]) { const n = Number(value); return isFinite(n) ? n : ''; }
  const s = String(value).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const id = String(data.session_id || '').replace(/[^a-z0-9]/gi, '').slice(0, 32);
    if (!id) return json_({ ok: false, error: 'missing session_id' });
    data.session_id = id;
    data.updated_at = new Date().toISOString();
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    const row = FIELDS.map(function (f) { return clean_(f, data[f]); });
    const found = sheet.getRange('A2:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (found) {
      const r = found.getRow();
      const old = sheet.getRange(r, 1, 1, FIELDS.length).getValues()[0];
      row[1] = old[1] || row[1];                       // keep the first start time
      if (Number(old[3]) > Number(row[3])) {           // keep the furthest step reached
        row[3] = old[3];
        row[4] = old[4];
      }
      sheet.getRange(r, 1, 1, FIELDS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// GET ?count=1 returns how many people have joined, for "You're number N".
function doGet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const last = sheet.getLastRow();
  let joined = 0;
  if (last > 1) {
    joined = sheet.getRange(2, EMAIL_COL, last - 1, 1).getValues()
      .filter(function (r) { return String(r[0]).indexOf('@') > 0; }).length;
  }
  return json_({ ok: true, joined: joined });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
