/** @OnlyCurrentDoc */
/**
 * Carryon waitlist receiver.
 * Paste into the "Carryon waitlist" sheet: Extensions > Apps Script. Select everything in the editor,
 * delete it, paste this whole file (nothing around it), and save.
 * Then Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy. The URL stays the same.
 */
const SHEET_NAME = 'Responses';
// Order matters: it is the column order of the Responses tab. Add new columns only at the end.
const FIELDS = ['session_id', 'started_at', 'updated_at', 'last_step', 'step_name', 'consent_variant',
  'price_variant', 'country', 'status', 'university', 'arrival', 'first_needs', 'experience',
  'home_score', 'home_score_band', 'pan', 'consent_given', 'share_report_early', 'price_choice',
  'email', 'heard_from', 'ref_in', 'ref_code', 'source',
  'qualified', 'test', 'deposit_link_opened', 'deposit_paid', 'report_sent', 'pan_entered',
  'pan_valid', 'rid', 'classmate'];
const NUMERIC = { last_step: true, price_variant: true };
// Columns the team fills in by hand. The page never sends them, so an update keeps what is there.
const TEAM_FIELDS = { deposit_paid: true, report_sent: true };
const EMAIL_COL = FIELDS.indexOf('email') + 1;
// Anything shaped like a PAN is dropped, as a second guard: the page never sends one.
const PAN_RE = /[A-Z]{5}[0-9]{4}[A-Z]/i;

// Text from the page is stored as plain text, never as a formula.
function clean_(key, value) {
  if (value === undefined || value === null) return '';
  if (NUMERIC[key]) { const n = Number(value); return isFinite(n) ? n : ''; }
  const s = String(value).slice(0, 500);
  if (PAN_RE.test(s)) return '';
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
    if (sheet.getMaxColumns() < FIELDS.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), FIELDS.length - sheet.getMaxColumns());
    }
    const row = FIELDS.map(function (f) { return TEAM_FIELDS[f] ? '' : clean_(f, data[f]); });
    const found = sheet.getRange('A2:A').createTextFinder(id).matchEntireCell(true).findNext();
    if (found) {
      const r = found.getRow();
      const old = sheet.getRange(r, 1, 1, FIELDS.length).getValues()[0];
      row[1] = old[1] || row[1];                       // keep the first start time
      if (Number(old[3]) > Number(row[3])) {           // keep the furthest step reached
        row[3] = old[3];
        row[4] = old[4];
      }
      FIELDS.forEach(function (f, i) { if (TEAM_FIELDS[f]) row[i] = old[i]; });
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
