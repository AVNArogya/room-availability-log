/**
 * Room Availability Log — Google Sheets backend (Google Apps Script).
 *
 * Setup:
 *   1. Create a Google Sheet, then Extensions → Apps Script, paste this file.
 *   2. Run `setup()` once (authorise when asked). It creates the Log, Lists and Rates tabs.
 *   3. Project Settings → Script properties → add APP_KEY = <any secret string>.
 *   4. Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone.
 *   5. Put the /exec URL and the APP_KEY in web/config.js.
 */

var LOG_SHEET = 'Room availability log';
var LISTS_SHEET = 'Lists';
var RATES_SHEET = 'Rates';

// Same columns as the Excel log, plus bookkeeping columns at the end.
var FIELDS = [
  ['enquiryDate', 'Enquiry date'],
  ['patientName', 'Patient name'],
  ['phone', 'Phone'],
  ['uhid', 'UHID (if registered)'],
  ['newOrReturning', 'New / returning'],
  ['city', 'City / district'],
  ['source', 'Source'],
  ['condition', 'Condition'],
  ['doctor', 'Advising doctor'],
  ['plannedStay', 'Planned stay (days)'],
  ['categoryAsked', 'Room category asked for'],
  ['dateWanted', 'Date patient wanted'],
  ['offeredInstead', 'Offered instead'],
  ['earliestDate', 'Earliest date we could give'],
  ['waitDays', 'Wait (days)'],
  ['status', 'Status'],
  ['admissionDate', 'Admission date'],
  ['categoryAdmitted', 'Category admitted to'],
  ['reasonNotAdmitted', 'Reason not admitted'],
  ['wentWhere', 'Where they went (if known)'],
  ['lastFollowUp', 'Last follow-up'],
  ['remarks', 'Remarks'],
  ['enteredBy', 'Entered by'],
  ['patientDaysLost', 'Patient-days lost'],
  ['revenueLost', 'Revenue lost (₹)'],
  ['month', 'Month'],
  ['id', 'Entry ID'],
  ['createdAt', 'Created at'],
  ['updatedAt', 'Updated at']
];
var DATE_FIELDS = ['enquiryDate', 'dateWanted', 'earliestDate', 'admissionDate', 'lastFollowUp'];
var CALCULATED = ['waitDays', 'patientDaysLost', 'revenueLost', 'month', 'id', 'createdAt', 'updatedAt'];
// Statuses that count as a lost admission for patient-days / revenue lost.
var LOST_STATUSES = ['Lost', 'Went elsewhere', 'Cancelled'];

var DEFAULT_LISTS = {
  newOrReturning: ['New', 'Returning'],
  source: ['Incoming call - Google', 'Incoming call - Website', 'Direct counseling', 'Zoho Bookings',
           'Website', 'Google Ad words', 'Referral', 'Walk-in'],
  doctor: ['Dr. Geethu', 'Dr. Sridevi', 'Dr. Chippi', 'Dr. Biju', 'Dr. Ramesh', 'Dr. Padmavathy', 'Dr. Jyothish'],
  roomCategory: ['Economy', 'Executive AC', 'Classic Villa'],
  offeredInstead: ['Nothing available', 'Economy', 'Executive AC', 'Classic Villa', 'Outside stay (hotel)'],
  status: ['Waiting', 'No answer', 'Booked', 'Admitted', 'Postponed', 'Lost', 'Went elsewhere', 'Cancelled'],
  reasonNotAdmitted: ['Room not available', 'Waiting time too long', 'Cost', 'Postponed by patient',
                      'Went elsewhere', 'Not reachable', 'Other'],
  enteredBy: []
};

/* ---------- setup ---------- */

function setup() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET, 0);
  if (log.getLastRow() === 0) {
    log.appendRow(FIELDS.map(function (f) { return f[1]; }));
    log.getRange(1, 1, 1, FIELDS.length).setFontWeight('bold').setBackground('#e8eef7');
    log.setFrozenRows(1);
    var calcCols = CALCULATED.map(colOf);
    calcCols.forEach(function (c) { log.getRange(1, c).setBackground('#d9d9d9'); });
  }

  var lists = ss.getSheetByName(LISTS_SHEET);
  if (!lists) {
    lists = ss.insertSheet(LISTS_SHEET);
    var keys = Object.keys(DEFAULT_LISTS);
    var depth = Math.max.apply(null, keys.map(function (k) { return DEFAULT_LISTS[k].length; }));
    var grid = [keys];
    for (var r = 0; r < depth; r++) {
      grid.push(keys.map(function (k) { return DEFAULT_LISTS[k][r] || ''; }));
    }
    lists.getRange(1, 1, grid.length, keys.length).setValues(grid);
    lists.getRange(1, 1, 1, keys.length).setFontWeight('bold');
    lists.setFrozenRows(1);
  }

  var rates = ss.getSheetByName(RATES_SHEET);
  if (!rates) {
    rates = ss.insertSheet(RATES_SHEET);
    var rows = [['Room category', 'Rate per day (₹)']].concat(
      DEFAULT_LISTS.roomCategory.map(function (c) { return [c, 0]; }));
    rates.getRange(1, 1, rows.length, 2).setValues(rows);
    rates.getRange(1, 1, 1, 2).setFontWeight('bold');
    rates.setFrozenRows(1);
  }
}

/* ---------- HTTP ---------- */

function doGet(e) {
  return handle_(function () {
    checkKey_(e.parameter.key);
    var action = e.parameter.action || 'config';
    if (action === 'config') return { lists: readLists_() };
    if (action === 'list') return { entries: listEntries_(Number(e.parameter.limit) || 200) };
    throw new Error('Unknown action: ' + action);
  });
}

// The web app posts JSON as text/plain to avoid a CORS preflight.
function doPost(e) {
  return handle_(function () {
    var body = JSON.parse(e.postData.contents || '{}');
    checkKey_(body.key);
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'add') return { entry: addEntry_(body.entry) };
      if (body.action === 'update') return { entry: updateEntry_(body.entry) };
      throw new Error('Unknown action: ' + body.action);
    } finally {
      lock.releaseLock();
    }
  });
}

function handle_(fn) {
  var out;
  try {
    out = { ok: true, data: fn() };
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkKey_(key) {
  var expected = PropertiesService.getScriptProperties().getProperty('APP_KEY');
  if (expected && key !== expected) throw new Error('Invalid app key');
}

/* ---------- data ---------- */

function addEntry_(entry) {
  if (!entry || !entry.patientName) throw new Error('Patient name is required');
  var log = logSheet_();
  // clientId makes retries from the offline queue idempotent.
  if (entry.clientId) {
    var existing = findRow_(log, entry.clientId);
    if (existing) return rowToEntry_(log.getRange(existing, 1, 1, FIELDS.length).getValues()[0]);
  }
  var now = new Date();
  entry.id = entry.clientId || Utilities.getUuid();
  entry.createdAt = now;
  entry.updatedAt = now;
  var row = entryToRow_(computed_(entry));
  log.appendRow(row);
  formatRow_(log, log.getLastRow());
  return rowToEntry_(row);
}

function updateEntry_(entry) {
  if (!entry || !entry.id) throw new Error('Entry ID is required for update');
  var log = logSheet_();
  var r = findRow_(log, entry.id);
  if (!r) throw new Error('Entry not found: ' + entry.id);
  var current = rowToEntry_(log.getRange(r, 1, 1, FIELDS.length).getValues()[0]);
  FIELDS.forEach(function (f) {
    var k = f[0];
    if (CALCULATED.indexOf(k) === -1 && entry.hasOwnProperty(k)) current[k] = entry[k];
  });
  current.updatedAt = new Date();
  var row = entryToRow_(computed_(current));
  log.getRange(r, 1, 1, FIELDS.length).setValues([row]);
  formatRow_(log, r);
  return rowToEntry_(row);
}

function listEntries_(limit) {
  var log = logSheet_();
  var last = log.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - limit + 1);
  var values = log.getRange(start, 1, last - start + 1, FIELDS.length).getValues();
  return values.map(rowToEntry_).filter(function (e) { return e.id; }).reverse();
}

function computed_(e) {
  var wanted = toDate_(e.dateWanted);
  var earliest = toDate_(e.earliestDate);
  e.waitDays = (wanted && earliest) ? Math.round((earliest - wanted) / 86400000) : '';

  var stay = parseFloat(e.plannedStay) || 0;
  var lost = LOST_STATUSES.indexOf(e.status) !== -1;
  e.patientDaysLost = lost ? stay : 0;
  var rate = readRates_()[String(e.categoryAsked || '').trim().toLowerCase()] || 0;
  e.revenueLost = e.patientDaysLost * rate;

  var enq = toDate_(e.enquiryDate);
  e.month = enq ? Utilities.formatDate(enq, Session.getScriptTimeZone(), 'MMM yyyy') : '';
  return e;
}

function entryToRow_(e) {
  return FIELDS.map(function (f) {
    var k = f[0], v = e[k];
    if (v === undefined || v === null) return '';
    if (DATE_FIELDS.indexOf(k) !== -1) return toDate_(v) || '';
    return v;
  });
}

function rowToEntry_(row) {
  var e = {};
  FIELDS.forEach(function (f, i) {
    var v = row[i];
    if (v instanceof Date) {
      v = DATE_FIELDS.indexOf(f[0]) !== -1
        ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : v.toISOString();
    }
    e[f[0]] = v;
  });
  return e;
}

function formatRow_(sheet, r) {
  DATE_FIELDS.forEach(function (k) { sheet.getRange(r, colOf(k)).setNumberFormat('dd-mmm-yyyy'); });
  sheet.getRange(r, colOf('phone')).setNumberFormat('@');
  sheet.getRange(r, colOf('revenueLost')).setNumberFormat('₹#,##0');
}

function findRow_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var ids = sheet.getRange(2, colOf('id'), last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === id) return i + 2;
  }
  return 0;
}

function readLists_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(LISTS_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return DEFAULT_LISTS;
  var grid = sheet.getDataRange().getValues();
  var out = {};
  grid[0].forEach(function (key, c) {
    if (!key) return;
    out[key] = grid.slice(1).map(function (r) { return String(r[c]).trim(); }).filter(String);
  });
  return out;
}

function readRates_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(RATES_SHEET);
  var out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0]) out[String(r[0]).trim().toLowerCase()] = Number(r[1]) || 0;
  });
  return out;
}

function logSheet_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET);
  if (!sheet) { setup(); sheet = SpreadsheetApp.getActive().getSheetByName(LOG_SHEET); }
  return sheet;
}

function colOf(key) {
  for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i][0] === key) return i + 1;
  throw new Error('No column for ' + key);
}

function toDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(v);
  return isNaN(d) ? null : d;
}
