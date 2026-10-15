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

// Same 26 columns as the Excel log. Each row's Entry ID is stored as hidden row
// metadata (not a column), so it moves with the row when the sheet is sorted.
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
  ['month', 'Month']
];
var ID_KEY = 'entryId';
var DATE_FIELDS = ['enquiryDate', 'dateWanted', 'earliestDate', 'admissionDate', 'lastFollowUp'];
var CALCULATED = ['waitDays', 'patientDaysLost', 'revenueLost', 'month'];
// Statuses that count as a lost admission: the whole planned stay is lost.
var LOST_STATUSES = ['Lost', 'Went elsewhere', 'Cancelled'];
// Statuses where a cheaper room than the one asked for counts as revenue lost.
var BOOKED_STATUSES = ['Booked', 'Admitted'];

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
    if (action === 'config') return { lists: readLists_(), rates: readRates_() };
    if (action === 'list') return withLock_(function () { return { entries: listEntries_(Number(e.parameter.limit) || 200) }; });
    throw new Error('Unknown action: ' + action);
  });
}

// The web app posts JSON as text/plain to avoid a CORS preflight.
function doPost(e) {
  return handle_(function () {
    var body = JSON.parse(e.postData.contents || '{}');
    checkKey_(body.key);
    return withLock_(function () {
      if (body.action === 'add') return { entry: addEntry_(body.entry) };
      if (body.action === 'update') return { entry: updateEntry_(body.entry) };
      throw new Error('Unknown action: ' + body.action);
    });
  });
}

// Listing can tag untagged rows with IDs, so reads and writes share one lock.
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
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
    if (existing) return readEntry_(log, existing, entry.clientId);
  }
  var id = entry.clientId || Utilities.getUuid();
  var row = entryToRow_(computed_(entry));
  log.appendRow(row);
  var r = log.getLastRow();
  tagRow_(log, r, id);
  formatRow_(log, r);
  return rowToEntry_(row, id);
}

function updateEntry_(entry) {
  if (!entry || !entry.id) throw new Error('Entry ID is required for update');
  var log = logSheet_();
  var r = findRow_(log, entry.id);
  if (!r) throw new Error('Entry not found: ' + entry.id);
  var current = readEntry_(log, r, entry.id);
  FIELDS.forEach(function (f) {
    var k = f[0];
    if (CALCULATED.indexOf(k) === -1 && entry.hasOwnProperty(k)) current[k] = entry[k];
  });
  var row = entryToRow_(computed_(current));
  log.getRange(r, 1, 1, FIELDS.length).setValues([row]);
  formatRow_(log, r);
  return rowToEntry_(row, entry.id);
}

function listEntries_(limit) {
  var log = logSheet_();
  var last = log.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - limit + 1);
  var values = log.getRange(start, 1, last - start + 1, FIELDS.length).getValues();
  var ids = rowIds_(log);
  var out = [];
  values.forEach(function (row, i) {
    var r = start + i;
    if (!row[colOf('patientName') - 1]) return;
    // Rows typed or pasted straight into the sheet get an ID so the app can edit them too.
    if (!ids[r]) { ids[r] = Utilities.getUuid(); tagRow_(log, r, ids[r]); }
    out.push(rowToEntry_(row, ids[r]));
  });
  return out.reverse();
}

function computed_(e) {
  var wanted = toDate_(e.dateWanted);
  var earliest = toDate_(e.earliestDate);
  e.waitDays = (wanted && earliest) ? Math.round((earliest - wanted) / 86400000) : '';

  var loss = revenueLoss_(e, readRates_());
  e.patientDaysLost = loss.days;
  e.revenueLost = loss.amount;

  var enq = toDate_(e.enquiryDate);
  e.month = enq ? Utilities.formatDate(enq, Session.getScriptTimeZone(), 'MMM yyyy') : '';
  return e;
}

// Revenue lost rules (web/app.js revenueLoss mirrors this for the on-screen estimate):
//   Lost / Went elsewhere / Cancelled: planned stay × rate of the category asked for.
//   Otherwise, admitted later than wanted: delay days × rate of the category asked for.
//   Booked / Admitted in a cheaper room: (rate asked − rate given) × planned stay.
//     The room given is "Category admitted to", or "Offered instead" when that is blank.
function revenueLoss_(e, rates) {
  var rateOf = function (c) { return rates[String(c || '').trim().toLowerCase()]; };
  var stay = parseFloat(e.plannedStay) || 0;
  var asked = rateOf(e.categoryAsked) || 0;
  if (LOST_STATUSES.indexOf(e.status) !== -1) return { days: stay, amount: stay * asked };

  var days = 0, amount = 0;
  var wanted = toDate_(e.dateWanted), admitted = toDate_(e.admissionDate);
  if (wanted && admitted && admitted > wanted) {
    days = Math.round((admitted - wanted) / 86400000);
    amount += days * asked;
  }
  if (BOOKED_STATUSES.indexOf(e.status) !== -1) {
    var given = rateOf(e.categoryAdmitted);
    if (given === undefined) given = rateOf(e.offeredInstead);
    if (given !== undefined && given < asked) amount += (asked - given) * stay;
  }
  return { days: days, amount: amount };
}

function entryToRow_(e) {
  return FIELDS.map(function (f) {
    var k = f[0], v = e[k];
    if (v === undefined || v === null) return '';
    if (DATE_FIELDS.indexOf(k) !== -1) return toDate_(v) || '';
    return v;
  });
}

function rowToEntry_(row, id) {
  var e = { id: id };
  FIELDS.forEach(function (f, i) {
    var v = row[i];
    if (v instanceof Date) {
      v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

function readEntry_(sheet, r, id) {
  return rowToEntry_(sheet.getRange(r, 1, 1, FIELDS.length).getValues()[0], id);
}

function tagRow_(sheet, r, id) {
  sheet.getRange(r + ':' + r).addDeveloperMetadata(ID_KEY, id);
}

function rowMetadata_(sheet) {
  return sheet.createDeveloperMetadataFinder()
    .withKey(ID_KEY)
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
    .find();
}

// { rowNumber: entryId } for every tagged row.
function rowIds_(sheet) {
  var out = {};
  rowMetadata_(sheet).forEach(function (m) { out[m.getLocation().getRow().getRow()] = m.getValue(); });
  return out;
}

function findRow_(sheet, id) {
  var found = sheet.createDeveloperMetadataFinder().withKey(ID_KEY).withValue(id).find();
  return found.length ? found[0].getLocation().getRow().getRow() : 0;
}

// Earlier versions wrote Entry ID / Created at / Updated at into columns AA:AC.
// Move those IDs into row metadata and clear the three columns.
function migrateIdColumns_(sheet) {
  var last = sheet.getLastRow();
  if (sheet.getMaxColumns() <= FIELDS.length || last < 1) return;
  var extra = sheet.getRange(1, FIELDS.length + 1, last, 3).getValues();
  var uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  var hasOldIds = extra[0][0] === 'Entry ID' || extra.some(function (r) { return uuid.test(String(r[0])); });
  if (!hasOldIds) return;
  var tagged = rowIds_(sheet);
  for (var i = 1; i < extra.length; i++) {
    var r = i + 1, id = String(extra[i][0]);
    if (uuid.test(id) && !tagged[r]) tagRow_(sheet, r, id);
  }
  sheet.getRange(1, FIELDS.length + 1, last, 3).clear();
}

function readLists_() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(LISTS_SHEET);
  if (!sheet || sheet.getLastRow() < 1) return DEFAULT_LISTS;
  var grid = sheet.getDataRange().getValues();
  var out = {};
  Object.keys(DEFAULT_LISTS).forEach(function (k) { out[k] = DEFAULT_LISTS[k]; });
  grid[0].forEach(function (key, c) {
    if (!key) return;
    var values = grid.slice(1).map(function (r) { return String(r[c]).trim(); }).filter(String);
    if (values.length) out[key] = values;
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
  migrateIdColumns_(sheet);
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
