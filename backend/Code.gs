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

// Columns are matched by their header text in row 1, so columns can be added, moved
// or removed in the sheet without breaking the app. Matching ignores case, spaces,
// punctuation and anything in brackets ("UHID (if registered)" matches "UHID").
// Each row's Entry ID is stored as hidden row metadata, not in a column.
// [key, header used when creating a new sheet, other accepted headers...]
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
  ['offeredRoom', 'Offerred Room', 'Offered Room', 'Room offered'],
  ['dateWanted', 'Date patient wanted'],
  ['earliestDate', 'Earliest date we could give'],
  ['waitDays', 'Wait (days)'],
  ['status', 'Status'],
  ['admissionDate', 'Admission date'],
  ['categoryAdmitted', 'Admission category', 'Category admitted to'],
  ['reasonNotAdmitted', 'Reason not admitted'],
  ['wentWhere', 'Where they went (if known)'],
  ['lastFollowUp', 'Last follow-up'],
  ['remarks', 'Remarks'],
  ['enteredBy', 'Entered by'],
  ['patientDaysLost', 'Patient-days lost'],
  ['revenueLost', 'Revenue lost (₹)'],
  ['month', 'Month'],
  ['updatedAt', 'Updated at']
];
// Header row for a brand-new sheet. "Month" is still filled in if you add that column.
var DEFAULT_HEADER = FIELDS.map(function (f) { return f[0]; }).filter(function (k) { return k !== 'month'; });
var ID_KEY = 'entryId';
var DATE_FIELDS = ['enquiryDate', 'dateWanted', 'earliestDate', 'admissionDate', 'lastFollowUp'];
var CALCULATED = ['waitDays', 'patientDaysLost', 'revenueLost', 'month', 'updatedAt'];
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
  offeredRoom: ['Economy', 'Executive AC', 'Classic Villa', 'Not offered'],
  status: ['Waiting', 'Booked', 'Admitted', 'Lost'],
  reasonNotAdmitted: ['Room not available', 'Waiting time too long', 'Cost', 'Postponed by patient',
                      'Went elsewhere', 'Not reachable', 'Other'],
  enteredBy: []
};

/* ---------- setup ---------- */

function setup() {
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET, 0);
  if (log.getLastRow() === 0) {
    log.appendRow(DEFAULT_HEADER.map(function (k) { return fieldOf_(k)[1]; }));
    log.getRange(1, 1, 1, DEFAULT_HEADER.length).setFontWeight('bold').setBackground('#e8eef7');
    log.setFrozenRows(1);
    DEFAULT_HEADER.forEach(function (k, i) {
      if (CALCULATED.indexOf(k) !== -1) log.getRange(1, i + 1).setBackground('#d9d9d9');
    });
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
  var cols = columns_(log);
  var e = computed_(entry);
  e.updatedAt = new Date();
  var row = [];
  for (var c = 0; c < cols.width; c++) row.push('');
  fillRow_(row, cols, e);
  log.appendRow(row);
  var r = log.getLastRow();
  tagRow_(log, r, id);
  formatRow_(log, r, cols);
  return readEntry_(log, r, id);
}

function updateEntry_(entry) {
  if (!entry || !entry.id) throw new Error('Entry ID is required for update');
  var log = logSheet_();
  var r = findRow_(log, entry.id);
  if (!r) throw new Error('Entry not found: ' + entry.id);
  var cols = columns_(log);
  var range = log.getRange(r, 1, 1, cols.width);
  var row = range.getValues()[0];
  var formulas = range.getFormulas()[0];
  var current = rowToEntry_(row, cols, entry.id);
  FIELDS.forEach(function (f) {
    var k = f[0];
    if (CALCULATED.indexOf(k) === -1 && entry.hasOwnProperty(k)) current[k] = entry[k];
  });
  var e = computed_(current);
  e.updatedAt = new Date();
  // Keep formulas in columns the app doesn't manage.
  formulas.forEach(function (f, i) { if (f) row[i] = f; });
  fillRow_(row, cols, e);
  range.setValues([row]);
  formatRow_(log, r, cols);
  return readEntry_(log, r, entry.id);
}

function listEntries_(limit) {
  var log = logSheet_();
  var last = log.getLastRow();
  if (last < 2) return [];
  var cols = columns_(log);
  var start = Math.max(2, last - limit + 1);
  var values = log.getRange(start, 1, last - start + 1, cols.width).getValues();
  var ids = rowIds_(log);
  var out = [];
  values.forEach(function (row, i) {
    var r = start + i;
    if (!row[cols.patientName - 1]) return;
    // Rows typed or pasted straight into the sheet get an ID so the app can edit them too.
    if (!ids[r]) { ids[r] = Utilities.getUuid(); tagRow_(log, r, ids[r]); }
    out.push(rowToEntry_(row, cols, ids[r]));
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
//     The room given is "Admission category", or "Offerred Room" when that is blank.
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
    if (given === undefined) given = rateOf(e.offeredRoom);
    if (given !== undefined && given < asked) amount += (asked - given) * stay;
  }
  return { days: days, amount: amount };
}

/* ---------- columns ---------- */

function normHeader_(h) {
  return String(h || '').replace(/\([^)]*\)/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function fieldOf_(key) {
  for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i][0] === key) return FIELDS[i];
  throw new Error('Unknown field ' + key);
}

// { key: columnNumber, ..., width } read from the header row.
function columns_(sheet) {
  var width = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(normHeader_);
  var cols = { width: width };
  FIELDS.forEach(function (f) {
    var names = f.slice(1).map(normHeader_);
    for (var c = 0; c < headers.length; c++) {
      if (headers[c] && names.indexOf(headers[c]) !== -1) { cols[f[0]] = c + 1; break; }
    }
  });
  if (!cols.patientName) {
    throw new Error('Column "Patient name" not found in row 1 of the "' + LOG_SHEET + '" tab');
  }
  return cols;
}

function fillRow_(row, cols, e) {
  FIELDS.forEach(function (f) {
    var k = f[0];
    if (!cols[k]) return;
    var v = e[k];
    if (v === undefined || v === null) v = '';
    else if (DATE_FIELDS.indexOf(k) !== -1) v = toDate_(v) || '';
    row[cols[k] - 1] = v;
  });
}

function rowToEntry_(row, cols, id) {
  var e = { id: id };
  FIELDS.forEach(function (f) {
    var k = f[0];
    if (!cols[k]) return;
    var v = row[cols[k] - 1];
    if (v instanceof Date) {
      v = Utilities.formatDate(v, Session.getScriptTimeZone(), k === 'updatedAt' ? "yyyy-MM-dd'T'HH:mm" : 'yyyy-MM-dd');
    }
    e[k] = v;
  });
  return e;
}

function formatRow_(sheet, r, cols) {
  var fmt = function (k, f) { if (cols[k]) sheet.getRange(r, cols[k]).setNumberFormat(f); };
  DATE_FIELDS.forEach(function (k) { fmt(k, 'dd-mmm-yyyy'); });
  fmt('phone', '@');
  fmt('revenueLost', '₹#,##0');
  fmt('updatedAt', 'dd-mmm-yyyy hh:mm');
}

function readEntry_(sheet, r, id) {
  var cols = columns_(sheet);
  return rowToEntry_(sheet.getRange(r, 1, 1, cols.width).getValues()[0], cols, id);
}

/* ---------- row IDs ---------- */

function tagRow_(sheet, r, id) {
  sheet.getRange(r + ':' + r).addDeveloperMetadata(ID_KEY, id);
}

// { rowNumber: entryId } for every tagged row.
function rowIds_(sheet) {
  var out = {};
  sheet.createDeveloperMetadataFinder()
    .withKey(ID_KEY)
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
    .find()
    .forEach(function (m) { out[m.getLocation().getRow().getRow()] = m.getValue(); });
  return out;
}

function findRow_(sheet, id) {
  var found = sheet.createDeveloperMetadataFinder().withKey(ID_KEY).withValue(id).find();
  return found.length ? found[0].getLocation().getRow().getRow() : 0;
}

// Earlier versions wrote Entry ID / Created at / Updated at into three columns at the end.
// Move those IDs into row metadata and clear the old values. Only columns whose header is
// blank, "Entry ID" or "Created at" are touched, and this runs once per spreadsheet.
function migrateIdColumns_(sheet) {
  var props = PropertiesService.getDocumentProperties();
  if (props.getProperty('idColumnsMigrated') === '2') return;
  var last = sheet.getLastRow(), width = sheet.getLastColumn();
  if (last >= 2 && width >= 1) {
    var data = sheet.getRange(1, 1, last, width).getValues();
    var head = data[0].map(normHeader_);
    var uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    var spare = function (c) { return c < width && ['', 'entryid', 'createdat'].indexOf(head[c]) !== -1; };
    var tagged = null;
    for (var c = 0; c < width; c++) {
      if (!spare(c)) continue;
      var hasIds = data.some(function (row, i) { return i > 0 && uuid.test(String(row[c])); });
      if (!hasIds && head[c] !== 'createdat') continue;
      if (hasIds) {
        tagged = tagged || rowIds_(sheet);
        for (var i = 1; i < data.length; i++) {
          var id = String(data[i][c]);
          if (uuid.test(id) && !tagged[i + 1]) { tagRow_(sheet, i + 1, id); tagged[i + 1] = id; }
        }
        // The old timestamps sat in the blank-headed columns right after the ID.
        for (var n = c + 1; n <= c + 2 && n < width && head[n] === ''; n++) {
          sheet.getRange(2, n + 1, last - 1, 1).clear();
        }
      }
      sheet.getRange(2, c + 1, last - 1, 1).clear();
      if (head[c]) sheet.getRange(1, c + 1).clear();
    }
  }
  props.setProperty('idColumnsMigrated', '2');
}

/* ---------- Lists & Rates ---------- */

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
  // The four statuses are fixed; a status column in the Lists tab is ignored.
  out.status = DEFAULT_LISTS.status;
  // Without its own column in Lists, "Offerred Room" offers the room categories.
  var hasOfferedRoom = grid[0].some(function (key, c) {
    return key === 'offeredRoom' && grid.slice(1).some(function (r) { return String(r[c]).trim(); });
  });
  if (!hasOfferedRoom) out.offeredRoom = out.roomCategory.concat(['Not offered']);
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

function toDate_(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  var m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  var d = new Date(v);
  return isNaN(d) ? null : d;
}
