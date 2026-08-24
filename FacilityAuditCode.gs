// =============================================================
//  FACILITY SELF AUDIT — CA MANAGEMENT SYSTEM
//  Container Supply Co. | FRM-018-002 / FRM-018-004
//  Apps Script backend — paste into your Google Sheet's
//  Apps Script editor (Extensions > Apps Script)
// =============================================================

// ── CONFIG ────────────────────────────────────────────────────
const FA = {
  caLogSheet:        'CA Log (Jul 2026+)',
  auditLogSheet:     'Audit Log (Jul 2026+)',
  refSheet:          'Reference',
  dataStartRow:      14,  // CA Log data starts at row 14
  auditDataStartRow: 5,   // Audit Log data starts at row 5
};

// ── DOCUMENT METADATA ─────────────────────────────────────────
const FA_META = {
  auditForm: { docNo:'FRM-018-002', rev:'0', revDate:'07/17/2026', title:'Facility Self Inspection Form' },
  caForm:    { docNo:'FRM-018-004', rev:'0', revDate:'07/17/2026', title:'Facility Self Audit Corrective Action Form' }
};
function getFADocMeta() { return FA_META; }

// ── COLUMN MAP — CA LOG (Jul 2026+) — 19 columns ─────────────
// A=CAR ID, B=Audit Date, C=NC Date, D=Zone, E=Department,
// F=Inspection Item, G=Location in Zone, H=NC Level,
// I=Description, J=Corrective Action Plan, K=Completed By,
// L=Assigned By, M=Due Date, N=Completion Date, O=Verified By,
// P=Verification Date, Q=Photo Link, R=Status, S=Deferred
var CA_COLS = {
  carId:             1,   // A
  auditDate:         2,   // B
  ncDate:            3,   // C
  zone:              4,   // D
  department:        5,   // E
  inspectionItem:    6,   // F
  locationInZone:    7,   // G
  ncLevel:           8,   // H
  description:       9,   // I
  correctiveAction:  10,  // J
  responsiblePerson: 11,  // K
  assignedBy:        12,  // L
  dueDate:           13,  // M
  completionDate:    14,  // N
  verifiedBy:        15,  // O
  verificationDate:  16,  // P
  photoLink:         17,  // Q
  status:            18,  // R
  deferred:          19,  // S
};
var CA_NUM_COLS = 19;

// ── INSPECTION ITEM → CATEGORY MAP (mirrors AuditForm.html's ITEMS/CAT_META) ──
// Used to re-derive Audit Log NC/Pass counts from live CA Log data — see
// recalcAuditLogFromCALog_() below.
var ITEM_CATEGORY = {
  'Building exterior & grounds':  'maint',
  'Roof, walls & doors':          'maint',
  'Lighting & ventilation':       'maint',
  'Equipment condition':          'maint',
  'Floors, walls & ceilings':     'sanit',
  'Waste containment & removal':  'sanit',
  'Handwashing & hygiene':        'sanit',
  'Pest control devices':         'gmp',
  'Food, drink & personal items': 'gmp',
  'Tools & equipment storage':    'gmp',
  'Personal protective equipment':'gmp',
  'Product & materials storage':  'gmp',
  'First aid kit':                'safety',
  'Eye wash station':             'safety',
  'Machine guards & covers':      'safety',
};
var CAT_ITEM_TOTAL = { maint: 4, sanit: 3, gmp: 5, safety: 3 };

// A raw comma-joined list of Drive URLs in one cell reads as a wall of
// text, so the Photo Link column stores a HYPERLINK() formula instead —
// the cell displays friendly "📎 Link N" labels while the formula text
// itself still carries the real URLs, which extractPhotoLinksFromFormula_
// pulls back out for every downstream consumer (photo loading, PDF/email
// generation). Legacy rows written before this change have a plain
// comma/newline-joined URL string with no formula — callers fall back
// to that automatically when getFormula() is empty.
function buildPhotoLinkFormula_(links) {
  return '=' + links.map(function(url, i) {
    return 'HYPERLINK("' + url.replace(/"/g, '""') + '","📎 Link ' + (i + 1) + '")';
  }).join(' & "   " & ');
}
function extractPhotoLinksFromFormula_(formula) {
  var urls = [];
  var re = /HYPERLINK\(\s*"((?:[^"]|"")*)"/g;
  var m;
  while ((m = re.exec(formula || '')) !== null) urls.push(m[1].replace(/""/g, '"'));
  return urls;
}

// ── COLUMN MAP — AUDIT LOG (Jul 2026+) — 20 columns ──────────
var AL_COLS = {
  auditDate:1, zone:2, department:3, coordinator:4, auditor:5,
  maintYes:6, maintNo:7, maintNA:8,
  sanitYes:9, sanitNo:10, sanitNA:11,
  gmpYes:12, gmpNo:13, gmpNA:14,
  safetyYes:15, safetyNo:16, safetyNA:17,
  totalYes:18, casGenerated:19, notes:20,
};
var AL_NUM_COLS = 20;

// =============================================================
//  MENU
// =============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Facility Audit')
    .addItem('📋  Start Inspection Form',   'openAuditForm')
    .addItem('📝  Update CA Record',        'openCAUpdateForm')
    .addSeparator()
    .addItem('📊  Manager Trend Report',    'openManagerReport')
    .addItem('📅  Schedule Next Audit',     'openScheduleAudit')
    .addSeparator()
    .addItem('🔁  Resync Audit Log Counts', 'resyncAuditLogCounts')
    .addItem('⚙️   Setup Config Block',     'setupConfigBlock')
    .addToUi();
}

// =============================================================
//  DEBUG — select debugSubmit > Run, then View > Execution log
// =============================================================
function debugSubmit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Spreadsheet: ' + ss.getName() + ' | ID: ' + ss.getId());
  var sheets = ss.getSheets().map(function(s){ return '[' + s.getName() + ']'; });
  Logger.log('All sheets: ' + sheets.join(', '));
  var cfg = getSettings();
  Logger.log('Settings: ' + JSON.stringify(cfg));
  var caName    = cfg['CA Log Sheet Name']    || FA.caLogSheet;
  var auditName = cfg['Audit Log Sheet Name'] || FA.auditLogSheet;
  var caSheet    = ss.getSheetByName(caName);
  var auditSheet = ss.getSheetByName(auditName);
  Logger.log('CA Log: ' + (caSheet ? 'FOUND, lastRow=' + caSheet.getLastRow() : 'NOT FOUND'));
  Logger.log('Audit Log: ' + (auditSheet ? 'FOUND, lastRow=' + auditSheet.getLastRow() : 'NOT FOUND'));
  if (caSheet) Logger.log('Next CA write row: ' + (getLastDataRow_(caSheet, 1, FA.dataStartRow) + 1));
  var result = submitAuditForm({
    auditInfo: { auditDate:'2026-07-23', zone:'3', coordinator:'Debug Test', designee:'Debug Test' },
    counts: { maint:{y:1,n:3,na:0}, sanit:{y:0,n:3,na:0}, gmp:{y:0,n:5,na:0}, safety:{y:0,n:3,na:0} },
    caRecords: [{
      carId:'CA-202607-DEBUG', auditDate:'2026-07-23', ncDate:'2026-07-23',
      zone:'3', inspectionItem:'Equipment condition', locationInZone:'Debug location',
      ncLevel:'Minor', description:'DEBUG TEST ROW — delete after testing',
      responsiblePerson:'',
      assignedBy:'', deferred:'No', photos:[]
    }]
  });
  Logger.log('submitAuditForm result: ' + JSON.stringify(result));
}

// =============================================================
//  HELPER — last row with real data in col, starting at minRow
// =============================================================
function getLastDataRow_(sheet, col, minRow) {
  var lastRow = sheet.getLastRow();
  if (lastRow < minRow) return minRow - 1;
  var values = sheet.getRange(minRow, col, lastRow - minRow + 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var v = String(values[i][0] || '').trim();
    if (v !== '' && v !== '0') return minRow + i;
  }
  return minRow - 1;
}

// =============================================================
//  OPEN AUDIT FORM
// =============================================================
function openAuditForm() {
  var html = HtmlService.createHtmlOutputFromFile('AuditForm')
    .setWidth(1300).setHeight(1400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Facility Self Inspection Form — FRM-018-002');
}

// Web app entry point — for mobile/standalone use
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('AuditForm')
    .setTitle('Facility Self Inspection — CSC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================
//  READ SETTINGS FROM Settings TAB
// =============================================================
function getSettings() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  var data = sheet.getRange(4, 1, 20, 2).getValues();
  var cfg  = {};
  data.forEach(function(row) {
    if (row[0]) cfg[String(row[0]).trim()] = String(row[1] || '').trim();
  });
  return cfg;
}

// =============================================================
//  READ ZONE MANAGERS FROM Settings TAB
// =============================================================
function getZoneManagers() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var mgrs = {};
  var inMgr = false;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'Zone' && String(data[i][1]).trim() === 'Department') {
      inMgr = true; continue;
    }
    if (inMgr && data[i][0] && String(data[i][0]).trim() !== '') {
      mgrs[String(data[i][0]).trim()] = {
        name:  String(data[i][2] || '').trim(),
        email: String(data[i][3] || '').trim(),
      };
    }
  }
  return mgrs;
}

// =============================================================
//  GET PILL LIBRARY — called from AuditForm on load
// =============================================================
function getPillLibrary() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var cfg       = getSettings();
  var sheetName = cfg['Pill Library Sheet'] || 'Pill Library';
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  var data  = sheet.getRange(4, 1, lastRow - 3, 4).getValues();
  var pills = [];
  data.forEach(function(row) {
    var num  = row[0];
    var pill = String(row[3] || '').trim();
    if (num && pill) pills.push({ item_num: parseInt(num), pill_text: pill });
  });
  return pills;
}

// =============================================================
//  GET NEXT CAR SEQUENCE NUMBER for current month
// =============================================================
function getNextCarSeq(zone) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var cfg       = getSettings();
  var sheetName = cfg['CA Log Sheet Name'] || FA.caLogSheet;
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) return 1;
  var today  = new Date();
  var ym     = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyyMM');
  var prefix = 'CA-' + ym + '-Z' + String(zone).trim() + '-';
  var lastRow = getLastDataRow_(sheet, 1, FA.dataStartRow);
  if (lastRow < FA.dataStartRow) return 1;
  var numCols = Math.max(CA_COLS.zone, CA_COLS.carId);
  var data = sheet.getRange(FA.dataStartRow, 1, lastRow - FA.dataStartRow + 1, numCols).getValues();
  var maxSeq = 0;
  data.forEach(function(row) {
    var id      = String(row[CA_COLS.carId - 1] || '');
    var rowZone = String(row[CA_COLS.zone - 1]  || '').trim();
    if (zone && rowZone !== String(zone).trim()) return;
    if (id.indexOf(prefix) === 0) {
      var seq = parseInt(id.replace(prefix, ''), 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  });
  return maxSeq + 1;
}

// =============================================================
//  WORKDAY DUE DATE CALCULATOR
// =============================================================
function addWorkdays(startDate, workdays) {
  if (workdays <= 0) return startDate;
  var d = new Date(startDate.getTime());
  var added = 0;
  while (added < workdays) {
    d.setDate(d.getDate() + 1);
    var day = d.getDay();
    if (day !== 0 && day !== 6) added++;
  }
  return d;
}

function getDueDate(ncDate, ncLevel) {
  var d = new Date(ncDate);
  if (ncLevel === 'Critical')    return d;
  if (ncLevel === 'Major')       return addWorkdays(d, 14);
  if (ncLevel === 'Minor')       return addWorkdays(d, 29);
  return null;
}


// =============================================================
//  FETCH MULTIPLE DRIVE PHOTOS AS BASE64 ARRAY
//  Called from CAUpdateForm.html for on-screen photo display
// =============================================================
function getCARPhotosBase64(photoLinkStr) {
  try {
    if (!photoLinkStr || !photoLinkStr.trim()) return [];
    // Split on either delimiter — older rows were saved newline-joined, new ones comma-joined.
    var links = photoLinkStr.split(/[\n,]+/).map(function(l){ return l.trim(); }).filter(Boolean);
    var results = [];
    links.forEach(function(url) {
      try {
        var m = url.match(/\/file\/d\/([^\/\?&]+)/);
        if (!m) return;
        var fileId   = m[1].replace(/\/$/, '');
        var file     = DriveApp.getFileById(fileId);
        var blob     = file.getBlob();
        var mimeType = blob.getContentType() || 'image/jpeg';
        var b64      = Utilities.base64Encode(blob.getBytes());
        results.push({
          dataUrl: 'data:' + mimeType + ';base64,' + b64,
          name:    file.getName(),
        });
      } catch(e) { Logger.log('Photo fetch error: ' + e.message); }
    });
    return results;
  } catch(e) {
    Logger.log('getCARPhotosBase64 error: ' + e.message);
    return [];
  }
}

// =============================================================
//  SUBMIT AUDIT FORM — main handler from AuditForm.html
// =============================================================
function submitAuditForm(payload) {
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet: ' + ss.getName() + ' | ID: ' + ss.getId());
    var cfg = getSettings();
    var tz  = Session.getScriptTimeZone();

    var caSheetName    = cfg['CA Log Sheet Name']    || FA.caLogSheet;
    var auditSheetName = cfg['Audit Log Sheet Name'] || FA.auditLogSheet;
    Logger.log('Using CA sheet: [' + caSheetName + '] | Audit sheet: [' + auditSheetName + ']');

    var caSheet    = ss.getSheetByName(caSheetName);
    var auditSheet = ss.getSheetByName(auditSheetName);
    if (!caSheet)    throw new Error('CA Log sheet not found: ' + caSheetName);
    if (!auditSheet) throw new Error('Audit Log sheet not found: ' + auditSheetName);

    var ai        = payload.auditInfo;
    var counts    = payload.counts;
    var caRecords = payload.caRecords || [];
    var auditDate = new Date(ai.auditDate);
    var deptMap   = {'1':'Shipping & Receiving','2':'Metals','3':'Plastics','4':'Lithography','5':'Miscellaneous'};
    var dept      = deptMap[ai.zone] || 'Zone ' + ai.zone;

    // ── Write Audit Log row ──────────────────────────────────
    var alRow    = getLastDataRow_(auditSheet, 1, FA.auditDataStartRow) + 1;
    Logger.log('Writing Audit Log to row: ' + alRow);
    var totalYes = counts.maint.y + counts.sanit.y + counts.gmp.y + counts.safety.y;
    var alValues = [
      Utilities.formatDate(auditDate, tz, 'MM/dd/yyyy'),
      ai.zone, dept, ai.coordinator, ai.designee,
      counts.maint.y,  counts.maint.n,  counts.maint.na,
      counts.sanit.y,  counts.sanit.n,  counts.sanit.na,
      counts.gmp.y,    counts.gmp.n,    counts.gmp.na,
      counts.safety.y, counts.safety.n, counts.safety.na,
      totalYes, caRecords.length, ''
    ];
    auditSheet.getRange(alRow, 1, 1, alValues.length).setValues([alValues]);

    // ── Write CA Log rows ────────────────────────────────────
    var caStartRow = getLastDataRow_(caSheet, 1, FA.dataStartRow) + 1;
    Logger.log('Writing CA records starting at row: ' + caStartRow);

    // Recalculate CAR IDs server-side with zone-scoped sequential numbers
    var seqStart = getNextCarSeq(ai.zone);
    var ymStr    = Utilities.formatDate(auditDate, tz, 'yyyyMM');
    caRecords.forEach(function(car, idx) {
      car.carId = 'CA-' + ymStr + '-Z' + ai.zone + '-' + String(seqStart + idx).padStart(3, '0');
    });

    caRecords.forEach(function(car, idx) {
      var r         = caStartRow + idx;
      var ncDate    = new Date(car.ncDate || ai.auditDate);
      var dueDate   = getDueDate(ncDate, car.ncLevel);
      var dueDateStr = dueDate ? Utilities.formatDate(dueDate, tz, 'MM/dd/yyyy') : '';

      var photoLinks = [];
      if (car.photos && car.photos.length > 0) {
        try {
          var folder = getOrCreateAuditPhotoFolder_(ai.zone, ai.auditDate);
          car.photos.forEach(function(p) {
            var base64Data = p.dataUrl.split(',')[1];
            var mimeType   = p.dataUrl.split(';')[0].replace('data:', '');
            var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, p.name || car.carId + '.jpg');
            var file = folder.createFile(blob);
            try {
              // Domain-restricted sharing — org policy commonly blocks ANYONE_WITH_LINK
              // outright, which would throw here and skip recording the link below.
              file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (shareErr) {
              Logger.log('Photo sharing warning (' + file.getName() + '): ' + shareErr.message);
            }
            photoLinks.push('https://drive.google.com/file/d/' + file.getId() + '/view');
          });
        } catch(photoErr) {
          Logger.log('Photo upload error: ' + photoErr.message);
        }
      }

      var rowVals = [
        car.carId,                                             // A  CAR ID
        Utilities.formatDate(auditDate, tz, 'MM/dd/yyyy'),     // B  Audit Date
        Utilities.formatDate(ncDate,    tz, 'MM/dd/yyyy'),     // C  NC Date
        ai.zone,                                               // D  Zone
        dept,                                                  // E  Department
        car.inspectionItem,                                    // F  Inspection Item
        car.locationInZone,                                    // G  Location in Zone
        car.ncLevel,                                           // H  NC Level
        car.description,                                       // I  Description
        '',                                                    // J  Corrective Action Plan
        car.responsiblePerson,                                 // K  Completed By
        car.assignedBy,                                        // L  Assigned By
        dueDateStr,                                            // M  Due Date
        '',                                                    // N  Completion Date
        '',                                                    // O  Verified By
        '',                                                    // P  Verification Date
        '',                                                    // Q  Photo Link (set below via formula if any)
        'Open',                                                // R  Status
        car.deferred,                                          // S  Deferred
      ];
      caSheet.getRange(r, 1, 1, rowVals.length).setValues([rowVals]);
      if (photoLinks.length) {
        caSheet.getRange(r, CA_COLS.photoLink).setFormula(buildPhotoLinkFormula_(photoLinks));
      }
      applyStatusColor_(caSheet, r, 'Open');
      applyNCLevelColor_(caSheet, r, car.ncLevel);
    });

    if (caRecords.length > 0) {
      applyZoneColorsToNewRows_(caSheet, caStartRow, caRecords.length, ai.zone);
    }

    SpreadsheetApp.flush();
    Logger.log('SUCCESS — Audit row: ' + alRow + ' | CA records: ' + caRecords.length);
    return { success: true, caCount: caRecords.length };

  } catch(e) {
    Logger.log('submitAuditForm ERROR: ' + e.message + '\n' + e.stack);
    return { success: false, message: e.message };
  }
}

// =============================================================
//  FETCH CA RECORDS — for CAR search sidebar
// =============================================================
function getCARecords() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var cfg       = getSettings();
  var sheet     = ss.getSheetByName(cfg['CA Log Sheet Name'] || FA.caLogSheet);
  if (!sheet) return [];
  var lastRow   = getLastDataRow_(sheet, 1, FA.dataStartRow);
  if (lastRow < FA.dataStartRow) return [];
  var numRows   = lastRow - FA.dataStartRow + 1;
  var numCols   = Math.min(CA_NUM_COLS, sheet.getLastColumn());
  var data      = sheet.getRange(FA.dataStartRow, 1, numRows, numCols).getDisplayValues();
  var photoFormulas = sheet.getRange(FA.dataStartRow, CA_COLS.photoLink, numRows, 1).getFormulas();
  var records   = [];
  data.forEach(function(row, i) {
    if (!row[0] && !row[1]) return;
    if ((row[0] || '').toString().toLowerCase().indexOf('debug') !== -1) return;
    var photoFormula = (photoFormulas[i] && photoFormulas[i][0]) || '';
    var photoLinks   = photoFormula
      ? extractPhotoLinksFromFormula_(photoFormula)
      : String(row[CA_COLS.photoLink - 1] || '').split(/[\n,]+/).map(function(l){ return l.trim(); }).filter(Boolean);
    records.push({
      _rowIndex:         FA.dataStartRow + i,
      carId:             row[CA_COLS.carId - 1]             || '',
      auditDate:         row[CA_COLS.auditDate - 1]         || '',
      ncDate:            row[CA_COLS.ncDate - 1]            || '',
      zone:              row[CA_COLS.zone - 1]              || '',
      department:        row[CA_COLS.department - 1]        || '',
      inspectionItem:    row[CA_COLS.inspectionItem - 1]    || '',
      locationInZone:    row[CA_COLS.locationInZone - 1]    || '',
      ncLevel:           row[CA_COLS.ncLevel - 1]           || '',
      description:       row[CA_COLS.description - 1]       || '',
      correctiveAction:  row[CA_COLS.correctiveAction - 1]  || '',
      responsiblePerson: row[CA_COLS.responsiblePerson - 1] || '',
      assignedBy:        row[CA_COLS.assignedBy - 1]        || '',
      dueDate:           row[CA_COLS.dueDate - 1]           || '',
      completionDate:    row[CA_COLS.completionDate - 1]    || '',
      verifiedBy:        row[CA_COLS.verifiedBy - 1]        || '',
      verificationDate:  row[CA_COLS.verificationDate - 1]  || '',
      photoLink:         photoLinks.join(', '),
      status:            row[CA_COLS.status - 1]            || '',
      deferred:          row[CA_COLS.deferred - 1]          || '',
    });
  });
  return records.reverse();
}

// =============================================================
//  FETCH AUDIT RECORDS — for manager trend report
// =============================================================
function getAuditRecords() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var cfg       = getSettings();
  var sheet     = ss.getSheetByName(cfg['Audit Log Sheet Name'] || FA.auditLogSheet);
  if (!sheet) return [];
  var lastRow   = getLastDataRow_(sheet, 1, FA.auditDataStartRow);
  if (lastRow < FA.auditDataStartRow) return [];
  var numRows   = lastRow - FA.auditDataStartRow + 1;
  var numCols   = Math.min(AL_NUM_COLS, sheet.getLastColumn());
  var data      = sheet.getRange(FA.auditDataStartRow, 1, numRows, numCols).getDisplayValues();
  var records   = [];
  data.forEach(function(row) {
    if (!row[0] && !row[1]) return;
    if ((row[0] || '').toString().toLowerCase().indexOf('example') !== -1) return;
    records.push({
      auditDate:    row[AL_COLS.auditDate - 1]    || '',
      zone:         row[AL_COLS.zone - 1]         || '',
      department:   row[AL_COLS.department - 1]   || '',
      coordinator:  row[AL_COLS.coordinator - 1]  || '',
      auditor:      row[AL_COLS.auditor - 1]      || '',
      maintYes:     Number(row[AL_COLS.maintYes - 1])  || 0,
      maintNo:      Number(row[AL_COLS.maintNo - 1])   || 0,
      maintNA:      Number(row[AL_COLS.maintNA - 1])   || 0,
      sanitYes:     Number(row[AL_COLS.sanitYes - 1])  || 0,
      sanitNo:      Number(row[AL_COLS.sanitNo - 1])   || 0,
      sanitNA:      Number(row[AL_COLS.sanitNA - 1])   || 0,
      gmpYes:       Number(row[AL_COLS.gmpYes - 1])    || 0,
      gmpNo:        Number(row[AL_COLS.gmpNo - 1])     || 0,
      gmpNA:        Number(row[AL_COLS.gmpNA - 1])     || 0,
      safetyYes:    Number(row[AL_COLS.safetyYes - 1]) || 0,
      safetyNo:     Number(row[AL_COLS.safetyNo - 1])  || 0,
      safetyNA:     Number(row[AL_COLS.safetyNA - 1])  || 0,
      totalYes:     Number(row[AL_COLS.totalYes - 1])  || 0,
      casGenerated: row[AL_COLS.casGenerated - 1] || '',
      notes:        row[AL_COLS.notes - 1]        || '',
    });
  });
  return records;
}

// =============================================================
//  KEEP AUDIT LOG COUNTS IN SYNC WITH LIVE CA LOG DATA
// =============================================================
// The Audit Log's per-category NC/Pass/N-A tallies are captured once,
// at submission time. If a CA record is later added or deleted by hand
// (entry error, etc.) those tallies go stale. This walks the live CA
// Log, re-derives NC counts + finding totals per audit (matched by
// zone + audit date), and rewrites the affected Audit Log columns.
// N/A counts are left untouched — an N/A item never produces a CA
// record, so its count can't be re-derived from the CA Log alone.
function recalcAuditLogFromCALog_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var cfg   = getSettings();
  var caSheet    = ss.getSheetByName(cfg['CA Log Sheet Name']    || FA.caLogSheet);
  var auditSheet = ss.getSheetByName(cfg['Audit Log Sheet Name'] || FA.auditLogSheet);
  if (!caSheet || !auditSheet) return 0;

  var auditLastRow = getLastDataRow_(auditSheet, 1, FA.auditDataStartRow);
  if (auditLastRow < FA.auditDataStartRow) return 0;

  var caLastRow = getLastDataRow_(caSheet, 1, FA.dataStartRow);
  var caData = (caLastRow >= FA.dataStartRow)
    ? caSheet.getRange(FA.dataStartRow, 1, caLastRow - FA.dataStartRow + 1, CA_NUM_COLS).getValues()
    : [];

  // key = "zone|auditDate" -> { items:{maint:{},sanit:{},gmp:{},safety:{}}, findingCount:n }
  var byAudit = {};
  caData.forEach(function(row) {
    var zone      = String(row[CA_COLS.zone - 1] || '').trim();
    var auditDate = normalizeDateKey_(row[CA_COLS.auditDate - 1]);
    var item      = String(row[CA_COLS.inspectionItem - 1] || '').trim();
    if (!zone || !auditDate) return;
    var key = zone + '|' + auditDate;
    if (!byAudit[key]) byAudit[key] = { items: { maint:{}, sanit:{}, gmp:{}, safety:{} }, findingCount: 0 };
    byAudit[key].findingCount++;
    var cat = ITEM_CATEGORY[item];
    if (cat) byAudit[key].items[cat][item] = true;
  });

  var numRows   = auditLastRow - FA.auditDataStartRow + 1;
  var auditVals = auditSheet.getRange(FA.auditDataStartRow, 1, numRows, AL_NUM_COLS).getValues();
  var changed = 0;

  auditVals.forEach(function(row) {
    var zone      = String(row[AL_COLS.zone - 1] || '').trim();
    var auditDate = normalizeDateKey_(row[AL_COLS.auditDate - 1]);
    if (!zone || !auditDate) return;
    var agg = byAudit[zone + '|' + auditDate] || { items: { maint:{}, sanit:{}, gmp:{}, safety:{} }, findingCount: 0 };

    var naMaint  = Number(row[AL_COLS.maintNA - 1])  || 0;
    var naSanit  = Number(row[AL_COLS.sanitNA - 1])  || 0;
    var naGmp    = Number(row[AL_COLS.gmpNA - 1])    || 0;
    var naSafety = Number(row[AL_COLS.safetyNA - 1]) || 0;

    var ncMaint  = Object.keys(agg.items.maint).length;
    var ncSanit  = Object.keys(agg.items.sanit).length;
    var ncGmp    = Object.keys(agg.items.gmp).length;
    var ncSafety = Object.keys(agg.items.safety).length;

    var passMaint  = Math.max(0, CAT_ITEM_TOTAL.maint  - naMaint  - ncMaint);
    var passSanit  = Math.max(0, CAT_ITEM_TOTAL.sanit  - naSanit  - ncSanit);
    var passGmp    = Math.max(0, CAT_ITEM_TOTAL.gmp    - naGmp    - ncGmp);
    var passSafety = Math.max(0, CAT_ITEM_TOTAL.safety - naSafety - ncSafety);
    var totalNC    = ncMaint + ncSanit + ncGmp + ncSafety;

    var before = [row[AL_COLS.maintYes-1], row[AL_COLS.maintNo-1], row[AL_COLS.sanitYes-1], row[AL_COLS.sanitNo-1],
                  row[AL_COLS.gmpYes-1], row[AL_COLS.gmpNo-1], row[AL_COLS.safetyYes-1], row[AL_COLS.safetyNo-1],
                  row[AL_COLS.totalYes-1], row[AL_COLS.casGenerated-1]].join('|');
    var after  = [ncMaint, passMaint, ncSanit, passSanit, ncGmp, passGmp, ncSafety, passSafety, totalNC, agg.findingCount].join('|');
    if (before === after) return;

    row[AL_COLS.maintYes-1]  = ncMaint;   row[AL_COLS.maintNo-1]  = passMaint;
    row[AL_COLS.sanitYes-1]  = ncSanit;   row[AL_COLS.sanitNo-1]  = passSanit;
    row[AL_COLS.gmpYes-1]    = ncGmp;     row[AL_COLS.gmpNo-1]    = passGmp;
    row[AL_COLS.safetyYes-1] = ncSafety;  row[AL_COLS.safetyNo-1] = passSafety;
    row[AL_COLS.totalYes-1]      = totalNC;
    row[AL_COLS.casGenerated-1]  = agg.findingCount;
    changed++;
  });

  if (changed > 0) {
    auditSheet.getRange(FA.auditDataStartRow, 1, numRows, AL_NUM_COLS).setValues(auditVals);
    SpreadsheetApp.flush();
  }
  return changed;
}

function normalizeDateKey_(v) {
  if (!v) return '';
  var d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d.getTime())) return String(v).trim();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM/dd/yyyy');
}

// Simple trigger — fires automatically on manual edits to the CA Log
// (typing a value, or adding/deleting a row through the sheet UI) so the
// Audit Log stays synced without any separate setup. Script-driven writes
// (e.g. submitAuditForm) don't re-fire this, so there's no risk of a loop.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var cfg = getSettings();
    var caSheetName = cfg['CA Log Sheet Name'] || FA.caLogSheet;
    if (sheet.getName() !== caSheetName) return;
    if (e.range.getRow() + e.range.getNumRows() - 1 < FA.dataStartRow) return; // header-only edit

    // Only recompute Due Date for rows where NC Date or NC Level itself was
    // just edited — that way a manually-overridden Due Date (e.g. a granted
    // extension) survives unrelated edits (status, corrective action, etc.)
    // to the same row, and only resets when the inputs driving it change.
    var colStart = e.range.getColumn();
    var colEnd   = colStart + e.range.getNumColumns() - 1;
    var touchesNcDate  = colStart <= CA_COLS.ncDate  && colEnd >= CA_COLS.ncDate;
    var touchesNcLevel = colStart <= CA_COLS.ncLevel && colEnd >= CA_COLS.ncLevel;

    if (touchesNcDate || touchesNcLevel) autofillDueDates_(sheet, e.range);
    if (touchesNcLevel) autofillNCLevelColors_(sheet, e.range);

    recalcAuditLogFromCALog_();
  } catch (err) {
    Logger.log('onEdit recalc error: ' + err.message);
  }
}

// Fills Due Date from NC Date + NC Level for whichever rows were just
// edited — same rule as getDueDate() used at audit submission time
// (Critical: same day, Major: 14 workdays, Minor: 29 workdays).
function autofillDueDates_(sheet, editedRange) {
  var rowStart = Math.max(editedRange.getRow(), FA.dataStartRow);
  var rowEnd   = editedRange.getRow() + editedRange.getNumRows() - 1;
  if (rowEnd < rowStart) return;
  var numRows = rowEnd - rowStart + 1;
  var tz = Session.getScriptTimeZone();

  var ncDates  = sheet.getRange(rowStart, CA_COLS.ncDate,  numRows, 1).getValues();
  var ncLevels = sheet.getRange(rowStart, CA_COLS.ncLevel, numRows, 1).getValues();
  var dueDates = sheet.getRange(rowStart, CA_COLS.dueDate, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var ncDateVal  = ncDates[i][0];
    var ncLevelVal = String(ncLevels[i][0] || '').trim();
    if (!ncDateVal || !ncLevelVal) continue;
    var due = getDueDate(ncDateVal, ncLevelVal);
    dueDates[i][0] = due ? Utilities.formatDate(due, tz, 'MM/dd/yyyy') : '';
  }
  sheet.getRange(rowStart, CA_COLS.dueDate, numRows, 1).setValues(dueDates);
}

// Re-applies the NC Level cell's color formatting (via applyNCLevelColor_,
// the same helper used when a record is first created) for whichever rows
// just had their NC Level manually edited — otherwise a cell changed from
// Minor to Major/Critical keeps its old color until someone notices.
function autofillNCLevelColors_(sheet, editedRange) {
  var rowStart = Math.max(editedRange.getRow(), FA.dataStartRow);
  var rowEnd   = editedRange.getRow() + editedRange.getNumRows() - 1;
  if (rowEnd < rowStart) return;
  var numRows = rowEnd - rowStart + 1;
  var levels  = sheet.getRange(rowStart, CA_COLS.ncLevel, numRows, 1).getValues();

  for (var i = 0; i < numRows; i++) {
    var level = String(levels[i][0] || '').trim();
    if (level) applyNCLevelColor_(sheet, rowStart + i, level);
  }
}

// Manual menu action — one-off resync, and a fallback for edit types that
// don't fire onEdit reliably (e.g. some bulk-paste/import paths).
function resyncAuditLogCounts() {
  var changed = recalcAuditLogFromCALog_();
  SpreadsheetApp.getUi().alert(changed > 0
    ? 'Audit Log updated — ' + changed + ' audit row' + (changed === 1 ? '' : 's') + ' resynced from the CA Log.'
    : 'Audit Log is already in sync with the CA Log.');
}

// =============================================================
//  UPDATE CA STATUS — called from CAR search sidebar
// =============================================================
function updateCAStatus(carId, newStatus) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var cfg   = getSettings();
    var sheet = ss.getSheetByName(cfg['CA Log Sheet Name'] || FA.caLogSheet);
    if (!sheet) return { success: false, message: 'CA Log sheet not found.' };
    var lastRow = getLastDataRow_(sheet, 1, FA.dataStartRow);
    if (lastRow < FA.dataStartRow) return { success: false, message: 'No data found.' };
    var ids = sheet.getRange(FA.dataStartRow, CA_COLS.carId, lastRow - FA.dataStartRow + 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(carId).trim()) {
        var rowNum = FA.dataStartRow + i;
        sheet.getRange(rowNum, CA_COLS.status).setValue(newStatus);
        applyStatusColor_(sheet, rowNum, newStatus);
        return { success: true };
      }
    }
    return { success: false, message: 'CAR ID not found: ' + carId };
  } catch(e) {
    return { success: false, message: String(e.message || e) };
  }
}

// =============================================================
//  COLOR HELPERS
// =============================================================
function applyStatusColor_(sheet, rowNum, status) {
  var cell = sheet.getRange(rowNum, CA_COLS.status);
  var sv   = String(status || '').toLowerCase();
  if      (sv.indexOf('closed')    !== -1) { cell.setBackground('#DCEEFB').setFontColor('#1B4F72'); }
  else if (sv.indexOf('complete')  !== -1) { cell.setBackground('#D6EFD8').setFontColor('#1A6B1A'); }
  else if (sv.indexOf('past due')  !== -1) { cell.setBackground('#FAD7D7').setFontColor('#8B0000'); }
  else if (sv.indexOf('removed')   !== -1) { cell.setBackground('#E8E8E8').setFontColor('#666666'); }
  else if (sv.indexOf('deferred')  !== -1) { cell.setBackground('#EDE8F8').setFontColor('#4A1A8A'); }
  else                                     { cell.setBackground('#FFF3CD').setFontColor('#7A5000'); }
  cell.setFontWeight('bold').setHorizontalAlignment('center');
}

function applyNCLevelColor_(sheet, row, level) {
  var NC_FILLS = {Critical:'FAD7D7', Major:'FFF3CD', Minor:'D6EAF8', Observation:'E8E8E8'};
  var NC_FONTS = {Critical:'8B0000', Major:'7A5000', Minor:'1A3A6B', Observation:'555555'};
  var cell = sheet.getRange(row, CA_COLS.ncLevel);
  if (NC_FILLS[level]) {
    cell.setBackground('#'+NC_FILLS[level]).setFontColor('#'+NC_FONTS[level])
        .setFontWeight('bold').setHorizontalAlignment('center');
  }
}

function applyZoneColorsToNewRows_(sheet, startRow, count, zone) {
  var FILLS = {'1':'C5D8EC','2':'C5DDD0','3':'FAFAC0','4':'E8DFB0','5':'E8C8E8'};
  var FONTS = {'1':'0D3A5C','2':'0D3D22','3':'5C5C00','4':'4A3C00','5':'4A0A4A'};
  if (!FILLS[zone]) return;
  for (var i = 0; i < count; i++) {
    sheet.getRange(startRow + i, 1, 1, 5)
      .setBackground('#' + FILLS[zone])
      .setFontColor('#' + FONTS[zone]);
  }
}

// =============================================================
//  DRIVE PHOTO FOLDER HELPER
// =============================================================
function getOrCreateAuditPhotoFolder_(zone, auditDate) {
  var rootName = 'CSC Facility Audit Photos';
  var roots    = DriveApp.getFoldersByName(rootName);
  var root     = roots.hasNext() ? roots.next() : DriveApp.createFolder(rootName);
  var month    = Utilities.formatDate(new Date(auditDate), Session.getScriptTimeZone(), 'yyyy-MM');
  var mFolders = root.getFoldersByName(month);
  var mFolder  = mFolders.hasNext() ? mFolders.next() : root.createFolder(month);
  var zoneName = 'Zone-' + zone;
  var zFolders = mFolder.getFoldersByName(zoneName);
  return zFolders.hasNext() ? zFolders.next() : mFolder.createFolder(zoneName);
}

// =============================================================
//  FETCH DRIVE IMAGE AS BASE64 — single-photo helper
// =============================================================
function getDriveImageBase64(url) {
  try {
    if (!url || !url.trim()) return null;
    var fileId = '';
    var m1 = url.match(/\/file\/d\/([^\/\?&]+)/);
    var m2 = url.match(/[?&]id=([^&]+)/);
    if (m1) fileId = m1[1].replace(/\/$/, '');
    else if (m2) fileId = m2[1];
    if (!fileId) return null;
    var file     = DriveApp.getFileById(fileId);
    var blob     = file.getBlob();
    var mimeType = blob.getContentType() || 'image/jpeg';
    return 'data:' + mimeType + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch(e) {
    Logger.log('getDriveImageBase64 error: ' + e.message);
    return null;
  }
}

// =============================================================
//  OPEN REPORT DIALOGS
// =============================================================
function openManagerReport() {
  var html = HtmlService.createHtmlOutputFromFile('ManagerReport').setWidth(1300).setHeight(1400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Facility Audit — Manager Trend Report');
}

function openCAUpdateForm() {
  var html = HtmlService.createHtmlOutputFromFile('CAUpdateForm')
    .setWidth(1300).setHeight(1400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Update Corrective Action — FRM-018-004');
}

function openScheduleAudit() {
  var html = HtmlService.createHtmlOutputFromFile('ScheduleAudit').setWidth(480).setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Schedule Next Audit');
}

// =============================================================
//  SCHEDULE AUDIT — writes to Audit Schedule tab + sends email
// =============================================================
function scheduleAudit(zone, scheduledDate, coordinatorEmail) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var cfg     = getSettings();
    var sheet   = ss.getSheetByName(cfg['Audit Schedule Sheet'] || 'Audit Schedule');
    if (!sheet) return { success:false, message:'Audit Schedule sheet not found' };
    var tz      = Session.getScriptTimeZone();
    var mgrs    = getZoneManagers();
    var mgr     = mgrs[zone] || {};
    var deptMap = {'1':'Shipping & Receiving','2':'Metals','3':'Plastics','4':'Lithography','5':'Miscellaneous'};
    var dept    = deptMap[zone] || 'Zone ' + zone;
    var lastRow = sheet.getLastRow();
    var r       = lastRow < 4 ? 4 : lastRow + 1;
    var schedId = 'SCH-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd') + '-' + String(r).padStart(3,'0');
    var dSched  = new Date(scheduledDate);
    sheet.getRange(r, 1, 1, 12).setValues([[
      schedId, zone, dept,
      (mgr.name||'') + ' <' + (mgr.email||'') + '>',
      Utilities.formatDate(dSched, tz, 'MM/dd/yyyy'),
      coordinatorEmail || '',
      'Not sent', 'Pending', '', '', '', ''
    ]]);
    var reminderDays = Number(cfg['Reminder Days Before']) || 3;
    sendAuditScheduledEmail_(r);
    scheduleReminderTrigger_(r, dSched, reminderDays);
    SpreadsheetApp.flush();
    return { success:true, schedId:schedId };
  } catch(e) { return { success:false, message:e.message }; }
}

// Sheets auto-converts the "MM/dd/yyyy" string written into the schedule
// row into a real Date value, so reading it back via getValues() returns a
// Date object — concatenating that directly into a string (as the email
// HTML did) triggers JS's default Date.toString(), producing something
// like "Sun Aug 23 2026 00:00:00 GMT-0700 (Pacific Daylight Time)".
function formatScheduleDate_(val) {
  var tz = Session.getScriptTimeZone();
  var d  = (val instanceof Date) ? val : new Date(val);
  if (isNaN(d.getTime())) return String(val || '');
  return Utilities.formatDate(d, tz, 'EEEE MMM d, yyyy');
}

// =============================================================
//  SEND "AUDIT SCHEDULED" CONFIRMATION EMAIL — fires immediately
//  when scheduleAudit() runs (separate from the later reminder)
// =============================================================
function sendAuditScheduledEmail_(scheduleRow) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var cfg     = getSettings();
    var sheet   = ss.getSheetByName(cfg['Audit Schedule Sheet'] || 'Audit Schedule');
    if (!sheet) return;
    var row         = sheet.getRange(scheduleRow, 1, 1, 12).getValues()[0];
    var schedId     = row[0];
    var zone        = row[1];
    var dept        = row[2];
    var emailMatch  = String(row[3]).match(/<(.+)>/);
    var leadEmail   = emailMatch ? emailMatch[1] : String(row[3]);
    var schedDate   = formatScheduleDate_(row[4]);
    var coordinator = row[5];
    var qaEmail     = cfg['CC Manager Email'] || '';
    var companyName = cfg['Company Name']     || 'Container Supply Co.';
    var reminderDays = Number(cfg['Reminder Days Before']) || 3;
    if (!leadEmail || leadEmail.indexOf('@') === -1) return;
    MailApp.sendEmail({
      to: leadEmail, cc: qaEmail,
      subject: companyName + ' — Facility Audit Scheduled: ' + dept + ' (' + schedDate + ')',
      htmlBody:
        '<div style="font-family:Arial;max-width:600px;color:#1E3A52">' +
        '<div style="background:#1E3A52;padding:14px 20px;color:#fff"><strong>' + companyName + '</strong><br>' +
        '<span style="font-size:11px;color:#B5D4F4">Facility Self Audit — Scheduled</span></div>' +
        '<div style="padding:20px;background:#F7F9FB;border:1px solid #D0D7E8">' +
        '<p>A facility self-inspection has been scheduled for <strong>' + dept + ' (Zone ' + zone + ')</strong>:</p>' +
        '<div style="background:#fff;border:1px solid #D0D7E8;padding:12px;margin:12px 0;font-size:16px;font-weight:bold;color:#2D5F82">' + schedDate + '</div>' +
        '<p style="font-size:12px;color:#555;margin-bottom:4px">Coordinator: ' + (coordinator||'TBD') + ' | ID: ' + schedId + '</p>' +
        '<p style="font-size:11px;color:#888">You will receive a reminder ' + reminderDays + ' day' + (reminderDays === 1 ? '' : 's') + ' before the scheduled date.</p>' +
        '</div></div>'
    });
  } catch(e) { Logger.log('sendAuditScheduledEmail_ error: ' + e.message); }
}

// =============================================================
//  SEND AUDIT REMINDER EMAIL
// =============================================================
function sendAuditReminder(scheduleRow) {
  try {
    var ss      = SpreadsheetApp.getActiveSpreadsheet();
    var cfg     = getSettings();
    var sheet   = ss.getSheetByName(cfg['Audit Schedule Sheet'] || 'Audit Schedule');
    if (!sheet) return;
    var row         = sheet.getRange(scheduleRow, 1, 1, 12).getValues()[0];
    var schedId     = row[0];
    var zone        = row[1];
    var dept        = row[2];
    var emailMatch  = String(row[3]).match(/<(.+)>/);
    var leadEmail   = emailMatch ? emailMatch[1] : String(row[3]);
    var schedDate   = formatScheduleDate_(row[4]);
    var coordinator = row[5];
    var qaEmail     = cfg['CC Manager Email'] || '';
    var companyName = cfg['Company Name']      || 'Container Supply Co.';
    if (!leadEmail || leadEmail.indexOf('@') === -1) return;
    MailApp.sendEmail({
      to: leadEmail, cc: qaEmail,
      subject: companyName + ' — Facility Audit Reminder: ' + dept + ' (' + schedDate + ')',
      htmlBody:
        '<div style="font-family:Arial;max-width:600px;color:#1E3A52">' +
        '<div style="background:#1E3A52;padding:14px 20px;color:#fff"><strong>' + companyName + '</strong><br>' +
        '<span style="font-size:11px;color:#B5D4F4">Facility Self Audit — Reminder</span></div>' +
        '<div style="padding:20px;background:#F7F9FB;border:1px solid #D0D7E8">' +
        '<p>Inspection for <strong>' + dept + ' (Zone ' + zone + ')</strong> is scheduled for:</p>' +
        '<div style="background:#fff;border:1px solid #D0D7E8;padding:12px;margin:12px 0;font-size:16px;font-weight:bold;color:#2D5F82">' + schedDate + '</div>' +
        '<p style="font-size:12px;color:#555;margin-bottom:16px">Coordinator: ' + (coordinator||'TBD') + ' | ID: ' + schedId + '</p>' +
        '<p style="font-size:12px;color:#555">If you need to reschedule this inspection, please reach out to the coordinator directly.</p>' +
        '</div></div>'
    });
    sheet.getRange(scheduleRow, 7).setValue('Sent ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MM/dd/yyyy'));
    SpreadsheetApp.flush();
  } catch(e) { Logger.log('sendAuditReminder error: ' + e.message); }
}

function scheduleReminderTrigger_(schedRow, auditDate, daysBefore) {
  var reminderDate = new Date(auditDate.getTime());
  reminderDate.setDate(reminderDate.getDate() - daysBefore);
  if (reminderDate > new Date()) {
    var trigger = ScriptApp.newTrigger('sendReminderFromTrigger_').timeBased().at(reminderDate).create();
    // Key by this trigger's own unique ID — a shared key would let two pending
    // reminders (e.g. for different zones) stomp on each other's row number.
    PropertiesService.getScriptProperties().setProperty('pending_reminder_' + trigger.getUniqueId(), String(schedRow));
  } else {
    sendAuditReminder(schedRow);
  }
}

function sendReminderFromTrigger_(e) {
  var uid  = e && e.triggerUid;
  var prop = uid ? 'pending_reminder_' + uid : null;
  var row  = prop ? PropertiesService.getScriptProperties().getProperty(prop) : null;
  if (row) {
    sendAuditReminder(parseInt(row, 10));
    PropertiesService.getScriptProperties().deleteProperty(prop);
  }
  // One-time trigger has fired — remove it so triggers don't accumulate.
  if (uid) {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getUniqueId() === uid) ScriptApp.deleteTrigger(t);
    });
  }
}

// =============================================================
//  FORMAT / SETUP HELPERS (informational only)
// =============================================================
function formatCALog() {
  SpreadsheetApp.getUi().alert(
    'CA Log settings:\n' +
    '• Sheet: ' + FA.caLogSheet + '\n' +
    '• Data starts at row: ' + FA.dataStartRow + '\n' +
    '• Columns: ' + CA_NUM_COLS + '\n\n' +
    'Column order: CAR ID | Audit Date | NC Date | Zone | Department | ' +
    'Inspection Item | Location | NC Level | Description | ' +
    'Corrective Action Plan | Completed By | Assigned By | Due Date | ' +
    'Completion Date | Verified By | Verification Date | Photo Link | Status | Deferred');
}

function formatAuditLog() {
  SpreadsheetApp.getUi().alert(
    'Audit Log settings:\n' +
    '• Sheet: ' + FA.auditLogSheet + '\n' +
    '• Data starts at row: ' + FA.auditDataStartRow + '\n' +
    '• Columns: ' + AL_NUM_COLS);
}

function setupConfigBlock() {
  SpreadsheetApp.getUi().alert(
    'Configuration is managed in the Settings tab.\n\n' +
    'Edit column B for:\n' +
    '  CC Manager Email\n' +
    '  CA Log Sheet Name\n' +
    '  Audit Log Sheet Name\n' +
    '  Pill Library Sheet\n' +
    '  Audit Schedule Sheet\n' +
    '  Reminder Days Before\n' +
    '  Past Due Grace Days\n' +
    '  Web App URL\n\n' +
    'Zone manager emails are in the ZONE MANAGERS table in Settings.');
}

// =============================================================
//  GET CA REPORTS FOLDER ID FROM SETTINGS
// =============================================================
function getCAReportsFolderId() {
  var cfg = getSettings();
  return cfg['CA Reports Folder ID'] || '';
}

// =============================================================
//  CA REPORT PDF — built via DocumentApp Table API (server-side,
//  reliable — no browser canvas rendering involved). Matches the
//  approved FRM-018-004 sections/colors/terminology.
// =============================================================

// Full-width colored banner as its own standalone table — used only where
// it can't be folded into the section's outer table (Section 2, whose row
// count varies with the number of photos).
function addSectionBanner_(body, text, bgColor) {
  var table = body.appendTable([['']]);
  table.setBorderWidth(0);
  fillBannerCell_(table.getCell(0, 0), text, bgColor);
  return table;
}

// Apps Script's Table API has no real cell-merge/colspan, so a "banner"
// spanning multiple columns of a data table always shows the grid lines
// running through it. The fix: every section's OUTER table is a single
// column, so a banner or a long label/value block naturally spans the
// full width with zero seams; multi-column data (CAR ID | NC Date | ...)
// lives in a NESTED table dropped into one cell of that outer column.

// Fills a cell with bold white banner text on a colored background.
function fillBannerCell_(cell, text, bgColor) {
  cell.setBackgroundColor(bgColor);
  cell.setPaddingTop(3).setPaddingBottom(3).setPaddingLeft(10).setPaddingRight(10);
  var para = cell.getChild(0).asParagraph();
  para.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  para.appendText(text);
  para.editAsText().setBold(true).setForegroundColor('#FFFFFF').setFontSize(9);
}

// Fills a cell with a small gray label line + a value line below it —
// used for full-width fields (Description, Corrective Action Plan).
function fillLongCell_(cell, label, value) {
  cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(8).setPaddingRight(8);
  var labelPara = cell.getChild(0).asParagraph();
  labelPara.setSpacingBefore(0).setSpacingAfter(1).setLineSpacing(1);
  labelPara.appendText(label.toUpperCase());
  labelPara.editAsText().setFontSize(7).setForegroundColor('#6B7280').setBold(true);
  var valuePara = cell.appendParagraph(value || '—');
  valuePara.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  valuePara.editAsText().setFontSize(10).setForegroundColor('#1E3A52');
}

// Fills row `rowIdx` of `table` with one label/value field per column.
function fillFieldRow_(table, rowIdx, fields, ncols) {
  for (var c = 0; c < ncols; c++) {
    var f = fields[c];
    var cell = table.getCell(rowIdx, c);
    cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(8).setPaddingRight(8);
    if (!f) continue;
    var labelPara = cell.getChild(0).asParagraph();
    labelPara.setSpacingBefore(0).setSpacingAfter(1).setLineSpacing(1);
    labelPara.appendText((f.label || '').toUpperCase());
    labelPara.editAsText().setFontSize(7).setForegroundColor('#6B7280').setBold(true);
    var valuePara = cell.appendParagraph(f.value || '—');
    valuePara.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    valuePara.editAsText().setFontSize(10).setForegroundColor(f.color || '#1E3A52').setBold(!!f.bold);
    if (f.mono) valuePara.editAsText().setFontFamily('Courier New');
  }
}

// Docs applies some default vertical spacing to a table itself (separate
// from cell padding), with no dedicated getter/setter on Table the way
// Paragraph has setSpacingBefore/After — this is the only lever available
// for it. Not confirmed to work in every case (may only apply to some
// element/attribute combos), but it's the best available shot at closing
// gaps around/between tables beyond what padding already controls.
function zeroTableSpacing_(table) {
  var attrs = {};
  attrs[DocumentApp.Attribute.MARGIN_TOP]     = 0;
  attrs[DocumentApp.Attribute.MARGIN_BOTTOM]  = 0;
  attrs[DocumentApp.Attribute.SPACING_BEFORE] = 0;
  attrs[DocumentApp.Attribute.SPACING_AFTER]  = 0;
  table.setAttributes(attrs);
  return table;
}

// Drops a bordered N-column field table into `parentCell` — this is how
// multi-column data rows end up flush inside a single-column outer table.
// `fieldRows` is an array of field-arrays — pass more than one row (e.g.
// CAR ID/NC Date/... AND Inspection Item/Location/...) to get them into
// ONE connected nested table instead of two separately-boxed ones.
function buildNestedFieldTable_(parentCell, fieldRows) {
  // Every cell starts with one empty paragraph; shrink it instead of
  // leaving a default-sized blank line sitting above the nested table.
  var pre = parentCell.getChild(0).asParagraph();
  pre.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  pre.editAsText().setFontSize(1);
  var ncols  = fieldRows[0].length;
  var initial = fieldRows.map(function(){
    var row = [];
    for (var i = 0; i < ncols; i++) row.push('');
    return row;
  });
  var nested = parentCell.appendTable(initial);
  fieldRows.forEach(function(fields, ri) { fillFieldRow_(nested, ri, fields, ncols); });
  nested.setBorderColor('#D0D7E8');
  zeroTableSpacing_(nested);
  return nested;
}

// Photo grid — real images embedded via DriveApp, 3 per row
function addPhotosSection_(body, photoLinks) {
  if (!photoLinks.length) {
    var p = body.appendParagraph('No photos attached.');
    p.setSpacingBefore(0).setSpacingAfter(0);
    p.editAsText().setItalic(true).setForegroundColor('#6B7280').setFontSize(9);
    return;
  }
  var perRow = 3;
  for (var i = 0; i < photoLinks.length; i += perRow) {
    var rowLinks = photoLinks.slice(i, i + perRow);
    var initial  = rowLinks.map(function(){ return ''; });
    var table    = body.appendTable([initial]);
    rowLinks.forEach(function(link, ci) {
      var cell = table.getCell(0, ci);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(2);
      var placeholderPara = cell.getChild(0).asParagraph();
      placeholderPara.setSpacingBefore(0).setSpacingAfter(0);
      try {
        var m      = link.match(/\/file\/d\/([^\/\?&]+)/) || link.match(/[?&]id=([^&]+)/);
        var fileId = m ? m[1].replace(/\/$/, '') : '';
        if (!fileId) throw new Error('Could not parse file ID from link');
        var blob = DriveApp.getFileById(fileId).getBlob();
        var img  = cell.appendImage(blob);
        var maxWidth = 160;
        if (img.getWidth() > maxWidth) {
          img.setHeight(Math.round(img.getHeight() * (maxWidth / img.getWidth())));
          img.setWidth(maxWidth);
        }
      } catch (err) {
        Logger.log('Photo embed error: ' + err.message);
        placeholderPara.appendText('[ Photo unavailable ]');
        placeholderPara.editAsText().setItalic(true).setForegroundColor('#991B1B').setFontSize(8);
      }
    });
    table.setBorderWidth(0);
  }
}

// 3-column report header — company / title / doc info. The doc-info column
// is a small nested table (Doc No. / Revision No. / Revision Date) so each
// row gets a real divider line and its own visible label. The outer table
// keeps a real border so the title cell reads as a complete box on all
// four sides — the nested doc-info table sits inset (padding) inside its
// own cell, so the two borders read as "a card inside a card," not a clash.
function addReportHeader_(body, companyName) {
  var table = body.appendTable([['', '', '']]);
  // Columns sum to 568pt — matching the page's usable width (612pt Letter
  // minus 22pt margins each side) so the header lines up flush with the
  // section tables below it, which auto-fill that same width.
  table.setColumnWidth(0, 100);
  table.setColumnWidth(1, 300);
  table.setColumnWidth(2, 168);
  table.setBorderColor('#D0D7E8');
  table.setBorderWidth(0.75);

  var c1 = table.getCell(0, 0);
  c1.setBackgroundColor('#0F2D5C');
  c1.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
  var p1 = c1.getChild(0).asParagraph();
  p1.setSpacingBefore(0).setSpacingAfter(2).setLineSpacing(1);
  p1.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p1.appendText(companyName);
  p1.editAsText().setBold(true).setForegroundColor('#FFFFFF').setFontSize(10);
  var p1b = c1.appendParagraph('Garden Grove, CA');
  p1b.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  p1b.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p1b.editAsText().setForegroundColor('#B5D4F4').setFontSize(7);

  var c2 = table.getCell(0, 1);
  c2.setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
  var p2 = c2.getChild(0).asParagraph();
  p2.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  p2.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p2.appendText('Facility Self Audit');
  p2.editAsText().setBold(true).setForegroundColor('#1E3A52').setFontSize(12);
  var p2b = c2.appendParagraph('Corrective Action Form');
  p2b.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  p2b.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p2b.editAsText().setBold(true).setForegroundColor('#1E3A52').setFontSize(12);

  var c3 = table.getCell(0, 2);
  // Zero padding — the nested table's own border then sits flush against
  // this cell's outer border (no gap between the two lines), so they read
  // as one continuous box instead of a visibly separate "box inside a box."
  c3.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(0).setPaddingRight(0);
  // This leading paragraph can't be removed (every cell needs at least one),
  // so shrink it to near-zero instead of leaving a default-sized blank line
  // sitting above the nested table.
  var c3Pre = c3.getChild(0).asParagraph();
  c3Pre.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
  c3Pre.editAsText().setFontSize(1);
  var infoTable = c3.appendTable([
    ['Doc No.',        'FRM-018-004'],
    ['Revision No.:',  '0'],
    ['Revision Date:', '07/17/2026'],
  ]);
  infoTable.setBorderColor('#D0D7E8');
  infoTable.setBorderWidth(0.75);
  zeroTableSpacing_(infoTable);
  for (var ri = 0; ri < 3; ri++) {
    var lblCell = infoTable.getCell(ri, 0);
    lblCell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(2);
    var lblPara = lblCell.getChild(0).asParagraph();
    lblPara.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    lblPara.editAsText().setFontSize(7).setForegroundColor('#6B7280');
    var valCell = infoTable.getCell(ri, 1);
    valCell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(2).setPaddingRight(4);
    var valPara = valCell.getChild(0).asParagraph();
    valPara.setSpacingBefore(0).setSpacingAfter(0).setLineSpacing(1);
    valPara.editAsText().setFontSize(8).setForegroundColor('#1E3A52').setBold(true);
  }

  return table;
}

// Builds one CA record's full report as a PDF blob. Reliable and server-side —
// no browser rendering involved, so nothing to go blank.
function buildCARReportPDF_(car, dept, companyName) {
  var tempDoc = DocumentApp.create('TEMP_CAR_' + car.carId);
  var docId   = tempDoc.getId();
  var body    = tempDoc.getBody();
  body.clear();
  body.setMarginTop(22).setMarginBottom(22).setMarginLeft(22).setMarginRight(22);

  addReportHeader_(body, companyName);

  // SECTION 1 — one single-column outer table: banner row, one nested
  // table holding BOTH data rows (CAR ID/... and Inspection Item/...) so
  // they're connected with a single shared border, then Description — all
  // flush, no gaps, and the banner/description genuinely span full width.
  var s1 = body.appendTable([[''],[''],['']]);
  s1.setBorderWidth(0);
  fillBannerCell_(s1.getCell(0, 0), 'SECTION 1 — NONCONFORMANCE IDENTIFICATION', '#4A7FA5');
  buildNestedFieldTable_(s1.getCell(1, 0), [
    [
      { label: 'CAR ID', value: car.carId, mono: true, bold: true, color: '#2D5F82' },
      { label: 'NC Date', value: car.ncDate },
      { label: 'Zone', value: 'Zone ' + car.zone + ' — ' + dept },
      { label: 'Audit Date', value: car.auditDate },
    ],
    [
      { label: 'Inspection Item', value: car.inspectionItem },
      { label: 'Location within zone', value: car.locationInZone },
      { label: 'NC Level', value: car.ncLevel },
      { label: 'Due Date', value: car.dueDate },
    ],
  ]);
  fillLongCell_(s1.getCell(2, 0), 'Description', car.description);

  addSectionBanner_(body, 'SECTION 2 — SUPPORTING IMAGES', '#2D5F82');
  var photoLinks = String(car.photoLink || '').split(/[\n,]+/).map(function(l){ return l.trim(); }).filter(Boolean);
  addPhotosSection_(body, photoLinks);

  // SECTION 3 — banner, Corrective Action Plan, then the Completed By /
  // Assigned By nested table, all rows of one single-column outer table.
  var s3 = body.appendTable([[''],[''],['']]);
  s3.setBorderWidth(0);
  zeroTableSpacing_(s3);
  fillBannerCell_(s3.getCell(0, 0), 'SECTION 3 — CORRECTIVE ACTION', '#1E3A52');
  fillLongCell_(s3.getCell(1, 0), 'Corrective Action Plan', car.correctiveAction);
  buildNestedFieldTable_(s3.getCell(2, 0), [[
    { label: 'Completed By', value: car.responsiblePerson },
    { label: 'Assigned By', value: car.assignedBy },
  ]]);

  // SECTION 4 — banner + the completion/verification nested table.
  var s4 = body.appendTable([[''],['']]);
  s4.setBorderWidth(0);
  zeroTableSpacing_(s4);
  fillBannerCell_(s4.getCell(0, 0), 'SECTION 4 — VERIFICATION AND CLOSURE', '#111E2E');
  buildNestedFieldTable_(s4.getCell(1, 0), [[
    { label: 'Completion Date', value: car.completionDate },
    { label: 'Verification Date', value: car.verificationDate },
    { label: 'Verified By', value: car.verifiedBy },
    { label: 'Status', value: car.status || 'Open' },
  ]]);

  var footer = body.appendParagraph('This document is confidential. Printed copies are uncontrolled.  ·  FRM-018-004  ·  ' + companyName);
  footer.setSpacingBefore(4).setSpacingAfter(0);
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setFontSize(7).setForegroundColor('#9CA3AF').setItalic(true);

  tempDoc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(docId).getAs('application/pdf');
  DriveApp.getFileById(docId).setTrashed(true);
  return pdfBlob;
}

// =============================================================
//  GENERATE CA PDFS AND EMAIL — main function
//  Called from CAUpdateForm.html's "Send Notification" flow after
//  the user confirms in the preview overlay
// =============================================================
function generateAndEmailCARReports(carIds, zone) {
  try {
    var cfg         = getSettings();
    var folderId    = cfg['CA Reports Folder ID'] || '';
    var companyName = cfg['Company Name'] || 'Container Supply Co.';
    var qaEmail     = cfg['CC Manager Email'] || '';

    if (!folderId) return { success: false, message: 'CA Reports Folder ID not set in Settings tab.' };

    // Get zone manager email
    var mgrs      = getZoneManagers();
    var mgr       = mgrs[String(zone).trim()] || {};
    var leadEmail = mgr.email || '';
    var leadName  = mgr.name  || ('Zone ' + zone + ' Lead');

    var zoneNames = {'1':'Shipping & Receiving','2':'Metals','3':'Plastics','4':'Lithography','5':'Miscellaneous'};
    var dept      = zoneNames[String(zone).trim()] || 'Zone ' + zone;

    if (!leadEmail || leadEmail.indexOf('@') === -1) {
      return { success: false, message: 'No valid email found for Zone ' + zone + ' in Settings tab.' };
    }

    // Fetch matching CA records
    var allRecords = getCARecords();
    var selected   = allRecords.filter(function(r) {
      return carIds.indexOf(r.carId) !== -1;
    });

    if (!selected.length) return { success: false, message: 'No matching records found for selected CAR IDs.' };

    // Create Drive subfolder for this batch
    var ss         = SpreadsheetApp.getActiveSpreadsheet();
    var tz         = Session.getScriptTimeZone();
    var batchDate  = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var batchName  = 'Zone-' + zone + '_' + batchDate;
    var rootFolder = DriveApp.getFolderById(folderId);
    var batchFolders = rootFolder.getFoldersByName(batchName);
    var batchFolder  = batchFolders.hasNext() ? batchFolders.next() : rootFolder.createFolder(batchName);

    // Generate each record's styled report PDF server-side and save it to
    // the batch folder — reliable, no browser rendering involved.
    var pdfBlobs = [];
    var pdfFiles = [];

    selected.forEach(function(car) {
      var filename = car.carId + '_' + dept.replace(/[^a-zA-Z0-9]/g,'_') + '.pdf';
      var pdfBlob  = buildCARReportPDF_(car, dept, companyName);
      pdfBlob.setName(filename);
      var savedPDF = batchFolder.createFile(pdfBlob);
      pdfFiles.push({ name: filename, url: savedPDF.getUrl() });
      pdfBlobs.push(pdfBlob);
    });

    var emailBody =
      '<div style="font-family:Arial,sans-serif;max-width:640px;color:#1E3A52">'
      + '<div style="background:#1E3A52;padding:14px 20px;color:#fff">'
      + '<strong style="font-size:15px">' + companyName + '</strong><br>'
      + '<span style="font-size:11px;color:#B5D4F4">Facility Self Audit — Corrective Action Reports</span>'
      + '</div>'
      + '<div style="padding:20px;background:#F7F9FB;border:1px solid #D0D7E8">'
      + '<p style="margin-bottom:12px">Dear ' + leadName + ',</p>'
      + '<p style="margin-bottom:12px">Please find attached the Corrective Action Report(s) for <strong>' + dept + ' (Zone ' + zone + ')</strong> resulting from the most recent facility self-inspection.</p>'
      + '<p style="margin-bottom:8px"><strong>' + selected.length + ' report(s) attached:</strong></p>'
      + '<ul style="margin-bottom:16px;padding-left:20px">'
      + selected.map(function(r) {
          return '<li style="margin-bottom:4px;font-size:12px"><strong>' + r.carId + '</strong> — ' + (r.inspectionItem||'') + ' (' + (r.ncLevel||'') + ')</li>';
        }).join('')
      + '</ul>'
      + '<p style="margin-bottom:12px">Please review each report, complete the <strong>Corrective Action</strong> section, sign, and return to the QA department by the due date indicated on each form.</p>'
      + '<p style="font-size:11px;color:#555;margin-bottom:4px">Reference: FRM-018-004 &nbsp;&middot;&nbsp; FRM-018-002</p>'
      + '<p style="font-size:10px;color:#888">This is an automated notification from ' + companyName + ' Quality Assurance. Contact ' + (qaEmail||'the QA team') + ' with any questions.</p>'
      + '</div></div>';

    // Send email with all PDFs attached
    MailApp.sendEmail({
      to:          leadEmail,
      cc:          qaEmail,
      subject:     companyName + ' — CA Reports: ' + dept + ' (' + batchDate + ')',
      htmlBody:    emailBody,
      attachments: pdfBlobs,
    });

    // Log sent date on each CA row
    var ss2   = SpreadsheetApp.getActiveSpreadsheet();
    var cfg2  = getSettings();
    var sheet = ss2.getSheetByName(cfg2['CA Log Sheet Name'] || FA.caLogSheet);
    if (sheet) {
      selected.forEach(function(car) {
        var lastRow = getLastDataRow_(sheet, 1, FA.dataStartRow);
        var ids = sheet.getRange(FA.dataStartRow, CA_COLS.carId,
                                 lastRow - FA.dataStartRow + 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]).trim() === car.carId) {
            var r = FA.dataStartRow + i;
            // Add note in Notes-adjacent column if available, or log to execution log
            Logger.log('CA Report emailed for ' + car.carId + ' on ' + batchDate);
            break;
          }
        }
      });
    }

    SpreadsheetApp.flush();
    return {
      success:   true,
      count:     selected.length,
      leadEmail: leadEmail,
      folderUrl: batchFolder.getUrl(),
    };

  } catch(e) {
    Logger.log('generateAndEmailCARReports ERROR: ' + e.message + '\n' + e.stack);
    return { success: false, message: e.message };
  }
}

function escHtml_(v) {
  return String(v || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// NC level → { label, color, bg } — mirrors the pill colors used throughout
// the app (CAUpdateForm.html's preview list, ManagerReport.html's donut
// chips) and the correction window from getDueDate() above, so the emailed
// table reads the same way the rest of the system does.
var NC_LEVEL_META = {
  Critical:    { color:'#7F1D1D', bg:'#FEF2F2', days:'Same day' },
  Major:       { color:'#78350F', bg:'#FFFBEB', days:'14 days' },
  Minor:       { color:'#1E3A8A', bg:'#EFF6FF', days:'29 days' },
  Observation: { color:'#4B5563', bg:'#F3F4F6', days:'N/A' },
};

// Builds the color-coded "what's due, and by when" table in the CA report email.
function buildCAEmailTable_(records) {
  var headerCells = ['CAR ID', 'Inspection Item', 'NC Level', 'Time to Correct', 'Due Date']
    .map(function(h) { return '<th style="padding:6px 8px;text-align:left;font-size:9px;text-transform:uppercase;letter-spacing:.04em">' + h + '</th>'; })
    .join('');

  var rows = records.map(function(r) {
    var meta = NC_LEVEL_META[r.ncLevel] || { color:'#4B5563', bg:'#F3F4F6', days:'N/A' };
    return '<tr style="border-bottom:1px solid #D0D7E8">'
      + '<td style="padding:6px 8px;font-weight:bold;color:#2D5F82;font-size:11px">' + escHtml_(r.carId) + '</td>'
      + '<td style="padding:6px 8px;font-size:11px">' + escHtml_(r.inspectionItem) + '</td>'
      + '<td style="padding:6px 8px"><span style="display:inline-block;padding:2px 8px;border-radius:3px;font-weight:bold;font-size:10px;color:' + meta.color + ';background:' + meta.bg + '">' + escHtml_(r.ncLevel || '—') + '</span></td>'
      + '<td style="padding:6px 8px;font-size:11px">' + meta.days + '</td>'
      + '<td style="padding:6px 8px;font-size:11px">' + escHtml_(r.dueDate || '—') + '</td>'
      + '</tr>';
  }).join('');

  return '<table style="width:100%;border-collapse:collapse;margin-bottom:16px;font-size:11px">'
    + '<thead><tr style="background:#1E3A52;color:#fff">' + headerCells + '</tr></thead>'
    + '<tbody>' + rows + '</tbody>'
    + '</table>';
}

// =============================================================
//  GENERATE CA PDFS AND EMAIL — from client-rendered PDFs
//  Same as generateAndEmailCARReports() above (Drive folder, zone-lead
//  lookup, email body, logging), but the PDF bytes are supplied by the
//  browser instead of being built server-side with DocumentApp. The
//  browser renders the exact same CSS-styled report used for Print/Save
//  (via html2canvas + jsPDF), so this PDF matches that look exactly —
//  DocumentApp's table/paragraph model can't reproduce that CSS layout.
//  pdfPayload: [{ carId, base64 }, ...] — one entry per selected record.
// =============================================================
function generateAndEmailCARReportsFromPDFs(carIds, zone, pdfPayload) {
  try {
    var cfg         = getSettings();
    var folderId    = cfg['CA Reports Folder ID'] || '';
    var companyName = cfg['Company Name'] || 'Container Supply Co.';
    var qaEmail     = cfg['CC Manager Email'] || '';

    if (!folderId) return { success: false, message: 'CA Reports Folder ID not set in Settings tab.' };

    var mgrs      = getZoneManagers();
    var mgr       = mgrs[String(zone).trim()] || {};
    var leadEmail = mgr.email || '';
    var leadName  = mgr.name  || ('Zone ' + zone + ' Lead');

    var zoneNames = {'1':'Shipping & Receiving','2':'Metals','3':'Plastics','4':'Lithography','5':'Miscellaneous'};
    var dept      = zoneNames[String(zone).trim()] || 'Zone ' + zone;

    if (!leadEmail || leadEmail.indexOf('@') === -1) {
      return { success: false, message: 'No valid email found for Zone ' + zone + ' in Settings tab.' };
    }

    var allRecords = getCARecords();
    var selected   = allRecords.filter(function(r) {
      return carIds.indexOf(r.carId) !== -1;
    });
    if (!selected.length) return { success: false, message: 'No matching records found for selected CAR IDs.' };

    var pdfByCarId = {};
    (pdfPayload || []).forEach(function(p) { pdfByCarId[p.carId] = p.base64; });

    var tz          = Session.getScriptTimeZone();
    var sendDateStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    var rootFolder  = DriveApp.getFolderById(folderId);

    // Name the folder after the audit's month, not the send date — a send
    // batches every NC from one audit together, e.g. "July 2026 Audit -
    // Zone 1 Shipping & Receiving". Falls back to the send date's month if
    // the first selected record has no parseable audit date.
    var auditMonthLabel = '';
    var firstAuditDate  = selected[0] && selected[0].auditDate;
    if (firstAuditDate) {
      var auditD = new Date(firstAuditDate);
      if (!isNaN(auditD.getTime())) auditMonthLabel = Utilities.formatDate(auditD, tz, 'MMMM yyyy');
    }
    if (!auditMonthLabel) auditMonthLabel = Utilities.formatDate(new Date(), tz, 'MMMM yyyy');

    var batchName   = auditMonthLabel + ' Audit - Zone ' + zone + ' ' + dept;
    var batchFolder = rootFolder.createFolder(batchName); // one new folder per send, by design

    var pdfBlobs = [];
    var sent     = [];

    selected.forEach(function(car) {
      var base64 = pdfByCarId[car.carId];
      if (!base64) return; // browser failed to render this one — skip rather than fail the whole batch
      var filename = car.carId + '_' + dept.replace(/[^a-zA-Z0-9]/g,'_') + '.pdf';
      var pdfBlob  = Utilities.newBlob(Utilities.base64Decode(base64), 'application/pdf', filename);
      batchFolder.createFile(pdfBlob);
      pdfBlobs.push(pdfBlob);
      sent.push(car);
    });

    if (!pdfBlobs.length) {
      return { success: false, message: 'No report PDFs were received from the browser — nothing to send.' };
    }

    var emailBody =
      '<div style="font-family:Arial,sans-serif;max-width:640px;color:#1E3A52">'
      + '<div style="background:#1E3A52;padding:14px 20px;color:#fff">'
      + '<strong style="font-size:15px">' + companyName + '</strong><br>'
      + '<span style="font-size:11px;color:#B5D4F4">Facility Self Audit — Corrective Action Reports</span>'
      + '</div>'
      + '<div style="padding:20px;background:#F7F9FB;border:1px solid #D0D7E8">'
      + '<p style="margin-bottom:12px">Dear ' + leadName + ',</p>'
      + '<p style="margin-bottom:12px">Please find attached the Corrective Action Report(s) for <strong>' + dept + ' (Zone ' + zone + ')</strong> resulting from the most recent facility self-inspection.</p>'
      + '<p style="margin-bottom:8px"><strong>' + sent.length + ' report(s) attached:</strong></p>'
      + buildCAEmailTable_(sent)
      + '<p style="margin-bottom:12px">Please review each report, complete the <strong>Corrective Action</strong> section, sign, and return to the QA department by the due date indicated on each form.</p>'
      + '<p style="font-size:11px;color:#555;margin-bottom:4px">Reference: FRM-018-004 &nbsp;&middot;&nbsp; FRM-018-002</p>'
      + '<p style="font-size:10px;color:#888">This is an automated notification from ' + companyName + ' Quality Assurance. Contact ' + (qaEmail||'the QA team') + ' with any questions.</p>'
      + '</div></div>';

    MailApp.sendEmail({
      to:          leadEmail,
      cc:          qaEmail,
      subject:     companyName + ' — CA Reports: ' + dept + ' (' + sendDateStr + ')',
      htmlBody:    emailBody,
      attachments: pdfBlobs,
    });

    Logger.log('CA reports emailed (client-rendered PDFs): ' + sent.map(function(r){ return r.carId; }).join(', ') + ' on ' + sendDateStr);

    SpreadsheetApp.flush();
    return {
      success:   true,
      count:     sent.length,
      leadEmail: leadEmail,
      folderUrl: batchFolder.getUrl(),
    };

  } catch(e) {
    Logger.log('generateAndEmailCARReportsFromPDFs ERROR: ' + e.message + '\n' + e.stack);
    return { success: false, message: e.message };
  }
}

// =============================================================
//  GET PREVIEW DATA FOR EMAIL CONFIRMATION
//  Called before sending — returns data to show in preview modal
// =============================================================
function getEmailPreviewData(carIds, zone) {
  try {
    var cfg       = getSettings();
    var mgrs      = getZoneManagers();
    var mgr       = mgrs[String(zone).trim()] || {};
    var allRecords = getCARecords();
    var selected   = allRecords.filter(function(r) {
      return carIds.indexOf(r.carId) !== -1;
    });
    var zoneNames = {'1':'Shipping & Receiving','2':'Metals','3':'Plastics','4':'Lithography','5':'Miscellaneous'};
    return {
      success:    true,
      leadName:   mgr.name  || 'Zone ' + zone + ' Lead',
      leadEmail:  mgr.email || '(not set in Settings)',
      ccEmail:    cfg['CC Manager Email'] || '',
      dept:       zoneNames[String(zone).trim()] || 'Zone ' + zone,
      zone:       zone,
      count:      selected.length,
      records:    selected.map(function(r) {
        return { carId: r.carId, inspectionItem: r.inspectionItem, ncLevel: r.ncLevel, dueDate: r.dueDate };
      }),
      folderSet:  !!(cfg['CA Reports Folder ID']),
    };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// =============================================================
//  UPDATE CA RECORD — called from CAUpdateForm.html
//  Lead fills in corrective action and returns form
// =============================================================
function updateCARecord(carId, data) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var cfg   = getSettings();
    var sheet = ss.getSheetByName(cfg['CA Log Sheet Name'] || FA.caLogSheet);
    if (!sheet) return { success: false, message: 'CA Log sheet not found.' };

    var lastRow = getLastDataRow_(sheet, 1, FA.dataStartRow);
    var ids = sheet.getRange(FA.dataStartRow, CA_COLS.carId,
                             lastRow - FA.dataStartRow + 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === String(carId).trim()) {
        var r = FA.dataStartRow + i;
        if (data.correctiveAction)  sheet.getRange(r, CA_COLS.correctiveAction).setValue(data.correctiveAction);
        if (data.responsiblePerson) sheet.getRange(r, CA_COLS.responsiblePerson).setValue(data.responsiblePerson);
        if (data.assignedBy)        sheet.getRange(r, CA_COLS.assignedBy).setValue(data.assignedBy);
        if (data.completionDate)    sheet.getRange(r, CA_COLS.completionDate).setValue(data.completionDate);
        if (data.verifiedBy)        sheet.getRange(r, CA_COLS.verifiedBy).setValue(data.verifiedBy);
        if (data.verificationDate)  sheet.getRange(r, CA_COLS.verificationDate).setValue(data.verificationDate);
        if (data.status) {
          sheet.getRange(r, CA_COLS.status).setValue(data.status);
          applyStatusColor_(sheet, r, data.status);
        }

        var photoUploadError = null;
        if (data.newPhotos && data.newPhotos.length > 0) {
          try {
            var zone      = String(sheet.getRange(r, CA_COLS.zone).getValue() || '').trim();
            var auditDate = String(sheet.getRange(r, CA_COLS.auditDate).getValue() || '').trim();
            var folder    = getOrCreateAuditPhotoFolder_(zone, auditDate || new Date().toISOString());
            var newLinks  = [];
            var total     = data.newPhotos.length;
            data.newPhotos.forEach(function(p, pi) {
              var base64Data = p.dataUrl.split(',')[1];
              var mimeType   = p.dataUrl.split(';')[0].replace('data:', '');
              var ext = (p.name || 'photo.jpg').split('.').pop().toLowerCase() || 'jpg';
              var fileName = carId + '_Photo-' + (pi+1) + '-of-' + total + '.' + ext;
              var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
              var file = folder.createFile(blob);
              try {
                // Domain-restricted sharing — org policy commonly blocks ANYONE_WITH_LINK
                // outright, which would throw here and skip recording the link below.
                file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
              } catch (shareErr) {
                Logger.log('Photo sharing warning (' + fileName + '): ' + shareErr.message);
              }
              newLinks.push('https://drive.google.com/file/d/' + file.getId() + '/view');
            });
            var photoCell        = sheet.getRange(r, CA_COLS.photoLink);
            var existingFormula  = photoCell.getFormula();
            var existingLinks    = existingFormula
              ? extractPhotoLinksFromFormula_(existingFormula)
              : String(photoCell.getValue() || '').split(/[\n,]+/).map(function(l){ return l.trim(); }).filter(Boolean);
            var allLinks = existingLinks.concat(newLinks);
            photoCell.setFormula(buildPhotoLinkFormula_(allLinks));
          } catch(photoErr) {
            // Surface this instead of swallowing it — a Drive upload can succeed
            // while the sheet write still fails (e.g. a protected range on the
            // CA Log sheet), which otherwise looks like a silent no-op to the user.
            Logger.log('Photo upload error: ' + photoErr.message);
            photoUploadError = photoErr.message;
          }
        }
        SpreadsheetApp.flush();

        return { success: true, photoError: photoUploadError };
      }
    }
    return { success: false, message: 'CAR ID not found: ' + carId };
  } catch(e) {
    Logger.log('updateCARecord ERROR: ' + e.message);
    return { success: false, message: e.message };
  }
}
