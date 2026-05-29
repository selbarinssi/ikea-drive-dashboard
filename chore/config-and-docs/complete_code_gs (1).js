// ================================================================
// D&A DRIVE — Complete code.gs
// ================================================================

// ---- Sheet names ----
var SHEET_VEHICLES    = 'Vehicles';
var SHEET_DAILY       = 'Daily Checks';
var SHEET_MAINTENANCE = 'Maintenance';
var SHEET_LOG         = 'MaintenanceLog';

// ================================================================
// GET — serves driver forms + dashboard API
// ================================================================
function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (action === 'getVehicles') return getVehicles_();
  if (action === 'getLogs')     return getLogs_();

  // --- existing form routing (unchanged) ---
  try {
    var page = e.parameter.page;
    Logger.log("Requested page: " + page);
    if (page === 'incidents') {
      Logger.log("Serving Incidents.html");
      return HtmlService.createHtmlOutputFromFile('Incidents')
        .setTitle('IKEA Incident Report')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    Logger.log("Serving checkForm.html");
    return HtmlService.createHtmlOutputFromFile('checkForm')
      .setTitle('IKEA Fleet Control')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    Logger.log("Error in doGet: " + err.message);
    return HtmlService.createHtmlOutput(
      '<h2>Error</h2><p>Unable to load the requested page: ' + err.message + '</p>'
    ).setTitle('Error - IKEA Fleet Control')
     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ================================================================
// POST — dashboard write actions
// ================================================================
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    if (action === 'updateVehicle')  return updateVehicle_(payload);
    if (action === 'logMaintenance') return logMaintenance_(payload);
    if (action === 'logOilChange')   return logOilChange_(payload);
    if (action === 'logDiagnostic')  return logDiagnostic_(payload);
    return cors_(ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action: ' + action })));
  } catch(err) {
    return cors_(ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.toString() })));
  }
}

// ================================================================
// DASHBOARD API — read functions
// ================================================================
function getVehicles_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sh      = ss.getSheetByName(SHEET_VEHICLES);
  var data    = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
      }
      obj[h] = v;
    });
    return obj;
  });
  return cors_(ContentService.createTextOutput(JSON.stringify({
    ok: true,
    lastModified: new Date().toISOString(),
    data: rows
  })));
}

function getLogs_() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = getOrCreateLogSheet_(ss);
  var data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return cors_(ContentService.createTextOutput(JSON.stringify({ ok: true, data: [] })));
  }
  var headers = data[0];
  var rows = data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      var v = row[i];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      }
      obj[h] = v;
    });
    return obj;
  });
  return cors_(ContentService.createTextOutput(JSON.stringify({ ok: true, data: rows })));
}

// ================================================================
// DASHBOARD API — write functions
// ================================================================
function updateVehicle_(payload) {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sh      = ss.getSheetByName(SHEET_VEHICLES);
  var data    = sh.getDataRange().getValues();
  var headers = data[0];

  var rowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(payload.vehicleId)) {
      rowIdx = i + 1;
      break;
    }
  }
  if (rowIdx === -1) {
    return cors_(ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Vehicle ID not found: ' + payload.vehicleId })));
  }

  var changed = [];
  Object.keys(payload.fields).forEach(function(fieldName) {
    var colIdx = headers.indexOf(fieldName);
    if (colIdx === -1) return;
    sh.getRange(rowIdx, colIdx + 1).setValue(payload.fields[fieldName]);
    changed.push(fieldName + ' → ' + payload.fields[fieldName]);
  });

  appendLog_(ss, {
    timestamp: new Date(),
    vehicleId: payload.vehicleId,
    action:    'Edit: ' + changed.join(', '),
    user:      payload.user || 'Unknown',
    note:      payload.note || ''
  });

  return cors_(ContentService.createTextOutput(JSON.stringify({ ok: true, updated: changed })));
}

function logOilChange_(payload) {
  var km        = Number(payload.currentKm);
  var nextOilKm = km + 15000;
  var today     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  updateVehicle_({
    vehicleId: payload.vehicleId,
    fields: {
      'Mileage':         km,
      'Last Oil Change': km,
      'Next Oil Change': nextOilKm,
      'Last Check':      today
    },
    user: payload.user,
    note: 'Oil change logged'
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAINTENANCE);
  if (sh) {
    sh.appendRow([new Date(), payload.vehicleId, 'Oil Change', km, payload.user || 'Unknown', payload.note || '']);
  }

  return cors_(ContentService.createTextOutput(JSON.stringify({ ok: true })));
}

function logDiagnostic_(payload) {
  updateVehicle_({
    vehicleId: payload.vehicleId,
    fields: { 'Annual Diagnostic': payload.newExpiryDate },
    user: payload.user,
    note: 'Annual diagnostic renewed'
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAINTENANCE);
  if (sh) {
    sh.appendRow([new Date(), payload.vehicleId, 'Annual Diagnostic', '', payload.user || 'Unknown', 'New expiry: ' + payload.newExpiryDate]);
  }

  return cors_(ContentService.createTextOutput(JSON.stringify({ ok: true })));
}

function logMaintenance_(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreateLogSheet_(ss);
  sh.appendRow([
    new Date(),
    payload.vehicleId   || '',
    payload.type        || 'General',
    payload.km          || '',
    payload.scheduledDate || '',
    payload.mechanic    || '',
    payload.user        || 'Unknown',
    payload.note        || ''
  ]);
  return cors_(ContentService.createTextOutput(JSON.stringify({ ok: true })));
}

// ================================================================
// HELPERS
// ================================================================
function cors_(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      'Access-Control-Allow-Origin':  '*',
      'Access-Control-Allow-Methods': 'GET,POST',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}

function getOrCreateLogSheet_(ss) {
  var sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LOG);
    sh.appendRow(['Timestamp', 'Vehicle ID', 'Type', 'KM at service', 'Scheduled date', 'Mechanic', 'User', 'Notes']);
    sh.setFrozenRows(1);
    sh.getRange('1:1').setFontWeight('bold').setBackground('#0047ab').setFontColor('#ffffff');
  }
  return sh;
}

function appendLog_(ss, entry) {
  var sh = getOrCreateLogSheet_(ss);
  sh.appendRow([entry.timestamp, entry.vehicleId, entry.action, '', '', '', entry.user, entry.note]);
}

// ================================================================
// EXISTING DRIVER FORM FUNCTIONS (unchanged)
// ================================================================
function uploadDailyCheck(formData) {
  try {
    Logger.log("uploadDailyCheck called with formData: " + JSON.stringify(formData));
    const rootFolder = DriveApp.getFolderById('19MizOwvGRKzuG9Ke2im7z7s7lsvRt1Eq');
    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const photoUrls = [];
    for (let i = 1; i <= 5; i++) {
      const photoData = formData[`photo${i}`];
      if (photoData && photoData.data) {
        const fileName = `${dateStr}_${formData.ikeaId}_${formData.matricule}_photo${i}.jpg`;
        const file = Utilities.newBlob(Utilities.base64Decode(photoData.data), photoData.mimeType, fileName);
        const savedFile = rootFolder.createFile(file);
        savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrls.push(savedFile.getUrl());
      } else {
        photoUrls.push('');
      }
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Daily Checks');
    if (!sheet) {
      sheet = ss.insertSheet('Daily Checks');
      sheet.appendRow(['Timestamp', 'IKEA ID', 'Vehicle ID', 'Kilometrage', 'Tires', 'Lights', 'Windows', 'Body', 'Notes', 'Latitude', 'Longitude', 'Accuracy', 'Photo 1 URL', 'Photo 2 URL', 'Photo 3 URL', 'Photo 4 URL', 'Photo 5 URL']);
    }
    sheet.appendRow([timestamp, formData.ikeaId, formData.matricule, formData.kilometrage, formData.tires, formData.lights, formData.windows, formData.body, formData.notes, formData.latitude, formData.longitude, formData.accuracy, ...photoUrls]);
    Logger.log("Data saved successfully for IKEA ID: " + formData.ikeaId);
    return { status: 'success' };
  } catch (e) {
    Logger.log("Error in uploadDailyCheck: " + e.message);
    throw new Error("Failed to submit daily check: " + e.message);
  }
}

function ensureIncidentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Incidents');
  if (!sheet) {
    sheet = ss.insertSheet('Incidents');
    sheet.appendRow(['Timestamp', 'IKEA ID', 'Vehicle ID', 'Incident Location', 'Incident Description', 'Physical Casualties', 'Photo 1 URL', 'Photo 2 URL', 'Photo 3 URL']);
  }
}

function uploadIncidentReport(formData) {
  try {
    Logger.log("uploadIncidentReport called with formData: " + JSON.stringify(formData));
    ensureIncidentsSheet();
    const rootFolder = DriveApp.getFolderById('19MizOwvGRKzuG9Ke2im7z7s7lsvRt1Eq');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Incidents');
    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd-MM-yyyy');
    const photoUrls = [];
    for (let i = 1; i <= 3; i++) {
      const photoData = formData[`photo${i}`];
      if (photoData && photoData.data) {
        const fileName = `${dateStr}_${formData.ikeaId}_${formData.matricule}_incident_photo${i}.jpg`;
        const file = Utilities.newBlob(Utilities.base64Decode(photoData.data), photoData.mimeType, fileName);
        const savedFile = rootFolder.createFile(file);
        savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrls.push(savedFile.getUrl());
      } else {
        photoUrls.push('');
      }
    }
    sheet.appendRow([timestamp, formData.ikeaId, formData.matricule, formData.incidentLocation, formData.incidentDescription, formData.casualties, ...photoUrls]);
    Logger.log("Incident report saved successfully for IKEA ID: " + formData.ikeaId);
    return { status: 'success' };
  } catch (e) {
    Logger.log("Error in uploadIncidentReport: " + e.message);
    throw new Error("Failed to submit incident report: " + e.message);
  }
}
