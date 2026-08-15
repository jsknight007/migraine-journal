/**
 * Migraine Journal - Backend
 * Google Apps Script web app for logging and retrieving migraine entries.
 */

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

function getSheet() {
  if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID not set in Script Properties");
  return SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
}

/**
 * Sheets auto-parses "yyyy-MM-dd" strings into Date objects, so a date cell
 * reads back as a Date rather than the string the app expects. Normalize to
 * a plain "yyyy-MM-dd" string regardless of how the cell is stored.
 */
function toDateStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(v);
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || "list";

    switch (action) {
      case "log": return handleLog(params);
      case "delete": return handleDelete(params);
      case "list": return handleList(params);
      default: return jsonResponse({ error: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || "log";

    switch (action) {
      case "log": return handleLog(body);
      case "delete": return handleDelete(body);
      default: return jsonResponse({ error: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function handleLog(params) {
  const date = params.date;
  const severity = parseInt(params.severity);
  const symptoms = params.symptoms || "";
  const triggers = params.triggers || "";
  const duration = params.duration || "";
  const medications = params.medications || "";
  const notes = params.notes || "";

  if (!date) return jsonResponse({ error: "Date is required" });
  if (isNaN(severity) || severity < 1 || severity > 5) {
    return jsonResponse({ error: "Severity must be 1-5" });
  }

  const sheet = getSheet();
  sheet.appendRow([
    date,
    severity,
    symptoms,
    triggers,
    duration,
    medications,
    notes,
    new Date()
  ]);

  return jsonResponse({ status: "success", date: date, severity: severity });
}

function handleDelete(params) {
  const date = params.date;
  if (!date) return jsonResponse({ error: "Date is required for delete" });

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    if (toDateStr(data[i][0]) === date) {
      sheet.deleteRow(i + 1);
      return jsonResponse({ status: "success", deleted: date });
    }
  }

  return jsonResponse({ error: "Entry not found for date: " + date });
}

function handleList(params) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const entries = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0] === "") continue;
    entries.push({
      date: toDateStr(row[0]),
      severity: row[1],
      symptoms: row[2],
      triggers: row[3],
      duration: row[4],
      medications: row[5],
      notes: row[6],
      logged: row[7]
    });
  }

  var months = parseInt(params.months) || 6;
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);
  var cutoffStr = Utilities.formatDate(cutoff, Session.getScriptTimeZone(), "yyyy-MM-dd");

  var filtered = entries.filter(function(e) { return e.date >= cutoffStr; });
  filtered.sort(function(a, b) { return a.date > b.date ? -1 : 1; });

  return jsonResponse({
    entries: filtered,
    total: entries.length,
    showing: filtered.length
  });
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
