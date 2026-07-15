/**
 * Our Walk Tracker - Backend
 * All operations via doGet for proper CORS handling.
 */

const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const WEEK_GOAL = 150;
const DAILY_GOAL_MULTIPLIER = 20;
const MAX_DURATION = 600;

function getSheet() {
  if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID not set in Script Properties");
  return SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
}

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || "stats";

    switch (action) {
      case "log": return handleLog(params);
      case "deleteLast": return handleDeleteLast();
      default: return handleStats();
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function handleDeleteLast() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return jsonResponse({ error: "No entries to delete" });
  }
  sheet.deleteRow(lastRow);
  return jsonResponse({ status: "success" });
}

function handleLog(params) {
  const duration = parseInt(params.duration);
  if (isNaN(duration) || duration <= 0) {
    return jsonResponse({ error: "Duration must be a positive number" });
  }
  if (duration > MAX_DURATION) {
    return jsonResponse({ error: "Duration exceeds maximum (" + MAX_DURATION + " minutes)" });
  }

  const dateStr = params.date || new Date().toLocaleDateString();
  const sheet = getSheet();

  sheet.appendRow([
    new Date(),
    "Couple",
    duration,
    dateStr
  ]);

  return jsonResponse({ status: "success", logged: duration });
}

function handleStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
  const monthGoal = daysInMonth * DAILY_GOAL_MULTIPLIER;

  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const startOfWeek = new Date(now);
  startOfWeek.setDate(now.getDate() - now.getDay());
  startOfWeek.setHours(0, 0, 0, 0);

  var stats = {
    currentWeek: 0,
    currentMonth: 0,
    monthGoal: monthGoal,
    weekGoal: WEEK_GOAL,
    recent: []
  };

  var walkRows = data.slice(1).filter(function(row) {
    return row[0] !== "" && row[2] !== "";
  });

  walkRows.forEach(function(row) {
    var rowDate = row[0];
    if (!(rowDate instanceof Date)) {
      rowDate = new Date(rowDate);
    }
    if (!isNaN(rowDate.getTime())) {
      var mins = Number(row[2]) || 0;
      if (rowDate >= startOfWeek) stats.currentWeek += mins;
      if (rowDate >= startOfMonth) stats.currentMonth += mins;
    }
  });

  var lastFive = walkRows.slice(-5).reverse();
  lastFive.forEach(function(row) {
    var d = row[0];
    if (!(d instanceof Date)) d = new Date(d);
    if (!isNaN(d.getTime())) {
      stats.recent.push({
        date: Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d"),
        mins: row[2]
      });
    }
  });

  return jsonResponse(stats);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
