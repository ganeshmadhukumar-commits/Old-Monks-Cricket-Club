
/**
 * Wicketkeeper — Google Sheets backend
 * ------------------------------------
 * Paste this entire file into the Apps Script editor bound to your
 * Google Sheet (Extensions > Apps Script), then deploy it as a Web App.
 * See the setup guide for step-by-step instructions.
 *
 * It stores three tabs in the spreadsheet — Players, Matches, Ledger —
 * and fully rewrites them on every save. That keeps the logic simple and
 * safe for a club-sized dataset (a few dozen players, a season of matches).
 */

var SHEET_CONFIG = {
  Players: ["id", "name", "contact", "balance"],
  Matches: ["id", "date", "day", "opponent", "fee", "present"],
  Ledger: ["id", "matchId", "playerId", "playerName", "type", "amount", "date", "description"]
};

function doGet(e) {
  var action = e.parameter.action;
  if (action === "getData") {
    return jsonResponse(getAllData());
  }
  return jsonResponse({ error: "Unknown GET action: " + action });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === "saveData") {
      saveAllData(body.data || {});
      return jsonResponse({ success: true });
    }
    return jsonResponse({ error: "Unknown POST action: " + body.action });
  } catch (err) {
    return jsonResponse({ error: String(err.message || err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(SHEET_CONFIG[name]);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  return values.slice(1)
    .filter(function (row) { return row.join("") !== ""; })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      return obj;
    });
}

function writeSheet(name, rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  sheet.clearContents();
  var headers = SHEET_CONFIG[name];
  sheet.appendRow(headers);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function getAllData() {
  var players = sheetToObjects(getOrCreateSheet("Players")).map(function (p) {
    return {
      id: String(p.id),
      name: String(p.name || ""),
      contact: String(p.contact || ""),
      balance: Number(p.balance) || 0
    };
  });

  var matches = sheetToObjects(getOrCreateSheet("Matches")).map(function (m) {
    var present = [];
    try { present = m.present ? JSON.parse(m.present) : []; } catch (e) { present = []; }
    return {
      id: String(m.id),
      date: formatDateValue(m.date),
      day: String(m.day || ""),
      opponent: String(m.opponent || ""),
      fee: Number(m.fee) || 0,
      present: present
    };
  });

  var ledger = sheetToObjects(getOrCreateSheet("Ledger")).map(function (l) {
    return {
      id: String(l.id),
      matchId: String(l.matchId || ""),
      playerId: String(l.playerId || ""),
      playerName: String(l.playerName || ""),
      type: String(l.type || ""),
      amount: Number(l.amount) || 0,
      date: formatDateValue(l.date),
      description: String(l.description || "")
    };
  });

  return { players: players, matches: matches, ledger: ledger };
}

function saveAllData(data) {
  var players = (data.players || []).map(function (p) {
    return [p.id, p.name, p.contact, p.balance];
  });
  var matches = (data.matches || []).map(function (m) {
    return [m.id, m.date, m.day, m.opponent, m.fee, JSON.stringify(m.present || [])];
  });
  var ledger = (data.ledger || []).map(function (l) {
    return [l.id, l.matchId || "", l.playerId, l.playerName, l.type, l.amount, l.date, l.description];
  });

  writeSheet("Players", players);
  writeSheet("Matches", matches);
  writeSheet("Ledger", ledger);
}

// Sheets sometimes hand back Date objects for date-like strings (e.g. "2026-08-15").
// Convert those back to a plain YYYY-MM-DD string so the client's date logic keeps working.
function formatDateValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value || "");
}
