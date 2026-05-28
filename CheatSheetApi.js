// ============================================================
// cheatsheet.gs
// Rhino-safe: var only, no arrow functions, no template literals
// ============================================================

var CHEATSHEET_TAB = "CheatSheet";

function _csSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(CHEATSHEET_TAB);
  if (!sh) throw new Error("Sheet \"" + CHEATSHEET_TAB + "\" not found.");
  return sh;
}

function _csRowToObject_(row) {
  return {
    id:      String(row[0] || "").trim(),
    issue:   String(row[1] || "").trim(),
    cat:     String(row[2] || "").trim().toLowerCase(),
    action:  String(row[3] || "").trim(),
    channel: String(row[4] || "").trim(),
    contact: String(row[5] || "").trim(),
    link:    String(row[6] || "").trim()
  };
}

function _csRequireAccess_() {
  var email = getCurrentUserEmail_();
  if (!canAccessPage("cheatsheet", email)) {
    throw new Error("You do not have permission to access this page.");
  }
}

function cheatsheet_list() {
  try {
    _csRequireAccess_();

    var sh      = _csSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };

    var values = sh.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    var out    = [];

    for (var i = 0; i < values.length; i++) {
      var id = String(values[i][0] || "").trim();
      if (!id) continue;
      out.push(_csRowToObject_(values[i]));
    }

    return { ok: true, data: out };
  } catch (e) {
    console.error("cheatsheet_list error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

function cheatsheet_add(payload) {
  try {
    _csRequireAccess_();

    var p = payload || {};

    var issue  = String(p.issue  || "").trim();
    var cat    = String(p.cat    || "").trim().toLowerCase();
    var action = String(p.action || "").trim();

    if (!issue)  throw new Error("Issue is required.");
    if (!cat)    throw new Error("Category is required.");
    if (!action) throw new Error("Action is required.");

    var sh      = _csSheet_();
    var lastRow = sh.getLastRow();

    // Generate next ID
    var newId = 1;
    if (lastRow >= 2) {
      var ids = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (var i = 0; i < ids.length; i++) {
        var n = parseInt(ids[i][0], 10);
        if (!isNaN(n) && n >= newId) newId = n + 1;
      }
    }

    sh.appendRow([
      newId,
      issue,
      cat,
      action,
      String(p.channel || "").trim(),
      String(p.contact || "").trim(),
      String(p.link    || "").trim()
    ]);

    return { ok: true, data: { id: String(newId) } };
  } catch (e) {
    console.error("cheatsheet_add error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

function cheatsheet_update(payload) {
  try {
    _csRequireAccess_();

    var p  = payload || {};
    var id = String(p.id || "").trim();
    if (!id) throw new Error("Missing ID.");

    var issue  = String(p.issue  || "").trim();
    var cat    = String(p.cat    || "").trim().toLowerCase();
    var action = String(p.action || "").trim();

    if (!issue)  throw new Error("Issue is required.");
    if (!cat)    throw new Error("Category is required.");
    if (!action) throw new Error("Action is required.");

    var sh      = _csSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error("No data in CheatSheet.");

    var ids = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    var rowIndex = -1;

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === id) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex < 0) throw new Error("Entry ID \"" + id + "\" not found.");

    sh.getRange(rowIndex, 1, 1, 7).setValues([[
      id,
      issue,
      cat,
      action,
      String(p.channel || "").trim(),
      String(p.contact || "").trim(),
      String(p.link    || "").trim()
    ]]);

    SpreadsheetApp.flush();
    return { ok: true, data: true };
  } catch (e) {
    console.error("cheatsheet_update error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

function cheatsheet_delete(id) {
  try {
    _csRequireAccess_();

    var target = String(id || "").trim();
    if (!target) throw new Error("Missing ID.");

    var sh      = _csSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) throw new Error("No data in CheatSheet.");

    var ids = sh.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    var rowIndex = -1;

    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === target) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex < 0) throw new Error("Entry ID \"" + target + "\" not found.");

    sh.deleteRow(rowIndex);
    return { ok: true, data: true };
  } catch (e) {
    console.error("cheatsheet_delete error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}