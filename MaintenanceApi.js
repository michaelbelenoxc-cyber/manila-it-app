// ============================================================
// maintenance.gs
// Backend for the Maintenance admin page.
// Fully self-contained — no dependencies on masterlist_report.gs.
//
// Public functions (called via google.script.run):
//   getSystemHealth()
//   getMaintenanceTriggers()
//   clearScriptCache()
//   forceDataRefresh()
//   resetSessionStorage()
//   exportMasterlistSnapshot()
//   findDuplicateSerials()
//   findMissingRequiredFields()
//   findInvalidWarrantyDates()
//   findOrphanedAssignments()
//   triggerWarrantyEmail()
//   createWeeklyBackupTrigger()
//   createDailyIntegrityScanTrigger()
//   deleteTrigger(payload)
// ============================================================


/* ============================================================
 * CONFIG
 * ============================================================ */

var WARRANTY_RECIPIENT_ = "manila-it@fbgphilippines.com";
var MAX_SCAN_ROWS_      = 10000;


/* ============================================================
 * LOCAL HELPERS  (self-contained, _maint_ prefixed)
 * ============================================================ */

function _maint_openSS_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function _maint_getMasterlistSheet_() {
  var sh = _maint_openSS_().getSheetByName(MASTERLIST_SHEET);
  if (!sh) throw new Error('Sheet "' + MASTERLIST_SHEET + '" not found.');
  return sh;
}

function _maint_isMeaningful_(v) {
  var s = String(v || "").trim().toLowerCase();
  return !!s && ["n/a", "na", "\u2014", "-", "none", "null", "(blank)"].indexOf(s) === -1;
}

function _maint_isRealRow_(r) {
  for (var i = 0; i < r.length; i++) {
    if (_maint_isMeaningful_(r[i])) return true;
  }
  return false;
}

function _maint_safeStr_(row, idx) {
  if (idx < 0) return "";
  return String(row[idx] != null ? row[idx] : "").trim();
}

function _maint_parseDate_(raw) {
  var s = String(raw || "").trim();
  if (!s) return null;

  var mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (mdy) {
    var mm = Number(mdy[1]);
    var dd = Number(mdy[2]);
    var yy = Number(mdy[3]);
    if (yy < 100) yy += 2000;
    var d1 = new Date(yy, mm - 1, dd);
    return isNaN(d1.getTime()) ? null : d1;
  }

  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    var d2 = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isNaN(d2.getTime()) ? null : d2;
  }

  var d3 = new Date(s);
  return isNaN(d3.getTime()) ? null : d3;
}

function _maint_headerIndex_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i] || "").trim().toLowerCase()] = i;
  }

  var pick = function() {
    for (var n = 0; n < arguments.length; n++) {
      var k = String(arguments[n]).toLowerCase();
      if (k in map) return map[k];
    }
    return -1;
  };

  return {
    type:       pick("type"),
    status:     pick("status", "asset status", "state"),
    warranty:   pick("end of warranty", "warranty end", "end warranty",
                     "warranty expiry", "warranty expiration",
                     "warranty expiry date", "warranty", "warranty_end"),
    assignee:   pick("assignee", "assigned to", "user", "employee", "assigned"),
    department: pick("department", "dept"),
    serialTag:  pick("serial tag"),
    serialNo:   pick("serial number", "serial no"),
    model:      pick("model")
  };
}

function _maint_deleteTriggers_(fnName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function _maint_getTriggerLastRun_(trigger) {
  try {
    var key = "trigger_lastrun_" + trigger.getHandlerFunction();
    var val = PropertiesService.getScriptProperties().getProperty(key);
    return val || "Unknown";
  } catch (e) {
    return "Unknown";
  }
}

function _maint_recordTriggerRun_(handlerName) {
  try {
    var stamp = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "MMM d, yyyy HH:mm"
    );
    PropertiesService.getScriptProperties()
      .setProperty("trigger_lastrun_" + handlerName, stamp);
  } catch (e) { /* non-fatal */ }
}


/* ============================================================
 * RBAC GUARD
 * ============================================================ */

function requireMaintenanceAccess_() {
  var email = Session.getActiveUser().getEmail();

  if (typeof canAccessPage === "function") {
    if (!canAccessPage("maintenance", email)) {
      throw new Error("You do not have permission to access the Maintenance page.");
    }
  } else {
    var ss      = _maint_openSS_();
    var editors = ss.getEditors();
    var allowed = false;
    for (var i = 0; i < editors.length; i++) {
      if (editors[i].getEmail().toLowerCase() === email.toLowerCase()) {
        allowed = true;
        break;
      }
    }
    if (!allowed) throw new Error("Maintenance access denied.");
  }
}


/* ============================================================
 * SYSTEM HEALTH
 * ============================================================ */

function getSystemHealth() {
  requireMaintenanceAccess_();

  var results = [];
  results.push(_maint_checkMasterlistSheet_());
  results.push(_maint_checkUsersSheet_());
  results.push(_maint_checkRbacSheet_());
  results.push(_maint_checkScriptCache_());
  results.push(_maint_checkGmailService_());
  results.push(_maint_checkWarrantyTrigger_());
  results.push(_maint_checkScriptQuota_());
  results.push(_maint_checkSpreadsheetSize_());
  return results;
}

function _maint_checkMasterlistSheet_() {
  try {
    var sh   = _maint_getMasterlistSheet_();
    var lr   = sh.getLastRow();
    var lc   = sh.getLastColumn();
    var rows = Math.max(0, lr - 1);

    if (lr < 2) {
      return { name: "Masterlist Sheet", status: "warn",
               detail: "Sheet exists but has no data rows." };
    }
    return { name: "Masterlist Sheet", status: "ok",
             detail: rows.toLocaleString() + " data rows \u00B7 " + lc + " columns" };
  } catch (e) {
    return { name: "Masterlist Sheet", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkUsersSheet_() {
  try {
    var ss = _maint_openSS_();
    var sh = ss.getSheetByName(EMPLOYEES_TAB_NAME);
    if (!sh) {
      return { name: "Employees Sheet", status: "warn",
               detail: 'Sheet "' + EMPLOYEES_TAB_NAME + '" not found.'};
    }

    var values  = sh.getDataRange().getDisplayValues();
    var headers = (values[0] || []).map(function(h) { return String(h).trim().toLowerCase(); });
    var roleCol = headers.indexOf("role");
    var rows    = [];
    for (var i = 1; i < values.length; i++) {
      if (_maint_isRealRow_(values[i])) rows.push(values[i]);
    }

    var noRole = 0;
    if (roleCol >= 0) {
      for (var j = 0; j < rows.length; j++) {
        if (!String(rows[j][roleCol] || "").trim()) noRole++;
      }
    }

    if (noRole > 0) {
      return { name: "Users Sheet", status: "warn",
               detail: rows.length + " users \u00B7 " + noRole + " missing role assignment" };
    }
    return { name: "Users Sheet", status: "ok",
             detail: rows.length + " registered users" };
  } catch (e) {
    return { name: "Users Sheet", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkRbacSheet_() {
  try {
    var ss = _maint_openSS_();
    var sh = ss.getSheetByName(RBAC_SHEET);
    if (!sh) {
      return { name: "RBAC Rules Sheet", status: "warn",
               detail: 'Sheet "' + RBAC_SHEET + '" not found.' };
    }
    var rows = Math.max(0, sh.getLastRow() - 1);
    return { name: "RBAC Rules Sheet", status: "ok",
             detail: rows + " permission rules" };
  } catch (e) {
    return { name: "RBAC Rules Sheet", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkScriptCache_() {
  try {
    var cache = CacheService.getScriptCache();
    cache.put("__health_probe__", "1", 10);
    var v = cache.get("__health_probe__");
    cache.remove("__health_probe__");
    if (v !== "1") throw new Error("Cache read-back mismatch");
    return { name: "Script Cache Service", status: "ok",
             detail: "Read/write probe passed" };
  } catch (e) {
    return { name: "Script Cache Service", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkGmailService_() {
  try {
    var remaining = MailApp.getRemainingDailyQuota();
    if (remaining < 10) {
      return { name: "Gmail Service", status: "warn",
               detail: "Only " + remaining + " emails remaining today" };
    }
    return { name: "Gmail Service", status: "ok",
             detail: remaining + " emails remaining today" };
  } catch (e) {
    return { name: "Gmail Service", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkWarrantyTrigger_() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var found    = false;
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "sendMonthlyWarrantyReminder") {
        found = true;
        break;
      }
    }
    if (!found) {
      return { name: "Warranty Trigger", status: "warn",
               detail: "Monthly warranty reminder trigger is NOT installed" };
    }
    return { name: "Warranty Trigger", status: "ok",
             detail: "Active \u00B7 fires 1st of every month at 08:00" };
  } catch (e) {
    return { name: "Warranty Trigger", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkScriptQuota_() {
  try {
    var count  = ScriptApp.getProjectTriggers().length;
    var pct    = Math.round((count / 20) * 100);
    var status = count >= 18 ? "warn" : "ok";
    return {
      name:   "Script Trigger Quota",
      status: status,
      detail: count + " / 20 triggers used (" + pct + "%)"
    };
  } catch (e) {
    return { name: "Script Trigger Quota", status: "error",
             detail: e.message || String(e) };
  }
}

function _maint_checkSpreadsheetSize_() {
  try {
    var ss         = _maint_openSS_();
    var sheets     = ss.getSheets();
    var totalCells = 0;

    for (var i = 0; i < sheets.length; i++) {
      totalCells += sheets[i].getMaxRows() * sheets[i].getMaxColumns();
    }

    var LIMIT  = 10000000;
    var pct    = ((totalCells / LIMIT) * 100).toFixed(1);
    var status = totalCells > 9000000 ? "error"
               : totalCells > 8000000 ? "warn"
               : "ok";

    return {
      name:   "Spreadsheet Size",
      status: status,
      detail: "~" + totalCells.toLocaleString() + " cells allocated \u00B7 " +
              pct + "% of 10M limit \u00B7 " + sheets.length + " sheet(s)"
    };
  } catch (e) {
    return { name: "Spreadsheet Size", status: "error",
             detail: e.message || String(e) };
  }
}


/* ============================================================
 * TRIGGER MANAGEMENT
 * ============================================================ */

function getMaintenanceTriggers() {
  requireMaintenanceAccess_();

  var KNOWN = [
    {
      name:            "Monthly Warranty Reminder",
      handlerFunction: "sendMonthlyWarrantyReminder",
      schedule:        "1st of every month \u00B7 08:00 AM"
    },
    {
      name:            "Weekly Data Backup",
      handlerFunction: "weeklyMasterlistBackup",
      schedule:        "Every Monday \u00B7 09:00 AM"
    },
    {
      name:            "Daily Integrity Scan",
      handlerFunction: "dailyIntegrityScan",
      schedule:        "Every day \u00B7 12:00 AM"
    }
  ];

  var installed = ScriptApp.getProjectTriggers();
  var result    = [];

  for (var i = 0; i < KNOWN.length; i++) {
    var def   = KNOWN[i];
    var match = null;

    for (var j = 0; j < installed.length; j++) {
      if (installed[j].getHandlerFunction() === def.handlerFunction) {
        match = installed[j];
        break;
      }
    }

    result.push({
      name:            def.name,
      handlerFunction: def.handlerFunction,
      schedule:        def.schedule,
      active:          !!match,
      lastRun:         match ? _maint_getTriggerLastRun_(match) : "Never"
    });
  }

  return result;
}

function createWeeklyBackupTrigger() {
  requireMaintenanceAccess_();
  _maint_deleteTriggers_("weeklyMasterlistBackup");

  ScriptApp.newTrigger("weeklyMasterlistBackup")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();

  Logger.log("[MAINTENANCE] Weekly backup trigger installed.");
  return { ok: true, message: "Weekly backup trigger installed (Mondays 09:00)" };
}

function createDailyIntegrityScanTrigger() {
  requireMaintenanceAccess_();
  _maint_deleteTriggers_("dailyIntegrityScan");

  ScriptApp.newTrigger("dailyIntegrityScan")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  Logger.log("[MAINTENANCE] Daily integrity scan trigger installed.");
  return { ok: true, message: "Daily integrity scan trigger installed (00:00)" };
}

function deleteMaintenanceTrigger(payload) {
  requireMaintenanceAccess_();

  var HANDLER_MAP = {
    "Monthly Warranty Reminder": "sendMonthlyWarrantyReminder",
    "Weekly Data Backup":        "weeklyMasterlistBackup",
    "Daily Integrity Scan":      "dailyIntegrityScan"
  };

  var name    = String((payload || {}).name || "").trim();
  var handler = HANDLER_MAP[name];

  if (!handler) {
    throw new Error('Unknown trigger name: "' + name + '"');
  }

  _maint_deleteTriggers_(handler);
  Logger.log("[MAINTENANCE] Trigger deleted: " + name);
  return { ok: true, message: 'Trigger "' + name + '" removed' };
}


/* ============================================================
 * CACHE MANAGEMENT
 * ============================================================ */

function clearScriptCache() {
  requireMaintenanceAccess_();

  try {
    var SHELL_PAGES = [
      "home", "masterlist", "employees", "events", "reports", "schedule",
      "aboutus", "admin", "users", "profile", "settings", "inventory",
      "request", "viewrequest", "myrequest", "companyengagement", "tasks",
      "assetreport", "vendor", "signedaf", "rbac", "kpi", "publish", "maintenance"
    ];

    var KNOWN_KEYS = [];
    for (var i = 0; i < SHELL_PAGES.length; i++) {
      KNOWN_KEYS.push("shell_v4_" + SHELL_PAGES[i]);
    }
    KNOWN_KEYS.push("rbac_rules", "__health_probe__");

    var sc = CacheService.getScriptCache();
    sc.removeAll(KNOWN_KEYS);

    var uc = CacheService.getUserCache();
    uc.removeAll(KNOWN_KEYS);

    Logger.log("[MAINTENANCE] Script cache cleared.");
    return { ok: true, message: "Cache cleared \u2014 all pages will reload fresh data on their next request" };
  } catch (e) {
    Logger.log("[MAINTENANCE] clearScriptCache error: " + e.message);
    throw e;
  }
}

function forceDataRefresh() {
  requireMaintenanceAccess_();

  try {
    clearScriptCache();

    var sh   = _maint_getMasterlistSheet_();
    var rows = Math.max(0, sh.getLastRow() - 1);

    Logger.log("[MAINTENANCE] forceDataRefresh \u2014 " + rows + " data rows found.");
    return {
      ok:      true,
      message: "Masterlist refreshed \u2014 " + rows.toLocaleString() + " data rows processed"
    };
  } catch (e) {
    Logger.log("[MAINTENANCE] forceDataRefresh error: " + e.message);
    throw e;
  }
}

function resetSessionStorage() {
  requireMaintenanceAccess_();

  try {
    clearScriptCache();

    var props       = PropertiesService.getScriptProperties().getProperties();
    var keys        = Object.keys(props);
    var sessionKeys = [];
    for (var i = 0; i < keys.length; i++) {
      if (keys[i].indexOf("session_") === 0) sessionKeys.push(keys[i]);
    }

    if (sessionKeys.length) {
      PropertiesService.getScriptProperties().deleteAllProperties();
      for (var j = 0; j < keys.length; j++) {
        if (keys[j].indexOf("session_") !== 0) {
          PropertiesService.getScriptProperties().setProperty(keys[j], props[keys[j]]);
        }
      }
    }

    Logger.log("[MAINTENANCE] Session storage reset.");
    return { ok: true, message: "Session tokens invalidated \u2014 users will receive fresh permissions on next load" };
  } catch (e) {
    Logger.log("[MAINTENANCE] resetSessionStorage error: " + e.message);
    throw e;
  }
}


/* ============================================================
 * SNAPSHOT / EXPORT
 * ============================================================ */

function exportMasterlistSnapshot() {
  requireMaintenanceAccess_();

  try {
    var ss    = _maint_openSS_();
    var src   = _maint_getMasterlistSheet_();
    var stamp = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm"
    );
    var name = "Masterlist_Snapshot_" + stamp;

    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);

    var dest = src.copyTo(ss);
    dest.setName(name);
    ss.setActiveSheet(dest);
    ss.moveActiveSheet(ss.getNumSheets());
    dest.setFrozenRows(1);

    var url = ss.getUrl() + "#gid=" + dest.getSheetId();
    Logger.log("[MAINTENANCE] Snapshot created: " + name);
    return { ok: true, message: "Snapshot created: " + name, url: url, sheetName: name };
  } catch (e) {
    Logger.log("[MAINTENANCE] exportMasterlistSnapshot error: " + e.message);
    throw e;
  }
}


/* ============================================================
 * DATA INTEGRITY SCANNERS
 * ============================================================ */

function findDuplicateSerials() {
  requireMaintenanceAccess_();

  try {
    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    var dataRows = values.slice(1);
    var seenTag  = {};
    var seenNo   = {};
    var issues   = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;
      var rowNum = i + 2;

      if (idx.serialTag >= 0) {
        var tag = String(row[idx.serialTag] || "").trim().toLowerCase();
        if (tag && tag !== "n/a" && tag !== "-") {
          if (seenTag[tag] !== undefined) {
            issues.push("Row " + rowNum + " \u00B7 " + String(row[idx.serialTag]).trim() +
                        " (Serial Tag \u2014 duplicate of row " + seenTag[tag] + ")");
          } else {
            seenTag[tag] = rowNum;
          }
        }
      }

      if (idx.serialNo >= 0) {
        var sn = String(row[idx.serialNo] || "").trim().toLowerCase();
        if (sn && sn !== "n/a" && sn !== "-") {
          if (seenNo[sn] !== undefined) {
            issues.push("Row " + rowNum + " \u00B7 " + String(row[idx.serialNo]).trim() +
                        " (Serial No. \u2014 duplicate of row " + seenNo[sn] + ")");
          } else {
            seenNo[sn] = rowNum;
          }
        }
      }
    }

    Logger.log("[MAINTENANCE] findDuplicateSerials \u2014 " + issues.length + " issue(s)");
    return { ok: true, count: issues.length, rows: issues };
  } catch (e) {
    Logger.log("[MAINTENANCE] findDuplicateSerials error: " + e.message);
    throw e;
  }
}

function findMissingRequiredFields() {
  requireMaintenanceAccess_();

  try {
    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    var dataRows = values.slice(1);
    var issues   = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;

      var rowNum  = i + 2;
      var missing = [];

      if (idx.type     >= 0 && !_maint_isMeaningful_(row[idx.type]))     missing.push("Type");
      if (idx.status   >= 0 && !_maint_isMeaningful_(row[idx.status]))   missing.push("Status");
      if (idx.assignee >= 0 && !_maint_isMeaningful_(row[idx.assignee])) missing.push("Assignee");

      if (missing.length) {
        var tag   = idx.serialTag >= 0 ? String(row[idx.serialTag] || "").trim() : "";
        var label = tag ? " \u00B7 " + tag : "";
        issues.push("Row " + rowNum + label + " \u00B7 Missing: " + missing.join(", "));
      }
    }

    Logger.log("[MAINTENANCE] findMissingRequiredFields \u2014 " + issues.length + " row(s)");
    return { ok: true, count: issues.length, rows: issues };
  } catch (e) {
    Logger.log("[MAINTENANCE] findMissingRequiredFields error: " + e.message);
    throw e;
  }
}

function findInvalidWarrantyDates() {
  requireMaintenanceAccess_();

  try {
    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    var dataRows = values.slice(1);

    if (idx.warranty < 0) {
      return { ok: true, count: 0, rows: [], note: "No warranty column found in headers." };
    }

    var SKIP   = ["n/a", "na", "\u2014", "-", "none", "null", "(blank)", "tbd"];
    var issues = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;

      var raw    = String(row[idx.warranty] || "").trim();
      var rowNum = i + 2;

      if (!raw || SKIP.indexOf(raw.toLowerCase()) !== -1) continue;

      if (!_maint_parseDate_(raw)) {
        var tag   = idx.serialTag >= 0 ? String(row[idx.serialTag] || "").trim() : "";
        var label = tag ? " \u00B7 " + tag : "";
        issues.push("Row " + rowNum + label + " \u00B7 \"" + raw + "\" is not a recognised date format");
      }
    }

    Logger.log("[MAINTENANCE] findInvalidWarrantyDates \u2014 " + issues.length + " row(s)");
    return { ok: true, count: issues.length, rows: issues };
  } catch (e) {
    Logger.log("[MAINTENANCE] findInvalidWarrantyDates error: " + e.message);
    throw e;
  }
}

function findOrphanedAssignments() {
  requireMaintenanceAccess_();

  try {
    var ss         = _maint_openSS_();
    var usersSh = ss.getSheetByName(EMPLOYEES_TAB_NAME);
    var knownUsers = {};

    if (usersSh) {
      var uValues  = usersSh.getDataRange().getDisplayValues();
      var uHeaders = (uValues[0] || []).map(function(h) { return String(h).trim().toLowerCase(); });
      var nameCol  = uHeaders.indexOf("display name");
      var emailCol = uHeaders.indexOf("email");

      for (var u = 1; u < uValues.length; u++) {
        if (nameCol  >= 0 && uValues[u][nameCol])  knownUsers[String(uValues[u][nameCol]).trim().toLowerCase()]  = true;
        if (emailCol >= 0 && uValues[u][emailCol]) knownUsers[String(uValues[u][emailCol]).trim().toLowerCase()] = true;
      }
    }

    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    var dataRows = values.slice(1);
    var issues   = [];

    if (idx.assignee < 0) {
      return { ok: true, count: 0, rows: [], note: "No assignee column found in headers." };
    }

    var hasKnown = Object.keys(knownUsers).length > 0;

    for (var i = 0; i < dataRows.length; i++) {
      var row      = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;

      var assignee = String(row[idx.assignee] || "").trim();
      if (!_maint_isMeaningful_(assignee)) continue;
      if (assignee.toLowerCase() === "unassigned") continue;

      if (hasKnown && !knownUsers[assignee.toLowerCase()]) {
        var rowNum = i + 2;
        var tag    = idx.serialTag >= 0 ? String(row[idx.serialTag] || "").trim() : "";
        var label  = tag ? " \u00B7 " + tag : "";
        issues.push("Row " + rowNum + label + " \u00B7 Assigned to \"" + assignee + "\" \u2014 not found in Employees sheet");
      }
    }

    Logger.log("[MAINTENANCE] findOrphanedAssignments \u2014 " + issues.length + " orphaned assignment(s)");
    return { ok: true, count: issues.length, rows: issues };
  } catch (e) {
    Logger.log("[MAINTENANCE] findOrphanedAssignments error: " + e.message);
    throw e;
  }
}


/* ============================================================
 * MANUAL TRIGGER — WARRANTY EMAIL
 * ============================================================ */

function triggerWarrantyEmail() {
  requireMaintenanceAccess_();

  try {
    sendMonthlyWarrantyReminder();
    _maint_recordTriggerRun_("sendMonthlyWarrantyReminder");
    return { ok: true, message: "Warranty reminder email sent to " + WARRANTY_RECIPIENT_ };
  } catch (e) {
    Logger.log("[MAINTENANCE] triggerWarrantyEmail error: " + e.message);
    throw e;
  }
}


/* ============================================================
 * SCHEDULED TRIGGER HANDLERS
 * ============================================================ */

function weeklyMasterlistBackup_manual() {
  requireMaintenanceAccess_();
  weeklyMasterlistBackup();
  return { ok: true, message: "Weekly backup executed manually." };
}

function weeklyMasterlistBackup() {
  try {
    var ss    = _maint_openSS_();
    var src   = _maint_getMasterlistSheet_();
    var stamp = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"
    );
    var name = "Weekly_Backup_" + stamp;

    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);

    var dest = src.copyTo(ss);
    dest.setName(name);
    ss.setActiveSheet(dest);
    ss.moveActiveSheet(ss.getNumSheets());
    dest.setFrozenRows(1);

    _maint_recordTriggerRun_("weeklyMasterlistBackup");
    Logger.log("[MAINTENANCE] Weekly backup created: " + name);
  } catch (e) {
    Logger.log("[MAINTENANCE] weeklyMasterlistBackup error: " + e.message);
  }
}

function dailyIntegrityScan_manual() {
  requireMaintenanceAccess_();
  dailyIntegrityScan();
  return { ok: true, message: "Daily integrity scan executed manually." };
}

function dailyIntegrityScan() {
  try {
    var dupes   = _maint_duplicates_noGuard_();
    var missing = _maint_missingFields_noGuard_();
    var dates   = _maint_invalidDates_noGuard_();
    var total   = dupes.count + missing.count + dates.count;

    Logger.log("[MAINTENANCE] Daily integrity scan \u2014 " +
               dupes.count + " duplicate(s), " +
               missing.count + " missing field(s), " +
               dates.count + " bad date(s)");

    if (total > 0) {
      _maint_sendScanSummary_(dupes, missing, dates);
    }

    _maint_recordTriggerRun_("dailyIntegrityScan");
  } catch (e) {
    Logger.log("[MAINTENANCE] dailyIntegrityScan error: " + e.message);
  }
}


/* ============================================================
 * NO-GUARD SCANNER VARIANTS  (for trigger-driven calls)
 * ============================================================ */

function _maint_duplicates_noGuard_() {
  try {
    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    var dataRows = values.slice(1);
    var seenTag  = {};
    var seenNo   = {};
    var issues   = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;
      var rowNum = i + 2;

      if (idx.serialTag >= 0) {
        var tag = String(row[idx.serialTag] || "").trim().toLowerCase();
        if (tag && tag !== "n/a" && tag !== "-") {
          if (seenTag[tag] !== undefined) {
            issues.push("Row " + rowNum + " \u00B7 " + String(row[idx.serialTag]).trim() +
                        " (Serial Tag duplicate of row " + seenTag[tag] + ")");
          } else {
            seenTag[tag] = rowNum;
          }
        }
      }

      if (idx.serialNo >= 0) {
        var sn = String(row[idx.serialNo] || "").trim().toLowerCase();
        if (sn && sn !== "n/a" && sn !== "-") {
          if (seenNo[sn] !== undefined) {
            issues.push("Row " + rowNum + " \u00B7 " + String(row[idx.serialNo]).trim() +
                        " (Serial No. duplicate of row " + seenNo[sn] + ")");
          } else {
            seenNo[sn] = rowNum;
          }
        }
      }
    }
    return { count: issues.length, rows: issues };
  } catch (e) {
    return { count: 0, rows: [] };
  }
}

function _maint_missingFields_noGuard_() {
  try {
    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    var dataRows = values.slice(1);
    var issues   = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row     = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;
      var missing = [];
      if (idx.type     >= 0 && !_maint_isMeaningful_(row[idx.type]))     missing.push("Type");
      if (idx.status   >= 0 && !_maint_isMeaningful_(row[idx.status]))   missing.push("Status");
      if (idx.assignee >= 0 && !_maint_isMeaningful_(row[idx.assignee])) missing.push("Assignee");
      if (missing.length) {
        issues.push("Row " + (i + 2) + " \u00B7 Missing: " + missing.join(", "));
      }
    }
    return { count: issues.length, rows: issues };
  } catch (e) {
    return { count: 0, rows: [] };
  }
}

function _maint_invalidDates_noGuard_() {
  try {
    var sh       = _maint_getMasterlistSheet_();
    var values   = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { count: 0, rows: [] };

    var headers  = values[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var idx      = _maint_headerIndex_(headers);
    if (idx.warranty < 0) return { count: 0, rows: [] };

    var SKIP     = ["n/a", "na", "\u2014", "-", "none", "null", "(blank)", "tbd"];
    var dataRows = values.slice(1);
    var issues   = [];

    for (var i = 0; i < dataRows.length; i++) {
      var row = dataRows[i];
      if (!_maint_isRealRow_(row)) continue;
      var raw = String(row[idx.warranty] || "").trim();
      if (!raw || SKIP.indexOf(raw.toLowerCase()) !== -1) continue;
      if (!_maint_parseDate_(raw)) {
        issues.push("Row " + (i + 2) + " \u00B7 \"" + raw + "\" is not a recognised date");
      }
    }
    return { count: issues.length, rows: issues };
  } catch (e) {
    return { count: 0, rows: [] };
  }
}

function _maint_sendScanSummary_(dupes, missing, dates) {
  try {
    var total = dupes.count + missing.count + dates.count;
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    var lines = [
      "Daily Integrity Scan \u2014 " + stamp,
      "Total issues found: " + total,
      ""
    ];

    var i;
    if (dupes.count) {
      lines.push("=== Duplicate Serials (" + dupes.count + ") ===");
      for (i = 0; i < dupes.rows.length; i++) lines.push("  " + dupes.rows[i]);
      lines.push("");
    }
    if (missing.count) {
      lines.push("=== Missing Required Fields (" + missing.count + ") ===");
      for (i = 0; i < missing.rows.length; i++) lines.push("  " + missing.rows[i]);
      lines.push("");
    }
    if (dates.count) {
      lines.push("=== Invalid Warranty Dates (" + dates.count + ") ===");
      for (i = 0; i < dates.rows.length; i++) lines.push("  " + dates.rows[i]);
      lines.push("");
    }

   var recipient = WARRANTY_RECIPIENT_;
    try {
      if (typeof getWarrantyReminderConfig === "function") {
        recipient = getWarrantyReminderConfig().to || WARRANTY_RECIPIENT_;
      }
    } catch(_) {}

    MailApp.sendEmail(
      recipient,
      "[IT Inventory] Daily Integrity Scan \u2014 " + total + " issue(s) found",
      lines.join("\n")
    );
  } catch (e) {
    Logger.log("[MAINTENANCE] _maint_sendScanSummary_ error: " + e.message);
  }
}


function exportRowsToFolder(payload) {
  requireMaintenanceAccess_();

  var title   = String((payload && payload.title)   || "Scan Results").trim();
  var headers = Array.isArray(payload && payload.headers) ? payload.headers : ["Issue"];
  var rows    = Array.isArray(payload && payload.rows)    ? payload.rows    : [];

  if (!rows.length) return { ok: false, error: "No rows to export." };

  try {
    var ss        = _maint_openSS_();
    var stamp     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var sheetName = title + " — " + stamp;

    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);

    var sh     = ss.insertSheet(sheetName);
    var values = [headers].concat(rows.map(function(r) {
      return Array.isArray(r) ? r.map(function(c) { return c == null ? "" : String(c); })
                              : [String(r || "")];
    }));

    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
    sh.getRange(1, 1, 1, values[0].length).setFontWeight("bold");
    sh.autoResizeColumns(1, values[0].length);

    var url = ss.getUrl() + "#gid=" + sh.getSheetId();
    Logger.log("[MAINTENANCE] exportRowsToFolder: " + sheetName);
    return { ok: true, url: url, sheetName: sheetName };
  } catch (e) {
    Logger.log("[MAINTENANCE] exportRowsToFolder error: " + e.message);
    throw e;
  }
}