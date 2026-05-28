// ============================================================
// procurement.gs — v3 (clean rewrite)
// Procurement Tracker backend for vendor.html
//
// Sheet: "Procurement"
// Columns: ID | Request No | Vendor | Department | Item Description |
//          Status | Priority | Amount | Qty | Request Date | Expected Date |
//          Requested By | PO Number | Notes | Created | Updated
// ============================================================

var PROC_SHEET_NAME = "Procurement";
var PROC_COLS = 18;  // 

var PROC_VALID_STATUSES = [
  "Quotation Requested",
  "PR for Approval",
  "PR Approved",
  "PO Requested",
  "Payment Received",
  "Order In Transit",
  "Completed"
];


/* ============================================================
 * RBAC GUARDS
 * ============================================================ */
function _procRequireView_() {
  var email = Session.getActiveUser().getEmail();
  if (typeof canAccessPage === "function" && !canAccessPage("vendor", email)) {
    throw new Error("You do not have permission to access Procurement.");
  }
}

function _procRequireEdit_() {
  var email = Session.getActiveUser().getEmail();
  if (typeof canAccessPage === "function" && !canAccessPage("vendor", email)) {
    throw new Error("You do not have permission to access Procurement.");
  }
  if (typeof canDoAction === "function" &&
      !canDoAction("vendor.add", email) &&
      !canDoAction("vendor.edit", email)) {
    throw new Error("You do not have permission to edit Procurement.");
  }
}

function _procRequireDelete_() {
  var email = Session.getActiveUser().getEmail();
  if (typeof canAccessPage === "function" && !canAccessPage("vendor", email)) {
    throw new Error("You do not have permission to access Procurement.");
  }
  if (typeof canDoAction === "function" && !canDoAction("vendor.delete", email)) {
    throw new Error("You do not have permission to delete Procurement records.");
  }
}


/* ============================================================
 * SHEET HELPERS
 * ============================================================ */
function _procSS_() {
  return (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function _procSheet_() {
  var ss = _procSS_();
  var sh = ss.getSheetByName(PROC_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(PROC_SHEET_NAME);
    var headers = ["ID","Request No","Vendor","Department","Item Description",
                   "Status","Priority","Amount","Qty","Request Date","Expected Date",
                   "Requested By","PO Number","Notes","Ticket URL","Created","Updated"];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function _procFindRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return -1;
}

function _procFmtDate_(v) {
  if (!v) return "";
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return "";
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(v).trim();
  var m = s.match(/^(\d{4}-\d{2}-\d{2})T/);
  if (m) return m[1];
  return s;
}

function _procNewId_() {
  return "PR-" + Utilities.getUuid().replace(/-/g, "").toUpperCase().slice(0, 12);
}

function _procNextRequestNo_(sh) {
  var year = new Date().getFullYear();
  var lastRow = sh.getLastRow();
  var nextSeq = lastRow < 2 ? 1 : lastRow;
  return "PR-" + year + "-" + String(nextSeq).padStart(3, "0");
}


/* ============================================================
 * PUBLIC API
 * ============================================================ */

function procurement_getAll() {
  try {
    _procRequireView_();
    var sh = _procSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: [] };

    var lastCol = Math.min(sh.getLastColumn(), PROC_COLS);
    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    var data = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var id = String(row[0] || "").trim();
      if (!id) continue;

      data.push({
  id:              id,
  requestNo:       String(row[1] || "").trim(),
  vendor:          String(row[2] || "").trim(),
  department:      String(row[3] || "").trim(),
  itemDesc:        String(row[4] || "").trim(),
  status:          String(row[5] || "").trim(),
  priority:        String(row[6] || "").trim(),
  amount:          parseFloat(row[7]) || 0,
  qty:             parseFloat(row[8]) || 0,
  requestDate:     _procFmtDate_(row[9]),
  expectedDate:    _procFmtDate_(row[10]),
  requestedBy:     String(row[11] || "").trim(),
  poNumber:        String(row[12] || "").trim(),
  notes:           String(row[13] || "").trim(),
  ticketUrl:       String(row[14] || "").trim(),
  linkedSerialTag: String(row[17] || "").trim()   // ← ADD (col 18, index 17)
});
    }

    return { ok: true, data: data };
  } catch (e) {
    Logger.log("[PROCUREMENT] getAll error: " + (e.message || e));
    return { ok: false, error: e.message || String(e) };
  }
}


function procurement_save(r) {
  try {
    _procRequireEdit_();

    var vendor    = String((r && r.vendor)    || "").trim();
    var dept      = String((r && r.department)|| "").trim();
    var itemDesc  = String((r && r.itemDesc)  || "").trim();

    if (!vendor)   throw new Error("Vendor is required.");
    if (!dept)     throw new Error("Department is required.");
    if (!itemDesc) throw new Error("Item description is required.");

    var status = String((r && r.status) || "Quotation Requested").trim();
    if (PROC_VALID_STATUSES.indexOf(status) < 0) status = "Quotation Requested";

    var id             = String((r && r.id)             || "").trim();
    var requestNo      = String((r && r.requestNo)      || "").trim();
    var priority       = String((r && r.priority)       || "Medium").trim();
    var amount         = parseFloat(r && r.amount)      || 0;
    var qty            = parseFloat(r && r.qty)         || 0;
    var reqDate        = String((r && r.requestDate)    || "").trim();
    var expDate        = String((r && r.expectedDate)   || "").trim();
    var reqBy          = String((r && r.requestedBy)    || "").trim();
    var poNumber       = String((r && r.poNumber)       || "").trim();
    var notes          = String((r && r.notes)          || "").trim();
    var ticketUrl      = String((r && r.ticketUrl)      || "").trim();
    var linkedSerialTag = String((r && r.linkedSerialTag) || "").trim();

    var sh  = _procSheet_();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");

    var rowNo = id ? _procFindRow_(sh, id) : -1;

    if (rowNo > 0) {
      // UPDATE — preserve Created date (col 16)
      var created = sh.getRange(rowNo, 16).getValue();
      var rowData = [
        id, requestNo, vendor, dept, itemDesc, status, priority, amount, qty,
        reqDate, expDate, reqBy, poNumber, notes, ticketUrl, created, now, linkedSerialTag
      ];
      sh.getRange(rowNo, 1, 1, PROC_COLS).setValues([rowData]);
    } else {
      // INSERT
      if (!id)        id        = _procNewId_();
      if (!requestNo) requestNo = _procNextRequestNo_(sh);
      var newRow = [
        id, requestNo, vendor, dept, itemDesc, status, priority, amount, qty,
        reqDate, expDate, reqBy, poNumber, notes, ticketUrl, now, now, linkedSerialTag
      ];
      sh.appendRow(newRow);
    }

    // ── Cross-reference into Masterlist ────────────────────────────
    // Always clear old links for this request number first (handles
    // edits where the user removed or changed serial tags), then
    // write the new ones. Both steps are non-fatal.
    if (requestNo) {
      try {
        _clearProcurementRefFromMasterlist_(requestNo);
      } catch (clearErr) {
        Logger.log("[PROCUREMENT] clear old links error: " + (clearErr.message || clearErr));
      }
    }

    if (linkedSerialTag && requestNo) {
      try {
        _writeProcurementRefToMasterlist_(linkedSerialTag, requestNo);
      } catch (linkErr) {
        Logger.log("[PROCUREMENT] masterlist link error: " + (linkErr.message || linkErr));
      }
    }
    // ──────────────────────────────────────────────────────────────

    return { ok: true, id: id, requestNo: requestNo };

  } catch (e) {
    Logger.log("[PROCUREMENT] save error: " + (e.message || e));
    return { ok: false, error: e.message || String(e) };
  }
}

function procurement_delete(id) {
  try {
    _procRequireDelete_();
    id = String(id || "").trim();
    if (!id) return { ok: false, error: "ID required." };
    var sh = _procSheet_();
    var row = _procFindRow_(sh, id);
    if (row < 0) return { ok: false, error: "Record not found." };
    sh.deleteRow(row);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}


function procurement_exportToSheet() {
  try {
    _procRequireEdit_();
    var ss = _procSS_();
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var name = "Procurement Export - " + stamp;
    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    var copy = _procSheet_().copyTo(ss);
    copy.setName(name);
    return { ok: true, sheetName: name, url: ss.getUrl() + "#gid=" + copy.getSheetId() };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}


function procurement_diagnose() {
  try {
    Logger.log("=== PROCUREMENT DIAGNOSTIC ===");
    var email = Session.getActiveUser().getEmail();
    Logger.log("User: " + email);
    var sh = _procSheet_();
    Logger.log("Sheet: " + sh.getName() + " | Rows: " + sh.getLastRow() + " | Cols: " + sh.getLastColumn());
    var result = procurement_getAll();
    Logger.log("getAll: ok=" + result.ok + " | data.length=" + (result.data ? result.data.length : "N/A") +
               " | error=" + (result.error || "none"));
    if (result.data && result.data.length) {
      Logger.log("First record: " + JSON.stringify(result.data[0]));
    }
    Logger.log("=== END ===");
  } catch (e) {
    Logger.log("Diagnostic error: " + (e.message || e));
  }
}


/* ============================================================
 * DRIVE — File Attachments
 * ============================================================ */

var PROC_DRIVE_ROOT = "17aSLJQzClJtvk_FRv1cGqJ2JeINo8fOz";
var PROC_MAX_FILE_BYTES = 20 * 1024 * 1024;

function _procDriveRoot_() {
  try { return DriveApp.getFolderById(PROC_DRIVE_ROOT); }
  catch (e) { throw new Error("Procurement Drive root folder not found. Check PROC_DRIVE_ROOT."); }
}

function _procSafeName_(s) {
  return String(s || "").replace(/[\/\\:*?"<>|]/g, "-").trim() || "Unknown";
}

function _procGetOrCreateFolder_(parent, name) {
  var iter = parent.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(name);
}

function _procRequestFolder_(vendorName, reqNo, itemDesc) {
  var root = _procDriveRoot_();
  var vendorFolder = _procGetOrCreateFolder_(root, _procSafeName_(vendorName));
  var reqFolderName = _procSafeName_((reqNo || "Untitled") + " - " + (itemDesc || "Request"));
  return _procGetOrCreateFolder_(vendorFolder, reqFolderName);
}

function procurement_uploadDoc(payload) {
  try {
    _procRequireEdit_();
    var p = payload || {};
    var reqId    = String(p.reqId    || "").trim();
    var fileName = String(p.fileName || "").trim();
    var mimeType = String(p.mimeType || "application/octet-stream").trim();
    var b64      = String(p.base64Data || "").trim();

    if (!reqId)    return { ok: false, error: "Request ID required." };
    if (!fileName) return { ok: false, error: "File name required." };
    if (!b64)      return { ok: false, error: "File data required." };

    var approxBytes = Math.floor(b64.length * 3 / 4);
    if (approxBytes > PROC_MAX_FILE_BYTES) {
      return { ok: false, error: "File exceeds 20 MB limit." };
    }

    var sh = _procSheet_();
    var rowNo = _procFindRow_(sh, reqId);
    if (rowNo < 0) return { ok: false, error: "Request not found." };

    var rowData = sh.getRange(rowNo, 1, 1, PROC_COLS).getValues()[0];
    var requestNo = String(rowData[1] || "").trim();
    var vendor    = String(rowData[2] || "").trim();
    var itemDesc  = String(rowData[4] || "").trim();

    if (!vendor) return { ok: false, error: "Request has no vendor — cannot create folder." };

    var folder = _procRequestFolder_(vendor, requestNo, itemDesc);

    try {
      var existing = folder.getFilesByName(fileName);
      if (existing.hasNext()) {
        var ex = existing.next();
        return { ok: true, file: _procFileMeta_(ex), duplicate: true };
      }
    } catch (eDupe) {
      Logger.log("[PROCUREMENT] dedupe check failed (non-fatal): " + (eDupe.message || eDupe));
    }

    var blob = Utilities.newBlob(Utilities.base64Decode(b64), mimeType, fileName);
    var file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (eShare) {
      Logger.log("[PROCUREMENT] setSharing failed (non-fatal): " + (eShare.message || eShare));
    }

    return { ok: true, file: _procFileMeta_(file) };
  } catch (e) {
    Logger.log("[PROCUREMENT] uploadDoc error: " + (e.message || e));
    return { ok: false, error: e.message || String(e) };
  }
}

function _procFileMeta_(file) {
  var out = { id: "", name: "", url: "", mimeType: "", size: 0 };
  try { out.id       = file.getId(); }      catch(_){}
  try { out.name     = file.getName(); }    catch(_){}
  try { out.url      = file.getUrl(); }     catch(_){}
  try { out.mimeType = file.getMimeType(); }catch(_){}
  try { out.size     = file.getSize(); }    catch(_){}
  return out;
}

function procurement_listDocs(payload) {
  try {
    _procRequireView_();
    var p = payload || {};
    var reqId = String(p.reqId || "").trim();
    if (!reqId) return { ok: true, files: [] };

    var sh = _procSheet_();
    var rowNo = _procFindRow_(sh, reqId);
    if (rowNo < 0) return { ok: true, files: [] };

    var rowData = sh.getRange(rowNo, 1, 1, PROC_COLS).getValues()[0];
    var requestNo = String(rowData[1] || "").trim();
    var vendor    = String(rowData[2] || "").trim();
    var itemDesc  = String(rowData[4] || "").trim();

    if (!vendor) return { ok: true, files: [] };

    var root = _procDriveRoot_();
    var vIter = root.getFoldersByName(_procSafeName_(vendor));
    if (!vIter.hasNext()) return { ok: true, files: [] };
    var vFolder = vIter.next();

    var reqFolderName = _procSafeName_((requestNo || "Untitled") + " - " + (itemDesc || "Request"));
    var rIter = vFolder.getFoldersByName(reqFolderName);
    if (!rIter.hasNext()) return { ok: true, files: [] };
    var rFolder = rIter.next();

    var files = [];
    var fIter = rFolder.getFiles();
    while (fIter.hasNext()) {
      var f = fIter.next();
      var size = 0;
      try { size = f.getSize(); } catch (_) {}
      files.push({
        id:          f.getId(),
        name:        f.getName(),
        url:         f.getUrl(),
        mimeType:    f.getMimeType(),
        size:        size,
        createdDate: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), "MMM d, yyyy")
      });
    }

    files.sort(function(a, b) { return a.createdDate < b.createdDate ? 1 : -1; });
    return { ok: true, files: files };
  } catch (e) {
    Logger.log("[PROCUREMENT] listDocs error: " + (e.message || e));
    return { ok: false, error: e.message || String(e) };
  }
}

function procurement_deleteDoc(fileId) {
  try {
    _procRequireDelete_();
    var id = String(fileId || "").trim();
    if (!id) return { ok: false, error: "File ID required." };
    DriveApp.getFileById(id).setTrashed(true);
    return { ok: true };
  } catch (e) {
    Logger.log("[PROCUREMENT] deleteDoc error: " + (e.message || e));
    return { ok: false, error: e.message || String(e) };
  }
}


/* ============================================================
 * ANALYTICS
 * ============================================================ */

function procurement_getAnalytics() {
  try {
    _procRequireView_();
    var sh = _procSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, data: _procEmptyAnalytics_() };

    var lastCol = Math.min(sh.getLastColumn(), PROC_COLS);
    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz = Session.getScriptTimeZone();
    var today = new Date(); today.setHours(0,0,0,0);

    var byStatus = {}, byPriority = {}, byDepartment = {}, byVendor = {};
    var monthly = {};
    var totalSpend = 0, overdueCount = 0, validCount = 0;

    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      if (!String(r[0] || "").trim()) continue;
      validCount++;

      var status   = String(r[5] || "").trim() || "—";
      var priority = String(r[6] || "Medium").trim();
      var dept     = String(r[3] || "").trim() || "—";
      var vendor   = String(r[2] || "").trim() || "—";
      var amount   = parseFloat(r[7]) || 0;
      var reqDate  = r[9];   // ← was r[8]
      var expDate  = r[10];  // ← was r[9]

      totalSpend += amount;

      if (!byStatus[status])      byStatus[status]      = { count: 0, spend: 0 };
      if (!byPriority[priority])  byPriority[priority]  = { count: 0, spend: 0 };
      if (!byDepartment[dept])    byDepartment[dept]    = { count: 0, spend: 0 };
      if (!byVendor[vendor])      byVendor[vendor]      = { count: 0, spend: 0 };

      byStatus[status].count++;        byStatus[status].spend     += amount;
      byPriority[priority].count++;    byPriority[priority].spend += amount;
      byDepartment[dept].count++;      byDepartment[dept].spend   += amount;
      byVendor[vendor].count++;        byVendor[vendor].spend     += amount;

      var rd = reqDate instanceof Date ? reqDate : new Date(reqDate);
      if (rd && !isNaN(rd.getTime())) {
        var ym = Utilities.formatDate(rd, tz, "yyyy-MM");
        if (!monthly[ym]) monthly[ym] = { count: 0, spend: 0 };
        monthly[ym].count++;
        monthly[ym].spend += amount;
      }

      var ed = expDate instanceof Date ? expDate : new Date(expDate);
      if (ed && !isNaN(ed.getTime()) && ed < today &&
          status !== "Payment Received" && status !== "Order In Transit" &&
          status !== "Completed") {
        overdueCount++;
      }
    }

    var trend = [];
    var now = new Date();
    for (var m = 11; m >= 0; m--) {
      var d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      var key = Utilities.formatDate(d, tz, "yyyy-MM");
      var bucket = monthly[key] || { count: 0, spend: 0 };
      trend.push({ label: Utilities.formatDate(d, tz, "MMM yy"), count: bucket.count, spend: bucket.spend });
    }

    return {
      ok: true,
      data: {
        totals: {
          totalSpend: totalSpend,
          totalCount: validCount,
          avgAmount:  validCount ? Math.round(totalSpend / validCount) : 0,
          overdueCount: overdueCount
        },
        byStatus:     _procToSortedArr_(byStatus),
        byPriority:   _procToSortedArr_(byPriority),
        byDepartment: _procToSortedArr_(byDepartment),
        byVendor:     _procToSortedArr_(byVendor).slice(0, 10),
        trend: trend
      }
    };
  } catch (e) {
    Logger.log("[PROCUREMENT] getAnalytics error: " + (e.message || e));
    return { ok: false, error: e.message || String(e) };
  }
}

function _procEmptyAnalytics_() {
  return {
    totals: { totalSpend: 0, totalCount: 0, avgAmount: 0, overdueCount: 0 },
    byStatus: [], byPriority: [], byDepartment: [], byVendor: [],
    trend: []
  };
}

function _procToSortedArr_(obj) {
  var arr = [];
  Object.keys(obj).forEach(function(k){
    arr.push({ label: k, count: obj[k].count, spend: obj[k].spend });
  });
  arr.sort(function(a,b){ return b.spend - a.spend; });
  return arr;
}


/* ============================================================
 * BULK ACTIONS
 * ============================================================ */

function procurement_bulkDelete(ids) {
  try {
    _procRequireDelete_();
    if (!Array.isArray(ids) || !ids.length) {
      return { ok: false, error: "No IDs provided." };
    }
    var sh = _procSheet_();
    var deleted = 0, failed = [];
    var rows = [];
    ids.forEach(function(id){
      var r = _procFindRow_(sh, id);
      if (r > 0) rows.push(r);
      else failed.push(id);
    });
    rows.sort(function(a,b){ return b - a; });
    rows.forEach(function(r){
      try { sh.deleteRow(r); deleted++; } catch (_) {}
    });
    return { ok: true, deleted: deleted, failed: failed.length };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function procurement_bulkUpdateStatus(payload) {
  try {
    _procRequireEdit_();
    var p = payload || {};
    var ids = p.ids;
    var newStatus = String(p.status || "").trim();
    if (!Array.isArray(ids) || !ids.length) return { ok: false, error: "No IDs provided." };
    if (PROC_VALID_STATUSES.indexOf(newStatus) < 0) return { ok: false, error: "Invalid status." };

    var sh = _procSheet_();
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var updated = 0;

    ids.forEach(function(id){
      var r = _procFindRow_(sh, id);
      if (r > 0) {
        sh.getRange(r, 6).setValue(newStatus);   // col 6 = Status (unchanged)
        sh.getRange(r, 17).setValue(now);        // ← was 15, now 16 = Updated
        updated++;
      }
    });

    return { ok: true, updated: updated };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function procurement_bulkImport(rows) {
  try {
    _procRequireEdit_();
    if (!Array.isArray(rows) || !rows.length) return { ok: false, error: "No rows provided." };

    var sh = _procSheet_();
    var year = new Date().getFullYear();
    var startSeq = Math.max(1, sh.getLastRow());
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    var inserted = 0, errors = [];

    var toAppend = [];
    rows.forEach(function(r, idx){
      var vendor = String((r && r.vendor) || "").trim();
      var dept   = String((r && r.department) || "").trim();
      var item   = String((r && r.itemDesc) || "").trim();
      if (!vendor || !dept || !item) {
        errors.push("Row " + (idx+2) + ": missing vendor/department/item");
        return;
      }
      var status   = String((r && r.status) || "Quotation Requested").trim();
      if (PROC_VALID_STATUSES.indexOf(status) < 0) status = "Quotation Requested";
      var priority = String((r && r.priority) || "Medium").trim();
      var amount   = parseFloat(r && r.amount) || 0;
      var qty      = parseFloat(r && r.qty) || 0;  // ← new

      var id = _procNewId_();
      var seq = startSeq + toAppend.length;
      var requestNo = "PR-" + year + "-" + String(seq).padStart(3, "0");

      toAppend.push([
        id, requestNo, vendor, dept, item, status, priority, amount, qty,
        String((r && r.requestDate)  || "").trim(),
        String((r && r.expectedDate) || "").trim(),
        String((r && r.requestedBy)  || "").trim(),
        String((r && r.poNumber)     || "").trim(),
        String((r && r.notes)        || "").trim(),
        String((r && r.ticketUrl)    || "").trim(),
        now, now
      ]);
      inserted++;
    });

    if (toAppend.length) {
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, toAppend.length, PROC_COLS).setValues(toAppend);
    }

    return { ok: true, inserted: inserted, errors: errors };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}


/* ============================================================
 * COMMENTS
 * ============================================================ */

var PROC_COMMENTS_SHEET = "Procurement_Comments";
var PROC_COMMENTS_COLS = 6;

function _procCommentsSheet_() {
  var ss = _procSS_();
  var sh = ss.getSheetByName(PROC_COMMENTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(PROC_COMMENTS_SHEET);
    sh.getRange(1, 1, 1, PROC_COMMENTS_COLS)
      .setValues([["ID","Request ID","Author Email","Author Name","Timestamp","Comment"]])
      .setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function procurement_listComments(reqId) {
  try {
    _procRequireView_();
    reqId = String(reqId || "").trim();
    if (!reqId) return { ok: true, comments: [] };

    var sh = _procCommentsSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, comments: [] };

    var values = sh.getRange(2, 1, lastRow - 1, PROC_COMMENTS_COLS).getValues();
    var tz = Session.getScriptTimeZone();
    var comments = [];

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      if (String(row[1]).trim() !== reqId) continue;

      var ts = row[4];
      var tsStr = "";
      if (ts instanceof Date && !isNaN(ts.getTime())) {
        tsStr = Utilities.formatDate(ts, tz, "MMM d, yyyy h:mm a");
      } else {
        tsStr = String(ts || "");
      }

      comments.push({
        id:      String(row[0] || "").trim(),
        reqId:   String(row[1] || "").trim(),
        email:   String(row[2] || "").trim(),
        author:  String(row[3] || "").trim() || String(row[2] || "").split("@")[0],
        ts:      tsStr,
        tsRaw:   ts instanceof Date ? ts.getTime() : 0,
        comment: String(row[5] || "")
      });
    }
    comments.sort(function(a,b){ return a.tsRaw - b.tsRaw; });
    return { ok: true, comments: comments };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function procurement_addComment(payload) {
  try {
    _procRequireView_();
    var p = payload || {};
    var reqId = String(p.reqId || "").trim();
    var text  = String(p.comment || "").trim();
    if (!reqId) return { ok: false, error: "Request ID required." };
    if (!text)  return { ok: false, error: "Comment cannot be empty." };
    if (text.length > 2000) return { ok: false, error: "Comment too long (max 2000 chars)." };

    var email  = Session.getActiveUser().getEmail() || "unknown";
    var author = email.split("@")[0];
    var now    = new Date();
    var id     = "C-" + Utilities.getUuid().replace(/-/g, "").toUpperCase().slice(0, 12);

    var sh = _procCommentsSheet_();
    sh.appendRow([id, reqId, email, author, now, text]);

    return {
      ok: true,
      comment: {
        id:      id,
        reqId:   reqId,
        email:   email,
        author:  author,
        ts:      Utilities.formatDate(now, Session.getScriptTimeZone(), "MMM d, yyyy h:mm a"),
        tsRaw:   now.getTime(),
        comment: text
      }
    };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function procurement_deleteComment(commentId) {
  try {
    _procRequireView_();
    var id = String(commentId || "").trim();
    if (!id) return { ok: false, error: "Comment ID required." };

    var sh = _procCommentsSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: false, error: "Comment not found." };

    var values = sh.getRange(2, 1, lastRow - 1, PROC_COMMENTS_COLS).getValues();
    var email = Session.getActiveUser().getEmail() || "";

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === id) {
        var authorEmail   = String(values[i][2]).trim().toLowerCase();
        var isAuthor      = authorEmail === email.toLowerCase();
        var canAdminDelete = (typeof canDoAction === "function") && canDoAction("vendor.delete", email);
        if (!isAuthor && !canAdminDelete) {
          return { ok: false, error: "You can only delete your own comments." };
        }
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
    return { ok: false, error: "Comment not found." };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}


/* ============================================================
 * SUPPLIERS
 * ============================================================ */

const SUPPLIERS_SHEET = 'Suppliers';

function suppliers_getAll() {
  try {
    const ss = _openSS_();
    let sh = ss.getSheetByName(SUPPLIERS_SHEET);
    if (!sh) return { ok: true, data: [] };

    const vals = sh.getDataRange().getDisplayValues();
    if (vals.length < 2) return { ok: true, data: [] };

    const headers = vals[0].map(h => String(h).trim().toLowerCase());
    const get = (row, name) => {
      const i = headers.indexOf(name);
      return i >= 0 ? String(row[i] || '').trim() : '';
    };

    const data = vals.slice(1)
      .filter(r => r.some(c => String(c).trim()))
      .map(r => ({
        id:           get(r, 'id'),
        name:         get(r, 'name'),
        category:     get(r, 'category'),
        contactName:  get(r, 'contactname'),
        email:        get(r, 'email'),
        phone:        get(r, 'phone'),
        website:      get(r, 'website'),
        status:       get(r, 'status') || 'Active',
        paymentTerms: get(r, 'paymentterms'),
        notes:        get(r, 'notes'),
      }));

    return { ok: true, data };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function suppliers_save(payload) {
  try {
    const ss = _openSS_();
    let sh = ss.getSheetByName(SUPPLIERS_SHEET);
    if (!sh) {
      sh = ss.insertSheet(SUPPLIERS_SHEET);
      sh.appendRow(['ID','Name','Category','ContactName','Email','Phone','Website','Notes','Status','PaymentTerms']);
    }

    const vals = sh.getDataRange().getValues();
    const headers = vals[0].map(h => String(h).trim().toLowerCase());
    const idCol = headers.indexOf('id');
    const totalCols = headers.length;

    const row = new Array(totalCols).fill('');
    const set = (headerName, value) => {
      const i = headers.indexOf(headerName);
      if (i >= 0) row[i] = value || '';
    };

    set('id',           payload.id || '');
    set('name',         payload.name || '');
    set('category',     payload.category || '');
    set('contactname',  payload.contactName || '');
    set('email',        payload.email || '');
    set('phone',        payload.phone || '');
    set('website',      payload.website || '');
    set('status',       payload.status || 'Active');
    set('paymentterms', payload.paymentTerms || '');
    set('notes',        payload.notes || '');

    if (payload.id) {
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][idCol]).trim() === payload.id) {
          sh.getRange(i + 1, 1, 1, totalCols).setValues([row]);
          return { ok: true, id: payload.id };
        }
      }
    }

    const newId = 'SUP-' + Date.now();
    row[idCol >= 0 ? idCol : 0] = newId;
    sh.appendRow(row);
    return { ok: true, id: newId };

  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function suppliers_delete(id) {
  try {
    const ss = _openSS_();
    const sh = ss.getSheetByName(SUPPLIERS_SHEET);
    if (!sh) return { ok: true };

    const vals = sh.getDataRange().getValues();
    const idCol = vals[0].map(h => String(h).trim().toLowerCase()).indexOf('id');
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][idCol]).trim() === id) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}