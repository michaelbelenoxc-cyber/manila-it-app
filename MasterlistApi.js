// ============================================================
// masterlist.gs
// Rhino-safe: var only, no arrow functions, no template literals,
// no optional chaining, no const/let, no shorthand properties,
// no for...of, no destructuring.
// ============================================================

var MASTERLIST_SHEET = "Masterlist";

var MASTERLIST_CACHE_KEY_     = "ml_display_v1";
var MASTERLIST_CACHE_TTL_     = 3 * 60; // 3 minutes in seconds
var EMPLOYEES_CACHE_KEY_      = "ml_employees_v1";
var EMPLOYEES_CACHE_TTL_      = 10 * 60; // 10 minutes

// masterlist_shared.gs  (new file, or paste at top of masterlist.gs)
// Rhino-safe: var, no arrow functions.

var STATUS_CANONICAL_ = {
  "assigned":         "Assigned",
  "in use":           "Assigned",
  "in-stock":         "In-stock",
  "in stock":         "In-stock",
  "instock":          "In-stock",
  "available":        "In-stock",
  "stock":            "In-stock",
  "spare":            "In-stock",
  "return":           "Return",
  "offboard - return":"Offboard - Return",
  "offboard-return":  "Offboard - Return",
  "retired":          "Retired",
  "disposed":         "Retired",
  "decommission":     "Retired",
  "defective":        "Defective",
  "faulty":           "Defective",
  "missing":          "Missing",
  "lost":             "Missing"
};

function normalizeStatusCanonical_(raw) {
  var s = String(raw || "").trim().toLowerCase();
  if (STATUS_CANONICAL_[s]) return STATUS_CANONICAL_[s];
  // partial-match fallback
  var keys = Object.keys(STATUS_CANONICAL_);
  for (var i = 0; i < keys.length; i++) {
    if (s.indexOf(keys[i]) >= 0) return STATUS_CANONICAL_[keys[i]];
  }
  return raw || "";
}
 
 
function _mlSetCachedDisplayValues_(values) {
  try {
    var cache = CacheService.getScriptCache();
    var json  = JSON.stringify(values);

    if (json.length < 90000) {
      // Fits in one entry — store normally, bump TTL to 10 min
      cache.put(MASTERLIST_CACHE_KEY_, json, 10 * 60);
      return;
    }

    // Chunked storage for large sheets (up to ~900KB across 10 entries)
    var CHUNK = 85000;
    var chunks = Math.ceil(json.length / CHUNK);
    var putMap = {};

    for (var i = 0; i < chunks; i++) {
      putMap[MASTERLIST_CACHE_KEY_ + '_chunk_' + i] = json.slice(i * CHUNK, (i + 1) * CHUNK);
    }
    putMap[MASTERLIST_CACHE_KEY_ + '_meta'] = JSON.stringify({ chunks: chunks, ts: Date.now() });

    // putAll is a single round-trip to the cache service
    cache.putAll(putMap, 10 * 60);
  } catch (e) {
    console.warn("_mlSetCachedDisplayValues_ write failed:", e && e.message);
  }
}

function _mlGetCachedDisplayValues_() {
  try {
    var cache = CacheService.getScriptCache();

    // Try single-entry path first
    var single = cache.get(MASTERLIST_CACHE_KEY_);
    if (single) {
      var parsed = JSON.parse(single);
      if (Array.isArray(parsed)) return parsed;
    }

    // Try chunked path
    var metaRaw = cache.get(MASTERLIST_CACHE_KEY_ + '_meta');
    if (!metaRaw) return null;

    var meta = JSON.parse(metaRaw);
    if (!meta || !meta.chunks) return null;

    var keys = [];
    for (var i = 0; i < meta.chunks; i++) {
      keys.push(MASTERLIST_CACHE_KEY_ + '_chunk_' + i);
    }

    var got = cache.getAll(keys);
    var assembled = '';
    for (var j = 0; j < keys.length; j++) {
      if (!got[keys[j]]) return null; // chunk expired — miss
      assembled += got[keys[j]];
    }

    var result = JSON.parse(assembled);
    return Array.isArray(result) ? result : null;
  } catch (e) {
    console.warn("_mlGetCachedDisplayValues_ read failed:", e && e.message);
    return null;
  }
}

function _mlBustCache_() {
  try {
    var cache = CacheService.getScriptCache();
    // Bust both single-entry and any chunk entries
    var metaRaw = cache.get(MASTERLIST_CACHE_KEY_ + '_meta');
    if (metaRaw) {
      try {
        var meta = JSON.parse(metaRaw);
        var toRemove = [MASTERLIST_CACHE_KEY_ + '_meta'];
        for (var i = 0; i < (meta.chunks || 0); i++) {
          toRemove.push(MASTERLIST_CACHE_KEY_ + '_chunk_' + i);
        }
        cache.removeAll(toRemove);
      } catch (_) {}
    }
    cache.remove(MASTERLIST_CACHE_KEY_);
  } catch (e) {}
}

function _mlGetDisplayValues_() {
  var cached = _mlGetCachedDisplayValues_();
  if (cached) return cached;

  var sh     = _mlSheet_();
  var values = sh.getDataRange().getDisplayValues();
  _mlSetCachedDisplayValues_(values);
  return values;
}

/* ============================================================
 * HEADER HELPERS
 * ============================================================ */

function _normHeader_(s) {
  return String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\s._-]+/g, " ")
    .replace(/[^a-z0-9 ]/gi, "")
    .trim()
    .toLowerCase();
}

function _getHeaderMap_(headerRow) {
  var headers = [];
  for (var i = 0; i < headerRow.length; i++) {
    headers.push(String(headerRow[i] == null ? "" : headerRow[i]));
  }

  var norm = [];
  for (var j = 0; j < headers.length; j++) {
    norm.push(_normHeader_(headers[j]));
  }

  function idxAny(names) {
    var arr = Array.isArray(names) ? names : [names];
    for (var a = 0; a < arr.length; a++) {
      var idx = norm.indexOf(_normHeader_(arr[a]));
      if (idx >= 0) return idx;
    }
    return -1;
  }

  function idx(name) {
    return idxAny([name]);
  }

  return { headers: headers, idx: idx, idxAny: idxAny };
}

function _safeGet_(row, i) {
  if (i < 0) return "";
  var v = row[i];
  return v == null ? "" : v;
}

function _fmtDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "MM/dd/yyyy");
  }
  return v || "";
}

function _requireCols_(colMap, requiredKeys, headersForDebug) {
  for (var i = 0; i < requiredKeys.length; i++) {
    var k = requiredKeys[i];
    if (colMap[k] == null || colMap[k] < 0) {
      console.error("Missing required header: " + k);
      if (headersForDebug) console.error("Headers found:", headersForDebug);
      return false;
    }
  }
  return true;
}

function _mlSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(MASTERLIST_SHEET);
  if (!sh) throw new Error("Sheet \"" + MASTERLIST_SHEET + "\" not found.");
  return sh;
}

function _toDateOrBlank_(v) {
  if (!v) return "";
  if (v instanceof Date) return v;

  var s = String(v).trim();
  if (!s) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var parts = s.split("-");
    var y = Number(parts[0]);
    var m = Number(parts[1]);
    var d = Number(parts[2]);
    return new Date(y, m - 1, d);
  }

  var mmddyyyy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (mmddyyyy) {
    var mm = Number(mmddyyyy[1]);
    var dd = Number(mmddyyyy[2]);
    var yy = Number(mmddyyyy[3]);
    if (yy < 100) yy += 2000;
    return new Date(yy, mm - 1, dd);
  }

  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? s : d2;
}

function _stripHtml_(v) {
  return String(v || "").replace(/<[^>]*>/g, "");
}

function _masterlistHeaderAliases_() {
  return {
    type:               ["Type", "Asset Type", "Device Type"],
    model:              ["Model"],
    serialTag:          ["Serial Tag"],
    serialNumber:       ["Serial Number"],
    ram:                ["RAM"],
    purchaseDate:       ["Purchase Date"],
    endOfWarranty:      ["End of Warranty"],
    status:             ["Status"],
    assignee:           ["Assignee"],
    email:              ["Email"],
    department:         ["Department"],
    accessories:        ["Accessories"],
    accountabilityForm: ["Accountability Form"],
    signedAf:           ["Signed AF"],
    afStatus:           ["AF Status"],
    afSignedDate:       ["AF Signed Date"],
    notes:              ["Notes"],
    procurementRef:     ["procurementRef", "Procurement Ref"]  // ← ADD THIS
  };
}

function _masterlistColMapFromHeaderRow_(headerRow) {
  var hm      = _getHeaderMap_(headerRow);
  var aliases = _masterlistHeaderAliases_();
  var col     = {};
  var keys    = Object.keys(aliases);

  for (var i = 0; i < keys.length; i++) {
    col[keys[i]] = hm.idxAny(aliases[keys[i]]);
  }

  return { headers: hm.headers, col: col };
}

function _masterlistObjectFromRow_(row, col, useFormatter) {
  function get(k) { return _safeGet_(row, col[k]); }

  return {
    type:               get("type"),
    model:              get("model"),
    serialTag:          get("serialTag"),
    serialNumber:       get("serialNumber"),
    ram:                get("ram"),
    purchaseDate:       useFormatter ? _fmtDate_(get("purchaseDate"))  : get("purchaseDate"),
    endOfWarranty:      useFormatter ? _fmtDate_(get("endOfWarranty")) : get("endOfWarranty"),
    status:             get("status"),
    assignee:           get("assignee"),
    email:              get("email"),
    department:         get("department"),
    accessories:        get("accessories"),
    accountabilityForm: get("accountabilityForm"),
    signedAf:           get("signedAf"),
    afStatus:           get("afStatus"),
    afSignedDate:       useFormatter ? _fmtDate_(get("afSignedDate")) : get("afSignedDate"),
    notes:              get("notes"),
    procurementRef:     get("procurementRef")   // ← ADD THIS
  };
}

function _mlHeaderMap_(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error("Masterlist sheet has no columns.");

  var headerRow = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var result    = _masterlistColMapFromHeaderRow_(headerRow);
  var col       = result.col;
  var map       = {};
  var keys      = Object.keys(col);

  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (col[k] >= 0) map[k] = col[k] + 1; // 1-based
  }

  return { headers: result.headers, map: map };
}

function _mlFindRowBySerialTag_(sh, serialTag) {
  var tag = String(serialTag || "").trim();
  if (!tag) return -1;

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;

  var hm  = _mlHeaderMap_(sh);
  var col = hm.map.serialTag;
  if (!col) throw new Error("Header \"Serial Tag\" not found.");

  var vals = sh.getRange(2, col, lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() === tag) return i + 2;
  }
  return -1;
}

function _mlNormalizeKey_(header) {
  var h   = String(header || "").trim();
  var map = {
    "Type":               "type",
    "Asset Type":         "type",
    "Device Type":        "type",
    "Model":              "model",
    "Serial Tag":         "serialTag",
    "Serial Number":      "serialNumber",
    "RAM":                "ram",
    "Purchase Date":      "purchaseDate",
    "End of Warranty":    "endOfWarranty",
    "Status":             "status",
    "Assignee":           "assignee",
    "Email":              "email",
    "Department":         "department",
    "Accessories":        "accessories",
    "Accountability Form":"accountabilityForm",
    "Signed AF":          "signedAf",
    "AF Status":          "afStatus",
    "AF Signed Date":     "afSignedDate",
    "Notes":              "notes",
    "procurementRef":     "procurementRef",
"Procurement Ref":    "procurementRef"
  };
  return map[h] || h;
}

function _mlRowToObject_(sh, row) {
  var lastCol = sh.getLastColumn();
  var rawHdrs = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var values  = sh.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
  var obj     = {};

  for (var i = 0; i < rawHdrs.length; i++) {
    var h = String(rawHdrs[i] || "").trim();
    if (!h) continue;
    obj[_mlNormalizeKey_(h)] = values[i];
  }

  var expectedKeys = [
  "type", "model", "serialTag", "serialNumber", "ram",
  "purchaseDate", "endOfWarranty", "status", "assignee",
  "email", "department", "accessories", "accountabilityForm",
  "signedAf", "afStatus", "afSignedDate", "notes", "procurementRef"  // ← add this
];

  for (var j = 0; j < expectedKeys.length; j++) {
    if (!(expectedKeys[j] in obj)) obj[expectedKeys[j]] = "";
  }

  return obj;
}

function _mlHeaderForKey_(key) {
  var map = {
    type:               "Type",
    model:              "Model",
    serialTag:          "Serial Tag",
    serialNumber:       "Serial Number",
    ram:                "RAM",
    purchaseDate:       "Purchase Date",
    endOfWarranty:      "End of Warranty",
    status:             "Status",
    assignee:           "Assignee",
    email:              "Email",
    department:         "Department",
    accessories:        "Accessories",
    accountabilityForm: "Accountability Form",
    signedAf:           "Signed AF",
    afStatus:           "AF Status",
    afSignedDate:       "AF Signed Date",
    notes:              "Notes",
    procurementRef: "procurementRef"
  };
  return map[key] || null;
}

function serverComparable_(item, key) {
  var v = (item && item[key] != null) ? item[key] : null;
  if (v == null) return "";

  if (key === "ram") {
    var num = parseFloat(String(v).replace(/[^0-9.]/g, ""));
    return isNaN(num) ? 0 : num;
  }

  if (key === "purchaseDate" || key === "endOfWarranty" || key === "afSignedDate") {
    var s = String(v).trim();
    if (!s) return 0;

    var parts = s.split("/");
    if (parts.length === 3) {
      var mm = Number(parts[0]);
      var dd = Number(parts[1]);
      var yy = Number(parts[2]);
      if (yy < 100) yy += 2000;
      var d = new Date(yy, mm - 1, dd);
      return d.getTime() || 0;
    }

    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) {
      var di = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
      return di.getTime() || 0;
    }

    return 0;
  }

  return String(v).toLowerCase();
}




/* ============================================================
 * PROCUREMENT → MASTERLIST LINK
 * ============================================================ */

/** Write a procurement request number into the asset row's procurementRef column. */
function _writeProcurementRefToMasterlist_(serialTagsCsv, requestNo) {
  var sh = _mlSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) {
    return String(h || '').trim().toLowerCase();
  });

  var refCol = headers.indexOf('procurementref');
  if (refCol < 0) {
    refCol = headers.length;
    sh.getRange(1, refCol + 1).setValue('procurementRef');
  }

  var stCol = headers.indexOf('serial tag');
  if (stCol < 0) return;

  // Parse comma-separated serial tags
  var tags = String(serialTagsCsv || '').split(',').map(function(t) {
    return t.trim();
  }).filter(Boolean);

  if (!tags.length) return;

  // Write requestNo into every matching asset row
  for (var i = 1; i < data.length; i++) {
    var rowTag = String(data[i][stCol] || '').trim();
    if (!rowTag) continue;
    for (var j = 0; j < tags.length; j++) {
      if (rowTag === tags[j]) {
        sh.getRange(i + 1, refCol + 1).setValue(requestNo);
        break;
      }
    }
  }

  SpreadsheetApp.flush();
}

/** Remove a procurement request number from ALL assets that currently reference it. */
function _clearProcurementRefFromMasterlist_(requestNo) {
  var sh = _mlSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) {
    return String(h || '').trim().toLowerCase();
  });

  var refCol = headers.indexOf('procurementref');
  if (refCol < 0) return; // column doesn't exist yet — nothing to clear

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][refCol] || '').trim() === String(requestNo || '').trim()) {
      sh.getRange(i + 1, refCol + 1).clearContent();
    }
  }

  SpreadsheetApp.flush();
}


/* ============================================================
 * READ
 * ============================================================ */

function getAssetMasterlistData() {
  try {
    var sheet = _mlSheet_();
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var result  = _masterlistColMapFromHeaderRow_(data[0]);
    var headers = result.headers;
    var col     = result.col;

    if (!_requireCols_(col, ["serialTag", "type"], headers)) return [];

    var out = [];
    for (var i = 1; i < data.length; i++) {
      var tag = String(_safeGet_(data[i], col.serialTag)).trim();
      if (!tag) continue;
      out.push(_masterlistObjectFromRow_(data[i], col, true));
    }
    return out;
  } catch (e) {
    console.error("Masterlist Read Error:", e);
    return [];
  }
}


/* ============================================================
 * CLEAR ASSIGNMENT FIELDS
 * ============================================================ */

function clearAssetFieldsInSheet(assetId, fieldsToClear, finalStatus) {
  var sheet = _mlSheet_();
  var data  = sheet.getDataRange().getValues();

  if (data.length < 2) throw new Error("Masterlist has no data rows.");

  var col = _masterlistColMapFromHeaderRow_(data[0]).col;

  if (typeof col.serialTag !== "number" || col.serialTag < 0) {
    throw new Error("Header \"Serial Tag\" not found.");
  }

  var tag = String(assetId || "").trim();
  if (!tag) throw new Error("Missing Serial Tag.");

  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][col.serialTag] || "").trim() === tag) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex < 0) throw new Error("Serial Tag \"" + tag + "\" not found.");

  var sheetRow = rowIndex + 1;

  var keyToCol = {
    assignee:           col.assignee,
    email:              col.email,
    department:         col.department,
    accessories:        col.accessories,
    accountabilityForm: col.accountabilityForm,
    signedAf:           col.signedAf,
    afStatus:           col.afStatus,
    afSignedDate:       col.afSignedDate
  };

  var fields = fieldsToClear || [];
  for (var f = 0; f < fields.length; f++) {
    var c = keyToCol[fields[f]];
    if (typeof c === "number" && c >= 0) {
      sheet.getRange(sheetRow, c + 1).clearContent();
    }
  }

  // Use the chosen status from the modal, default to In-stock
  if (typeof col.status === "number" && col.status >= 0) {
    var statusToSet = String(finalStatus || "In-stock").trim() || "In-stock";
    sheet.getRange(sheetRow, col.status + 1).setValue(statusToSet);
  }

  SpreadsheetApp.flush();
  return true;
}


/* ============================================================
 * DELETE ROW BY SERIAL TAG
 * ============================================================ */

function deleteAssetFromMasterlist(assetId) {
  try {
    var sheet = _mlSheet_();
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) return false;

    var col = _masterlistColMapFromHeaderRow_(data[0]).col;
    if (col.serialTag < 0) return false;

    var target   = String(assetId || "").trim();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][col.serialTag]).trim() === target) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex < 0) return false;

    sheet.deleteRow(rowIndex + 1);
    return true;
  } catch (e) {
    console.error("Deletion Error:", e);
    return false;
  }
}


/* ============================================================
 * ADD NEW ASSET ROW
 * ============================================================ */

function addAssetMasterlistRow(payload) {
  payload = payload || {};

  // Required field validation
  var type         = String(payload.type         || "").trim();
  var model        = String(payload.model        || "").trim();
  var serialTag    = String(payload.serialTag    || "").trim();
  var serialNumber = String(payload.serialNumber || "").trim();

  if (!type)         throw new Error("Type is required.");
  if (!model)        throw new Error("Model is required.");
  if (!serialTag)    throw new Error("Serial Tag is required.");
  if (!serialNumber) throw new Error("Serial Number is required.");

  var sh   = _mlSheet_();
  var data = sh.getDataRange().getValues();
  if (!data.length) throw new Error("Masterlist sheet is empty (missing headers).");

  var hm        = _getHeaderMap_(data[0]);
  var serialCol = hm.idx("Serial Tag");
  if (serialCol < 0) throw new Error("Missing header: \"Serial Tag\"");

  var existingCount = Math.max(0, sh.getLastRow() - 1);
  if (existingCount > 0) {
    var existing = sh.getRange(2, serialCol + 1, existingCount, 1).getDisplayValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0] || "").trim() === serialTag) {
        throw new Error("Serial Tag already exists.");
      }
    }
  }

 var row = [];
  for (var j = 0; j < hm.headers.length; j++) row.push("");

  function set(headerName, value) {
    var c = hm.idx(headerName);
    if (c >= 0) row[c] = (value == null ? "" : value);
  }

  set("Type",               payload.type);
  set("Model",              payload.model);
  set("Serial Tag",         serialTag);
  set("Serial Number",      payload.serialNumber);
  set("RAM",                payload.ram);
  set("Purchase Date",      _toDateOrBlank_(payload.purchaseDate));
  set("End of Warranty",    _toDateOrBlank_(payload.endOfWarranty));
  set("Status",             payload.status || "In-stock");
  set("Assignee",           payload.assignee);
  set("Email",              payload.email);
  set("Department",         payload.department);
  set("Accessories",        payload.accessories);
  set("Accountability Form",payload.accountabilityForm);
  set("Signed AF",          payload.signedAf);
  set("AF Status",          payload.afStatus);
  set("AF Signed Date",     _toDateOrBlank_(payload.afSignedDate));
  set("Notes",               payload.notes);

  sh.appendRow(row);
  return true;
}


/* ============================================================
 * EXPORT
 * ============================================================ */

function generateMasterlistExport(config) {
  config = config || {};

  var ss = SpreadsheetApp.create(
    "Export - " + (config.title || "Masterlist") + " (" +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") + ")"
  );
  var sheet = ss.getSheets()[0];

  sheet.appendRow(config.headers || []);

  var configData    = config.data    || [];
  var configKeys    = config.keys    || [];
  var configHeaders = config.headers || [];
  var rows          = [];

  for (var i = 0; i < configData.length; i++) {
    var item = configData[i] || {};
    var r    = [];
    for (var j = 0; j < configKeys.length; j++) {
      var val = item[configKeys[j]];
      r.push(_stripHtml_(val));
    }
    rows.push(r);
  }

  if (rows.length && configHeaders.length) {
    sheet.getRange(2, 1, rows.length, configHeaders.length).setValues(rows);
  }

  var lastCol = Math.max(1, configHeaders.length);
  var lastRow = sheet.getLastRow();

  sheet.getRange(1, 1, 1, lastCol)
    .setBackground("#1c1c1e")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, lastCol);

  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=ISEVEN(ROW())")
      .setBackground("#F5F5F5")
      .setRanges([dataRange])
      .build();
    sheet.setConditionalFormatRules([rule]);
  }

  return ss.getUrl();
}


/* ============================================================
 * DEBUG HELPER
 * ============================================================ */

function debugMasterlistHeaders_() {
  var sh      = _mlSheet_();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var out     = [];
  for (var i = 0; i < headers.length; i++) {
    out.push("[" + headers[i] + "]");
  }
  Logger.log(out.join(" | "));
}


/* ============================================================
 * WRAPPERS — called by masterlist.html
 * ============================================================ */

function masterlist_list() {
  try {
    return { ok: true, data: getAssetMasterlistData() };
  } catch (e) {
    console.error("masterlist_list error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

function masterlist_add(payload) {
  try {
    var email = getCurrentUserEmail_();
    requireAssetReportAdd_(email);
    addAssetMasterlistRow(payload)
    _mlBustCache_();

     assetLogWrite({
      action:    "ADD",
      serialTag: (payload && payload.serialTag) || "",
      entity:    "Masterlist",
      field:     "*",
      oldValue:  "",
      // Surface key fields so the log table is scannable without
      // expanding the raw JSON every time.
      newValue:  [
        String((payload && payload.type)      || ""),
        String((payload && payload.model)     || ""),
        String((payload && payload.serialTag) || "")
      ].filter(Boolean).join(" | "),
      details: { payload: payload }
    });

    // ── Slack notification ──────────────────────────────
    try {
      var actor = email;
      var p     = payload || {};

      sendSlackNotification_(null, [
        {
          type: "section",
          text: {
            type: "mrkdwn",
            text: "*\uD83D\uDCBB New Asset Added to Manila IT Inventory*"
          }
        },
        {
          type: "section",
          fields: [
            { type: "mrkdwn", text: "*Serial Tag:*\n" + (String(p.serialTag    || "").trim() || "\u2014") },
            { type: "mrkdwn", text: "*Type:*\n"       + (String(p.type         || "").trim() || "\u2014") },
            { type: "mrkdwn", text: "*Model:*\n"      + (String(p.model        || "").trim() || "\u2014") },
            { type: "mrkdwn", text: "*Serial No:*\n"  + (String(p.serialNumber || "").trim() || "\u2014") },
            { type: "mrkdwn", text: "*RAM:*\n"        + (String(p.ram          || "").trim() || "\u2014") },
            { type: "mrkdwn", text: "*Status:*\n"     + (String(p.status       || "In-stock").trim())     },
            { type: "mrkdwn", text: "*Department:*\n" + (String(p.department   || "").trim() || "\u2014") },
            { type: "mrkdwn", text: "*Added by:*\n"   + actor                                             }
          ]
        },
        {
          type: "actions",
          elements: [
            {
              type: "button",
              text: { type: "plain_text", text: "Open Masterlist", emoji: true },
              url:  ScriptApp.getService().getUrl() + "?page=masterlist",
              style: "primary"
            }
          ]
        },
        {
          type: "context",
          elements: [
            {
              type: "mrkdwn",
              text: "\uD83D\uDD50 " + new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) + " PHT"
            }
          ]
        }
      ], "#3b82f6"); // blue sidebar
    } catch (slackErr) {
      console.error("Slack notify failed (non-blocking):", slackErr);
    }
    // ────────────────────────────────────────────────────

    return { ok: true, data: true };
  } catch (e) {
    console.error("masterlist_add error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


function canAddAsset(email) {
  return canAccessPage("assetreport", email) && canDoAction("assetreport.add", email);
}

function requireAssetReportAdd_(email) {
  if (!canAccessPage("assetreport", email)) {
    throw new Error("You do not have permission to access Asset Report.");
  }
  if (!canDoAction("assetreport.add", email)) {
    throw new Error("You do not have permission to add assets.");
  }
}

function requireMasterlistEdit_(email) {
  if (!canAccessPage("assetreport", email)) {
    throw new Error("You do not have permission to access Asset Report.");
  }
  if (!canDoAction("masterlist.edit", email)) {
    throw new Error("You do not have permission to edit assets.");
  }
}

function requireMasterlistDelete_(email) {
  if (!canAccessPage("assetreport", email)) {
    throw new Error("You do not have permission to access Asset Report.");
  }
  if (!canDoAction("masterlist.delete", email)) {
    throw new Error("You do not have permission to delete assets.");
  }
}


function getAssetReportPermissions() {
  var email = Session.getActiveUser().getEmail();
  return {
    canView:   canAccessPage("assetreport", email),
    canAdd:    canDoAction("assetreport.add", email),
    canEdit:   canDoAction("masterlist.edit", email),
    canDelete: canDoAction("masterlist.delete", email),
    canExport: canDoAction("assetreport.export", email)
  };
}


function masterlist_delete(serialTag) {
  try {
    var email = getCurrentUserEmail_();
    requireMasterlistDelete_(email);

    var before = null;
    try {
      var g = masterlist_get(serialTag);
      if (g && g.ok) before = g.data;
    } catch (ignErr) {}

    var ok = deleteAssetFromMasterlist(serialTag);

    if (ok) {
      _mlBustCache_();
     assetLogWrite({
        action:    "DELETE",
        serialTag: serialTag,
        entity:    "Masterlist",
        field:     "*",
        oldValue:  before
          ? [
              String(before.type      || ""),
              String(before.model     || ""),
              String(before.serialTag || ""),
              String(before.assignee  || "")
            ].filter(Boolean).join(" | ")
          : serialTag,
        newValue:  "DELETED",
        details: {
          before:       before,
          affectedUser: String((before && before.assignee) || "").trim()
        }
      });

      // ── Slack notification ──────────────────────────────
      try {
        var actor = email;
        var b     = before || {};

        sendSlackNotification_(null, [
          {
            type: "section",
            text: {
              type: "mrkdwn",
              text: "*\uD83D\uDDD1\uFE0F Asset Deleted from Manila IT Inventory*"
            }
          },
          {
            type: "section",
            fields: [
              { type: "mrkdwn", text: "*Serial Tag:*\n"  + (String(b.serialTag    || serialTag || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Type:*\n"        + (String(b.type         || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Model:*\n"       + (String(b.model        || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Serial No:*\n"   + (String(b.serialNumber || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Status:*\n"      + (String(b.status       || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Assignee:*\n"    + (String(b.assignee     || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Department:*\n"  + (String(b.department   || "").trim() || "\u2014") },
              { type: "mrkdwn", text: "*Deleted by:*\n"  + actor                                             }
            ]
          },
          {
            type: "context",
            elements: [
              {
                type: "mrkdwn",
                text: "\uD83D\uDD50 " + new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) + " PHT"
              }
            ]
          }
        ], "#ef4444"); // red sidebar
      } catch (slackErr) {
        console.error("Slack notify failed (non-blocking):", slackErr);
      }
      // ────────────────────────────────────────────────────
    }

    return { ok: !!ok, data: !!ok, error: ok ? null : "Serial Tag not found." };
  } catch (e) {
    console.error("masterlist_delete error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


function masterlist_clearAssignment(serialTag, actorEmail, chosenStatus) {
  try {
    var email = getCurrentUserEmail_();
    requireMasterlistEdit_(email);

    // Use chosenStatus from modal, default to In-stock
    var finalStatus = String(chosenStatus || "In-stock").trim() || "In-stock";

    var before = null;
    try {
      var g = masterlist_get(serialTag);
      if (g && g.ok && g.data) {
        before = {};
        var bKeys = Object.keys(g.data);
        for (var b = 0; b < bKeys.length; b++) {
          before[bKeys[b]] = g.data[bKeys[b]];
        }
      }
    } catch (e) {
      console.error("Failed to get before state:", e);
    }

    // ── Stamp RETURNED watermark BEFORE clearing AF link ────────
// Must run before clearAssetFieldsInSheet wipes accountabilityForm
var returnedAfUrl = '';
try {
  var wResult = _stampReturnedWatermarkOnAf_(serialTag);
  if (wResult.ok) {
    returnedAfUrl = wResult.url;
    console.log('[WATERMARK] RETURNED AF generated:', returnedAfUrl);
  } else if (wResult.reason !== 'no_af_link') {
    console.warn('[WATERMARK] Stamp failed:', wResult.error);
  }
} catch (watermarkErr) {
  console.error('[WATERMARK] Error (non-blocking):', watermarkErr);
}

var ok = clearAssetFieldsInSheet(serialTag, [
  "assignee", "email", "department", "accessories",
  "accountabilityForm", "signedAf", "afStatus", "afSignedDate"
], finalStatus);

var taskResult = null;

if (ok) {
  _mlBustCache_();
  // ── Task creation ───────────────────────────────────
  try {
    taskResult = _createTaskForClearingUser_(
          before || { serialTag: serialTag },
          actorEmail
        );
        console.log("clearAssignment taskResult:", JSON.stringify(taskResult));
      } catch (taskErr) {
        console.error("clearAssignment task creation failed:", taskErr);
      }

      // ── FYI email to former assignee ────────────────────
      if (before && String(before.email || "").trim()) {
        try {
          _sendAssetClearedEmail_(before, actorEmail);
        } catch (mailErr) {
          console.error("clearAssignment FYI email failed:", mailErr);
        }
      }
      // ── Per-field audit entries (Asset Logs) ────────────
      // Log each cleared field individually so the audit trail
      // shows exactly what changed (e.g. assignee: "Juan" → "")
      // rather than one opaque CLEAR_ASSIGNMENT entry.
      var CLEARED_FIELDS = [
        "assignee", "email", "department", "accessories",
        "accountabilityForm", "signedAf", "afStatus", "afSignedDate"
      ];

      var affectedUser = String((before && before.assignee) || "").trim();

      for (var cf = 0; cf < CLEARED_FIELDS.length; cf++) {
        var cfKey    = CLEARED_FIELDS[cf];
        var cfOldVal = String((before && before[cfKey]) || "").trim();
        if (!cfOldVal) continue; // skip fields already blank — no noise

        assetLogWrite({
          action:    "CLEAR_ASSIGNMENT",
          serialTag: serialTag,
          entity:    "Masterlist",
          field:     cfKey,
          oldValue:  cfOldVal,
          newValue:  "",
          details: {
            clearedBy:    String(actorEmail || "").trim(),
            finalStatus:  finalStatus,
            affectedUser: affectedUser
          }
        });
      }

      // One entry for the status change itself
      assetLogWrite({
        action:    "CLEAR_ASSIGNMENT",
        serialTag: serialTag,
        entity:    "Masterlist",
        field:     "status",
        oldValue:  String((before && before.status) || "").trim(),
        newValue:  finalStatus,
        details: {
          clearedBy:    String(actorEmail || "").trim(),
          affectedUser: affectedUser,
          taskResult:   taskResult
        }
      });

      // ── Legacy auditWrite (AuditLogs sheet) ────────────
      // Kept for backward compatibility with the admin audit page.
      auditWrite({
        action:    "CLEAR_ASSIGNMENT",
        serialTag: serialTag,
        entity:    "Masterlist",
        field:     "assignment",
        oldValue:  String((before && before.assignee) || "").trim() || "",
        newValue:  finalStatus,
        details: {
          before:       before,
          clearedBy:    String(actorEmail || "").trim(),
          finalStatus:  finalStatus,
          affectedUser: affectedUser,
          taskResult:   taskResult
        }
      });

      // ── Slack notification ──────────────────────────────
      try {
        var cleared          = before || {};
        var actorClear       = String(actorEmail || "").trim() || "\u2014";
        var finalStatusLC    = finalStatus.toLowerCase();
        var clearTagVal      = String(cleared.serialTag  || serialTag || "").trim() || "\u2014";
        var clearTypeVal     = String(cleared.type       || "").trim()              || "\u2014";
        var clearModelVal    = String(cleared.model      || "").trim()              || "\u2014";
        var clearAssigneeVal = String(cleared.assignee   || "").trim()              || "\u2014";
        var clearDeptVal     = String(cleared.department || "").trim()              || "\u2014";

        var clearTitle;
        var clearColor;

        if (finalStatusLC === "return") {
          clearTitle = "*\u21A9\uFE0F Asset Returned in Manila IT Inventory*";
          clearColor = "#6366f1";
        } else if (finalStatusLC === "offboard - return") {
          clearTitle = "*\uD83D\uDEAA Asset Offboard - Return in Manila IT Inventory*";
          clearColor = "#f59e0b";
        } else {
          clearTitle = "*\uD83D\uDD13 Asset Assignment Cleared in Manila IT Inventory*";
          clearColor = "#f59e0b";
        }

        sendSlackNotification_(null, [
          {
            type: "section",
            text: { type: "mrkdwn", text: clearTitle }
          },
          {
            type: "section",
            fields: [
              { type: "mrkdwn", text: "*Serial Tag:*\n"      + clearTagVal      },
              { type: "mrkdwn", text: "*Type:*\n"            + clearTypeVal     },
              { type: "mrkdwn", text: "*Model:*\n"           + clearModelVal    },
              { type: "mrkdwn", text: "*Was Assigned to:*\n" + clearAssigneeVal },
              { type: "mrkdwn", text: "*Department:*\n"      + clearDeptVal     },
              { type: "mrkdwn", text: "*New Status:*\n"      + finalStatus      },
              { type: "mrkdwn", text: "*Cleared by:*\n"      + actorClear       },
              { type: "mrkdwn", text: "*Cleared via:*\n"     + "Clear Assignment button" }
            ]
          },
          {
            type: "actions",
            elements: [
              {
                type: "button",
                text: { type: "plain_text", text: "Open Masterlist", emoji: true },
                url:  ScriptApp.getService().getUrl() + "?page=masterlist",
                style: "primary"
              }
            ]
          },
          {
            type: "context",
            elements: [
              {
                type: "mrkdwn",
                text: "\uD83D\uDD50 " + new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) + " PHT"
              }
            ]
          }
        ], clearColor);
      } catch (slackErr) {
        console.error("Slack notify failed (non-blocking):", slackErr);
      }
      // ────────────────────────────────────────────────────
    }

    return { ok: !!ok, data: !!ok, taskResult: taskResult };
  } catch (e) {
    console.error("masterlist_clearAssignment error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


function masterlist_generateAccountabilityForm(serialTag) {
  try {
    var url = generateAccountabilityForm(serialTag);
    _mlBustCache_();
    return { ok: true, data: url };
  } catch (e) {
    console.error("masterlist_generateAccountabilityForm error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


/* ============================================================
 * EDIT SUPPORT
 * ============================================================ */

function masterlist_get(serialTag) {
  try {
    var sh  = _mlSheet_();
    var row = _mlFindRowBySerialTag_(sh, serialTag);
    if (row < 0) return { ok: false, error: "Serial Tag \"" + serialTag + "\" not found." };
    return { ok: true, data: _mlRowToObject_(sh, row) };
  } catch (e) {
    console.error("masterlist_get error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

function masterlist_update(payload) {
  try {
    var email = getCurrentUserEmail_();
    requireMasterlistEdit_(email);

    var p           = payload || {};
    var originalTag = String(p.originalSerialTag || "").trim();
    if (!originalTag) return { ok: false, error: "Missing originalSerialTag." };

    var beforeResp = masterlist_get(originalTag);
    var before     = (beforeResp && beforeResp.ok) ? (beforeResp.data || {}) : {};

    var res = (function doUpdate_() {
      var newTag = String(p.serialTag || "").trim();
      if (!newTag) return { ok: false, error: "Serial Tag is required." };

      var sh  = _mlSheet_();
      var row = _mlFindRowBySerialTag_(sh, originalTag);
      if (row < 0) return { ok: false, error: "Serial Tag \"" + originalTag + "\" not found." };

      if (newTag !== originalTag) {
        var dupRow = _mlFindRowBySerialTag_(sh, newTag);
        if (dupRow > 0) return { ok: false, error: "Serial Tag \"" + newTag + "\" already exists." };
      }

      var hm   = _mlHeaderMap_(sh);
      var map  = hm.map;
      var keys = [
        "type", "model", "serialTag", "serialNumber", "ram",
        "purchaseDate", "endOfWarranty", "status",
        "assignee", "email", "department", "accessories",
        "accountabilityForm", "signedAf", "afStatus", "afSignedDate", "notes"
      ];

      for (var i = 0; i < keys.length; i++) {
        var k   = keys[i];
        var col = map[k];
        if (!col) continue;

        var value = (p[k] == null ? "" : p[k]);
        if (k === "purchaseDate" || k === "endOfWarranty" || k === "afSignedDate") {
          value = _toDateOrBlank_(value);
        } else {
          value = String(value).trim();
        }

        sh.getRange(row, col).setValue(value);
      }
      _mlBustCache_(); 

      return { ok: true, data: true };
    })();

    if (!res.ok) return res;

    var afterResp = masterlist_get(String(p.serialTag || originalTag));
    var after     = (afterResp && afterResp.ok) ? (afterResp.data || {}) : {};

    var auditKeys = [
      "type", "model", "serialTag", "serialNumber", "ram",
      "purchaseDate", "endOfWarranty", "status",
      "assignee", "email", "department", "accessories",
      "accountabilityForm", "signedAf", "afStatus", "afSignedDate"
    ];

    for (var ai = 0; ai < auditKeys.length; ai++) {
      var ak  = auditKeys[ai];
      var ov  = (before[ak] == null ? "" : before[ak]);
      var nv  = (after[ak]  == null ? "" : after[ak]);
      if (String(ov) !== String(nv)) {
        assetLogWrite({
          action:    "UPDATE",
          serialTag: after.serialTag || originalTag,
          entity:    "Masterlist",
          field:     ak,
          oldValue:  ov,
          newValue:  nv,
          details: {
            originalSerialTag: originalTag,
            affectedUser:      String(before.assignee || "").trim()
          }
        });
      }
    }

    var statusBefore  = String(before.status || "").trim().toLowerCase();
    var statusAfter   = String(after.status  || "").trim().toLowerCase();
    var payloadStatus = String(p.status      || "").trim().toLowerCase();

    // ── Slack: status changed to Assigned ──────────────────────────
    if (
      (statusAfter === "assigned" || payloadStatus === "assigned") &&
      statusBefore !== "assigned"
    ) {
      try {
        var actor       = email;
        var notifTag    = String(after.serialTag    || p.serialTag    || "").trim() || "\u2014";
        var notifType   = String(after.type         || p.type         || "").trim() || "\u2014";
        var notifModel  = String(after.model        || p.model        || "").trim() || "\u2014";
        var notifAssign = String(after.assignee     || p.assignee     || "").trim() || "\u2014";
        var notifDept   = String(after.department   || p.department   || "").trim() || "\u2014";
        var notifEmail  = String(after.email        || p.email        || "").trim() || "\u2014";
        var notifSerial = String(after.serialNumber || p.serialNumber || "").trim() || "\u2014";

        sendSlackNotification_(null, [
          {
            type: "section",
            text: { type: "mrkdwn", text: "*\uD83D\uDCCB Asset Assigned in Manila IT Inventory*" }
          },
          {
            type: "section",
            fields: [
              { type: "mrkdwn", text: "*Serial Tag:*\n"  + notifTag    },
              { type: "mrkdwn", text: "*Type:*\n"        + notifType   },
              { type: "mrkdwn", text: "*Model:*\n"       + notifModel  },
              { type: "mrkdwn", text: "*Assigned to:*\n" + notifAssign },
              { type: "mrkdwn", text: "*Department:*\n"  + notifDept   },
              { type: "mrkdwn", text: "*Email:*\n"       + notifEmail  },
              { type: "mrkdwn", text: "*Serial No:*\n"   + notifSerial },
              { type: "mrkdwn", text: "*Updated by:*\n"  + actor       },
              { type: "mrkdwn", text: "*Updated via:*\n" + "Edit Asset modal" }
            ]
          },
          {
            type: "actions",
            elements: [
              {
                type: "button",
                text: { type: "plain_text", text: "Open Masterlist", emoji: true },
                url:  ScriptApp.getService().getUrl() + "?page=masterlist",
                style: "primary"
              }
            ]
          },
          {
            type: "context",
            elements: [
              {
                type: "mrkdwn",
                text: "\uD83D\uDD50 " + new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) + " PHT"
              }
            ]
          }
        ], "#22c55e");
      } catch (slackErr) {
        console.error("Slack notify failed (non-blocking):", slackErr);
      }
    }

    // ── Return / Offboard-Return: email + Slack ─────────────────────
    var isReturn   = (statusAfter === "return"            || payloadStatus === "return");
    var isOffboard = (statusAfter === "offboard - return" || payloadStatus === "offboard - return");

    if ((isReturn || isOffboard) && statusBefore !== statusAfter) {

      // Declare actor once — used by both email and Slack below
      var actorReturn = email;

      // ── Return email to assignee (non-blocking) ─────────────────
      try {
        _sendAssetReturnedEmail_(before, after, actorReturn, isOffboard);
      } catch (mailErr) {
        console.error("Return email failed (non-blocking):", mailErr);
      }

      // ── Slack notification ──────────────────────────────────────
      try {
        var retifTag    = String(after.serialTag   || p.serialTag   || "").trim() || "\u2014";
        var retifType   = String(after.type        || p.type        || "").trim() || "\u2014";
        var retifModel  = String(after.model       || p.model       || "").trim() || "\u2014";
        var retifAssign = String(before.assignee   || "").trim()                  || "\u2014";
        var retifDept   = String(before.department || "").trim()                  || "\u2014";
        var retifStatus = isOffboard ? "Offboard - Return" : "Return";

        sendSlackNotification_(null, [
          {
            type: "section",
            text: {
              type: "mrkdwn",
              text: isOffboard
                ? "*\uD83D\uDEAA Asset Offboard - Return in Manila IT Inventory*"
                : "*\u21A9\uFE0F Asset Returned in Manila IT Inventory*"
            }
          },
          {
            type: "section",
            fields: [
              { type: "mrkdwn", text: "*Serial Tag:*\n"      + retifTag    },
              { type: "mrkdwn", text: "*Type:*\n"            + retifType   },
              { type: "mrkdwn", text: "*Model:*\n"           + retifModel  },
              { type: "mrkdwn", text: "*Was Assigned to:*\n" + retifAssign },
              { type: "mrkdwn", text: "*Department:*\n"      + retifDept   },
              { type: "mrkdwn", text: "*New Status:*\n"      + retifStatus },
              { type: "mrkdwn", text: "*Updated by:*\n"      + actorReturn },
              { type: "mrkdwn", text: "*Updated via:*\n"     + "Edit Asset modal" }
            ]
          },
          {
            type: "actions",
            elements: [
              {
                type: "button",
                text: { type: "plain_text", text: "Open Masterlist", emoji: true },
                url:  ScriptApp.getService().getUrl() + "?page=masterlist",
                style: "primary"
              }
            ]
          },
          {
            type: "context",
            elements: [
              {
                type: "mrkdwn",
                text: "\uD83D\uDD50 " + new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) + " PHT"
              }
            ]
          }
        ], isOffboard ? "#f59e0b" : "#6366f1");
      } catch (slackErr) {
        console.error("Slack notify failed (non-blocking):", slackErr);
      }
    }
    // ───────────────────────────────────────────────────────────────

    return res;
  } catch (e) {
    console.error("masterlist_update error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


/* ============================================================
 * EMPLOYEES DROPDOWN
 * ============================================================ */

function employees_dropdown_list() {
  try {
    // ── Cache check ──────────────────────────────
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(EMPLOYEES_CACHE_KEY_);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        if (parsed && parsed.ok) return parsed;
      } catch (_) {}
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("Employees");
    if (!sh) return { ok: false, error: "Employees sheet not found." };

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) return { ok: true, data: [] };

    var rawHeaders = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var headers    = [];
    for (var h = 0; h < rawHeaders.length; h++) {
      headers.push(String(rawHeaders[h] || "").trim());
    }

    var nameCol  = headers.indexOf("Display Name") + 1;
    var emailCol = headers.indexOf("Email")        + 1;
    var deptCol  = headers.indexOf("Department")   + 1;

    if (!nameCol)  return { ok: false, error: "Employees header \"Display Name\" not found." };
    if (!emailCol) return { ok: false, error: "Employees header \"Email\" not found." };
    if (!deptCol)  return { ok: false, error: "Employees header \"Department\" not found." };

    // ── One batch read instead of three separate calls ──
    var allData  = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    var nameIdx  = nameCol  - 1;
    var emailIdx = emailCol - 1;
    var deptIdx  = deptCol  - 1;

    var out = [];
    for (var i = 0; i < allData.length; i++) {
      var displayName = String(allData[i][nameIdx]  || "").trim();
      var email       = String(allData[i][emailIdx] || "").trim();
      var department  = String(allData[i][deptIdx]  || "").trim();
      if (!displayName) continue;
      out.push({ displayName: displayName, email: email, department: department });
    }

    out.sort(function(a, b) {
      return a.displayName.localeCompare(b.displayName);
    });

    var result = { ok: true, data: out };

    // ── Cache for 10 minutes ────────────────────
    try {
      var json = JSON.stringify(result);
      if (json.length < 90000) cache.put(EMPLOYEES_CACHE_KEY_, json, EMPLOYEES_CACHE_TTL_);
    } catch (_) {}

    return result;
  } catch (e) {
    console.error("employees_dropdown_list error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


/* ============================================================
 * SERVER-SIDE PAGED LIST
 * ============================================================ */

function masterlist_list_paged(options) {
  try {
    var opt            = options || {};
    var page           = Math.max(1, parseInt(opt.page, 10) || 1);
    var pageSize       = Math.max(1, Math.min(200, parseInt(opt.pageSize, 10) || 25));
    var query          = String(opt.query      || "").trim().toLowerCase();
    var typeFilter     = String(opt.type       || "").trim().toLowerCase();
    var statusFilter   = String(opt.status     || "").trim().toLowerCase();
    var deptFilter     = String(opt.department || "").trim().toLowerCase();
    var modelFilter    = String(opt.model      || "").trim().toLowerCase();
    var sortKey        = String(opt.sortKey    || "serialTag").trim();
    var asc            = (opt.asc !== false);

    var values = _mlGetDisplayValues_();
if (values.length < 2) {
  return { ok: true, data: [], total: 0, page: page, pageSize: pageSize };
}

    var col = _masterlistColMapFromHeaderRow_(values[0]).col;
    if (col.serialTag < 0) throw new Error("Header \"Serial Tag\" not found.");

    var allRows = [];
    for (var i = 1; i < values.length; i++) {
      var obj = _masterlistObjectFromRow_(values[i], col, false);
      if (String(obj.serialTag || "").trim()) allRows.push(obj);
    }

    var filtered = allRows;

    if (typeFilter) {
      var tmp = [];
      for (var t = 0; t < filtered.length; t++) {
        if (String(filtered[t].type || "").trim().toLowerCase() === typeFilter) tmp.push(filtered[t]);
      }
      filtered = tmp;
    }

    if (statusFilter) {
  var tmpS = [];
  var normTarget = normalizeStatusCanonical_(statusFilter).toLowerCase();
  for (var s = 0; s < filtered.length; s++) {
    var sv = String(filtered[s].status || "").trim().toLowerCase();
    if (normalizeStatusCanonical_(sv).toLowerCase() === normTarget) tmpS.push(filtered[s]);
  }
  filtered = tmpS;
}

    if (deptFilter) {
      var tmpD = [];
      for (var d = 0; d < filtered.length; d++) {
        if (String(filtered[d].department || "").trim().toLowerCase() === deptFilter) tmpD.push(filtered[d]);
      }
      filtered = tmpD;
    }

    if (modelFilter) {
      var tmpM = [];
      for (var m = 0; m < filtered.length; m++) {
        if (String(filtered[m].model || "").trim().toLowerCase() === modelFilter) tmpM.push(filtered[m]);
      }
      filtered = tmpM;
    }

    if (query) {
      var tmpQ = [];
      for (var q = 0; q < filtered.length; q++) {
        var x    = filtered[q];
        var blob = [
  x.type, x.model, x.serialTag, x.serialNumber, x.ram,
  x.purchaseDate, x.endOfWarranty, x.status,
  x.assignee, x.email, x.department, x.accessories,
  x.accountabilityForm, x.signedAf, x.afStatus, x.afSignedDate,
  x.procurementRef   // ← add this
].join(" ").toLowerCase();
        if (blob.indexOf(query) >= 0) tmpQ.push(x);
      }
      filtered = tmpQ;
    }

    filtered.sort(function(a, b) {
      var A = serverComparable_(a, sortKey);
      var B = serverComparable_(b, sortKey);
      if (A < B) return asc ? -1 : 1;
      if (A > B) return asc ? 1  : -1;
      return 0;
    });

    var total  = filtered.length;
    var start  = (page - 1) * pageSize;
    var paged  = filtered.slice(start, start + pageSize);

    return { ok: true, data: paged, total: total, page: page, pageSize: pageSize };
  } catch (err) {
    console.error("masterlist_list_paged error:", err);
    return { ok: false, error: (err && err.message) || String(err) };
  }
}


/* ============================================================
 * EXPORT ALL FILTERED
 * ============================================================ */

function masterlist_export_all(req) {
  try {
    req            = req || {};
    var query      = String(req.query      || "").toLowerCase().trim();
    var type       = String(req.type       || "").toLowerCase().trim();
    var status     = String(req.status     || "").toLowerCase().trim();
    var department = String(req.department || "").toLowerCase().trim();
    var model      = String(req.model      || "").toLowerCase().trim();
    var sortKey    = String(req.sortKey    || "serialTag").trim();
    var asc        = (req.asc !== false);

    var sh     = _mlSheet_();
    var values = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, data: "" };

    var col = _masterlistColMapFromHeaderRow_(values[0]).col;

    var allRows = [];
    for (var i = 1; i < values.length; i++) {
      var obj = _masterlistObjectFromRow_(values[i], col, false);
      if (String(obj.serialTag || "").trim()) allRows.push(obj);
    }

    var filtered = allRows;

    if (type) {
      var tmpT = [];
      for (var t = 0; t < filtered.length; t++) {
        if (String(filtered[t].type || "").trim().toLowerCase() === type) tmpT.push(filtered[t]);
      }
      filtered = tmpT;
    }

    if (status) {
      var tmpS = [];
      for (var s = 0; s < filtered.length; s++) {
        var sv = String(filtered[s].status || "").trim().toLowerCase();
        var matchS = false;
        matchS = (normalizeStatusCanonical_(sv).toLowerCase() === normalizeStatusCanonical_(status).toLowerCase());
        if (matchS) tmpS.push(filtered[s]);
      }
      filtered = tmpS;
    }

    if (department) {
      var tmpD = [];
      for (var d = 0; d < filtered.length; d++) {
        if (String(filtered[d].department || "").trim().toLowerCase() === department) tmpD.push(filtered[d]);
      }
      filtered = tmpD;
    }

    if (model) {
      var tmpM = [];
      for (var m = 0; m < filtered.length; m++) {
        if (String(filtered[m].model || "").trim().toLowerCase() === model) tmpM.push(filtered[m]);
      }
      filtered = tmpM;
    }

    if (query) {
      var tmpQ = [];
      for (var q = 0; q < filtered.length; q++) {
        var x   = filtered[q];
        var hay = [
          x.type, x.model, x.serialTag, x.serialNumber, x.ram,
          x.purchaseDate, x.endOfWarranty, x.status,
          x.assignee, x.email, x.department, x.accessories,
          x.accountabilityForm, x.signedAf, x.afStatus, x.afSignedDate
        ].join(" ").toLowerCase();
        if (hay.indexOf(query) >= 0) tmpQ.push(x);
      }
      filtered = tmpQ;
    }

    filtered.sort(function(a, b) {
      var A = serverComparable_(a, sortKey);
      var B = serverComparable_(b, sortKey);
      if (A < B) return asc ? -1 : 1;
      if (A > B) return asc ? 1  : -1;
      return 0;
    });

    var exportConfig = {
      title:   String(req.title || "Hardware Masterlist (All Results)"),
      headers: Array.isArray(req.headers) ? req.headers : [],
      keys:    Array.isArray(req.keys)    ? req.keys    : [],
      data:    filtered
    };

    var url = generateMasterlistExport(exportConfig);
    return { ok: true, data: url };
  } catch (err) {
    console.error("masterlist_export_all error:", err);
    return { ok: false, error: (err && err.message) || String(err) };
  }
}


/* ============================================================
 * RESOLVE EMAIL FROM EMPLOYEE
 * ============================================================ */

function _resolveEmailFromEmployee_(assigneeName) {
  var name = String(assigneeName || "").trim();
  if (!name) return "";

  try {
    var list = employees_dropdown_list();
    if (!list.ok) return "";

    var data = list.data || [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i].displayName || "").trim() === name) {
        return String(data[i].email || "").trim();
      }
    }
    return "";
  } catch (e) {
    console.error("resolve email error:", e);
    return "";
  }
}


/* ============================================================
 * SEND ASSET CLEARED EMAIL
 * ============================================================ */

function _sendAssetClearedEmail_(asset, clearedByEmail) {
  if (!asset) return false;

  var to = String(asset.email || "").trim();
  if (!to) to = _resolveEmailFromEmployee_(asset.assignee);

  var assignee     = String(asset.assignee     || "").trim() || "User";
  var serialTag    = String(asset.serialTag    || "").trim();
  var type         = String(asset.type         || "").trim() || "-";
  var model        = String(asset.model        || "").trim() || "-";
  var serialNumber = String(asset.serialNumber || "").trim() || "-";
  var accessories  = String(asset.accessories  || "").trim() || "None";
  var clearedBy    = String(clearedByEmail      || "Manila IT").trim();

  if (!to || !serialTag) {
    console.warn("Email NOT sent — missing:", { to: to, serialTag: serialTag });
    return false;
  }

  var FROM_ALIAS  = "manila-it@fbgphilippines.com";
  var SENDER_NAME = "FBG Manila IT";
  var CC          = "manila-it@fbgphilippines.com";
  var subject     = "Asset Assignment Cleared - " + serialTag;

  var clearedDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MMMM d, yyyy 'at' h:mm a"
  );

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g,  "&amp;")
      .replace(/</g,  "&lt;")
      .replace(/>/g,  "&gt;")
      .replace(/"/g,  "&quot;")
      .replace(/'/g,  "&#39;");
  }

  // Build accessories display — chips if there are items, italic placeholder otherwise.
var accList = (accessories === "None")
  ? []
  : String(accessories || "").split(",").map(function(s) { return s.trim(); }).filter(Boolean);

var accessoriesHtml;
if (!accList.length) {
  accessoriesHtml = '<span style="color:#9a9aa0;font-style:italic;font-weight:400;">None recorded</span>';
} else {
  accessoriesHtml = accList.map(function(item) {
    return (
      '<span style="display:inline-block;margin:0 6px 6px 0;padding:5px 11px;' +
        'background:rgba(10,132,255,0.14);border:1px solid rgba(10,132,255,0.30);' +
        'border-radius:999px;font-size:12px;font-weight:700;color:#60a5fa;">' +
        '&#10003; ' + esc(item) +
      '</span>'
    );
  }).join('');
}

  var htmlBody =
    "<div style=\"margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;\">" +
      "<div style=\"max-width:560px;margin:28px auto;padding:0 16px;\">" +
        "<div style=\"background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);\">" +
          "<div style=\"padding:18px 20px;background:linear-gradient(135deg,#0A84FF,#2997FF);color:#fff;\">" +
            "<div style=\"font-size:13px;opacity:.92;\">Manila IT Inventory</div>" +
            "<div style=\"font-size:18px;font-weight:800;margin-top:4px;\">Asset Assignment Cleared</div>" +
          "</div>" +
          "<div style=\"padding:20px;color:#f5f5f7;\">" +
            "<p style=\"margin:0 0 12px;font-size:14px;line-height:1.5;color:#d1d1d6;\">" +
              "Hi <b style=\"color:#fff;\">" + esc(assignee) + "</b>, this is an FYI that the asset below " +
              "has been cleared from your assignment record." +
            "</p>" +
            "<div style=\"margin:14px 0;padding:16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;\">" +
              "<table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" width=\"100%\" style=\"border-collapse:collapse;\">" +
                "<tr>" +
                  "<td style=\"width:50%;padding:0 12px 14px 0;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Serial Tag</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(serialTag) + "</div>" +
                  "</td>" +
                  "<td style=\"width:50%;padding:0 0 14px 12px;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Type</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(type) + "</div>" +
                  "</td>" +
                "</tr>" +
                "<tr>" +
                  "<td style=\"width:50%;padding:0 12px 14px 0;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Model</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(model) + "</div>" +
                  "</td>" +
                  "<td style=\"width:50%;padding:0 0 14px 12px;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Serial Number</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(serialNumber) + "</div>" +
                  "</td>" +
                "</tr>" +
                "<tr>" +
  "<td colspan=\"2\" style=\"padding:0 0 14px 0;vertical-align:top;\">" +
    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;\">Accessories Returned</div>" +
    "<div style=\"font-size:14px;color:#ffffff;line-height:1.6;\">" + accessoriesHtml + "</div>" +
  "</td>" +
"</tr>" +
              "</table>" +
            "</div>" +
            "<div style=\"margin:14px 0 0;padding:13px 16px;background:rgba(10,132,255,0.10);border:1px solid rgba(10,132,255,0.25);border-radius:10px;\">" +
              "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;\">Cleared by</div>" +
              "<div style=\"font-size:13px;color:#ffffff;font-weight:700;\">" + esc(clearedBy) + "</div>" +
              "<div style=\"font-size:11px;color:#9a9aa0;margin-top:3px;\">" + esc(clearedDate) + "</div>" +
            "</div>" +
            "<p style=\"margin:14px 0 0;font-size:12px;color:#9a9aa0;line-height:1.4;\">" +
              "This is for your information only. No action is required unless instructed by IT." +
            "</p>" +
          "</div>" +
        "</div>" +
        "<div style=\"text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;\">" +
          "&#169; " + new Date().getFullYear() + " Manila IT Inventory" +
        "</div>" +
      "</div>" +
    "</div>";

  var plainBody =
    "Hi " + assignee + ",\n\n" +
    "This is an FYI that the asset below has been cleared from your assignment record.\n\n" +
    "Serial Tag:    " + serialTag    + "\n" +
    "Type:          " + type         + "\n" +
    "Model:         " + model        + "\n" +
    "Serial Number: " + serialNumber + "\n" +
    "Accessories:   " + (accList.length ? accList.join(", ") : "None recorded") + "\n\n" +
    "Cleared by:    " + clearedBy    + "\n" +
    "Cleared on:    " + clearedDate  + "\n\n" +
    "No action is required unless instructed by IT.\n\n" +
    "Thanks,\n" +
    "Manila IT Team";

  try {
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody: htmlBody,
      name:     SENDER_NAME,
      from:     FROM_ALIAS,
      replyTo:  FROM_ALIAS,
      cc:       CC
    });
  } catch (e) {
    console.error("_sendAssetClearedEmail_ alias send failed:", e);
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody: htmlBody,
      name:     SENDER_NAME,
      cc:       CC
    });
  }

  return true;
}


/* ============================================================
 * USER RECORD LOOKUP
 * ============================================================ */

function _getUserRecordByEmail_(email) {
  var target = String(email || "").trim().toLowerCase();
  if (!target) return null;

  try {
    var users = getUsersData() || [];
    for (var i = 0; i < users.length; i++) {
      var u        = users[i] || {};
      var rowEmail = String(u.email || "").trim().toLowerCase();
      if (rowEmail === target) {
        return {
          username: String(u.username || u.name || "").trim(),
          email:    String(u.email    || "").trim(),
          role:     String(u.role     || "").trim()
        };
      }
    }
  } catch (e) {
    console.error("_getUserRecordByEmail_ error:", e);
  }

  return null;
}


/* ============================================================
 * CREATE TASK FOR CLEARING USER
 * ============================================================ */

function _createTaskForClearingUser_(asset, actorEmail) {
  var email = String(actorEmail || "").trim();
  if (!email) return { ok: false, error: "Missing actorEmail." };

  var user = _getUserRecordByEmail_(email);
  if (!user || !user.username) {
    return { ok: false, error: "Clearing user not found in Users sheet." };
  }

  var serialTag = String((asset && asset.serialTag) || "").trim();
  var model     = String((asset && asset.model)     || "").trim();
  var type      = String((asset && asset.type)      || "").trim();

  if (!serialTag) return { ok: false, error: "Missing asset serialTag." };

  var sh      = getTasksSheet_();
  var lastRow = sh.getLastRow();

  if (lastRow >= 2) {
    var rows         = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    var hasDuplicate = false;

    for (var i = 0; i < rows.length; i++) {
      var employee   = String(rows[i][0] || "").trim().toLowerCase();
      var taskDesc   = String(rows[i][1] || "").trim().toLowerCase();
      var taskStatus = String(rows[i][4] || "").trim().toLowerCase();

      var sameUser   = (employee === user.username.toLowerCase());
      var sameAsset  = (taskDesc.indexOf(serialTag.toLowerCase()) >= 0);
      var openStatus = (["completed","done","closed","cancelled"].indexOf(taskStatus) < 0);

      if (sameUser && sameAsset && openStatus) {
        hasDuplicate = true;
        break;
      }
    }

    if (hasDuplicate) {
      return { ok: true, skipped: true, reason: "Duplicate open task already exists." };
    }
  }

  var dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + 30);

  var typeModel  = "";
  if (type || model) {
    var parts = [];
    if (type)  parts.push(type);
    if (model) parts.push(model);
    typeModel = " (" + parts.join(" - ") + ")";
  }

  var taskPayload = {
    employee:        user.username,
    taskDescription: "Pending wipe after clearance: " + serialTag + typeModel,
    dateRequired:    Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    dateSubmitted:   "",
    status:          "Pending",
    note:            "Auto-created after clearing assignment for asset " + serialTag + "."
  };

  addTask(taskPayload);
  return { ok: true, data: taskPayload };
}


/* ============================================================
 * EXPORT ROWS TO FOLDER
 * ============================================================ */

function exportRowsToFolder(payload) {
  try {
    const email = getCurrentUserEmail_();
    requireAssetReportExport_(email);

    if (!payload || !Array.isArray(payload.rows) || !Array.isArray(payload.headers)) {
      return { ok: false, error: "Invalid payload" };
    }

    var stamp = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"
    );
    var title = (payload.title || "Maintenance Export") + " - " + stamp;

    var ss    = SpreadsheetApp.create(title);
    var sheet = ss.getSheets()[0];

    var values = [payload.headers];
    for (var i = 0; i < payload.rows.length; i++) {
      var r   = payload.rows[i];
      var row = [];
      for (var j = 0; j < r.length; j++) {
        row.push(r[j] == null ? "" : String(r[j]));
      }
      values.push(row);
    }

    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    sheet.getRange(1, 1, 1, values[0].length).setFontWeight("bold");
    sheet.autoResizeColumns(1, values[0].length);

    try {
      var folderName   = "Manila IT Exports";
      var folderSearch = DriveApp.getFoldersByName(folderName);
      var folder       = folderSearch.hasNext()
        ? folderSearch.next()
        : DriveApp.createFolder(folderName);
      DriveApp.getFileById(ss.getId()).moveTo(folder);
    } catch (folderErr) {
      console.warn("exportRowsToFolder: folder move failed:", folderErr && folderErr.message);
    }

    return { ok: true, url: ss.getUrl(), sheetName: title };
  } catch (e) {
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

function _sendAssetReturnedEmail_(before, after, actorEmail, isOffboard) {
  if (!before) return false;

  var to = String(before.email || "").trim();
  if (!to) to = _resolveEmailFromEmployee_(before.assignee);
  if (!to) return false;

  var assignee     = String(before.assignee    || "").trim() || "User";
  var serialTag    = String(before.serialTag   || "").trim();
  var type         = String(before.type        || "").trim() || "-";
  var model        = String(before.model       || "").trim() || "-";
  var serialNumber = String(before.serialNumber|| "").trim() || "-";
  var accessories  = String(before.accessories || "").trim();
  var clearedBy    = String(actorEmail         || "Manila IT").trim();
  var statusLabel  = isOffboard ? "Offboard - Return" : "Return";

  var FROM_ALIAS  = "manila-it@fbgphilippines.com";
  var SENDER_NAME = "FBG Manila IT";
  var CC          = "manila-it@fbgphilippines.com";
  var subject     = (isOffboard ? "Offboard Asset Return" : "Asset Return") + " - " + serialTag;

  var returnedDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MMMM d, yyyy 'at' h:mm a"
  );

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // Build accessories pending checklist
  var accList = accessories
    ? accessories.split(",").map(function(s) { return s.trim(); }).filter(Boolean)
    : [];

  var accHtml = accList.length
    ? accList.map(function(item) {
        return (
          '<div style="display:flex;align-items:center;gap:10px;padding:8px 0;' +
          'border-bottom:1px solid rgba(255,255,255,.06);">' +
            '<span style="width:18px;height:18px;border:2px solid rgba(255,255,255,.25);' +
            'border-radius:4px;display:inline-block;flex-shrink:0;"></span>' +
            '<span style="font-size:14px;color:#f5f5f7;">' + esc(item) + '</span>' +
          '</div>'
        );
      }).join('')
    : '<span style="color:#9a9aa0;font-style:italic;">No accessories on record</span>';

  var headerColor = isOffboard
    ? "linear-gradient(135deg,#f59e0b,#d97706)"
    : "linear-gradient(135deg,#6366f1,#4f46e5)";

  var headerLabel = isOffboard
    ? "Offboard Asset Return"
    : "Asset Return";

  var introText = isOffboard
    ? "Hi <b style=\"color:#fff;\">" + esc(assignee) + "</b>, your asset has been marked for <b>offboard return</b>. Please ensure the following accessories are returned to IT along with the device."
    : "Hi <b style=\"color:#fff;\">" + esc(assignee) + "</b>, your asset has been marked as <b>returned</b>. Please confirm the following accessories were included.";

  var htmlBody =
    "<div style=\"margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;\">" +
      "<div style=\"max-width:560px;margin:28px auto;padding:0 16px;\">" +
        "<div style=\"background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);\">" +
          "<div style=\"padding:18px 20px;background:" + headerColor + ";color:#fff;\">" +
            "<div style=\"font-size:13px;opacity:.92;\">Manila IT Inventory</div>" +
            "<div style=\"font-size:18px;font-weight:800;margin-top:4px;\">" + headerLabel + "</div>" +
          "</div>" +
          "<div style=\"padding:20px;color:#f5f5f7;\">" +
            "<p style=\"margin:0 0 16px;font-size:14px;line-height:1.6;color:#d1d1d6;\">" + introText + "</p>" +

            // Asset details
            "<div style=\"margin:0 0 16px;padding:16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;\">" +
              "<table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" width=\"100%\" style=\"border-collapse:collapse;\">" +
                "<tr>" +
                  "<td style=\"width:50%;padding:0 12px 12px 0;vertical-align:top;\">" +
                    "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Serial Tag</div>" +
                    "<div style=\"font-size:15px;color:#fff;font-weight:700;margin-top:4px;\">" + esc(serialTag) + "</div>" +
                  "</td>" +
                  "<td style=\"width:50%;padding:0 0 12px;vertical-align:top;\">" +
                    "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Type</div>" +
                    "<div style=\"font-size:15px;color:#fff;font-weight:700;margin-top:4px;\">" + esc(type) + "</div>" +
                  "</td>" +
                "</tr>" +
                "<tr>" +
                  "<td style=\"width:50%;padding:0 12px 0 0;vertical-align:top;\">" +
                    "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Model</div>" +
                    "<div style=\"font-size:15px;color:#fff;font-weight:700;margin-top:4px;\">" + esc(model) + "</div>" +
                  "</td>" +
                  "<td style=\"width:50%;padding:0;vertical-align:top;\">" +
                    "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Serial Number</div>" +
                    "<div style=\"font-size:15px;color:#fff;font-weight:700;margin-top:4px;\">" + esc(serialNumber) + "</div>" +
                  "</td>" +
                "</tr>" +
              "</table>" +
            "</div>" +

            // Accessories checklist
            "<div style=\"margin:0 0 16px;\">" +
              "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px;\">Accessories to Return</div>" +
              "<div style=\"padding:4px 0;\">" + accHtml + "</div>" +
            "</div>" +

            // Notice
            "<div style=\"padding:13px 16px;background:rgba(99,102,241,0.10);border:1px solid rgba(99,102,241,0.25);border-radius:10px;\">" +
              "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;\">Processed by</div>" +
              "<div style=\"font-size:13px;color:#fff;font-weight:700;\">" + esc(clearedBy) + "</div>" +
              "<div style=\"font-size:11px;color:#9a9aa0;margin-top:3px;\">" + esc(returnedDate) + "</div>" +
            "</div>" +

            "<p style=\"margin:14px 0 0;font-size:12px;color:#9a9aa0;line-height:1.4;\">" +
              "IT will verify receipt of all accessories and finalize the return. " +
              "A confirmation email will be sent once the process is complete." +
            "</p>" +
          "</div>" +
        "</div>" +
        "<div style=\"text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;\">" +
          "&#169; " + new Date().getFullYear() + " Manila IT Inventory" +
        "</div>" +
      "</div>" +
    "</div>";

  var plainBody =
    "Hi " + assignee + ",\n\n" +
    (isOffboard
      ? "Your asset has been marked for offboard return."
      : "Your asset has been marked as returned.") + "\n\n" +
    "Serial Tag:    " + serialTag    + "\n" +
    "Type:          " + type         + "\n" +
    "Model:         " + model        + "\n" +
    "Serial Number: " + serialNumber + "\n\n" +
    "Accessories to return:\n" +
    (accList.length ? accList.map(function(a) { return "  [ ] " + a; }).join("\n") : "  None on record") + "\n\n" +
    "Processed by:  " + clearedBy   + "\n" +
    "On:            " + returnedDate + "\n\n" +
    "IT will verify receipt and send a confirmation once complete.\n\n" +
    "Thanks,\nManila IT Team";

  try {
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody: htmlBody,
      name:     SENDER_NAME,
      from:     FROM_ALIAS,
      replyTo:  FROM_ALIAS,
      cc:       CC
    });
  } catch (e) {
    console.error("_sendAssetReturnedEmail_ alias failed:", e);
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody: htmlBody,
      name:     SENDER_NAME,
      cc:       CC
    });
  }

  return true;
}

function sendDailyActivityDigest() {
  try {
    var logSh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Asset Logs");
    if (!logSh) return;

    var now       = new Date();
    var yesterday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    var todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    todayStart.setHours(0, 0, 0, 0);

    var data    = logSh.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h || "").trim().toLowerCase(); });
    var tsIdx   = headers.indexOf("timestamp");
    var actIdx  = headers.indexOf("action");
    var tagIdx  = headers.indexOf("serial tag");
    var usrIdx  = headers.indexOf("user");
    var fldIdx  = headers.indexOf("field");
    var oldIdx  = headers.indexOf("old value");
    var newIdx  = headers.indexOf("new value");

    if (tsIdx < 0 || actIdx < 0) {
      Logger.log("[DIGEST] Required columns not found in Asset Logs.");
      return;
    }

    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var rawTs = data[i][tsIdx];
      var d     = (rawTs instanceof Date) ? rawTs : new Date(rawTs);
      if (isNaN(d.getTime())) continue;
      if (d >= yesterday && d < todayStart) rows.push(data[i]);
    }

    if (!rows.length) {
      Logger.log("[DIGEST] No activity yesterday. Email skipped.");
      return;
    }

    // Group by action
    var groups = { ADD: [], UPDATE: [], DELETE: [], CLEAR_ASSIGNMENT: [] };
    for (var j = 0; j < rows.length; j++) {
      var action = String(rows[j][actIdx] || "").trim().toUpperCase();
      if (groups[action]) groups[action].push(rows[j]);
      else groups["UPDATE"].push(rows[j]); // catch-all
    }

    function esc(s) {
      return String(s == null ? "" : s)
        .replace(/&/g,"&amp;").replace(/</g,"&lt;")
        .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
    }

    function makeSection(title, color, items) {
      if (!items.length) return "";
      var rowsHtml = items.map(function(r) {
        return (
          "<tr>" +
          "<td style='padding:7px 10px;border-bottom:1px solid #eee;font-family:monospace;font-size:12px;'>" + esc(r[tagIdx] || "—") + "</td>" +
          "<td style='padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;'>" + esc(r[usrIdx] || "—") + "</td>" +
          "<td style='padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;'>" + esc(r[fldIdx] || "—") + "</td>" +
          "<td style='padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;color:#888;'>" + esc(r[oldIdx] || "") + "</td>" +
          "<td style='padding:7px 10px;border-bottom:1px solid #eee;font-size:12px;font-weight:700;'>" + esc(r[newIdx] || "") + "</td>" +
          "</tr>"
        );
      }).join("");

      return (
        "<h3 style='margin:20px 0 6px;font-size:13px;color:" + color + ";text-transform:uppercase;letter-spacing:.06em;'>" +
        title + " (" + items.length + ")</h3>" +
        "<table style='width:100%;border-collapse:collapse;font-size:13px;'>" +
        "<thead><tr style='background:#f5f5f5;'>" +
        "<th style='padding:7px 10px;text-align:left;font-size:11px;'>Serial Tag</th>" +
        "<th style='padding:7px 10px;text-align:left;font-size:11px;'>Actor</th>" +
        "<th style='padding:7px 10px;text-align:left;font-size:11px;'>Field</th>" +
        "<th style='padding:7px 10px;text-align:left;font-size:11px;'>Old</th>" +
        "<th style='padding:7px 10px;text-align:left;font-size:11px;'>New</th>" +
        "</tr></thead><tbody>" + rowsHtml + "</tbody></table>"
      );
    }

    var dateLabel = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "MMMM d, yyyy");

    var htmlBody =
      "<div style='font-family:Arial,sans-serif;max-width:700px;margin:0 auto;'>" +
      "<div style='background:#1c1c1e;color:#fff;padding:16px 20px;border-radius:8px 8px 0 0;'>" +
      "<div style='font-size:11px;opacity:.7;margin-bottom:4px;'>Manila IT Inventory</div>" +
      "<div style='font-size:18px;font-weight:700;'>Daily Activity Digest</div>" +
      "<div style='font-size:12px;opacity:.6;margin-top:4px;'>" + dateLabel + "</div>" +
      "</div>" +
      "<div style='padding:16px 20px;border:1px solid #eee;border-top:none;border-radius:0 0 8px 8px;'>" +
      "<div style='display:flex;gap:16px;margin-bottom:16px;'>" +
      "<div style='flex:1;text-align:center;padding:12px;background:#f9f9f9;border-radius:8px;'>" +
      "<div style='font-size:22px;font-weight:700;color:#16a34a;'>" + groups.ADD.length + "</div>" +
      "<div style='font-size:11px;color:#888;text-transform:uppercase;'>Added</div>" +
      "</div>" +
      "<div style='flex:1;text-align:center;padding:12px;background:#f9f9f9;border-radius:8px;'>" +
      "<div style='font-size:22px;font-weight:700;color:#2857a4;'>" + groups.UPDATE.length + "</div>" +
      "<div style='font-size:11px;color:#888;text-transform:uppercase;'>Updated</div>" +
      "</div>" +
      "<div style='flex:1;text-align:center;padding:12px;background:#f9f9f9;border-radius:8px;'>" +
      "<div style='font-size:22px;font-weight:700;color:#b45309;'>" + groups.CLEAR_ASSIGNMENT.length + "</div>" +
      "<div style='font-size:11px;color:#888;text-transform:uppercase;'>Cleared</div>" +
      "</div>" +
      "<div style='flex:1;text-align:center;padding:12px;background:#f9f9f9;border-radius:8px;'>" +
      "<div style='font-size:22px;font-weight:700;color:#c94040;'>" + groups.DELETE.length + "</div>" +
      "<div style='font-size:11px;color:#888;text-transform:uppercase;'>Deleted</div>" +
      "</div>" +
      "</div>" +
      makeSection("Added",            "#16a34a", groups.ADD)              +
      makeSection("Updated",          "#2857a4", groups.UPDATE)           +
      makeSection("Cleared",          "#b45309", groups.CLEAR_ASSIGNMENT) +
      makeSection("Deleted",          "#c94040", groups.DELETE)           +
      "<p style='margin-top:16px;font-size:11px;color:#aaa;'>" +
      "This is an automated digest. Total events: " + rows.length + "." +
      "</p>" +
      "</div></div>";

    GmailApp.sendEmail(
      "manila-it@fbgphilippines.com",
      "Daily Inventory Digest — " + dateLabel,
      "Yesterday's inventory activity: " + rows.length + " event(s). See HTML version for details.",
      { htmlBody: htmlBody, name: "Manila IT Inventory", from: "manila-it@fbgphilippines.com" }
    );

    Logger.log("[DIGEST] Sent. Events: " + rows.length);
  } catch (e) {
    Logger.log("[DIGEST] Error: " + (e && e.message ? e.message : String(e)));
    throw e;
  }
}

function createDailyDigestTrigger() {
  _deleteTriggers_("sendDailyActivityDigest");
  ScriptApp.newTrigger("sendDailyActivityDigest")
    .timeBased()
    .everyDays(1)
    .atHour(7) // 7 AM — arrives before the workday
    .create();
  Logger.log("[TRIGGER] Daily digest trigger created.");
}

function _stampReturnedWatermarkOnAf_(serialTag) {
  try {
    var sh      = _mlSheet_();
    var values  = sh.getDataRange().getDisplayValues();
    var headers = values[0].map(function(h) {
      return String(h || '').trim().toLowerCase();
    });

    var cSerial = headers.indexOf('serial tag');
    var cForm   = headers.indexOf('accountability form');

    if (cSerial < 0 || cForm < 0) return { ok: false, error: 'Missing required columns.' };

    var rowIdx = -1;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][cSerial] || '').trim() === serialTag) {
        rowIdx = i;
        break;
      }
    }

    if (rowIdx < 0) return { ok: false, error: 'Serial tag not found.' };

    var afUrl = String(values[rowIdx][cForm] || '').trim();
    if (!afUrl) return { ok: false, reason: 'no_af_link' };

    var fileId = _extractDriveFileId_(afUrl);
    if (!fileId) return { ok: false, error: 'Could not extract file ID.' };

    var origFile = DriveApp.getFileById(fileId);
    var folder   = origFile.getParents().hasNext()
      ? origFile.getParents().next()
      : DriveApp.getRootFolder();
    var origName = origFile.getName().replace(/\.pdf$/i, '');

    var token     = ScriptApp.getOAuthToken();
    var thumbResp = UrlFetchApp.fetch(
      'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1654',
      { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );

    if (thumbResp.getResponseCode() !== 200) {
      return { ok: false, error: 'Thumbnail fetch failed: ' + thumbResp.getResponseCode() };
    }

    var imgBlob = thumbResp.getBlob().setName('af_page.png');

    var pres   = SlidesApp.create('RETURNED_' + serialTag + '_' + Date.now());
    var presId = pres.getId();
    var slide  = pres.getSlides()[0];
    var pageW  = pres.getPageWidth();
    var pageH  = pres.getPageHeight();

    var shapes = slide.getShapes();
    for (var s = 0; s < shapes.length; s++) {
      shapes[s].remove();
    }

    // ── Full-page AF image (original size preserved) ────────────────
    slide.insertImage(imgBlob, 0, 0, pageW, pageH);

    // ── RETURNED stamp ──────────────────────────────────────────────
    // Fit within 55% of page width, positioned in the lower-center
    // area over the signature section — least disruptive to content
    var tbWidth  = pageW * 0.55;
    var tbHeight = pageH * 0.16;
    var tbLeft   = (pageW - tbWidth) / 2;
    var tbTop    = pageH * 0.55; // starts at 55% down the page

    var textBox = slide.insertTextBox('RETURNED', tbLeft, tbTop, tbWidth, tbHeight);
    var tf      = textBox.getText();

    tf.getTextStyle()
      .setFontSize(60)
      .setBold(true)
      .setForegroundColor('#CC0000');

    tf.getParagraphs()[0]
      .getRange()
      .getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    textBox.getFill().setTransparent();
    textBox.getBorder().setTransparent();
    textBox.setRotation(345); // -15 degrees

    // ── Info line — directly below stamp, no rotation ───────────────
    var clearedDate = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy"
    );

    var infoW    = pageW * 0.55;
    var infoH    = 20;
    var infoLeft = (pageW - infoW) / 2;
    var infoTop  = tbTop + tbHeight + 6;

    var infoBox = slide.insertTextBox(
      serialTag + '   \u2022   ' + clearedDate,
      infoLeft, infoTop, infoW, infoH
    );

    var itf = infoBox.getText();
    itf.getTextStyle()
      .setFontSize(8)
      .setBold(false)
      .setForegroundColor('#CC0000');

    itf.getParagraphs()[0]
      .getRange()
      .getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    infoBox.getFill().setTransparent();
    infoBox.getBorder().setTransparent();

    pres.saveAndClose();

    var mergedBlob = DriveApp.getFileById(presId)
      .getAs(MimeType.PDF)
      .setName(origName + ' [RETURNED].pdf');

    var mergedFile = folder.createFile(mergedBlob);
    var mergedUrl  = mergedFile.getUrl();

    try { DriveApp.getFileById(presId).setTrashed(true); } catch(_) {}

    return { ok: true, url: mergedUrl };

  } catch (e) {
    console.error('_stampReturnedWatermarkOnAf_ error:', e);
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}



function _mergePdfsViaSlides_(fileIds, outputName) {
  var presentation = SlidesApp.create('TEMP_MERGE_' + Date.now());
  var presId = presentation.getId();

  var slides = presentation.getSlides();
  if (slides.length) slides[0].remove();

  for (var fi = 0; fi < fileIds.length; fi++) {
    var id = fileIds[fi];
    try {
      var url  = 'https://drive.google.com/thumbnail?id=' + id + '&sz=w1240';
      var resp = UrlFetchApp.fetch(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });

      if (resp.getResponseCode() !== 200) {
        console.warn('Thumbnail fetch failed for', id, resp.getResponseCode());
        return;
      }

      var imgBlob = resp.getBlob().setName('page_' + id + '.png');
      var slide   = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      var pageW   = presentation.getPageWidth();
      var pageH   = presentation.getPageHeight();

      slide.insertImage(imgBlob, 0, 0, pageW, pageH);
    } catch (e) {
      console.error('Slide insert failed for', id, e);
    }
  }

  presentation.save();

  var mergedBlob = DriveApp.getFileById(presId)
    .getAs(MimeType.PDF)
    .setName(outputName || 'merged.pdf');

  try { DriveApp.getFileById(presId).setTrashed(true); } catch(_) {}

  return mergedBlob;
}


function testStampOnly() {
  // Tests just the stamp — no sheet changes, no email, no cache bust
  var result = _stampReturnedWatermarkOnAf_('FBG-H6LQFL7MT4');
  Logger.log('ok: ' + result.ok);
  Logger.log('url: ' + (result.url || '—'));
  Logger.log('error: ' + (result.error || '—'));
  Logger.log('reason: ' + (result.reason || '—'));
}