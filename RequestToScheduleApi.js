var SHIFTS_SHEET_NAME   = "Shifts";
var REQUESTS_SHEET_NAME = "Requests";
var SHIFTS_HEADERS = ["Role","Name","Mode","Group","Date","Start","End"];

// Requests headers (keep AttachmentName even if UI removed attachments; it can stay blank)
var REQUESTS_HEADERS = [
  "ID",
  "SubmittedAt",
  "Email",
  "Type",
  "Priority",
  "StartDate",
  "EndDate",
  "TimeorHours",
  "Contact",
  "Reason",
  "AttachmentName",
  "Status",
  "ReviewedBy",
  "ReviewedAt",
  "ScheduleRowIds" // <-- NEW: stores JSON array of created shift row numbers
];

// ---------- Spreadsheet ----------
function getSpreadsheet_(){
  // Prefer SHEET_ID if you use a bound-less deployment
  if (typeof SHEET_ID !== "undefined" && SHEET_ID) {
    return SpreadsheetApp.openById(SHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ---------- Generic helpers ----------
function normalizeHeader_(h){ return String(h || "").trim(); }

function getHeaderMap_(sh) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  var map = {};
  headers.forEach(function(name, idx) {
    if (name) {
      map[name] = idx;                    // original case
      map[name.toLowerCase()] = idx;     // lowercase fallback
    }
  });
  return { headers: headers, map: map };
}

function ensureSheetAndHeaders_(sheetName, expectedHeaders){
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  var lastCol = Math.max(sh.getLastColumn(), expectedHeaders.length);
  var existing = sh.getRange(1,1,1,lastCol).getValues()[0].map(normalizeHeader_);

  // Detect if row 1 looks like data (protect existing)
  var rowLooksLikeData =
    existing.some(function(v){ return v.indexOf("@") !== -1; }) ||
    sh.getRange(1,1,1,lastCol).getValues()[0].some(function(v){ return v instanceof Date; }) ||
    (existing[0] && existing[0] !== normalizeHeader_(expectedHeaders[0]));

  var match = expectedHeaders.every(function(h, i){ return existing[i] === h; });

  if (!match){
    if (rowLooksLikeData) sh.insertRowBefore(1);
    sh.getRange(1,1,1,expectedHeaders.length).setValues([expectedHeaders]);
  }

  return sh;
}


function getShiftsSheet_(){
  return ensureSheetAndHeaders_(SHIFTS_SHEET_NAME, SHIFTS_HEADERS);
}

function safeLower_(s){ return String(s || "").trim().toLowerCase(); }

// Parse date input from HTML (yyyy-mm-dd) into Date object, or return "".
function toDateOrEmpty_(v) {
  if (!v) return "";
  if (v instanceof Date) return v;
  const s = String(v).trim();
  if (!s) return "";

  // YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const parts = s.split("-");
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  }

  // MM/DD/YYYY  ← add this
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) {
    return new Date(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2]));
  }

  const d = new Date(s);
  return isNaN(d.getTime()) ? "" : d;
}

// Convert sheet cell to Date object at midnight (best for schedule Date column)
function toMidnightDate_(cell) {
  if (!cell) return "";
  
  let d;
  if (cell instanceof Date) {
    d = new Date(cell);
  } else {
    const s = String(cell).trim();
    // MM/DD/YYYY
    const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (mdy) {
      d = new Date(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2]));
    } else {
      d = new Date(s);
    }
  }

  if (isNaN(d.getTime())) return "";
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Inclusive date loop
function eachDateInclusive_(startDate, endDate, fn){
  var d = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  var end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  while (d.getTime() <= end.getTime()){
    fn(new Date(d));
    d.setDate(d.getDate() + 1);
  }
}

// ---------- Users lookup (Role/Name/Group) ----------
function lookupUserInfoByEmail_(email){
  email = safeLower_(email);
  var local = email.split("@")[0] || "";

  var sh = getUsersSheet_();
  if (!sh){
    return {
      Role: "",
      Name: capitalizeWords_(local.replace(/[._\-]/g," ")),
      Group: ""
    };
  }

  var hm = getHeaderMap_(sh);
  var map = hm.map;
  var lastRow = sh.getLastRow();
  if (lastRow < 2){
    return { Role:"", Name: capitalizeWords_(local.replace(/[._\-]/g," ")), Group:"" };
  }

  var values = sh.getRange(2,1,lastRow-1,hm.headers.length).getValues();

  // allow common header variants
  var emailCol = (map["Email"] != null) ? map["Email"] :
                 (map["email"] != null) ? map["email"] :
                 (map["EMAIL"] != null) ? map["EMAIL"] : -1;

  var nameCol  = (map["username"] != null) ? map["username"] :
                 (map["Username"] != null) ? map["Username"] :
                 (map["Name"] != null) ? map["Name"] : -1;

  var roleCol  = (map["Role"] != null) ? map["Role"] :
                 (map["role"] != null) ? map["role"] : -1;

  var groupCol = (map["Group"] != null) ? map["Group"] :
                 (map["group"] != null) ? map["group"] : -1;

  for (var i = 0; i < values.length; i++){
    var row = values[i];
    var rowEmail = emailCol >= 0 ? safeLower_(row[emailCol]) : "";
    if (rowEmail === email){
      return {
        Role: String(roleCol >= 0 ? row[roleCol] : "" || ""),
        Name: String(nameCol >= 0 ? row[nameCol] : "" || local),
        Group: String(groupCol >= 0 ? row[groupCol] : "" || "")
      };
    }
  }

  return { Role:"", Name: capitalizeWords_(local.replace(/[._\-]/g," ")), Group:"" };
}

function capitalizeWords_(s){
  s = String(s || "").trim();
  if (!s) return "";
  return s.split(/\s+/).map(function(w){
    return w ? (w.charAt(0).toUpperCase() + w.slice(1)) : "";
  }).join(" ");
}

// ---------- Requests API ----------
function submitEmployeeRequest(payload) {
  if (!payload) return "missing_payload";
 
  var email = safeLower_(payload.email);
  if (!email) return "missing_email";
 
  // --- RBAC check ---
  try {
    requireRequestsSubmit_(email);
  } catch (e) {
    return (e && e.message) ? e.message : String(e);
  }
 
  var sh          = getRequestsSheet_();
  var id          = Utilities.getUuid();
  var submittedAt = new Date();
 
  var type        = String(payload.type        || "").trim();
  var priority    = String(payload.priority    || "Normal").trim();
  var startDate   = toDateOrEmpty_(payload.startDate);
  var endDate     = toDateOrEmpty_(payload.endDate);
  var timeOrHours = String(payload.timeOrHours || "").trim();
  var contact     = String(payload.contact     || "").trim();
  var reason      = String(payload.reason      || "").trim();
 
  if (!type)      return "missing_type";
  if (!startDate) return "missing_startDate";
  if (!endDate)   return "missing_endDate";
  if (!reason)    return "missing_reason";
 
  // --- Handle attachment ---
  var attachmentUrl = "";
 
  var b64      = String(payload.attachmentBase64 || "").trim();
  var mime     = String(payload.attachmentMime   || "").trim();
  var origName = String(payload.attachmentName   || "").trim();
 
  if (b64 && origName) {
    try {
      attachmentUrl = uploadRequestAttachment_(email, b64, mime || "application/octet-stream", origName);
    } catch (uploadErr) {
      // Log but don't block submission — store error note instead
      console.error("Attachment upload failed:", uploadErr);
      attachmentUrl = "Upload failed: " + ((uploadErr && uploadErr.message) || String(uploadErr));
    }
  }
 
  // --- Write row ---
  sh.appendRow([
    id,
    submittedAt,
    email,
    type,
    priority,
    startDate,
    endDate,
    timeOrHours,
    contact,
    reason,
    attachmentUrl,   // AttachmentName column — stores Drive URL (or empty)
    "Pending",
    "",              // ReviewedBy
    "",              // ReviewedAt
    ""               // ScheduleRowIds
  ]);
 
  return "ok";
}
 

function buildMapFromHm_(hm) {
  return {
    "Type":           hm.type,
    "ScheduleRowIds": hm.scheduleRowIds,
    "Email":          hm.email,
    "StartDate":      hm.startDate,
    "EndDate":        hm.endDate,
    "TimeorHours":    hm.timeOrHours,
    "Status":         hm.status,
    "ReviewedBy":     hm.reviewedBy,
    "ReviewedAt":     hm.reviewedAt,
    // lowercase fallbacks so createScheduleIfNeeded_ finds them either way
    "type":           hm.type,
    "schedulerowids": hm.scheduleRowIds,
    "email":          hm.email,
    "startdate":      hm.startDate,
    "enddate":        hm.endDate,
    "timeorhours":    hm.timeOrHours
  };
}

function getUsersSheet_() {
  try {
    const ss = (typeof SHEET_ID !== "undefined" && SHEET_ID)
      ? SpreadsheetApp.openById(SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
    return ss.getSheetByName("Users") || null;
  } catch(e) {
    return null;
  }
}

// ---------- Schedule creation ----------
function createScheduleIfNeeded_(reqSheet, reqHeaderMeta, reqRow) {
  var map = reqHeaderMeta.map;

  var typeColIdx = map["Type"] != null ? map["Type"] : map["type"];
  if (typeColIdx == null) return;

  var type = String(reqSheet.getRange(reqRow, typeColIdx + 1).getValue() || "").trim().toUpperCase();
  var scheduleTypes = ["VL", "SL", "WFH"];
  if (scheduleTypes.indexOf(type) === -1) return;

  // Scan actual header row to find ScheduleRowIds — never trust the map alone
  var scheduleCol = null;
  var headerRow = reqSheet.getRange(1, 1, 1, reqSheet.getLastColumn()).getValues()[0];
  for (var h = 0; h < headerRow.length; h++) {
    if (String(headerRow[h] || "").trim().toLowerCase() === "schedulerowids") {
      scheduleCol = h + 1;
      break;
    }
  }

  // Only create the column if it truly doesn't exist
  if (!scheduleCol) {
    scheduleCol = reqSheet.getLastColumn() + 1;
    reqSheet.getRange(1, scheduleCol).setValue("ScheduleRowIds");
  }

  var existing = String(reqSheet.getRange(reqRow, scheduleCol).getValue() || "").trim();
  if (existing) return; // idempotency

  var emailColIdx  = map["Email"]      != null ? map["Email"]      : map["email"];
  var startColIdx  = map["StartDate"]  != null ? map["StartDate"]  : map["startdate"];
  var endColIdx    = map["EndDate"]    != null ? map["EndDate"]    : map["enddate"];
  var timeColIdx   = map["TimeorHours"]!= null ? map["TimeorHours"]: map["timeorhours"];

  var email       = emailColIdx  != null ? String(reqSheet.getRange(reqRow, emailColIdx  + 1).getValue() || "").trim() : "";
  var startRaw    = startColIdx  != null ? reqSheet.getRange(reqRow, startColIdx  + 1).getValue() : null;
  var endRaw      = endColIdx    != null ? reqSheet.getRange(reqRow, endColIdx    + 1).getValue() : null;
  var timeOrHours = timeColIdx   != null ? String(reqSheet.getRange(reqRow, timeColIdx   + 1).getValue() || "").trim() : "";

  var startDate = toMidnightDate_(startRaw);
  var endDate   = toMidnightDate_(endRaw) || startDate;

  if (!(startDate instanceof Date) || isNaN(startDate.getTime())) {
    console.error("createScheduleIfNeeded_: invalid StartDate at row " + reqRow, startRaw);
    return;
  }

  var user = lookupUserInfoByEmail_(email);

  var outStart = "";
  var outEnd   = "";
  if (timeOrHours && timeOrHours.indexOf("-") !== -1) {
    var parts = timeOrHours.split("-");
    outStart = (parts[0] || "").trim();
    outEnd   = (parts[1] || "").trim();
  } else {
    outStart = timeOrHours || "";
    outEnd   = "";
  }

  var shifts      = getShiftsSheet_();
  var createdRows = [];

  eachDateInclusive_(startDate, endDate, function(day) {
    shifts.appendRow([
      user.Role  || "",
      user.Name  || email,
      type,
      user.Group || "",
      new Date(day.getFullYear(), day.getMonth(), day.getDate()),
      outStart,
      outEnd
    ]);
    createdRows.push(shifts.getLastRow());
  });

  reqSheet.getRange(reqRow, scheduleCol).setValue(JSON.stringify(createdRows));
  SpreadsheetApp.flush();

  console.log("createScheduleIfNeeded_: created rows " + JSON.stringify(createdRows) + " for request row " + reqRow);
}