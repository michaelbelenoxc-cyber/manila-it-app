/******************** REQUEST API (RBAC VERSION) ********************/
const REQUESTS_SHEET_NAME_ = "Requests";

const REQUEST_HEADERS = [
  "ID", "SubmittedAt", "Email", "Type", "Priority",
  "StartDate", "EndDate", "TimeorHours", "Contact",
  "Reason", "AttachmentName", "Status", "ReviewedBy", "ReviewedAt",
  "ScheduleRowIds"  
];
/* =========================================================
 * RBAC
 * =======================================================*/
function canViewRequests(email) {
  return canAccessPage("request", email);
}

function canSubmitRequests(email) {
  return canAccessPage("request", email) && canDoAction("request.submit", email);
}

function requireRequestsView_(email) {
  if (!canAccessPage("request", email)) {
    throw new Error("You do not have permission to access Requests.");
  }
}

function requireRequestsSubmit_(email) {
  if (!canAccessPage("request", email)) {
    throw new Error("You do not have permission to access Requests.");
  }
  if (!canDoAction("request.submit", email)) {
    throw new Error("You do not have permission to submit or manage requests.");
  }
}

/* =========================================================
 * SHEET / ACTOR HELPERS
 * =======================================================*/
function getRequestsSpreadsheet_() {
  return (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}



function ensureRequestsHeaders_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
    return;
  }

  const lastCol = Math.max(sh.getLastColumn(), REQUEST_HEADERS.length);
  const row1 = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());

  const matches = REQUEST_HEADERS.every((h, i) => row1[i] === h);
  const rowLooksLikeData = row1.some(v => v.includes("@")) || row1[0] !== "ID";

  if (!matches) {
    if (rowLooksLikeData) sh.insertRowBefore(1);
    sh.getRange(1, 1, 1, REQUEST_HEADERS.length).setValues([REQUEST_HEADERS]);
  }
}

function getRequestActorEmail_(fallbackEmail) {
  const fb = normalizeEmail_(fallbackEmail || "");
  if (fb) return fb;

  try {
    const active = normalizeEmail_(Session.getActiveUser().getEmail() || "");
    if (active) return active;
  } catch (e) {}

  try {
    const effective = normalizeEmail_(Session.getEffectiveUser().getEmail() || "");
    if (effective) return effective;
  } catch (e) {}

  return "";
}

function lockRequests_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(8000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 * NORMALIZE / UTILS
 * =======================================================*/
function normRequest_(v) {
  return String(v ?? "").trim();
}

function lowerRequest_(v) {
  return String(v ?? "").trim().toLowerCase();
}

function normHeaderRequest_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getRequestHeaderMap_(headers) {
  const normHeaders = (headers || []).map(normHeaderRequest_);

  function idx(name) {
    return normHeaders.indexOf(normHeaderRequest_(name));
  }

  function idxAny() {
    for (let i = 0; i < arguments.length; i++) {
      const pos = idx(arguments[i]);
      if (pos > -1) return pos;
    }
    return -1;
  }

  return {
    id:             idxAny("id", "requestid", "reqid"),
    submittedAt:    idxAny("submittedat", "timestamp", "createdat", "datecreated"),
    email:          idxAny("email", "emailaddress", "requesteremail"),
    type:           idxAny("type", "requesttype"),
    priority:       idxAny("priority", "urgency"),
    startDate:      idxAny("startdate", "fromdate", "datefrom"),
    endDate:        idxAny("enddate", "todate", "dateto"),
    timeOrHours:    idxAny("timeorhours", "hours", "time"),
    contact:        idxAny("contact", "contactnumber", "phone", "mobilenumber"),
    reason:         idxAny("reason", "details", "description", "purpose"),
    attachmentName: idxAny("attachmentname", "attachment", "filename", "file"),
    status:         idxAny("status", "requeststatus"),
    reviewedBy:     idxAny("reviewedby", "approver", "handledby"),
    reviewedAt:     idxAny("reviewedat", "reviewdate", "approvedat", "updatedat"),
    scheduleRowIds: idxAny("schedulerowids", "schedulerows", "shiftrowids")
  };
}

function parseMMDDYYYY_(v) {
  const s = String(v || "").trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;

  const mm = Number(m[1]);
  const dd = Number(m[2]);
  const yyyy = Number(m[3]);

  const d = new Date(yyyy, mm - 1, dd);
  if (
    d.getFullYear() !== yyyy ||
    d.getMonth() !== mm - 1 ||
    d.getDate() !== dd
  ) {
    return null;
  }

  return d;
}

function formatDateRequest_(v) {
  const tz = Session.getScriptTimeZone();
  return v instanceof Date ? Utilities.formatDate(v, tz, "MM/dd/yyyy") : String(v || "");
}

function formatDateTimeRequest_(v) {
  const tz = Session.getScriptTimeZone();
  return v instanceof Date ? Utilities.formatDate(v, tz, "MM/dd/yyyy HH:mm") : String(v || "");
}

function buildRequestObjectFromRow_(row, hm) {
  const safe = (idx) => (idx > -1 ? row[idx] : "");

  return {
    id: normRequest_(safe(hm.id)),
    submittedAt: formatDateTimeRequest_(safe(hm.submittedAt)),
    email: normRequest_(safe(hm.email)),
    type: normRequest_(safe(hm.type)),
    priority: normRequest_(safe(hm.priority)) || "Normal",
    startDate: formatDateRequest_(safe(hm.startDate)),
    endDate: formatDateRequest_(safe(hm.endDate)),
    timeOrHours: normRequest_(safe(hm.timeOrHours)),
    contact: normRequest_(safe(hm.contact)),
    reason: normRequest_(safe(hm.reason)),
    attachmentName: normRequest_(safe(hm.attachmentName)),
    status: normRequest_(safe(hm.status)) || "Pending",
    reviewedBy: normRequest_(safe(hm.reviewedBy)),
    reviewedAt: formatDateTimeRequest_(safe(hm.reviewedAt))
  };
}

function validateRequestPayload_(payload) {
  if (!payload) throw new Error("Missing payload.");

  const type = normRequest_(payload.type);
  const startDate = normRequest_(payload.startDate);
  const endDate = normRequest_(payload.endDate);
  const reason = normRequest_(payload.reason);

  if (!type) throw new Error("Missing request type.");
  if (!startDate) throw new Error("Missing start date.");
  if (!endDate) throw new Error("Missing end date.");
  if (!reason) throw new Error("Missing reason.");
}



/* =========================================================
 * UPDATE REQUEST
 * Used by myrequest.html owner edit flow
 * =======================================================*/
function updateRequestById(requestId, payload, callerEmail) {
  requestId = normRequest_(requestId);
  if (!requestId) return { ok: false, error: "missing_id" };

  payload = payload || {};
  const actorEmail = normalizeEmail_(getRequestActorEmail_(callerEmail || payload._actorEmail));
  if (!actorEmail) return { ok: false, error: "no_email" };

  try {
    if (typeof requireMyRequestsEdit_ === "function") {
      requireMyRequestsEdit_(actorEmail);
    } else {
      requireRequestsSubmit_(actorEmail);
    }
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }

  return lockRequests_(() => {
    const sh = getRequestsSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { ok: false, error: "empty" };

    const hm = getRequestHeaderMap_(values[0]);
    if (hm.id === -1 || hm.email === -1) return { ok: false, error: "missing_headers" };

    const rowIndex = values.findIndex((row, i) =>
      i > 0 && normRequest_(row[hm.id]) === requestId
    );
    if (rowIndex < 0) return { ok: false, error: "not_found" };

    const rowEmail = normalizeEmail_(values[rowIndex][hm.email]);
    if (rowEmail !== actorEmail) {
      return {
        ok: false,
        error: "not_owner",
        debug: { actorEmail: actorEmail, rowEmail: rowEmail, requestId: requestId }
      };
    }

    const currentStatus = hm.status > -1 ? normRequest_(values[rowIndex][hm.status]) : "";
    const currentStatusLower = lowerRequest_(currentStatus);
    if (currentStatusLower === "approved" || currentStatusLower === "rejected") {
      return { ok: false, error: "locked_status" };
    }

    validateRequestPayload_(payload);

    const startDate = parseMMDDYYYY_(payload.startDate);
    const endDate = parseMMDDYYYY_(payload.endDate);

    if (!startDate) return { ok: false, error: "missing_startDate" };
    if (!endDate) return { ok: false, error: "missing_endDate" };
    if (endDate.getTime() < startDate.getTime()) {
      return { ok: false, error: "invalid_date_range" };
    }

    const sheetRow = rowIndex + 1;

    const setIf = (idx, value) => {
      if (idx > -1) sh.getRange(sheetRow, idx + 1).setValue(value);
    };

    setIf(hm.type, normRequest_(payload.type));
    setIf(hm.priority, normRequest_(payload.priority) || "Normal");
    setIf(hm.startDate, startDate);
    setIf(hm.endDate, endDate);
    setIf(hm.timeOrHours, normRequest_(payload.timeOrHours));
    setIf(hm.contact, normRequest_(payload.contact));
    setIf(hm.reason, normRequest_(payload.reason));
    setIf(hm.attachmentName, normRequest_(payload.attachmentName));

    const updatedValues = sh.getRange(sheetRow, 1, 1, sh.getLastColumn()).getValues()[0];
    const updatedRow = buildRequestObjectFromRow_(updatedValues, hm);

    return { ok: true, updatedRow: updatedRow };
  });
}

/* =========================================================
 * GET MY REQUESTS
 * Used by myrequest.html
 * =======================================================*/
function getMyRequests(callerEmail) {
  const actorEmail = normalizeEmail_(getRequestActorEmail_(callerEmail));

  if (!actorEmail) {
    return {
      ok: false,
      email: "",
      matched: 0,
      rows: [],
      who: { active: "", effective: "" }
    };
  }

  try {
    if (typeof requireMyRequestsView_ === "function") {
      requireMyRequestsView_(actorEmail);
    } else {
      requireRequestsView_(actorEmail);
    }
  } catch (e) {
    return {
      ok: false,
      email: actorEmail,
      matched: 0,
      rows: [],
      error: e.message || String(e)
    };
  }

  const who = {
    active: (() => {
      try { return normalizeEmail_(Session.getActiveUser().getEmail() || ""); } catch (e) { return ""; }
    })(),
    effective: (() => {
      try { return normalizeEmail_(Session.getEffectiveUser().getEmail() || ""); } catch (e) { return ""; }
    })()
  };

  const sh = getRequestsSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) {
    return { ok: true, email: actorEmail, matched: 0, rows: [], who: who };
  }

  const hm = getRequestHeaderMap_(data[0]);
  if (hm.email === -1) {
    return {
      ok: false,
      email: actorEmail,
      matched: 0,
      rows: [],
      who: who,
      error: "missing_email_header"
    };
  }

  const rows = data.slice(1)
    .filter(r => normalizeEmail_(r[hm.email]) === actorEmail)
    .map(r => buildRequestObjectFromRow_(r, hm));

  return {
    ok: true,
    email: actorEmail,
    matched: rows.length,
    rows: rows,
    who: who
  };
}

/* =========================================================
 * MAP REQUESTS
 * =======================================================*/
function mapRequests_(values) {
  if (!values || !values.length) return [];

  const hm = getRequestHeaderMap_(values[0]);

  return values
    .slice(1)
    .filter(r => r.some(v => String(v || "").trim() !== ""))
    .map(r => buildRequestObjectFromRow_(r, hm));
}

function getRequestsSheet_() {
  if (typeof ensureSheetAndHeaders_ === "function") {
    return ensureSheetAndHeaders_(REQUESTS_SHEET_NAME_, REQUEST_HEADERS);
  }
  const ss = getRequestsSpreadsheet_();
  let sh = ss.getSheetByName(REQUESTS_SHEET_NAME_);
  if (!sh) sh = ss.insertSheet(REQUESTS_SHEET_NAME_);
  ensureRequestsHeaders_(sh);
  return sh;
}

