// ============================================================
// requests.gs
// Handles reading, updating, and emailing employee requests
// (VL, SL, WFH, etc.)
// Rhino-safe: var only, no arrow functions, no template literals,
// no optional chaining, no const/let, no shorthand properties.
// ============================================================


/* ============================================================
 * RBAC GUARDS
 * ============================================================ */

function canViewRequestReview(email) {
  return canAccessPage("viewrequest", email);
}

function canReviewRequests(email) {
  return canAccessPage("viewrequest", email) && canDoAction("request.review", email);
}

function requireViewRequestReview_(email) {
  if (!canViewRequestReview(email)) {
    throw new Error("You do not have permission to access View Requests.");
  }
}

function requireReviewRequests_(email) {
  if (!canAccessPage("viewrequest", email)) {
    throw new Error("You do not have permission to access View Requests.");
  }
  if (!canDoAction("request.review", email)) {
    throw new Error("You do not have permission to review requests.");
  }
}


/* ============================================================
 * PRIVATE HELPERS
 * ============================================================ */

function normRequest_(v) {
  return String(v == null ? "" : v).trim();
}

function lowerRequest_(v) {
  return normRequest_(v).toLowerCase();
}

// NOTE: normalizeEmail_ is also defined in request_api.gs with the same implementation.
// Both are kept intentionally identical — do not diverge them.
function normalizeEmail_(v) {
  return String(v == null ? "" : v).trim().toLowerCase();
}

function getRequestActorEmail_(callerEmail) {
  if (callerEmail && String(callerEmail).trim()) {
    return String(callerEmail).trim();
  }
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function getRequestsSheet_() {
  var ss = (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  var sheetName = (typeof REQUESTS_SHEET_NAME !== "undefined" && REQUESTS_SHEET_NAME)
    ? REQUESTS_SHEET_NAME
    : "Requests";

  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    var headers = (typeof REQUEST_HEADERS !== "undefined" && Array.isArray(REQUEST_HEADERS))
      ? REQUEST_HEADERS
      : ["ID","SubmittedAt","Email","Type","Priority","StartDate","EndDate",
         "TimeorHours","Contact","Reason","AttachmentName","Status",
         "ReviewedBy","ReviewedAt","ScheduleRowIds"];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function lockRequests_(fn) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Build a header-index map from the first row of values.
 * Returns an object with named indices; -1 means not found.
 */
function getRequestHeaderMap_(headerRow) {
  var map = {};
  var headers = headerRow || [];

  headers.forEach(function(h, i) {
    map[String(h == null ? "" : h).trim().toLowerCase()] = i;
  });

  function pick() {
    for (var n = 0; n < arguments.length; n++) {
      var k = String(arguments[n]).toLowerCase();
      if (k in map) return map[k];
    }
    return -1;
  }

  return {
    id:          pick("id", "request id", "requestid"),
    status:      pick("status"),
    email:       pick("email", "employee email"),
    type:        pick("type", "request type"),
    priority:    pick("priority"),
    startDate:   pick("start date", "startdate", "start"),
    endDate:     pick("end date", "enddate", "end"),
    timeOrHours: pick("time", "hours", "time / hours", "timeorHours", "timeothours"),
    contact:     pick("contact"),
    reason:      pick("reason", "details", "reason / details"),
    reviewedBy:  pick("reviewed by", "reviewedby"),
    reviewedAt:  pick("reviewed at", "reviewedat"),
    submittedAt: pick("submitted at", "submittedat", "submitted"),
    attachmentName: pick("attachment", "attachment name", "attachmentname")
  };
}

/**
 * Build a flat request object from a single row of values.
 */
function buildRequestObjectFromRow_(row, hm) {
  function get(idx) {
    if (idx < 0 || idx >= row.length) return "";
    var v = row[idx];
    if (v instanceof Date) {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    }
    return String(v == null ? "" : v).trim();
  }

  return {
    id:             get(hm.id),
    status:         get(hm.status),
    email:          get(hm.email),
    type:           get(hm.type),
    priority:       get(hm.priority),
    startDate:      get(hm.startDate),
    endDate:        get(hm.endDate),
    timeOrHours:    get(hm.timeOrHours),
    contact:        get(hm.contact),
    reason:         get(hm.reason),
    reviewedBy:     get(hm.reviewedBy),
    reviewedAt:     get(hm.reviewedAt),
    submittedAt:    get(hm.submittedAt),
    attachmentName: get(hm.attachmentName)
  };
}

/**
 * Map all data rows to request objects.
 */
function mapRequests_(values) {
  if (!values || values.length < 2) return [];

  var hm   = getRequestHeaderMap_(values[0]);
  var rows = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var hasData = row.some(function(cell) {
      return String(cell == null ? "" : cell).trim() !== "";
    });
    if (!hasData) continue;
    rows.push(buildRequestObjectFromRow_(row, hm));
  }

  return rows;
}

/**
 * Build a column-key map from a header map (for createScheduleIfNeeded_).
 */
function buildMapFromHm_(hm) {
  var result = {};
  var keys = Object.keys(hm);
  for (var i = 0; i < keys.length; i++) {
    result[keys[i]] = hm[keys[i]];
  }
  return result;
}

/**
 * HTML escape helper — used in email construction.
 */
function _escHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g,  "&amp;")
    .replace(/</g,  "&lt;")
    .replace(/>/g,  "&gt;")
    .replace(/"/g,  "&quot;")
    .replace(/'/g,  "&#39;");
}


/* ============================================================
 * PUBLIC API — READ
 * ============================================================ */

function getRequestsData(callerEmail) {
  var actorEmail = normalizeEmail_(getRequestActorEmail_(callerEmail));

  if (!actorEmail) {
    return { ok: false, rows: [], error: "no_email" };
  }

  try {
    requireViewRequestReview_(actorEmail);
  } catch (e) {
    return { ok: false, rows: [], error: e.message || String(e) };
  }

  try {
    var sh     = getRequestsSheet_();
    var values = sh.getDataRange().getValues();

    if (values.length < 2) {
      return { ok: true, rows: [] };
    }

    return {
      ok:   true,
      rows: mapRequests_(values)
    };
  } catch (e) {
    Logger.log("[REQUESTS] getRequestsData error: " + (e.message || String(e)));
    return { ok: false, rows: [], error: e.message || String(e) };
  }
}


/* ============================================================
 * PUBLIC API — UPDATE STATUS
 * ============================================================ */

function updateRequestStatus(requestId, status, note, callerEmail) {
  requestId = normRequest_(requestId);
  var nextStatus  = normRequest_(status);
  var actorEmail  = normalizeEmail_(getRequestActorEmail_(callerEmail));

  if (!requestId)  return { ok: false, error: "missing_id" };
  if (!actorEmail) return { ok: false, error: "no_email" };

  var allowedStatuses = ["Acknowledged", "Approved", "Rejected"];
  var statusAllowed = false;
  for (var i = 0; i < allowedStatuses.length; i++) {
    if (allowedStatuses[i] === nextStatus) { statusAllowed = true; break; }
  }
  if (!statusAllowed) return { ok: false, error: "invalid_status" };

  try {
    requireReviewRequests_(actorEmail);
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }

  return lockRequests_(function() {
    var sh     = getRequestsSheet_();
    var values = sh.getDataRange().getValues();

    if (values.length < 2) return { ok: false, error: "empty" };

    var hm = getRequestHeaderMap_(values[0]);
    if (hm.id     === -1) return { ok: false, error: "missing_id_header" };
    if (hm.status === -1) return { ok: false, error: "missing_status_header" };

    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (normRequest_(values[i][hm.id]) === requestId) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex < 0) return { ok: false, error: "not_found" };

    var currentStatus      = hm.status > -1 ? normRequest_(values[rowIndex][hm.status]) : "";
    var currentStatusLower = lowerRequest_(currentStatus);

    if (currentStatusLower === "approved" || currentStatusLower === "rejected") {
      return { ok: false, error: "locked_status" };
    }

    var sheetRow = rowIndex + 1;

    if (hm.status     > -1) sh.getRange(sheetRow, hm.status     + 1).setValue(nextStatus);
    if (hm.reviewedBy > -1) sh.getRange(sheetRow, hm.reviewedBy + 1).setValue(actorEmail);
    if (hm.reviewedAt > -1) sh.getRange(sheetRow, hm.reviewedAt + 1).setValue(new Date());

    SpreadsheetApp.flush();

    // Create schedule rows if approved — non-blocking
    if (nextStatus.toLowerCase() === "approved") {
      try {
        if (typeof createScheduleIfNeeded_ === "function") {
          var freshValues = sh.getDataRange().getValues();
          var freshHm     = getRequestHeaderMap_(freshValues[0]);
          createScheduleIfNeeded_(sh, { map: buildMapFromHm_(freshHm) }, sheetRow);
          SpreadsheetApp.flush();
        }
      } catch (schedErr) {
        Logger.log("[REQUESTS] createScheduleIfNeeded_ failed (non-blocking): " + (schedErr.message || String(schedErr)));
      }
    }

    var updatedValues = sh.getRange(sheetRow, 1, 1, sh.getLastColumn()).getValues()[0];
    var updatedRow    = buildRequestObjectFromRow_(updatedValues, hm);

   // Send notification email — non-blocking
    try {
      if (nextStatus.toLowerCase() === "approved") {
        _sendRequestApprovalEmail_(updatedRow, actorEmail);
      } else if (nextStatus.toLowerCase() === "rejected") {
        _sendRequestRejectionEmail_(updatedRow, actorEmail);
      }
    } catch (e) {
      Logger.log("[REQUESTS] Notification email failed: " + (e.message || String(e)));
    }

    return { ok: true, updatedRow: updatedRow };
  });
}


/* ============================================================
 * EMAIL — APPROVAL NOTIFICATION
 * ============================================================ */

function _sendRequestApprovalEmail_(request, reviewerEmail) {
  var to = String((request && request.email) ? request.email : "").trim();
  if (!to) return;

  var type      = String((request && request.type)      ? request.type      : "Request").trim();
  var startDate = String((request && request.startDate) ? request.startDate : "—").trim();
  var endDate   = String((request && request.endDate)   ? request.endDate   : "—").trim();
  var priority  = String((request && request.priority)  ? request.priority  : "Normal").trim();
  var reason    = String((request && request.reason)    ? request.reason    : "—").trim();
  var reviewer  = String(reviewerEmail                  ? reviewerEmail     : "Manila IT").trim();
  var subject   = "Your " + type + " Request Has Been Approved";

  // Local escape helper so this function is fully self-contained
  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g,  "&amp;")
      .replace(/</g,  "&lt;")
      .replace(/>/g,  "&gt;")
      .replace(/"/g,  "&quot;")
      .replace(/'/g,  "&#39;");
  }

  // Build a table row
  function tableRow(label, value, shaded) {
    var bg = shaded ? " style=\"background:#f9f9f9;\"" : "";
    return "<tr" + bg + ">" +
      "<td style=\"padding:10px 14px;font-size:12px;font-weight:700;text-transform:uppercase;" +
          "letter-spacing:.06em;color:#888;width:38%;border-bottom:1px solid #eee;\">" +
        esc(label) +
      "</td>" +
      "<td style=\"padding:10px 14px;font-size:13px;font-weight:600;color:#111;" +
          "border-bottom:1px solid #eee;\">" +
        esc(value) +
      "</td>" +
    "</tr>";
  }

  var htmlBody =
    "<div style=\"font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1a1a1a;\">" +

      "<div style=\"background:#0f0f0f;border-radius:12px 12px 0 0;padding:28px 32px;\">" +
        "<div style=\"font-size:22px;font-weight:800;color:#ffffff;letter-spacing:-.02em;\">" +
          "Request Approved &#10003;" +
        "</div>" +
        "<div style=\"font-size:13px;color:rgba(255,255,255,.55);margin-top:4px;\">" +
          "Manila IT Inventory System" +
        "</div>" +
      "</div>" +

      "<div style=\"border:1px solid #e5e5e5;border-top:none;border-radius:0 0 12px 12px;padding:28px 32px;\">" +

        "<p style=\"margin:0 0 20px;font-size:14px;line-height:1.6;color:#333;\">" +
          "Hi, your <strong>" + esc(type) + "</strong> request has been " +
          "<strong style=\"color:#16a34a;\">approved</strong>. Here are the details:" +
        "</p>" +

        "<table style=\"width:100%;border-collapse:collapse;margin-bottom:24px;\">" +
          tableRow("Request Type", type,      true)  +
          tableRow("Start Date",   startDate, false) +
          tableRow("End Date",     endDate,   true)  +
          tableRow("Priority",     priority,  false) +
          "<tr style=\"background:#f9f9f9;\">" +
            "<td style=\"padding:10px 14px;font-size:12px;font-weight:700;text-transform:uppercase;" +
                "letter-spacing:.06em;color:#888;\">Reason</td>" +
            "<td style=\"padding:10px 14px;font-size:13px;color:#111;\">" + esc(reason) + "</td>" +
          "</tr>" +
        "</table>" +

        "<div style=\"background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;" +
            "padding:14px 16px;margin-bottom:24px;\">" +
          "<div style=\"font-size:13px;font-weight:700;color:#15803d;margin-bottom:4px;\">" +
            "&#10003; Approved" +
          "</div>" +
          "<div style=\"font-size:12px;color:#166534;\">" +
            "Reviewed by: " + esc(reviewer) +
          "</div>" +
        "</div>" +

        "<p style=\"font-size:12px;color:#aaa;margin:0;line-height:1.6;\">" +
          "This is an automated message from the Manila IT Inventory system. " +
          "Please do not reply to this email. For questions, contact " +
          "<a href=\"mailto:manila-it@fbgphilippines.com\" style=\"color:#2563eb;\">" +
            "manila-it@fbgphilippines.com" +
          "</a>." +
        "</p>" +

      "</div>" +
    "</div>";

  var plainBody =
    "Your " + type + " request has been approved.\n" +
    "\n" +
    "Type:        " + type      + "\n" +
    "Start Date:  " + startDate + "\n" +
    "End Date:    " + endDate   + "\n" +
    "Priority:    " + priority  + "\n" +
    "Reason:      " + reason    + "\n" +
    "\n" +
    "Reviewed by: " + reviewer  + "\n" +
    "\n" +
    "This is an automated message from Manila IT Inventory.\n" +
    "For questions contact: manila-it@fbgphilippines.com";

  GmailApp.sendEmail(to, subject, plainBody, {
    htmlBody: htmlBody,
    from:     "manila-it@fbgphilippines.com",
    name:     "Manila IT Inventory",
    replyTo:  "manila-it@fbgphilippines.com"
  });

  Logger.log("[REQUESTS] Approval email sent to " + to + " — type: " + type);
}

function _sendRequestRejectionEmail_(request, reviewerEmail) {
  var to = String((request && request.email) ? request.email : "").trim();
  if (!to) return;

  var type      = String((request && request.type)      ? request.type      : "Request").trim();
  var startDate = String((request && request.startDate) ? request.startDate : "—").trim();
  var endDate   = String((request && request.endDate)   ? request.endDate   : "—").trim();
  var reviewer  = String(reviewerEmail                  ? reviewerEmail     : "Manila IT").trim();
  var subject   = "Your " + type + " Request Was Not Approved";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#39;");
  }

  var htmlBody =
    "<div style=\"font-family:Arial,sans-serif;max-width:560px;margin:0 auto;color:#1a1a1a;\">" +
      "<div style=\"background:#0f0f0f;border-radius:12px 12px 0 0;padding:28px 32px;\">" +
        "<div style=\"font-size:22px;font-weight:800;color:#ffffff;letter-spacing:-.02em;\">Request Not Approved</div>" +
        "<div style=\"font-size:13px;color:rgba(255,255,255,.55);margin-top:4px;\">Manila IT Inventory System</div>" +
      "</div>" +
      "<div style=\"border:1px solid #e5e5e5;border-top:none;border-radius:0 0 12px 12px;padding:28px 32px;\">" +
        "<p style=\"margin:0 0 20px;font-size:14px;line-height:1.6;color:#333;\">" +
          "Hi, your <strong>" + esc(type) + "</strong> request has been " +
          "<strong style=\"color:#c94040;\">rejected</strong>. Please reach out to your manager for more details." +
        "</p>" +
        "<table style=\"width:100%;border-collapse:collapse;margin-bottom:24px;\">" +
          "<tr><td style=\"padding:10px 14px;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#888;width:38%;border-bottom:1px solid #eee;\">Request Type</td>" +
          "<td style=\"padding:10px 14px;font-size:13px;font-weight:600;color:#111;border-bottom:1px solid #eee;\">" + esc(type) + "</td></tr>" +
          "<tr style=\"background:#f9f9f9;\"><td style=\"padding:10px 14px;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#888;border-bottom:1px solid #eee;\">Start Date</td>" +
          "<td style=\"padding:10px 14px;font-size:13px;font-weight:600;color:#111;border-bottom:1px solid #eee;\">" + esc(startDate) + "</td></tr>" +
          "<tr><td style=\"padding:10px 14px;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:#888;\">End Date</td>" +
          "<td style=\"padding:10px 14px;font-size:13px;font-weight:600;color:#111;\">" + esc(endDate) + "</td></tr>" +
        "</table>" +
        "<div style=\"background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:14px 16px;margin-bottom:24px;\">" +
          "<div style=\"font-size:13px;font-weight:700;color:#c94040;margin-bottom:4px;\">&#10007; Rejected</div>" +
          "<div style=\"font-size:12px;color:#991b1b;\">Reviewed by: " + esc(reviewer) + "</div>" +
        "</div>" +
        "<p style=\"font-size:12px;color:#aaa;margin:0;line-height:1.6;\">This is an automated message from the Manila IT Inventory system. " +
          "For questions, contact <a href=\"mailto:manila-it@fbgphilippines.com\" style=\"color:#2563eb;\">manila-it@fbgphilippines.com</a>.</p>" +
      "</div>" +
    "</div>";

  var plainBody =
    "Your " + type + " request has been rejected.\n\n" +
    "Type:       " + type      + "\n" +
    "Start Date: " + startDate + "\n" +
    "End Date:   " + endDate   + "\n\n" +
    "Reviewed by: " + reviewer + "\n\n" +
    "Please contact your manager for more information.\n" +
    "For questions contact: manila-it@fbgphilippines.com";

  GmailApp.sendEmail(to, subject, plainBody, {
    htmlBody: htmlBody,
    from:     "manila-it@fbgphilippines.com",
    name:     "Manila IT Inventory",
    replyTo:  "manila-it@fbgphilippines.com"
  });

  Logger.log("[REQUESTS] Rejection email sent to " + to + " — type: " + type);
}