const EMPLOYEES_TAB_NAME = "Employees";
const EMPLOYEE_DATE_FMT = "MM/dd/yyyy";

const EMPLOYEE_FIELDS = [
  "employeeId",
  "displayName",
  "email",
  "alias",
  "department",
  "manager",
  "title",
  "location",
  "status",
  "startDate",
  "eod",
  "personalEmail"
];

const EMPLOYEE_REQUIRED_FIELDS = [
  "displayName",
  "email",
  "department",
  "status"
];

const EMPLOYEE_ALIASES = {
  employeeId:    ["Employee ID", "EmployeeID", "Emp ID", "ID"],
  displayName:   ["Display Name", "Employee Name", "Name"],
  email:         ["Email", "Work Email", "Company Email"],
  alias:         ["Alias"],
  department:    ["Department", "Dept"],
  manager:       ["Manager", "Supervisor"],
  title:         ["Title", "Job Title", "Position"],
  location:      ["Location", "Site"],
  status:        ["Status"],
  startDate:     ["Start Date", "Hire Date", "StartDate"],
  eod:           ["EOD", "End Date", "End Of Date"],
  personalEmail: ["Personal Email", "PersonalEmail", "Alt Email"]
};

const EMPLOYEE_EXPORT_LABELS = {
  employeeId: "Employee ID",
  displayName: "Display Name",
  email: "Email",
  alias: "Alias",
  department: "Department",
  manager: "Manager",
  title: "Title",
  location: "Location",
  status: "Status",
  startDate: "Start Date",
  eod: "EOD",
  personalEmail: "Personal Email"
};


// Add these near the top of employees.gs alongside the other constants
var EMPLOYEES_DATASET_CACHE_KEY_ = "emp_dataset_v1";
var EMPLOYEES_DATASET_CACHE_TTL_ = 3 * 60; // 3 minutes

function _empGetCachedDataset_() {
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get(EMPLOYEES_DATASET_CACHE_KEY_);
    if (cached) {
      var parsed = JSON.parse(cached);
      if (parsed && Array.isArray(parsed)) return parsed;
    }
  } catch (e) {}
  return null;
}

function _empSetCachedDataset_(displayRows) {
  try {
    var cache = CacheService.getScriptCache();
    var json  = JSON.stringify(displayRows);
    if (json.length < 90000) {
      cache.put(EMPLOYEES_DATASET_CACHE_KEY_, json, EMPLOYEES_DATASET_CACHE_TTL_);
    }
  } catch (e) {}
}

function _empBustCache_() {
  try {
    CacheService.getScriptCache().remove(EMPLOYEES_DATASET_CACHE_KEY_);
  } catch (e) {}
}

function employees_ok_(data, extra) {
  return Object.assign({ ok: true, data: data }, extra || {});
}

function employees_fail_(err, extra) {
  return Object.assign(
    {
      ok: false,
      error: err && err.message ? err.message : String(err)
    },
    extra || {}
  );
}

/* =========================================================
 * BASIC HELPERS
 * =======================================================*/
function employees_norm_(s) {
  return String(s ?? "")
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/[\s._-]+/g, " ")
    .trim()
    .toLowerCase();
}

function employees_trim_(v) {
  return String(v ?? "").trim();
}

function employees_norm_id_(v) {
  return String(v ?? "").trim();
}

function employees_norm_email_(v) {
  return String(v ?? "").trim().toLowerCase();
}

function employees_is_valid_email_(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v ?? "").trim());
}

function employees_to_positive_int_(v, fallback) {
  const n = parseInt(String(v ?? ""), 10);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}

function employees_format_date_(v) {
  if (!v) return "";

  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), EMPLOYEE_DATE_FMT);
  }

  const s = String(v).trim();
  if (!s) return "";

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), EMPLOYEE_DATE_FMT);
  }

  return s;
}

function employees_parse_date_input_(v) {
  if (v === null || v === undefined) return "";
  if (v instanceof Date && !isNaN(v.getTime())) return v;

  const s = String(v).trim();
  if (!s) return "";

  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) {
    const mm = parseInt(m[1], 10);
    const dd = parseInt(m[2], 10);
    const yy = parseInt(m[3], 10);
    const d = new Date(yy, mm - 1, dd);
    if (!isNaN(d.getTime())) return d;
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  return s;
}

function employees_date_to_epoch_(v) {
  if (!v) return NaN;

  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getTime();
  }

  const s = String(v).trim();
  if (!s) return NaN;

  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(s);
  if (m) {
    const mm = parseInt(m[1], 10);
    const dd = parseInt(m[2], 10);
    const yy = parseInt(m[3], 10);
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? NaN : d.getTime();
  }

  const t = new Date(s).getTime();
  return isNaN(t) ? NaN : t;
}

function employees_cell_(row, idx) {
  return idx >= 0 ? row[idx] ?? "" : "";
}

function employees_strip_html_(value) {
  return String(value ?? "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}

/* =========================================================
 * AUDIT
 * =======================================================*/
function employees_audit_(action, resource, status) {
  try {
    // Parse resource like "Employees:FBG-00169" into entity + serialTag
    const parts     = String(resource || "").split(":");
    const entity    = parts[0]?.trim() || "Employees";
    const serialTag = parts[1]?.trim() || "";

    auditWrite({
      action:    action,
      entity:    entity,
      serialTag: serialTag,
      field:     "",
      oldValue:  "",
      newValue:  status || "",
      details:   {}
    });
  } catch (_) {}
}

/* =========================================================
 * RBAC
 * =======================================================*/
function employees_get_rbac_() {
  let email = "";
  let role = "viewer";

  try {
    email = typeof getCurrentUserEmail_ === "function"
      ? String(getCurrentUserEmail_() || "").trim().toLowerCase()
      : String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  } catch (_) {}

  try {
    if (typeof getCurrentUserRole_ === "function") {
      role = String(getCurrentUserRole_() || "viewer").trim().toLowerCase();
    }
  } catch (_) {}

  let pageAllowed = true;
  let canManage = false;

  if (typeof canAccessPage === "function") {
    // Try exact key first, then capitalized, then case-insensitive scan
    pageAllowed = !!(
      canAccessPage("employees", email) ||
      canAccessPage("Employees", email) ||
      canAccessPage("EMPLOYEES", email)
    );
  } else if (typeof isAllowed_ === "function") {
    pageAllowed = !!isAllowed_(role, "page", "employees");
  }

  if (typeof canDoAction === "function") {
    canManage = !!(
      canDoAction("employees.manage", email) ||
      canDoAction("Employees.delete", email)
    );
  } else if (typeof isAllowed_ === "function") {
    canManage = !!isAllowed_(role, "action", "employees.manage");
  }

  return { email, role, pageAllowed, canManage };
}

function employees_assert_page_access_() {
  const p = employees_get_rbac_();
  if (!p.pageAllowed) throw new Error("You do not have permission to access Employees.");
  return p;
}

function employees_assert_manage_() {
  const p = employees_get_rbac_();
  if (!p.pageAllowed) {
    throw new Error("You do not have permission to access Employees.");
  }
  if (!p.canManage) {
    throw new Error("You do not have permission to manage employees.");
  }
  return p;
}

function employees_assert_delete_() {
  const p = employees_get_rbac_();
  if (!p.pageAllowed) {
    throw new Error("You do not have permission to access Employees.");
  }
  if (!canDoAction("employees.delete", p.email)) {
    throw new Error("You do not have permission to delete employees.");
  }
  return p;
}

/* =========================================================
 * SHEET / DATA LOADING
 * =======================================================*/
function employees_require_sheet_() {
  if (typeof SHEET_ID === "undefined" || !SHEET_ID) {
    throw new Error("SHEET_ID is not defined. Define SHEET_ID in a global .gs file.");
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(EMPLOYEES_TAB_NAME);
  if (!sh) {
    throw new Error('Employees sheet not found: "' + EMPLOYEES_TAB_NAME + '"');
  }

  return sh;
}

function employees_read_header_row_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error("Employees sheet has no columns.");
  return sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) {
    return String(h ?? "");
  });
}

function employees_build_col_map_(headerRow) {
  const rawHeaders = headerRow.map(function (h) { return String(h ?? ""); });
  const normalized = rawHeaders.map(employees_norm_);

  function findIndex_(aliases) {
    for (let i = 0; i < aliases.length; i++) {
      const idx = normalized.indexOf(employees_norm_(aliases[i]));
      if (idx >= 0) return idx;
    }
    return -1;
  }

  const col = {};
  EMPLOYEE_FIELDS.forEach(function (key) {
    col[key] = findIndex_(EMPLOYEE_ALIASES[key] || [key]);
  });

  const missing = EMPLOYEE_REQUIRED_FIELDS.filter(function (key) {
    return col[key] < 0;
  });

  if (missing.length) {
    throw new Error(
      "Missing required Employees sheet columns: " +
      missing.join(", ") +
      "\nHeaders found: " +
      rawHeaders.join(" | ")
    );
  }

  return { rawHeaders: rawHeaders, col: col };
}

function employees_load_sheet_context_() {
  const sh = employees_require_sheet_();
  const headerRow = employees_read_header_row_(sh);
  const map = employees_build_col_map_(headerRow);

  return {
    sh: sh,
    rawHeaders: map.rawHeaders,
    col: map.col,
    lastRow: sh.getLastRow(),
    lastCol: sh.getLastColumn()
  };
}

function employees_load_dataset_() {
  const ctx = employees_load_sheet_context_();
  const sh = ctx.sh;
  const rawHeaders = ctx.rawHeaders;
  const col = ctx.col;
  const lastRow = ctx.lastRow;
  const lastCol = ctx.lastCol;

  if (lastRow < 2) {
    return {
      sh: sh,
      rawHeaders: rawHeaders,
      col: col,
      valueRows: [],
      displayRows: [],
      employees: [],
      idToIndex: {}
    };
  }

 var cachedDisplay = _empGetCachedDataset_();
var range = sh.getRange(2, 1, lastRow - 1, lastCol);
var valueRows = range.getValues();
var displayRows = cachedDisplay || range.getDisplayValues();
if (!cachedDisplay) _empSetCachedDataset_(displayRows);

  const employees = [];
  const idToIndex = {};

 for (let i = 0; i < displayRows.length; i++) {
  const rowObj = employees_row_to_object_(displayRows[i], col);

  // Skip completely empty rows
  if (!EMPLOYEE_FIELDS.some(function(k) { return String(rowObj[k] || "").trim(); })) continue;

  rowObj._rowIndex = i; // attach sheet row index for fallback lookup

  employees.push(rowObj);

  const id = employees_norm_id_(rowObj.employeeId);
  if (id && !Object.prototype.hasOwnProperty.call(idToIndex, id)) {
    idToIndex[id] = i;
  }
}

  return {
    sh: sh,
    rawHeaders: rawHeaders,
    col: col,
    valueRows: valueRows,
    displayRows: displayRows,
    employees: employees,
    idToIndex: idToIndex
  };
}

/* =========================================================
 * TRANSFORM / VALIDATION
 * =======================================================*/
function employees_row_to_object_(row, col) {
  return {
    employeeId: employees_norm_id_(employees_cell_(row, col.employeeId)),
    displayName: employees_trim_(employees_cell_(row, col.displayName)),
    email: employees_trim_(employees_cell_(row, col.email)),
    alias: employees_trim_(employees_cell_(row, col.alias)),
    department: employees_trim_(employees_cell_(row, col.department)),
    manager: employees_trim_(employees_cell_(row, col.manager)),
    title: employees_trim_(employees_cell_(row, col.title)),
    location: employees_trim_(employees_cell_(row, col.location)),
    status: employees_trim_(employees_cell_(row, col.status)),
    startDate: employees_format_date_(employees_cell_(row, col.startDate)),
    eod: employees_format_date_(employees_cell_(row, col.eod)),
    personalEmail: employees_trim_(employees_cell_(row, col.personalEmail))
  };
}

function employees_sanitize_payload_(payload) {
  const src = payload || {};
  const out = {};

  EMPLOYEE_FIELDS.forEach(function (key) {
    if (!Object.prototype.hasOwnProperty.call(src, key)) return;

    if (key === "employeeId") {
      out[key] = employees_norm_id_(src[key]);
      return;
    }

    if (key === "email" || key === "personalEmail") {
      out[key] = employees_trim_(src[key]);
      return;
    }

    if (key === "startDate" || key === "eod") {
      out[key] = employees_parse_date_input_(src[key]);
      return;
    }

    out[key] = employees_trim_(src[key]);
  });

  return out;
}

function employees_validate_payload_(payload, opts) {
  const options = opts || {};
  const requireAllRequired = !!options.requireAllRequired;
  const p = payload || {};
  

  if (requireAllRequired) {
    EMPLOYEE_REQUIRED_FIELDS.forEach(function (key) {
      const val = String(p[key] ?? "").trim();
      if (!val) throw new Error(key + " is required.");
    });
  }

  if (Object.prototype.hasOwnProperty.call(p, "email")) {
    const email = employees_trim_(p.email);
    if (!email) {
      if (requireAllRequired) throw new Error("email is required.");
    } else if (!employees_is_valid_email_(email)) {
      throw new Error("Invalid email.");
    }
  }

  if (Object.prototype.hasOwnProperty.call(p, "personalEmail")) {
    const personalEmail = employees_trim_(p.personalEmail);
    if (personalEmail && !employees_is_valid_email_(personalEmail)) {
      throw new Error("Invalid personalEmail.");
    }
  }
}

function employees_write_payload_to_row_(rowArray, col, payload, onlyDefined) {
  const row = rowArray.slice();

  EMPLOYEE_FIELDS.forEach(function (key) {
    const idx = col[key];
    if (idx < 0) return;

    if (onlyDefined && !Object.prototype.hasOwnProperty.call(payload || {}, key)) return;

    let value = payload && Object.prototype.hasOwnProperty.call(payload, key)
      ? payload[key]
      : "";

    if (key === "employeeId") {
      value = employees_norm_id_(value);
    } else if (key === "email" || key === "personalEmail") {
      value = employees_trim_(value);
    } else if (key === "startDate" || key === "eod") {
      value = employees_parse_date_input_(value);
    } else {
      value = employees_trim_(value);
    }

    row[idx] = value;
  });

  return row;
}

/* =========================================================
 * FILTER / SORT
 * =======================================================*/
function employees_filter_sort_rows_(rows, query, sortKey, asc) {
  const q = employees_norm_(query);
  const key = EMPLOYEE_FIELDS.indexOf(sortKey) >= 0 ? sortKey : "displayName";
  const directionAsc = asc !== false;

  let out = Array.isArray(rows) ? rows.slice() : [];

  if (q) {
    out = out.filter(function (row) {
      const hay = EMPLOYEE_FIELDS.map(function (field) {
        return employees_norm_(row[field]);
      }).join(" ");
      return hay.indexOf(q) >= 0;
    });
  }

  out.sort(function (a, b) {
    const A = a && a[key] != null ? a[key] : "";
    const B = b && b[key] != null ? b[key] : "";

    if (key === "startDate" || key === "eod") {
      const ea = employees_date_to_epoch_(A);
      const eb = employees_date_to_epoch_(B);

      const aValid = Number.isFinite(ea);
      const bValid = Number.isFinite(eb);

      if (!aValid && !bValid) return 0;
      if (!aValid) return 1;
      if (!bValid) return -1;

      return directionAsc ? ea - eb : eb - ea;
    }

    const sa = employees_norm_(A);
    const sb = employees_norm_(B);

    if (sa < sb) return directionAsc ? -1 : 1;
    if (sa > sb) return directionAsc ? 1 : -1;

    const sidA = employees_norm_(a.employeeId);
    const sidB = employees_norm_(b.employeeId);
    if (sidA < sidB) return -1;
    if (sidA > sidB) return 1;
    return 0;
  });

  return out;
}

/* =========================================================
 * EXPORT
 * =======================================================*/
function generateEmployeeExport(config) {
  const headers = Array.isArray(config && config.headers) ? config.headers : [];
  const keys    = Array.isArray(config && config.keys)    ? config.keys    : [];
  const data    = Array.isArray(config && config.data)    ? config.data    : [];

  if (!headers.length) throw new Error("Export failed: headers are missing.");
  if (!keys.length)    throw new Error("Export failed: keys are missing.");
  if (headers.length !== keys.length) {
    throw new Error("Export failed: headers and keys length mismatch.");
  }

  const title      = String(config && config.title ? config.title : "Employee Directory");
  const exportName = "Export - " + title + " (" +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") + ")";

  // Trash any existing file with the same name before creating
  try {
    const existing = DriveApp.getFilesByName(exportName);
    while (existing.hasNext()) existing.next().setTrashed(true);
  } catch (e) {
    console.warn("generateEmployeeExport cleanup failed:", e?.message);
  }

  const ss    = SpreadsheetApp.create(exportName);
  const sheet = ss.getSheets()[0];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (data.length) {
    const rows = data.map(function (item) {
      return keys.map(function (key) {
        const val = item ? item[key] : "";
        if (val instanceof Date && !isNaN(val.getTime())) {
          return Utilities.formatDate(val, Session.getScriptTimeZone(), EMPLOYEE_DATE_FMT);
        }
        return employees_strip_html_(val);
      });
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  const lastCol = Math.max(1, headers.length);
  const lastRow = sheet.getLastRow();

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
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=ISEVEN(ROW())")
      .setBackground("#F5F5F5")
      .setRanges([dataRange])
      .build();
    sheet.setConditionalFormatRules([rule]);
  }

  // Move to shared exports folder to keep Drive tidy
  try {
    const folderName   = "Manila IT Exports";
    const folderSearch = DriveApp.getFoldersByName(folderName);
    const folder       = folderSearch.hasNext()
      ? folderSearch.next()
      : DriveApp.createFolder(folderName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  } catch (e) {
    console.warn("generateEmployeeExport folder move failed:", e?.message);
  }

  return ss.getUrl();
}

/* =========================================================
 * CORE API
 * =======================================================*/
const EmployeesApi = (function () {
  function permissions() {
    try {
      const p = employees_get_rbac_();
      return employees_ok_({
        email: p.email,
        role: p.role,
        pageAllowed: p.pageAllowed,
        canManage: p.canManage
      });
    } catch (e) {
      return employees_fail_(e);
    }
  }

  function ping() {
    try {
      employees_assert_page_access_();

      const sh = employees_require_sheet_();

      employees_audit_("EMPLOYEES_PING", "Employees", "OK");
      return employees_ok_(true, {
        tab: sh.getName(),
        lastRow: sh.getLastRow(),
        lastCol: sh.getLastColumn()
      });
    } catch (e) {
      employees_audit_("EMPLOYEES_PING", "Employees", "ERROR");
      return employees_fail_(e);
    }
  }

  function list() {
    try {
      employees_assert_page_access_();

      const data = employees_load_dataset_();
      employees_audit_("EMPLOYEES_LIST", "Employees rows=" + data.employees.length, "OK");

      return employees_ok_(data.employees);
    } catch (e) {
      employees_audit_("EMPLOYEES_LIST", "Employees", "ERROR");
      return employees_fail_(e);
    }
  }

  function get(payload) {
    try {
      employees_assert_page_access_();

      // payload may be a plain string ID (legacy) or an object { employeeId, _rowIndex }
      const isObj = payload && typeof payload === "object";
      const id    = employees_norm_id_(isObj ? payload.employeeId : payload);
      const rowIndexHint = isObj && payload._rowIndex != null
        ? parseInt(payload._rowIndex, 10)
        : NaN;

      const data = employees_load_dataset_();

      let idx0 = -1;
      if (id && Object.prototype.hasOwnProperty.call(data.idToIndex, id)) {
        idx0 = data.idToIndex[id];
      } else if (Number.isFinite(rowIndexHint) && rowIndexHint >= 0 && rowIndexHint < data.displayRows.length) {
        idx0 = rowIndexHint;
      }

      if (idx0 < 0) {
        throw new Error("Employee not found: " + (id || "no id"));
      }

      const emp = employees_row_to_object_(data.displayRows[idx0], data.col);
      emp._rowIndex = idx0; // always return so frontend can store it

      employees_audit_("EMPLOYEE_GET", "Employees:" + (id || "rowIndex=" + idx0), "OK");
      return employees_ok_(emp);
    } catch (e) {
      employees_audit_("EMPLOYEE_GET", "Employees:" + String(
        (payload && typeof payload === "object" ? payload.employeeId : payload) || ""
      ), "ERROR");
      return employees_fail_(e);
    }
  }
  

 function add(payload) {
    try {
      employees_assert_manage_();

      const data = employees_load_dataset_();
      const sh = data.sh;
      const rawHeaders = data.rawHeaders;
      const col = data.col;

     const clean = employees_sanitize_payload_(payload);
employees_validate_payload_(clean, { requireAllRequired: true });

const id = employees_norm_id_(clean.employeeId);

// Duplicate Employee ID check
if (id && Object.prototype.hasOwnProperty.call(data.idToIndex, id)) {
  throw new Error("Employee ID already exists: " + id);
}

// Duplicate email check — scan all existing employees
const incomingEmail = employees_norm_email_(clean.email);
if (incomingEmail) {
  const emailConflict = data.employees.find(function(emp) {
    return employees_norm_email_(emp.email) === incomingEmail;
  });
  if (emailConflict) {
    throw new Error(
      "Email already exists: " + clean.email +
      (emailConflict.employeeId
        ? " (Employee ID: " + emailConflict.employeeId + ")"
        : " (row " + (emailConflict._rowIndex + 2) + ")")
    );
  }
}

const newRow = new Array(rawHeaders.length).fill("");
      const finalRow = employees_write_payload_to_row_(newRow, col, clean, false);

      sh.appendRow(finalRow);
      _empBustCache_();

      // ── Slack notification ──────────────────────────────
      try {
        const actor = String(Session.getActiveUser().getEmail() || "unknown").trim();
        sendSlackNotification_(null, [
          {
            type: "section",
            text: {
              type: "mrkdwn",
              text: `*🧑‍💼 New Employee Added to Manila IT Inventory*`
            }
          },
          {
            type: "section",
            fields: [
              { type: "mrkdwn", text: `*Employee ID:*\n${id}`                                        },
              { type: "mrkdwn", text: `*Name:*\n${clean.displayName || "—"}`                         },
              { type: "mrkdwn", text: `*Email:*\n${clean.email || "—"}`                              },
              { type: "mrkdwn", text: `*Department:*\n${clean.department || "—"}`                    },
              { type: "mrkdwn", text: `*Title:*\n${clean.title || "—"}`                              },
              { type: "mrkdwn", text: `*Status:*\n${clean.status || "—"}`                            },
              { type: "mrkdwn", text: `*Start Date:*\n${employees_format_date_(clean.startDate) || "—"}` },
              { type: "mrkdwn", text: `*Added by:*\n${actor}`                                        }
            ]
          },
          {
            type: "actions",
            elements: [
              {
                type: "button",
                text: { type: "plain_text", text: "Open Employees Page", emoji: true },
                url: ScriptApp.getService().getUrl() + "?page=employees",
                style: "primary"
              }
            ]
          },
          {
            type: "context",
            elements: [
              {
                type: "mrkdwn",
                text: `🕐 ${new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" })} PHT`
              }
            ]
          }
        ], "#14b8a6"); // teal sidebar — matches employees page accent
      } catch (slackErr) {
        console.error("Slack notify failed (non-blocking):", slackErr);
      }
      // ────────────────────────────────────────────────────

      employees_audit_("EMPLOYEE_ADD", "Employees:" + id, "OK");
      return employees_ok_(true, { employeeId: id });
    } catch (e) {
      employees_audit_("EMPLOYEE_ADD", "Employees:" + (payload && payload.employeeId || ""), "ERROR");
      return employees_fail_(e);
    }
  }



  function update(payload) {
    try {
      employees_assert_manage_();

      const data = employees_load_dataset_();
      const sh = data.sh;
      const rawHeaders = data.rawHeaders;
      const col = data.col;

        const originalId = employees_norm_id_(
  payload && (payload._originalEmployeeId || payload.employeeId)
);

// Resolve once — reused for both the before-snapshot and the write.
// Prefer ID lookup; fall back to _rowIndex for employees with no ID.
const idx0 = (() => {
  if (originalId && Object.prototype.hasOwnProperty.call(data.idToIndex, originalId)) {
    return data.idToIndex[originalId];
  }
  if (payload && payload._rowIndex != null) {
    const ri = parseInt(payload._rowIndex, 10);
    return Number.isFinite(ri) && ri >= 0 && ri < data.displayRows.length ? ri : -1;
  }
  return -1;
})();

// Fail early — before sanitize/validate — so we don't do unnecessary work
if (idx0 < 0) {
  throw new Error("Employee not found. Cannot update without a valid ID or row reference.");
}

// Capture status before update for the Slack offboarding notification
const empBefore    = employees_row_to_object_(data.displayRows[idx0], data.col);
const statusBefore = String(empBefore.status || "").trim().toLowerCase();

const clean = employees_sanitize_payload_(payload);
employees_validate_payload_(clean, { requireAllRequired: true });


      const newId = employees_norm_id_(clean.employeeId) || originalId;
      if (
        newId !== originalId &&
        Object.prototype.hasOwnProperty.call(data.idToIndex, newId)
      ) {
        throw new Error("Cannot change ID. Employee ID already exists: " + newId);
      }

      const baseRow = data.valueRows[idx0] ? data.valueRows[idx0].slice() : new Array(rawHeaders.length).fill("");
      const updatedRow = employees_write_payload_to_row_(baseRow, col, clean, false);

      sh.getRange(idx0 + 2, 1, 1, rawHeaders.length).setValues([updatedRow]);
      _empBustCache_(); 

      // ── Slack notification — only when status changes to Offboarded ──
      const statusAfter = String(clean.status || "").trim().toLowerCase();
      if (statusAfter === "offboarded" && statusBefore !== "offboarded") {
        try {
          const actor = String(Session.getActiveUser().getEmail() || "unknown").trim();
          sendSlackNotification_(null, [
            {
              type: "section",
              text: {
                type: "mrkdwn",
                text: `*🚪 Employee Offboarded from Manila IT Inventory*`
              }
            },
            {
              type: "section",
              fields: [
                { type: "mrkdwn", text: `*Employee ID:*\n${newId}`                     },
                { type: "mrkdwn", text: `*Name:*\n${clean.displayName || "—"}`         },
                { type: "mrkdwn", text: `*Email:*\n${clean.email || "—"}`              },
                { type: "mrkdwn", text: `*Department:*\n${clean.department || "—"}`   },
                { type: "mrkdwn", text: `*Title:*\n${clean.title || "—"}`             },
                { type: "mrkdwn", text: `*Previous Status:*\n${empBefore ? empBefore.status : "—"}` },
                { type: "mrkdwn", text: `*EOD:*\n${employees_format_date_(clean.eod) || "—"}` },
                { type: "mrkdwn", text: `*Updated by:*\n${actor}`                     }
              ]
            },
            {
              type: "actions",
              elements: [
                {
                  type: "button",
                  text: { type: "plain_text", text: "Open Employees Page", emoji: true },
                  url: ScriptApp.getService().getUrl() + "?page=employees",
                  style: "primary"
                }
              ]
            },
            {
              type: "context",
              elements: [
                {
                  type: "mrkdwn",
                  text: `🕐 ${new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" })} PHT`
                }
              ]
            }
          ], "#f59e0b"); // amber sidebar — offboarding event
        } catch (slackErr) {
          console.error("Slack notify failed (non-blocking):", slackErr);
        }
      }
      // ─────────────────────────────────────────────────────────────────

      employees_audit_("EMPLOYEE_UPDATE", "Employees:" + originalId + "→" + newId, "OK");
      return employees_ok_(true, { employeeId: newId });
    } catch (e) {
      employees_audit_(
        "EMPLOYEE_UPDATE",
        "Employees:" + ((payload && (payload._originalEmployeeId || payload.employeeId)) || ""),
        "ERROR"
      );
      return employees_fail_(e);
    }
  }

  

function remove(payload) {
  try {
    employees_assert_delete_();

    // Accept plain ID string (legacy) or object { employeeId, _rowIndex }
    const isObj        = payload && typeof payload === "object";
    const id           = employees_norm_id_(isObj ? payload.employeeId : payload);
    const rowIndexHint = isObj && payload._rowIndex != null
      ? parseInt(payload._rowIndex, 10) : NaN;

    const data = employees_load_dataset_();
    const sh   = data.sh;

    let idx0 = -1;
    if (id && Object.prototype.hasOwnProperty.call(data.idToIndex, id)) {
      idx0 = data.idToIndex[id];
    } else if (Number.isFinite(rowIndexHint) &&
               rowIndexHint >= 0 &&
               rowIndexHint < data.displayRows.length) {
      idx0 = rowIndexHint;
    }

    if (idx0 < 0) {
      employees_audit_("EMPLOYEE_SOFT_DELETE", "Employees:" + (id || "rowIndex=" + rowIndexHint), "NOT_FOUND");
      throw new Error("Employee not found.");
    }

    // ── Soft delete: mark Offboarded + stamp EOD instead of deleting the row ──
    const empNow = employees_row_to_object_(data.displayRows[idx0], data.col);

    // Build update payload — preserve all existing fields, only change status + eod
    const softPayload = Object.assign({}, empNow, {
      status: "Offboarded",
      eod:    empNow.eod && String(empNow.eod).trim()
                ? empNow.eod   // keep existing EOD if already set
                : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy"),
      _originalEmployeeId: id || empNow.employeeId,
      _rowIndex:           idx0
    });

    const baseRow     = data.valueRows[idx0]
      ? data.valueRows[idx0].slice()
      : new Array(data.rawHeaders.length).fill("");
    const updatedRow  = employees_write_payload_to_row_(baseRow, data.col, softPayload, false);

    sh.getRange(idx0 + 2, 1, 1, data.rawHeaders.length).setValues([updatedRow]);
    _empBustCache_();

    employees_audit_("EMPLOYEE_SOFT_DELETE", "Employees:" + (id || "rowIndex=" + idx0), "OK");
    return employees_ok_(true, { employeeId: id, softDeleted: true });

  } catch (e) {
    employees_audit_("EMPLOYEE_SOFT_DELETE", "Employees:" + String(
      (payload && typeof payload === "object" ? payload.employeeId : payload) || ""
    ), "ERROR");
    return employees_fail_(e);
  }
}


function purge(payload) {
  try {
    employees_assert_delete_();

    // Require explicit confirmation flag to prevent accidental hard deletes
    if (!payload || payload.confirm !== true) {
      throw new Error(
        "Hard delete requires confirm: true in the payload. " +
        "Use remove() for soft delete instead."
      );
    }

    const isObj        = payload && typeof payload === "object";
    const id           = employees_norm_id_(payload.employeeId);
    const rowIndexHint = payload._rowIndex != null
      ? parseInt(payload._rowIndex, 10) : NaN;

    const data = employees_load_dataset_();
    const sh   = data.sh;

    let idx0 = -1;
    if (id && Object.prototype.hasOwnProperty.call(data.idToIndex, id)) {
      idx0 = data.idToIndex[id];
    } else if (Number.isFinite(rowIndexHint) &&
               rowIndexHint >= 0 &&
               rowIndexHint < data.displayRows.length) {
      idx0 = rowIndexHint;
    }

    if (idx0 < 0) {
      employees_audit_("EMPLOYEE_HARD_DELETE", "Employees:" + (id || "rowIndex=" + rowIndexHint), "NOT_FOUND");
      throw new Error("Employee not found.");
    }

    sh.deleteRow(idx0 + 2);
    _empBustCache_();

    employees_audit_("EMPLOYEE_HARD_DELETE", "Employees:" + (id || "rowIndex=" + idx0), "OK");
    return employees_ok_(true, { employeeId: id, hardDeleted: true });

  } catch (e) {
    employees_audit_("EMPLOYEE_HARD_DELETE", "Employees:" + String(
      (payload && payload.employeeId) || ""
    ), "ERROR");
    return employees_fail_(e);
  }
}



  function listPaged(payload) {
    try {
      employees_assert_page_access_();

      const page = employees_to_positive_int_(payload && payload.page, 1);
      const pageSize = employees_to_positive_int_(payload && payload.pageSize, 25);
      const query = payload && payload.query ? payload.query : "";
      const sortKey = String(payload && payload.sortKey || "displayName").trim();
      const asc = !(payload && payload.asc === false);

      const data = employees_load_dataset_();
      const rows = employees_filter_sort_rows_(data.employees, query, sortKey, asc);

      const total = rows.length;
      const start = (page - 1) * pageSize;
      const pageRows = rows.slice(start, start + pageSize);

      employees_audit_(
        "EMPLOYEES_LIST_PAGED",
        "q=" + (employees_norm_(query) ? "1" : "0") +
          " page=" + page +
          " size=" + pageSize +
          " total=" + total,
        "OK"
      );

      return employees_ok_(pageRows, {
        total: total,
        page: page,
        pageSize: pageSize
      });
    } catch (e) {
      employees_audit_("EMPLOYEES_LIST_PAGED", "Employees", "ERROR");
      return employees_fail_(e);
    }
  }



  function exportAll(exportRequest) {
    try {
      employees_assert_page_access_();

      const req = exportRequest || {};
      const query = req.query ?? "";
      const sortKey = String(req.sortKey || "displayName").trim();
      const asc = req.asc !== false;

      const dataset = employees_load_dataset_();
      const rows = employees_filter_sort_rows_(dataset.employees, query, sortKey, asc);

      const requestedKeys = Array.isArray(req.keys) ? req.keys : [];
      const keys = requestedKeys.filter(function (k) {
        return EMPLOYEE_FIELDS.indexOf(k) >= 0;
      });

      const finalKeys = keys.length ? keys : EMPLOYEE_FIELDS.slice();
      const headers = finalKeys.map(function (k) {
        return EMPLOYEE_EXPORT_LABELS[k] || k;
      });

      const url = generateEmployeeExport({
        title: req.title || "Employee Directory (All Results)",
        headers: headers,
        keys: finalKeys,
        data: rows
      });

      employees_audit_(
        "EMPLOYEES_EXPORT_ALL",
        "q=" + (employees_norm_(query) ? "1" : "0") + " rows=" + rows.length,
        "OK"
      );

     return employees_ok_(url);
    } catch (e) {
      employees_audit_("EMPLOYEES_EXPORT_ALL", "Employees", "ERROR");
      return employees_fail_(e);
    }
  }

  return {
    permissions: permissions,
    ping: ping,
    list: list,
    get: get,
    add: add,
    update: update,
    remove: remove,
    purge: purge,
    listPaged: listPaged,
    exportAll: exportAll
  };
})();

/* =========================================================
 * GLOBAL WRAPPERS FOR google.script.run
 * =======================================================*/
function employees_get_permissions() {
  return EmployeesApi.permissions();
}

function employees_ping() {
  return EmployeesApi.ping();
}

function employees_list() {
  return EmployeesApi.list();
}

function employees_get(employeeId) {
  return EmployeesApi.get(employeeId);
}

function employees_add(payload) {
  return EmployeesApi.add(payload);
}

function employees_update(payload) {
  return EmployeesApi.update(payload);
}

function employees_remove(payload) {
  return EmployeesApi.remove(payload);
}

function employees_list_paged(payload) {
  return EmployeesApi.listPaged(payload);
}

function employees_export_all(exportRequest) {
  return EmployeesApi.exportAll(exportRequest);
}

function employees_purge(payload) {
  return EmployeesApi.purge(payload);
}