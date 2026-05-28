/************************************************************
 * Events.gs
 * Manila IT Inventory - Events Kanban Backend
 * Clean rewrite with safer sheet handling and better structure
 ************************************************************/

const EVENTS_SHEET = "Events";
const EVENTS_HEADERS = ["Name", "Date", "Category", "Assignee", "Status", "Notes"];
const AUTO_CREATE_EVENTS_SHEET = false;

const EVENT_STATUS = {
  UPCOMING: "Upcoming",
  IN_PROGRESS: "In Progress",
  COMPLETED: "Completed"
};

const EVENT_CATEGORY_MAP = {
  "meeting / deployment": "tag-blue",
  "holiday / off-day": "tag-purple",
  "inventory / audit": "tag-orange",
  "other": "tag-green"
};

/* =========================================================
 * CORE SHEET HELPERS
 * =======================================================*/
function getSpreadsheetTimeZone_() {
  return getSS_().getSpreadsheetTimeZone();
}

function getEventsSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(EVENTS_SHEET);

  if (!sh && AUTO_CREATE_EVENTS_SHEET) {
    sh = ss.insertSheet(EVENTS_SHEET);
    sh.getRange(1, 1, 1, EVENTS_HEADERS.length).setValues([EVENTS_HEADERS]);
    sh.setFrozenRows(1);
  }

  if (!sh) {
    throw new Error(`Sheet "${EVENTS_SHEET}" not found.`);
  }

  ensureEventsHeader_(sh);
  return sh;
}

function ensureEventsHeader_(sheet) {
  const needed = EVENTS_HEADERS.slice();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), needed.length);

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, needed.length).setValues([needed]);
    sheet.setFrozenRows(1);
    return;
  }

  const current = sheet.getRange(1, 1, 1, needed.length).getValues()[0];
  const mismatch = needed.some((h, i) => String(current[i] || "").trim() !== h);

  if (mismatch) {
    sheet.getRange(1, 1, 1, needed.length).setValues([needed]);
  }

  if (sheet.getFrozenRows() !== 1) {
    sheet.setFrozenRows(1);
  }
}

function getEventsDataRangeValues_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, EVENTS_HEADERS.length).getValues();
}

/* =========================================================
 * NORMALIZERS / FORMATTERS
 * =======================================================*/
function normalizeText_(value) {
  return String(value == null ? "" : value)
    .replace(/\u00A0/g, " ")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function safeCell_(value) {
  return value == null ? "" : value;
}

function formatDateKey_(value) {
  const tz = getSpreadsheetTimeZone_();

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, "yyyy-MM-dd");
  }

  const raw = String(value == null ? "" : value).trim();
  if (!raw) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return raw;
  }

  const parsed = new Date(raw);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
  }

  return "";
}

function normalizeStatus_(status) {
  const value = normalizeText_(status);

  if (value === "in progress" || value === "progress" || value === "inprogress") {
    return EVENT_STATUS.IN_PROGRESS;
  }

  if (value === "completed" || value === "done" || value === "complete") {
    return EVENT_STATUS.COMPLETED;
  }

  return EVENT_STATUS.UPCOMING;
}

function normalizeCategoryLabel_(value) {
  const raw = normalizeText_(value);

  if (!raw) return "Other";
  if (raw.includes("meeting") || raw.includes("deploy")) return "Meeting / Deployment";
  if (raw.includes("holiday") || raw.includes("off")) return "Holiday / Off-day";
  if (raw.includes("inventory") || raw.includes("audit")) return "Inventory / Audit";
  return "Other";
}

function getCategoryTag_(categoryLabel) {
  const key = normalizeText_(normalizeCategoryLabel_(categoryLabel));
  return EVENT_CATEGORY_MAP[key] || "tag-green";
}

function escapeEventNameKey_(value) {
  return normalizeText_(value);
}

/* =========================================================
 * VALIDATION
 * =======================================================*/
function validateEventPayload_(data) {
  const name = String(data && data.name != null ? data.name : "").trim();
  const dateKey = formatDateKey_(data && data.date);
  const categoryLabel = normalizeCategoryLabel_(data && data.categoryLabel);
  const assignee = String(data && data.assignee != null ? data.assignee : "").trim();
  const notes = String(data && data.notes != null ? data.notes : "").trim();
  const oldName = String(data && data.oldName != null ? data.oldName : "").trim();

  if (!name) throw new Error("Event name is required.");
  if (!dateKey) throw new Error("Valid event date is required.");

  return {
    name,
    date: dateKey,
    categoryLabel,
    assignee,
    notes,
    oldName
  };
}

/* =========================================================
 * ROW HELPERS
 * =======================================================*/
function getEventRowObject_(row) {
  const name = String(row[0] == null ? "" : row[0]).trim();
  if (!name) return null;

  const categoryLabel = normalizeCategoryLabel_(row[2]);

  return {
    name: name,
    date: formatDateKey_(row[1]),
    categoryLabel: categoryLabel,
    assignee: String(row[3] == null ? "" : row[3]).trim(),
    status: normalizeStatus_(row[4]),
    notes: String(row[5] == null ? "" : row[5]).trim(),
    category: getCategoryTag_(categoryLabel)
  };
}

function findEventRowByName_(sheet, eventName) {
  const key = escapeEventNameKey_(eventName);
  if (!key) return -1;

  const values = getEventsDataRangeValues_(sheet);
  if (!values.length) return -1;

  for (let i = 0; i < values.length; i++) {
    if (escapeEventNameKey_(values[i][0]) === key) {
      return i + 2;
    }
  }

  return -1;
}

function eventExistsByName_(sheet, eventName, excludeRow) {
  const key = escapeEventNameKey_(eventName);
  if (!key) return false;

  const values = getEventsDataRangeValues_(sheet);
  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    if (excludeRow && rowNumber === excludeRow) continue;

    if (escapeEventNameKey_(values[i][0]) === key) {
      return true;
    }
  }

  return false;
}

function sortEventsOutput_(items) {
  const statusOrder = {
    "Upcoming": 1,
    "In Progress": 2,
    "Completed": 3
  };

  return items.sort((a, b) => {
    const sa = statusOrder[a.status] || 99;
    const sb = statusOrder[b.status] || 99;
    if (sa !== sb) return sa - sb;

    const da = a.date || "";
    const db = b.date || "";
    if (da !== db) return da.localeCompare(db);

    return String(a.name || "").localeCompare(String(b.name || ""));
  });
}

/* =========================================================
 * EVENTS API
 * =======================================================*/
function getEventsFromSheet() {
  const sh = getEventsSheet_();
  const values = getEventsDataRangeValues_(sh);
  const output = [];

  for (let i = 0; i < values.length; i++) {
    const item = getEventRowObject_(values[i]);
    if (item) output.push(item);
  }

  return sortEventsOutput_(output);
}

function saveEventToSheet(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sh = getEventsSheet_();
    const payload = validateEventPayload_(data);

    const lookupName = payload.oldName || payload.name;
    const hitRow = findEventRowByName_(sh, lookupName);

    if (hitRow > 0) {
      if (payload.oldName && payload.name !== payload.oldName) {
        if (eventExistsByName_(sh, payload.name, hitRow)) {
          throw new Error("Another event with the same name already exists.");
        }
      }

      const currentStatus = normalizeStatus_(sh.getRange(hitRow, 5).getValue());

      sh.getRange(hitRow, 1, 1, EVENTS_HEADERS.length).setValues([[
        safeCell_(payload.name),
        safeCell_(payload.date),
        safeCell_(payload.categoryLabel),
        safeCell_(payload.assignee),
        safeCell_(currentStatus),
        safeCell_(payload.notes)
      ]]);

      return {
        ok: true,
        mode: "update",
        name: payload.name
      };
    }

    if (eventExistsByName_(sh, payload.name)) {
      throw new Error("An event with the same name already exists.");
    }

    sh.appendRow([
      safeCell_(payload.name),
      safeCell_(payload.date),
      safeCell_(payload.categoryLabel),
      safeCell_(payload.assignee),
      EVENT_STATUS.UPCOMING,
      safeCell_(payload.notes)
    ]);

    return {
      ok: true,
      mode: "create",
      name: payload.name
    };
  } finally {
    lock.releaseLock();
  }
}

function updateEventStatus(eventName, newStatus) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sh = getEventsSheet_();
    const row = findEventRowByName_(sh, eventName);

    if (row < 2) {
      return false;
    }

    sh.getRange(row, 5).setValue(normalizeStatus_(newStatus));
    return true;
  } finally {
    lock.releaseLock();
  }
}

function deleteEventFromSheet(eventName) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sh = getEventsSheet_();
    const row = findEventRowByName_(sh, eventName);

    if (row < 2) {
      return false;
    }

    sh.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/* =========================================================
 * USERS DROPDOWN API
 * =======================================================*/
function getUserNames() {
  return getUsernamesForDropdown();
}

function getUsernamesForDropdown() {
  const sh = getUsersSheet_();
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues().flat();

  return [...new Set(
    values
      .map(v => String(v == null ? "" : v).trim())
      .filter(Boolean)
  )].sort((a, b) => a.localeCompare(b));
}