/* ============================================================
 * tasks.gs
 * Full task management — CRUD, bulk, reminders
 * ============================================================ */

const TASKS_SHEET_NAME       = "Tasks";
const REMIND_DAYS_BEFORE     = 1;
const SENDER_ALIAS_EMAIL     = "manila-it@fbgphilippines.com";
const SENDER_ALIAS_NAME      = "FBG Manila IT";
const REPLY_TO_EMAIL         = "manila-it@fbgphilippines.com";

const TASKS_HEADERS = [
  "Employee",
  "Task Description",
  "Date Required",
  "Date Submitted",
  "Status",
  "Note",
  "ID",
  "ReminderSentAt",
  "ReminderKey"
];

const TASKS_REQUIRED_HEADERS = [
  "Employee",
  "Task Description",
  "Date Required",
  "Date Submitted",
  "Status",
  "Note"
];

const TASKS_STATUS = { COMPLETED: "completed" };


/* ============================================================
 * SPREADSHEET BOOTSTRAP
 * Self-contained — no dependency on _openSS_ from other files.
 * ============================================================ */
if (typeof _ssInstance_ === "undefined") var _ssInstance_ = null;

function _getTasksSS_() {
  if (_ssInstance_) return _ssInstance_;
  _ssInstance_ = (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  return _ssInstance_;
}


/* ============================================================
 * CURRENT USER HELPER
 * Falls back to Session if getCurrentUserEmail_ is unavailable.
 * ============================================================ */
function _getTasksActorEmail_(passedEmail) {
  if (passedEmail && String(passedEmail).trim()) return String(passedEmail).trim();
  if (typeof getCurrentUserEmail_ === "function") return getCurrentUserEmail_();
  try { return Session.getActiveUser().getEmail() || ""; } catch (_) { return ""; }
}


/* ============================================================
 * HEADER MAP HELPER
 * Returns { headerName: zeroBasedColumnIndex }
 * Safe to call even if getHeaderMap_ is defined elsewhere —
 * Apps Script uses the last definition, so we guard with typeof.
 * ============================================================ */
if (typeof getHeaderMap_ === "undefined") {
  function getHeaderMap_(sh) {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const map = {};
    headers.forEach((h, i) => {
      const key = String(h || "").trim();
      if (key) map[key] = i;
    });
    return map;
  }
}


/* ============================================================
 * RBAC
 * ============================================================ */
function canViewTasks(email) {
  return canAccessPage("tasks", email);
}

function canManageTasks(email) {
  return canViewTasks(email) && canDoAction("tasks.manage", email);
}

function canBulkTasks(email) {
  return canManageTasks(email) && canDoAction("tasks.bulk", email);
}

function canSendTaskReminders(email) {
  return canManageTasks(email) && canDoAction("tasks.reminders", email);
}

function requireTasksView_(email) {
  if (!canViewTasks(email))
    throw new Error("You do not have permission to access Tasks.");
}

function requireTasksManage_(email) {
  if (!canViewTasks(email))
    throw new Error("You do not have permission to access Tasks.");
  if (!canDoAction("tasks.manage", email))
    throw new Error("You do not have permission to manage Tasks.");
}

function requireTasksBulk_(email) {
  requireTasksManage_(email);
  if (!canDoAction("tasks.bulk", email))
    throw new Error("You do not have permission to use bulk task assignment.");
}

function requireTaskReminders_(email) {
  requireTasksManage_(email);
  if (!canDoAction("tasks.reminders", email))
    throw new Error("You do not have permission to send task reminders.");
}


/* ============================================================
 * SHEET / HEADER HELPERS
 * ============================================================ */
function getTasksSheet_() {
  const ss = _getTasksSS_();
  let sh = ss.getSheetByName(TASKS_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(TASKS_SHEET_NAME);
  ensureTasksHeaders_(sh);
  return sh;
}

function ensureTasksHeaders_(sh) {
  const lastCol   = Math.max(1, sh.getLastColumn());
  const headerRow = sh.getLastRow() >= 1
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim())
    : [];

  const hasAnyHeader = headerRow.some(Boolean);

  if (!hasAnyHeader) {
    sh.getRange(1, 1, 1, TASKS_HEADERS.length).setValues([TASKS_HEADERS]);
    return;
  }

  const existing = new Set(headerRow.map(h => String(h || "").trim()));
  const missing  = TASKS_HEADERS.filter(h => !existing.has(h));

  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function requireHeaders_(headerMap, names) {
  const missing = names.filter(n => !(n in headerMap));
  if (missing.length)
    throw new Error("Tasks sheet missing header(s): " + missing.join(", "));
}

function getTaskColumns_() {
  const sh = getTasksSheet_();
  const hm = getHeaderMap_(sh);

  requireHeaders_(hm, TASKS_REQUIRED_HEADERS);

  return {
    sh,
    hm,
    cEmployee: hm["Employee"],
    cTaskDesc: hm["Task Description"],
    cDateReq:  hm["Date Required"],
    cDateSub:  hm["Date Submitted"],
    cStatus:   hm["Status"],
    cNote:     hm["Note"],
    cId:       ("ID" in hm)             ? hm["ID"]             : null,
    cRemSent:  ("ReminderSentAt" in hm) ? hm["ReminderSentAt"] : null,
    cRemKey:   ("ReminderKey"    in hm) ? hm["ReminderKey"]    : null,
  };
}

function getTaskSheetDataRows_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function ensureValidTaskRowId_(sh, rowId) {
  const n = Number(rowId);
  if (!n || n < 2 || n > sh.getLastRow())
    throw new Error("Invalid task id.");
  return n;
}


/* ============================================================
 * GENERIC HELPERS
 * ============================================================ */
function norm_(value) {
  return String(value || "").trim().toLowerCase();
}

function parseEmployees_(cellValue) {
  return String(cellValue || "").trim()
    .split(/,|\n/g).map(v => v.trim()).filter(Boolean);
}

function joinEmployees_(list) {
  return (Array.isArray(list) ? list : [])
    .map(v => String(v || "").trim()).filter(Boolean).join(", ");
}

function asDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const t = String(value).trim();
  if (!t) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
    const [y, m, d] = t.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  const mdy = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) return new Date(Number(mdy[3]), Number(mdy[1]) - 1, Number(mdy[2]));
  const parsed = new Date(t);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function toIsoDate_(value) {
  if (!value) return "";
  if (value instanceof Date && !isNaN(value.getTime()))
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const t = String(value).trim();
  if (!t) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  const mdy = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) return `${mdy[3]}-${mdy[1].padStart(2,"0")}-${mdy[2].padStart(2,"0")}`;
  const parsed = new Date(t);
  if (isNaN(parsed.getTime())) return "";
  return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function isoToDate_(value) {
  const t = String(value || "").trim();
  if (!t) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
    const [y, m, d] = t.split("-").map(n => parseInt(n, 10));
    return new Date(y, m - 1, d);
  }
  return "";
}

function normalizeTaskRow_(row, sheetRowNumber) {
  return {
    id:              Number(sheetRowNumber),
    employee:        String(row[0] ?? "").trim(),
    taskDescription: String(row[1] ?? "").trim(),
    dateRequired:    toIsoDate_(row[2]),
    dateSubmitted:   toIsoDate_(row[3]),
    status:          String(row[4] ?? "").trim(),
    note:            String(row[5] ?? "").trim(),
  };
}

function buildTaskWriteRow_(payload) {
  const employee        = String(payload?.employee        || "").trim();
  const taskDescription = String(payload?.taskDescription || "").trim();
  const dateRequiredIso = toIsoDate_(payload?.dateRequired);
  const dateSubmittedIso= toIsoDate_(payload?.dateSubmitted);
  const status          = String(payload?.status          || "").trim();
  const note            = String(payload?.note            || "").trim();

  if (!employee || !taskDescription || !dateRequiredIso || !status)
    throw new Error("Missing required fields: Employee, Task Description, Date Required, Status.");

  return {
    employee,
    taskDescription,
    dateRequired:  isoToDate_(dateRequiredIso),
    dateSubmitted: dateSubmittedIso ? isoToDate_(dateSubmittedIso) : "",
    status,
    note,
  };
}

function getUserEmailMap_() {
  const users = (typeof getUsersData === "function" ? getUsersData() : null) || [];
  const map   = {};
  users.forEach(user => {
    const name  = String(user.username || user.name  || "").trim();
    const email = String(user.email                  || "").trim();
    if (name && email) map[norm_(name)] = email;
  });
  return map;
}

function getTaskRowById_(taskId) {
  const cols  = getTaskColumns_();
  const rowId = ensureValidTaskRowId_(cols.sh, taskId);
  const row   = cols.sh.getRange(rowId, 1, 1, cols.sh.getLastColumn()).getValues()[0];
  return { cols, rowId, row };
}

function syncTaskIds_() {
  const sh        = getTasksSheet_();
  const headerMap = getHeaderMap_(sh);
  if (!("ID" in headerMap)) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  const values = [];
  for (let row = 2; row <= lastRow; row++) values.push([row]);
  sh.getRange(2, headerMap["ID"] + 1, values.length, 1).setValues(values);
}


/* ============================================================
 * PUBLIC CRUD
 * ============================================================ */

/** Called by tasks.html — no email arg needed, resolved server-side. */
function getTasksData(email) {
  const e = _getTasksActorEmail_(email);
  requireTasksView_(e);

  const sh      = getTasksSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  return sh.getRange(2, 1, lastRow - 1, 6).getValues()
    .map((row, i) => normalizeTaskRow_(row, i + 2));
}

function addTask(payload) {
  const e = _getTasksActorEmail_(payload?.actorEmail);
  requireTasksManage_(e);

  const sh     = getTasksSheet_();
  const record = buildTaskWriteRow_(payload);

  sh.appendRow([
    record.employee,
    record.taskDescription,
    record.dateRequired  || "",
    record.dateSubmitted || "",
    record.status,
    record.note,
    "", "", ""
  ]);

  const newRow    = sh.getLastRow();
  const headerMap = getHeaderMap_(sh);
  if ("ID" in headerMap) sh.getRange(newRow, headerMap["ID"] + 1).setValue(newRow);

  return { ok: true, id: newRow };
}

function updateTask(payload) {
  const e = _getTasksActorEmail_(payload?.actorEmail);
  requireTasksManage_(e);

  const sh        = getTasksSheet_();
  const rowId     = ensureValidTaskRowId_(sh, payload?.id);
  const record    = buildTaskWriteRow_(payload);
  const headerMap = getHeaderMap_(sh);

  sh.getRange(rowId, 1, 1, 6).setValues([[
    record.employee,
    record.taskDescription,
    record.dateRequired  || "",
    record.dateSubmitted || "",
    record.status,
    record.note,
  ]]);

  if ("ID" in headerMap) sh.getRange(rowId, headerMap["ID"] + 1).setValue(rowId);

  return { ok: true, id: rowId };
}

function deleteTask(rowId, actorEmail) {
  const e = _getTasksActorEmail_(actorEmail);
  requireTasksManage_(e);

  const sh         = getTasksSheet_();
  const validRowId = ensureValidTaskRowId_(sh, rowId);
  sh.deleteRow(validRowId);
  syncTaskIds_();

  return { ok: true };
}

function addTaskBulkOneRow(payload) {
  const e = _getTasksActorEmail_(payload?.actorEmail);
  requireTasksBulk_(e);

  return addTask({
    employee:        joinEmployees_(payload?.employees),
    taskDescription: payload?.taskDescription,
    dateRequired:    payload?.dateRequired,
    dateSubmitted:   payload?.dateSubmitted,
    status:          payload?.status,
    note:            payload?.note,
    actorEmail:      e,
  });
}

function updateTaskBulkOneRow(payload) {
  const e = _getTasksActorEmail_(payload?.actorEmail);
  requireTasksBulk_(e);

  return updateTask({
    id:              payload?.id,
    employee:        joinEmployees_(payload?.employees),
    taskDescription: payload?.taskDescription,
    dateRequired:    payload?.dateRequired,
    dateSubmitted:   payload?.dateSubmitted,
    status:          payload?.status,
    note:            payload?.note,
    actorEmail:      e,
  });
}


/* ============================================================
 * REMINDERS
 * ============================================================ */

/** Scheduled trigger handler — no RBAC, runs as service account. */
function sendTaskDueReminders() {
  return sendTaskDueRemindersWithStats_(REMIND_DAYS_BEFORE);
}

function installTaskReminderTrigger(actorEmail) {
  const e = _getTasksActorEmail_(actorEmail);
  requireTaskReminders_(e);

  const handler = "sendTaskDueReminders";
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === handler)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(9).create();
  return { ok: true };
}

function sendTaskDueRemindersNow(daysBefore, actorEmail) {
  const e = _getTasksActorEmail_(actorEmail);
  requireTaskReminders_(e);

  const n = Number(daysBefore);
  return sendTaskDueRemindersWithStats_(Number.isFinite(n) && n >= 0 ? n : REMIND_DAYS_BEFORE);
}

function sendTaskDueRemindersWithStats_(daysBefore) {
  const cols = getTaskColumns_();
  const sh   = cols.sh;

  if (cols.cRemKey == null)
    throw new Error("Tasks sheet missing header: ReminderKey");

  const rows = getTaskSheetDataRows_(sh);
  if (!rows.length) return { sentCount: 0, rowsUpdated: 0, debug: { reason: "No data rows" } };

  const tz         = Session.getScriptTimeZone() || "Asia/Manila";
  const today      = new Date();
  const targetDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + Number(daysBefore || 0));
  const targetKey  = Utilities.formatDate(targetDate, tz, "yyyy-MM-dd");

  const emailMap         = getUserEmailMap_();
  const recipientBucket  = {};
  const pendingUpdates   = [];
  const missingEmailNames= new Set();
  let matchedDueRows     = 0;

  rows.forEach((row, index) => {
    const rowNumber = index + 2;
    const status    = norm_(row[cols.cStatus]);
    if (status === TASKS_STATUS.COMPLETED) return;

    const dueDate = asDate_(row[cols.cDateReq]);
    if (!dueDate) return;

    const dueKey = Utilities.formatDate(dueDate, tz, "yyyy-MM-dd");
    if (dueKey !== targetKey) return;

    matchedDueRows++;

    const wantedKey = `${dueKey}|${daysBefore}`;
    if (String(row[cols.cRemKey] || "").trim() === wantedKey) return;

    const employees       = parseEmployees_(row[cols.cEmployee]);
    if (!employees.length) return;

    const taskDescription = String(row[cols.cTaskDesc] || "").trim();
    const note            = cols.cNote != null ? String(row[cols.cNote] || "").trim() : "";
    let queued            = false;

    employees.forEach(name => {
      const email = emailMap[norm_(name)];
      if (!email) { missingEmailNames.add(name); return; }
      if (!recipientBucket[email]) recipientBucket[email] = [];
      recipientBucket[email].push(
        `• ${taskDescription}${note ? ` (Note: ${note})` : ""} — Due: ${dueKey}`
      );
      queued = true;
    });

    if (queued) pendingUpdates.push({ rowNumber, reminderKeyWanted: wantedKey });
  });

  let sentCount = 0;
  Object.entries(recipientBucket).forEach(([recipientEmail, lines]) => {
    if (!lines?.length) return;
    sendEmailFromAliasSafe_(
      recipientEmail,
      `Task Reminder: Due in ${daysBefore} day(s)`,
      `Hi,\n\nThis is a reminder for task(s) due on ${targetKey}:\n\n${lines.join("\n")}\n\n— Manila IT`
    );
    sentCount++;
  });

  let rowsUpdated = 0;
  pendingUpdates.forEach(u => {
    if (cols.cRemSent != null) sh.getRange(u.rowNumber, cols.cRemSent + 1).setValue(new Date());
    sh.getRange(u.rowNumber, cols.cRemKey + 1).setValue(u.reminderKeyWanted);
    rowsUpdated++;
  });

  return {
    sentCount,
    rowsUpdated,
    debug: {
      targetKey,
      matchedDueRows,
      distinctEmails: Object.keys(recipientBucket).length,
      missingEmailNames: Array.from(missingEmailNames).slice(0, 25),
    },
  };
}

function sendTaskReminderForTask(taskId, actorEmail) {
  const e        = _getTasksActorEmail_(actorEmail);
  requireTaskReminders_(e);

  const { cols, rowId, row } = getTaskRowById_(taskId);
  const sh = cols.sh;

  const employeeCell    = String(row[cols.cEmployee] || "").trim();
  const taskDescription = String(row[cols.cTaskDesc] || "").trim();
  const status          = String(row[cols.cStatus]   || "").trim();
  const note            = cols.cNote != null ? String(row[cols.cNote] || "").trim() : "";

  if (!employeeCell)    throw new Error("Selected task has no employee assigned.");
  if (!taskDescription) throw new Error("Selected task has no task description.");

  const dueDate = asDate_(row[cols.cDateReq]);
  const dueKey  = dueDate
    ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
    : "";

  const employees = parseEmployees_(employeeCell);
  if (!employees.length) throw new Error("Selected task has no valid employee names.");

  const emailMap          = getUserEmailMap_();
  const missingEmailNames = [];
  const recipients        = [];

  employees.forEach(name => {
    const email = emailMap[norm_(name)];
    email ? recipients.push({ name, email }) : missingEmailNames.push(name);
  });

  if (!recipients.length)
    throw new Error("No email found for: " + missingEmailNames.join(", "));

  const subject = `Task Reminder: ${taskDescription}`;
  const body    =
    `Hi,\n\nThis is a reminder for the following task:\n\n` +
    `Task: ${taskDescription}\n` +
    `Assigned to: ${employees.join(", ")}\n` +
    `Due Date: ${dueKey || "N/A"}\n` +
    `Status: ${status || "Pending"}\n` +
    `${note ? `Note: ${note}\n` : ""}\n` +
    `— Manila IT`;

  let sentCount = 0;
  recipients.forEach(r => { sendEmailFromAliasSafe_(r.email, subject, body); sentCount++; });

  const stamp      = new Date();
  const reminderKey= `manual|row:${rowId}|${Utilities.formatDate(stamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")}`;

  if (cols.cRemSent != null) sh.getRange(rowId, cols.cRemSent + 1).setValue(stamp);
  if (cols.cRemKey  != null) sh.getRange(rowId, cols.cRemKey  + 1).setValue(reminderKey);

  return {
    ok: true, taskId: rowId, sentCount, taskDescription,
    recipients: recipients.map(r => r.email),
    missingEmailNames,
  };
}

function getTaskReminderOptions(actorEmail) {
  const e = _getTasksActorEmail_(actorEmail);
  requireTaskReminders_(e);

  const sh   = getTasksSheet_();
  const rows = getTaskSheetDataRows_(sh);
  if (!rows.length) return [];

  const options = rows.reduce((acc, row, index) => {
    const rowId  = index + 2;
    const status = norm_(row[4]);
    if (status === TASKS_STATUS.COMPLETED) return acc;

    const employee        = String(row[0] || "").trim();
    const taskDescription = String(row[1] || "").trim();
    const due             = toIsoDate_(row[2]);
    const note            = String(row[5] || "").trim();

    acc.push({
      id: rowId,
      employee,
      taskDescription,
      dateRequired: due,
      status: String(row[4] || "").trim(),
      note,
      label: `${taskDescription} — ${employee}${due ? ` — Due: ${due}` : ""}`,
    });
    return acc;
  }, []);

  options.sort((a, b) => {
    const tA = asDate_(a.dateRequired)?.getTime() || 0;
    const tB = asDate_(b.dateRequired)?.getTime() || 0;
    return tA - tB;
  });

  return options;
}


/* ============================================================
 * EMAIL HELPER
 * ============================================================ */
function sendEmailFromAliasSafe_(to, subject, body) {
  const opts = { name: SENDER_ALIAS_NAME, replyTo: REPLY_TO_EMAIL };
  if (SENDER_ALIAS_EMAIL && String(SENDER_ALIAS_EMAIL).trim())
    opts.from = String(SENDER_ALIAS_EMAIL).trim();

  try {
    GmailApp.sendEmail(to, subject, body, opts);
  } catch (_) {
    MailApp.sendEmail(to, subject, body);
  }
}


/* ============================================================
 * CURRENT USER — MY TASKS (profile page / personal view)
 * ============================================================ */
function getMyAssignedTasks(username, email) {
  const e  = _getTasksActorEmail_(email);
  requireTasksView_(e);

  const me = norm_(username);
  if (!me) return [];

  const sh      = getTasksSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const output = sh.getRange(2, 1, lastRow - 1, 6).getValues().reduce((acc, row, index) => {
    const employees = parseEmployees_(row[0]);
    if (employees.some(n => norm_(n) === me)) acc.push(normalizeTaskRow_(row, index + 2));
    return acc;
  }, []);

  output.sort((a, b) => {
    const tA = asDate_(a.dateRequired)?.getTime() || 0;
    const tB = asDate_(b.dateRequired)?.getTime() || 0;
    return tA - tB;
  });

  return output;
}


/* ============================================================
 * USERNAME DROPDOWN  (used by tasks.html modal)
 * ============================================================ */
function getUsernamesForDropdown() {
  const users = (typeof getUsersData === "function" ? getUsersData() : null) || [];
  return users
    .map(u => String(u.username || u.name || "").trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
}