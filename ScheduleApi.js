const AUDIT_SHEET = "AuditLogs";
const SHIFTS_SHEET = "Shifts";
const SHIFT_HEADERS = ["Role", "Name", "Mode", "Group", "Date", "Start", "End", "Notes"];
const USERS_SHEET = "Users";

const SHIFT_COLS = {
  ROLE:  0,
  NAME:  1,
  MODE:  2,
  GROUP: 3,
  DATE:  4,
  START: 5,
  END:   6,
  NOTES: 7
};


// =========================
// COMMON HELPERS
// =========================
function _getDbSs_() {
  return (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function _actorEmail_(fallback) {
  return String(fallback || Session.getActiveUser().getEmail() || "unknown").trim();
}

function _dateKey_(v) {
  const tz = Session.getScriptTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, "yyyy-MM-dd");
  return String(v || "").trim();
}

function _timeStr_(v) {
  const tz = Session.getScriptTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, "hh:mm a");
  return String(v || "").trim();
}

function toDateKey(date) {
  const tz = Session.getScriptTimeZone();
  const d = (date instanceof Date) ? date : new Date(date);
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

function _dateObjFromCell_(cell) {
  const key = _dateKey_(cell);
  if (!key) return null;
  const d = new Date(key + "T00:00:00");
  d.setHours(0, 0, 0, 0);
  return d;
}

function _normName_(s) {
  return String(s || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function _ensureSheetHeaders_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const lastCol = headers.length;

  if (lastRow < 1) {
    sheet.getRange(1, 1, 1, lastCol).setValues([headers]);
    return;
  }

  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || "").trim());
  const need = headers.some((h, i) => current[i] !== h);
  if (need) {
    sheet.getRange(1, 1, 1, lastCol).setValues([headers]);
  }
}

function _getShiftsSheet_() {
  const ss = _getDbSs_();
  let sh = ss.getSheetByName(SHIFTS_SHEET);
  if (!sh) sh = ss.insertSheet(SHIFTS_SHEET);
  _ensureSheetHeaders_(sh, SHIFT_HEADERS);
  return sh;
}

function _ensureAuditSheet_() {
  const ss = _getDbSs_();
  let sh = ss.getSheetByName(AUDIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AUDIT_SHEET);
    sh.appendRow(["Timestamp", "User", "Action", "Resource", "Status"]);
  }
  return sh;
}



function _withScheduleLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// =========================
// CACHE HELPERS
// =========================
function _cacheGetJson_(key) {
  const raw = CacheService.getScriptCache().get(key);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (_) {
    return null;
  }
}

function _cachePutJson_(key, value, seconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), seconds || 30);
  } catch (_) {}
}

function clearScheduleCaches_() {
  const cache = CacheService.getScriptCache();
  [
    "schedule_shifts_v1",
    "schedule_team_v1",
    "schedule_bootstrap_v1"
  ].forEach(k => cache.remove(k));
}

// =========================
// FAST SHEET READERS
// =========================
function _readShiftRows_() {
  const sh = _getShiftsSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, SHIFT_HEADERS.length).getValues();
}

function _readUserRows_() {
  const ss = _getDbSs_();
  const sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) return { headers: [], rows: [] };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { headers: [], rows: [] };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim().toLowerCase());
  const rows    = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return { headers, rows };
}

function _buildShiftIndex_(rows) {
  const index = Object.create(null);
  rows.forEach((row, i) => {
    const key = `${String(row[SHIFT_COLS.NAME] || "").trim()}|${_dateKey_(row[SHIFT_COLS.DATE])}`;
    if (key !== "|") index[key] = i;
  });
  return index;
}

// =========================
// READ SHIFTS
// =========================
function getEmployeesFromShifts() {
  const cached = _cacheGetJson_("schedule_shifts_v1");
  if (cached) return cached;

  const rows = _readShiftRows_();
  const out = rows
    .map(row => ({
      role:  String(row[SHIFT_COLS.ROLE]  || ""),
      name:  String(row[SHIFT_COLS.NAME]  || ""),
      mode:  String(row[SHIFT_COLS.MODE]  || ""),
      group: String(row[SHIFT_COLS.GROUP] || "Morning"),
      date:  _dateKey_(row[SHIFT_COLS.DATE]),
      start: _timeStr_(row[SHIFT_COLS.START]),
      end:   _timeStr_(row[SHIFT_COLS.END]),
      notes: String(row[SHIFT_COLS.NOTES] || "")
    }))
    .filter(r => r.name);

  _cachePutJson_("schedule_shifts_v1", out, 30);
  return out;
}

// =========================
// USERS / TEAM
// =========================
function getTeamList() {
  const cached = _cacheGetJson_("schedule_team_v1");
  if (cached) return cached;

  const { headers, rows } = _readUserRows_();

  // Resolve column indices from headers — falls back to legacy fixed indices
  const col = (names, fallback) => {
    for (const n of names) {
      const i = headers.indexOf(n.toLowerCase());
      if (i > -1) return i;
    }
    return fallback;
  };

  const iName     = col(["username", "display name", "name"],  0);
  const iEmail    = col(["email"],                              1);
  const iStatus   = col(["status"],                            5);
  const iRole     = col(["role"],                              6);
  const iPosition = col(["position"],                          8);
  const iGroup    = col(["group", "shift group"],              9);
  const iAvatar   = col(["avatarurl", "avatar url", "avatar"], 11);

  const out = rows
    .map(row => ({
      name:      String(row[iName]     || "").trim(),
      email:     String(row[iEmail]    || "").trim(),
      status:    String(row[iStatus]   || "").trim(),
      role:      String(row[iRole]     || "Staff").trim(),
      position:  String(row[iPosition] || "").trim(),
      group:     String(row[iGroup]    || "Morning").trim(),
      avatarUrl: String(row[iAvatar]   || "").trim()
    }))
    .filter(u => u.name && String(u.status).toLowerCase() === "active")
    .sort((a, b) => {
      const weight = { morning: 1, graveyard: 2 };
      const wa = weight[String(a.group || "").toLowerCase()] || 99;
      const wb = weight[String(b.group || "").toLowerCase()] || 99;
      if (wa !== wb) return wa - wb;
      return String(a.name || "").localeCompare(String(b.name || ""));
    });

  _cachePutJson_("schedule_team_v1", out, 60);
  return out;
}

function getMyScheduleIdentity() {
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!email) return { email: "", name: "" };

  const users = getTeamList();
  const me = users.find(u => String(u.email || "").trim().toLowerCase() === email);

  return {
    email,
    name:  String(me?.name  || "").trim(),
    group: String(me?.group || "").trim()
  };
}

function getScheduleBootstrap() {
  const cached = _cacheGetJson_("schedule_bootstrap_v1");
  if (cached) return cached;

  const payload = {
    perms: getMyPermissions(),
    me: getMyScheduleIdentity(),
    users: getTeamList(),
    shifts: getEmployeesFromShifts()
  };

  _cachePutJson_("schedule_bootstrap_v1", payload, 20);
  return payload;
}

// =========================
// WRITE HELPERS
// =========================
function _trimExtraRows_(sheet, desiredLastRow) {
  const currentLastRow = sheet.getLastRow();
  if (currentLastRow > desiredLastRow) {
    sheet.deleteRows(desiredLastRow + 1, currentLastRow - desiredLastRow);
  }
}

function _writeAllShiftRows_(rows) {
  const sh = _getShiftsSheet_();
  const lastCol = SHIFT_HEADERS.length;
  const existingLastRow = sh.getLastRow();

  sh.getRange(1, 1, 1, lastCol).setValues([SHIFT_HEADERS]);

  if (existingLastRow > 1) {
    sh.getRange(2, 1, existingLastRow - 1, lastCol).clearContent();
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, lastCol).setValues(rows);
  }

  _trimExtraRows_(sh, rows.length + 1);
  clearScheduleCaches_();
}

function _upsertShifts_(entries) {
  const sh = _getShiftsSheet_();
  const rows = _readShiftRows_();
  const index = _buildShiftIndex_(rows);

  const updates = [];
  const appends = [];

  entries.forEach(entry => {
    const key = `${String(entry[SHIFT_COLS.NAME] || "").trim()}|${String(entry[SHIFT_COLS.DATE] || "").trim()}`;
    if (index[key] !== undefined) {
      updates.push({ rowNumber: index[key] + 2, values: entry });
    } else {
      appends.push(entry);
    }
  });

  updates.forEach(u => {
    sh.getRange(u.rowNumber, 1, 1, SHIFT_HEADERS.length).setValues([u.values]);
  });

  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, SHIFT_HEADERS.length).setValues(appends);
  }

  clearScheduleCaches_();
  return true;
}

// =========================
// SAVE / UPSERT SINGLE SHIFT
// =========================
function saveShiftToSheet(data) {
  return _withScheduleLock_(() => {
    const actor = _actorEmail_(data?.actorEmail);

    try {
     const row = [
        String(data.role  || ""),
        String(data.name  || ""),
        String(data.mode  || ""),
        String(data.group || "Morning"),
        String(data.date  || ""),
        String(data.start || ""),
        String(data.end   || ""),
        String(data.notes || "")
      ];

      _upsertShifts_([row]);
      auditWrite(actor, "Save Shift", `${data.name} @ ${data.date}`, "Success");
      return true;
    } catch (e) {
      auditWrite(actor, "Save Shift", `${data?.name || ""} @ ${data?.date || ""}`, "Failed");
      throw e;
    }
  });
}

// =========================
// BULK APPLY SHIFTS
// =========================
function saveBulkShiftsToSheet(batchData) {
  return _withScheduleLock_(() => {
    if (!canDoAction("schedule.manage")) {
      throw new Error("Access denied: schedule.manage");
    }

    const actor = _actorEmail_(batchData?.actorEmail);

    try {
      const nameKey = String(batchData.name || "").trim();
      const dates = Array.isArray(batchData.dates) ? batchData.dates.filter(Boolean) : [];

      const entries = dates.map(dateKey => ([
        String(batchData.role  || ""),
        nameKey,
        String(batchData.mode  || ""),
        String(batchData.group || "Morning"),
        String(dateKey || "").trim(),
        String(batchData.start || ""),
        String(batchData.end   || ""),
        String(batchData.notes || "")
      ]));

      _upsertShifts_(entries);
      auditWrite(actor, "Bulk Apply Shifts", `${nameKey} (${entries.length} day(s))`, "Success");
      return true;
    } catch (e) {
      auditWrite(actor, "Bulk Apply Shifts", `${batchData?.name || ""}`, "Failed");
      throw e;
    }
  });
}

// =========================
// CLEAR RANGE
// =========================
function clearScheduleForRange(config) {
  return _withScheduleLock_(() => {
    if (!canDoAction("schedule.clear")) {
      throw new Error("Access denied: schedule.clear");
    }

    const actor = _actorEmail_(config?.actorEmail);

    try {
      const start = new Date(config.startDate);
      start.setHours(0, 0, 0, 0);

      const daysToClear = (config.view === "month")
        ? new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate()
        : 7;

      const end = new Date(start);
      end.setDate(start.getDate() + daysToClear - 1);
      end.setHours(23, 59, 59, 999);

      const kept = _readShiftRows_().filter(row => {
        const d = _dateObjFromCell_(row[SHIFT_COLS.DATE]);
        return !(d && d >= start && d <= end);
      });

      _writeAllShiftRows_(kept);
      auditWrite(actor, "Clear Schedule Range", `${config.view || "week"} starting ${toDateKey(start)}`, "Success");
      return true;
    } catch (err) {
      auditWrite(actor, "Clear Schedule Range", `${config?.view || ""} starting ${config?.startDate || ""}`, "Failed");
      throw err;
    }
  });
}

// =========================
// CLEAR EMPLOYEE
// =========================
function clearScheduleForEmployee(config) {
  return _withScheduleLock_(() => {
    if (!canDoAction("schedule.clear")) {
      throw new Error("Access denied: schedule.clear");
    }

    const actor = _actorEmail_(config?.actorEmail);

    try {
      const targetName = String(config.name || "").trim();
      const start = new Date(String(config.startDate || "") + "T00:00:00");
      start.setHours(0, 0, 0, 0);

      const daysToClear = (config.view === "month")
        ? new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate()
        : 7;

      const end = new Date(start);
      end.setDate(start.getDate() + daysToClear - 1);
      end.setHours(23, 59, 59, 999);

      const kept = _readShiftRows_().filter(row => {
        const rowName = String(row[SHIFT_COLS.NAME] || "").trim();
        if (rowName !== targetName) return true;

        const d = _dateObjFromCell_(row[SHIFT_COLS.DATE]);
        return !(d && d >= start && d <= end);
      });

      _writeAllShiftRows_(kept);
      auditWrite(actor, "Clear Schedule Employee", `${targetName} starting ${config.startDate} (${config.view || "week"})`, "Success");
      return true;
    } catch (err) {
      auditWrite(actor, "Clear Schedule Employee", `${config?.name || ""} starting ${config?.startDate || ""}`, "Failed");
      throw new Error("Server failed to clear schedule: " + err.message);
    }
  });
}

// =========================
// GET / DELETE SHIFT HELPERS
// =========================
function getShiftDetails(name, date) {
  const targetKey = `${String(name || "").trim()}|${String(date || "").trim()}`;
  const rows = _readShiftRows_();
  const index = _buildShiftIndex_(rows);
  const rowIndex = index[targetKey];

  if (rowIndex === undefined) return null;
  const row = rows[rowIndex];

  return {
    role: row[SHIFT_COLS.ROLE],
    mode: row[SHIFT_COLS.MODE],
    group: row[SHIFT_COLS.GROUP],
    start: row[SHIFT_COLS.START],
    end: row[SHIFT_COLS.END]
  };
}

function deleteExistingShift(name, date) {
  return _withScheduleLock_(() => {
    const targetName = String(name || "").trim();
    const targetDate = String(date || "").trim();

    const kept = _readShiftRows_().filter(row => {
      return !(
        String(row[SHIFT_COLS.NAME] || "").trim() === targetName &&
        _dateKey_(row[SHIFT_COLS.DATE]) === targetDate
      );
    });

    _writeAllShiftRows_(kept);
    return true;
  });
}

// =========================
// BULK COPY / MOVE
// =========================
function copyShiftsBulk(details) {
  return _withScheduleLock_(() => {
    const actor = _actorEmail_(details?.actorEmail);

    try {
      const rows = _readShiftRows_();
      const index = _buildShiftIndex_(rows);

      const sourceShifts = Array.isArray(details?.sourceShifts) ? details.sourceShifts : [];
      if (!sourceShifts.length) {
        auditWrite(actor, "Copy Shifts Bulk", "No source shifts", "Failed");
        return false;
      }

      const oldAnchor = new Date(sourceShifts[0].date + "T00:00:00");
      const newAnchor = new Date(details.targetDate + "T00:00:00");
      const dayOffset = Math.round((newAnchor - oldAnchor) / 86400000);

      sourceShifts.forEach(source => {
        const srcKey = `${String(source.name || "").trim()}|${String(source.date || "").trim()}`;
        const srcIdx = index[srcKey];
        if (srcIdx === undefined) return;

        const src = rows[srcIdx];
        const targetD = new Date(String(source.date) + "T00:00:00");
        targetD.setDate(targetD.getDate() + dayOffset);
        const targetDateStr = toDateKey(targetD);

        const targetKey = `${String(details.targetName || "").trim()}|${targetDateStr}`;
        const newRow = [
          src[SHIFT_COLS.ROLE],
          String(details.targetName || "").trim(),
          src[SHIFT_COLS.MODE],
          src[SHIFT_COLS.GROUP],
          targetDateStr,
          src[SHIFT_COLS.START],
          src[SHIFT_COLS.END]
        ];

        if (index[targetKey] !== undefined) {
          rows[index[targetKey]] = newRow;
        } else {
          index[targetKey] = rows.length;
          rows.push(newRow);
        }
      });

      _writeAllShiftRows_(rows);
      auditWrite(actor, "Copy Shifts Bulk", `${sourceShifts.length} shift(s) → ${details.targetName}@${details.targetDate}`, "Success");
      return true;
    } catch (e) {
      auditWrite(actor, "Copy Shifts Bulk", `${details?.targetName || ""}`, "Failed");
      throw e;
    }
  });
}

function moveShiftsBulk(details) {
  return _withScheduleLock_(() => {
    const actor = _actorEmail_(details?.actorEmail);

    try {
      let rows = _readShiftRows_();
      const sourceShifts = Array.isArray(details?.sourceShifts) ? details.sourceShifts : [];
      if (!sourceShifts.length) {
        auditWrite(actor, "Move Shifts Bulk", "No source shifts", "Failed");
        return false;
      }

      const oldAnchor = new Date(sourceShifts[0].date + "T00:00:00");
      const newAnchor = new Date(details.targetDate + "T00:00:00");
      const dayJump = Math.round((newAnchor - oldAnchor) / 86400000);

      const sourceKeySet = new Set(
        sourceShifts.map(s => `${String(s.name || "").trim()}|${String(s.date || "").trim()}`)
      );

      const index = _buildShiftIndex_(rows);
      const movedRows = [];

      sourceShifts.forEach(source => {
        const srcKey = `${String(source.name || "").trim()}|${String(source.date || "").trim()}`;
        const srcIdx = index[srcKey];
        if (srcIdx === undefined) return;

        const src = rows[srcIdx];
        const targetD = new Date(String(source.date) + "T00:00:00");
        targetD.setDate(targetD.getDate() + dayJump);

        movedRows.push([
          src[SHIFT_COLS.ROLE],
          String(details.targetName || "").trim(),
          src[SHIFT_COLS.MODE],
          src[SHIFT_COLS.GROUP],
          toDateKey(targetD),
          src[SHIFT_COLS.START],
          src[SHIFT_COLS.END]
        ]);
      });

      rows = rows.filter(row => {
        const key = `${String(row[SHIFT_COLS.NAME] || "").trim()}|${_dateKey_(row[SHIFT_COLS.DATE])}`;
        return !sourceKeySet.has(key);
      });

      const newIndex = _buildShiftIndex_(rows);
      movedRows.forEach(row => {
        const key = `${String(row[SHIFT_COLS.NAME] || "").trim()}|${String(row[SHIFT_COLS.DATE] || "").trim()}`;
        if (newIndex[key] !== undefined) rows[newIndex[key]] = row;
        else rows.push(row);
      });

      _writeAllShiftRows_(rows);
      auditWrite(actor, "Move Shifts Bulk", `${sourceShifts.length} shift(s) → ${details.targetName}@${details.targetDate}`, "Success");
      return true;
    } catch (e) {
      auditWrite(actor, "Move Shifts Bulk", `${details?.targetName || ""}`, "Failed");
      throw e;
    }
  });
}

// =========================
// COPY PREVIOUS SCHEDULE
// =========================
function copyPreviousSchedule(params) {
  return _withScheduleLock_(() => {
    if (!canDoAction("schedule.copy")) {
      throw new Error("Access denied: schedule.copy");
    }

    const actor = _actorEmail_(params?.actorEmail);

    try {
      const rows = _readShiftRows_();
      const _rawDate = String(params.targetStart || "").split("T")[0];
      const targetStart = new Date(_rawDate + "T12:00:00");

      const isMonth = params.period === "month";
      const sourceStart = new Date(targetStart);
      if (isMonth) sourceStart.setMonth(sourceStart.getMonth() - 1);
      else sourceStart.setDate(sourceStart.getDate() - 7);

      const rangeLimit = isMonth ? 31 : 7;
      const templateMap = Object.create(null);

      rows.forEach(row => {
        const rowDate = _dateObjFromCell_(row[SHIFT_COLS.DATE]);
        if (!rowDate) return;
        rowDate.setHours(12, 0, 0, 0);

        const diffDays = Math.round((rowDate.getTime() - sourceStart.getTime()) / 86400000);
        if (diffDays < 0 || diffDays >= rangeLimit) return;

        const name = String(row[SHIFT_COLS.NAME] || "").trim();
        const dow = rowDate.getDay();

        if (!templateMap[name]) templateMap[name] = {};
        templateMap[name][dow] = row;
      });

      const daysInTarget = isMonth
        ? new Date(targetStart.getFullYear(), targetStart.getMonth() + 1, 0).getDate()
        : 7;

      const teamList = getTeamList();
      const replaceMap = Object.create(null);

      teamList.forEach(emp => {
        const empName = String(emp.name || "").trim();
        const empTemplate = templateMap[empName];
        if (!empTemplate) return;

        for (let i = 0; i < daysInTarget; i++) {
          const tDate = new Date(targetStart);
          tDate.setDate(targetStart.getDate() + i);

          const targetDateKey = toDateKey(tDate);
          const sourceMatch = empTemplate[tDate.getDay()];
          if (!sourceMatch) continue;

          replaceMap[`${empName}|${targetDateKey}`] = [
            emp.role || sourceMatch[SHIFT_COLS.ROLE],
            empName,
            sourceMatch[SHIFT_COLS.MODE],
            emp.group || sourceMatch[SHIFT_COLS.GROUP],
            targetDateKey,
            sourceMatch[SHIFT_COLS.START] || "",
            sourceMatch[SHIFT_COLS.END] || ""
          ];
        }
      });

      const kept = rows.filter(row => {
        const key = `${String(row[SHIFT_COLS.NAME] || "").trim()}|${_dateKey_(row[SHIFT_COLS.DATE])}`;
        return !replaceMap[key];
      });

      const finalRows = kept.concat(Object.values(replaceMap));
      _writeAllShiftRows_(finalRows);

      auditWrite(actor, "Copy Previous Schedule", `${params.period || "week"} starting ${toDateKey(targetStart)}`, "Success");
      return Object.keys(replaceMap).length;
    } catch (e) {
      auditWrite(actor, "Copy Previous Schedule", `${params?.period || ""} starting ${params?.targetStart || ""}`, "Failed");
      throw e;
    }
  });
}

// =========================
// WEEKEND OFF APPLY/CLEAR
// =========================
function applyFixedWeekendOff(config) {
  return _withScheduleLock_(() => {
    if (!canDoAction("schedule.weekend")) {
      throw new Error("Access denied: schedule.weekend");
    }

    const actor = _actorEmail_(config?.actorEmail);

    try {
      const start = new Date(config.startDate);
      start.setHours(12, 0, 0, 0);

      const days = (config.view === "month")
        ? new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate()
        : 7;

      const isRemoving = config.removeMode === true || String(config.mode || "").toUpperCase() === "CLEAR";

      const weekendKeys = new Set();
      for (let i = 0; i < days; i++) {
        const d = new Date(start);
        d.setDate(start.getDate() + i);
        if (d.getDay() === 0 || d.getDay() === 6) weekendKeys.add(toDateKey(d));
      }

      const teamList = getTeamList();
      const targetNames = Array.isArray(config.employeeNames) && config.employeeNames.length
        ? new Set(config.employeeNames.map(_normName_))
        : null;

      const targetEmployees = targetNames
        ? teamList.filter(e => targetNames.has(_normName_(e.name)))
        : teamList;

      const targetNameSet = new Set(targetEmployees.map(e => _normName_(e.name)));
      let rows = _readShiftRows_();
      let affected = 0;

      if (isRemoving) {
        rows = rows.filter(row => {
          const dateKey = _dateKey_(row[SHIFT_COLS.DATE]);
          const shift = String(row[SHIFT_COLS.MODE] || "").trim().toUpperCase();
          const nameKey = _normName_(row[SHIFT_COLS.NAME]);

          const shouldRemove = weekendKeys.has(dateKey) && shift === "OFF" && targetNameSet.has(nameKey);
          if (shouldRemove) affected++;
          return !shouldRemove;
        });
      } else {
        const index = _buildShiftIndex_(rows);
        targetEmployees.forEach(emp => {
          weekendKeys.forEach(dateKey => {
            const key = `${String(emp.name || "").trim()}|${dateKey}`;
            if (index[key] !== undefined) return;

            rows.push([
              emp.role || "Staff",
              emp.name,
              "OFF",
              emp.group || "Morning",
              dateKey,
              "",
              "",
              ""
            ]);
            index[key] = rows.length - 1;
            affected++;
          });
        });
      }

      _writeAllShiftRows_(rows);
      auditWrite(actor, "Weekend OFF Apply/Clear", `${config.view || "week"} starting ${toDateKey(start)}`, "Success");
      return affected;
    } catch (e) {
      auditWrite(actor, "Weekend OFF Apply/Clear", `${config?.view || ""} starting ${config?.startDate || ""}`, "Failed");
      throw e;
    }
  });
}

// =========================
// COPY WEEK TO REMAINING WEEKS OF MONTH
// =========================
function copyWeekToRemainingWeeksOfMonth(payload) {
  return _withScheduleLock_(() => {
    payload = payload || {};
    const overwrite = payload.overwrite !== false;

    const sourceStartKey = _dateKey_(payload.sourceWeekStart);
    if (!sourceStartKey) return { weeksCopied: 0, shiftsCopied: 0 };

    const sourceWeekStart = new Date(sourceStartKey + "T00:00:00");
    sourceWeekStart.setHours(0, 0, 0, 0);

    const sourceEnd = new Date(sourceWeekStart);
    sourceEnd.setDate(sourceEnd.getDate() + 6);
    sourceEnd.setHours(23, 59, 59, 999);

    const monthStart = new Date(sourceWeekStart.getFullYear(), sourceWeekStart.getMonth(), 1);
    const monthEnd = new Date(sourceWeekStart.getFullYear(), sourceWeekStart.getMonth() + 1, 0);
    monthStart.setHours(0, 0, 0, 0);
    monthEnd.setHours(23, 59, 59, 999);

    const targetWeekStarts = [];
    const cursor = new Date(sourceWeekStart);
    cursor.setDate(cursor.getDate() + 7);

    while (cursor.getMonth() === sourceWeekStart.getMonth() && cursor.getFullYear() === sourceWeekStart.getFullYear()) {
      targetWeekStarts.push(new Date(cursor));
      cursor.setDate(cursor.getDate() + 7);
    }

    if (!targetWeekStarts.length) return { weeksCopied: 0, shiftsCopied: 0 };

    let rows = _readShiftRows_();

    const sourceShifts = rows.filter(r => {
      const d = _dateObjFromCell_(r[SHIFT_COLS.DATE]);
      return d && d >= sourceWeekStart && d <= sourceEnd;
    });

    if (!sourceShifts.length) return { weeksCopied: targetWeekStarts.length, shiftsCopied: 0 };

    if (overwrite) {
      rows = rows.filter(r => {
        const d = _dateObjFromCell_(r[SHIFT_COLS.DATE]);
        if (!d) return true;

        return !targetWeekStarts.some(ws => {
          const we = new Date(ws);
          we.setDate(ws.getDate() + 6);
          we.setHours(23, 59, 59, 999);
          return d >= ws && d <= we;
        });
      });
    }

    const toAppend = [];
    targetWeekStarts.forEach(targetStart => {
      const deltaDays = Math.round((targetStart.getTime() - sourceWeekStart.getTime()) / 86400000);

      sourceShifts.forEach(r => {
        const od = _dateObjFromCell_(r[SHIFT_COLS.DATE]);
        if (!od) return;

        const nd = new Date(od);
        nd.setDate(nd.getDate() + deltaDays);

        if (nd < monthStart || nd > monthEnd) return;

        const newRow = r.slice();
        newRow[SHIFT_COLS.DATE] = toDateKey(nd);
        toAppend.push(newRow);
      });
    });

    _writeAllShiftRows_(rows.concat(toAppend));
    return { weeksCopied: targetWeekStarts.length, shiftsCopied: toAppend.length };
  });
}

// =========================
// EXPORT
// =========================
function generateRosterExport(config) {
  const timeZone  = Session.getScriptTimeZone();
  const allShifts = Array.isArray(config?.shiftData) ? config.shiftData : getEmployeesFromShifts();

  const raw   = new Date(config.startDate);
  let start   = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate(), 12, 0, 0);
  while (start.getDay() !== 1) start.setDate(start.getDate() - 1);

  const days = (config.view === "month")
    ? new Date(start.getFullYear(), start.getMonth() + 1, 0).getDate()
    : 7;

  const exportDates = [];
  for (let i = 0; i < days; i++) {
    const d = new Date(start);
    d.setDate(start.getDate() + i);
    exportDates.push(d);
  }

  const teamList = getTeamList();
  const shiftMap = Object.create(null);
  allShifts.forEach(s => {
    shiftMap[`${String(s.name || "").trim()}|${String(s.date || "").trim()}`] = s;
  });

  // Build output data in memory first
  const headers = ["Role", "Shift Group", "Employee"];
  exportDates.forEach(d => headers.push(Utilities.formatDate(d, timeZone, "EEE MM/dd")));
  headers.push("Totals");

  const COUNTED_MODES = ["RTO","WFH","SL","VL","MAT","PAT","BL","HOL","OFF","FH ON-CALL"];

  const output = teamList.map(emp => {
    const row = [emp.role || "Staff", emp.group || "Morning", emp.name];
    const modeCounts = {};

    exportDates.forEach(d => {
      const dateKey = Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
      const shift   = shiftMap[`${emp.name}|${dateKey}`];

      if (!shift) { row.push("-"); return; }

      const mode    = String(shift.mode || "").toUpperCase();
      const hasTime = shift.start && shift.end && !String(shift.start).includes("undefined");

      modeCounts[mode] = (modeCounts[mode] || 0) + 1;

      row.push(
        hasTime && ["RTO", "WFH", "FH ON-CALL"].includes(mode)
          ? `${mode}\n${shift.start} - ${shift.end}`
          : mode
      );
    });

    // Totals column
    const summary = COUNTED_MODES
      .filter(m => modeCounts[m])
      .map(m => `${m}: ${modeCounts[m]}`)
      .join("\n") || "-";
    row.push(summary);

    return row;
  });

  // Clean up old exports in Drive before creating a new one
  const exportPrefix = "Manila IT Roster - ";
  const exportTitle  = exportPrefix + (config?.title || "Export");

  // Create spreadsheet first
  const ss    = SpreadsheetApp.create(exportTitle);
  const sheet = ss.getSheets()[0];

  sheet.appendRow(headers);

  if (output.length) {
    sheet.getRange(2, 1, output.length, headers.length).setValues(output);
  }

  const lastRow = sheet.getLastRow();
  const lastCol = headers.length;

  // Styling
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground("#1C1C1E")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily("Arial")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#D1D1D6", SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 4, lastRow - 1, days);
    dataRange.setWrap(true).setHorizontalAlignment("center");

    sheet.setConditionalFormatRules([
      rule("RTO",      "#FFF3E0", "#E65100", dataRange),
      rule("WFH",      "#E3F2FD", "#1565C0", dataRange),
      rule("FH ON-CALL","#FFEBEE","#C62828", dataRange),
      rule("MAT",      "#FCE4EC", "#C2185B", dataRange),
      rule("PAT",      "#E1F5FE", "#0288D1", dataRange),
      rule("SL",       "#E0F2F1", "#00796B", dataRange),
      rule("VL",       "#E8F5E9", "#2E7D32", dataRange),
      rule("HOL",      "#F3E5F5", "#7B1FA2", dataRange),
      rule("OFF",      "#EEEEEE", "#9E9E9E", dataRange)
    ]);
  }

 sheet.autoResizeColumns(1, 3);
  for (let c = 4; c <= lastCol - 1; c++) sheet.setColumnWidth(c, 135);
  sheet.autoResizeColumns(lastCol, 1);
  sheet.getRange(1, lastCol).setBackground("#1C1C1E").setFontColor("#FFFFFF").setFontWeight("bold");
  if (lastRow > 1) {
    sheet.getRange(2, lastCol, lastRow - 1, 1)
      .setBackground("#f8f8f8")
      .setFontWeight("bold")
      .setWrap(true)
      .setVerticalAlignment("middle");
  }

  // Move to a dedicated export folder to keep Drive tidy
  try {
    const folderName   = "Manila IT Exports";
    const folderSearch = DriveApp.getFoldersByName(folderName);
    const folder       = folderSearch.hasNext()
      ? folderSearch.next()
      : DriveApp.createFolder(folderName);

    const newFile = DriveApp.getFileById(ss.getId());
    newFile.moveTo(folder);

    // Clean up older exports with the same title in the same folder
    try {
      const dupes = folder.getFilesByName(exportTitle);
      while (dupes.hasNext()) {
        const f = dupes.next();
        if (f.getId() !== newFile.getId()) f.setTrashed(true);
      }
    } catch (ce) {
      console.warn("generateRosterExport: cleanup failed:", ce?.message);
    }

  } catch (e) {
    console.warn("generateRosterExport: folder move failed:", e?.message);
  }

  return ss.getUrl();
}



function rule(text, bg, color, range) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextStartsWith(text)
    .setBackground(bg)
    .setFontColor(color)
    .setRanges([range])
    .build();
}


// ============================================================
// ADD THIS to your existing request_api.gs (or schedule.gs)
// It replaces / extends submitEmployeeRequest with Drive upload.
// ============================================================
 
/* =========================================================
 * DRIVE FOLDER CONFIG
 * =========================================================
 * Parent folder ID for all request attachments.
 * Sub-folders are created per user (email username).
 * ======================================================= */
var REQUEST_ATTACHMENTS_FOLDER_ID = "1HYCSODXKM4D789VOh7JOnPvtoLGyrwrS";
 
/* =========================================================
 * UPLOAD ATTACHMENT TO DRIVE
 *
 * Creates (or reuses) a sub-folder named after the user's
 * email prefix inside the parent REQUEST_ATTACHMENTS_FOLDER_ID,
 * then saves the file and returns its Drive URL.
 * ======================================================= */
function uploadRequestAttachment_(email, base64Data, mimeType, fileName) {
  var userEmail  = String(email    || "").trim().toLowerCase();
  var mime       = String(mimeType || "application/octet-stream").trim();
  var name       = String(fileName || "attachment").trim();

  var b64 = String(base64Data || "").trim();

  // Strip any accidental data-URL prefix
  var commaIdx = b64.indexOf(",");
  if (commaIdx >= 0) b64 = b64.slice(commaIdx + 1);

  // Remove whitespace
  b64 = b64.replace(/\s/g, "");

  // Convert URL-safe base64 back to standard base64
  b64 = b64.replace(/-/g, "+").replace(/_/g, "/");

  // Re-pad to a multiple of 4
  while (b64.length % 4 !== 0) b64 += "=";

  if (!b64 || !userEmail) return "";

  var folderName = userEmail.split("@")[0] || userEmail;

  var parent;
  try {
    parent = DriveApp.getFolderById(REQUEST_ATTACHMENTS_FOLDER_ID);
  } catch (e) {
    throw new Error("Cannot access attachments folder: " + e.message);
  }

  var userFolder;
  var existing = parent.getFoldersByName(folderName);
  if (existing.hasNext()) {
    userFolder = existing.next();
  } else {
    userFolder = parent.createFolder(folderName);
  }

  var decoded = Utilities.base64Decode(b64);
  var blob    = Utilities.newBlob(decoded, mime, name);

  var file = userFolder.createFile(blob);

  // setSharing can fail if Workspace domain policies restrict external sharing
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e1) {
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e2) {
      // Sharing restricted by admin — file still accessible to folder members
      console.warn("uploadRequestAttachment_: setSharing failed (non-blocking):", e2.message);
    }
  }

  return file.getUrl();
}

function copyEmployeeWeekForward(payload) {
  return _withScheduleLock_(() => {
    if (!canDoAction("schedule.copy")) {
      throw new Error("Access denied: schedule.copy");
    }

    const name      = String(payload?.name || "").trim();
    const sourceKey = String(payload?.sourceWeekStart || "").trim();
    if (!name || !sourceKey) return { shiftsCopied: 0 };

    const sourceStart = new Date(sourceKey + "T00:00:00");
    sourceStart.setHours(0, 0, 0, 0);
    const sourceEnd = new Date(sourceStart);
    sourceEnd.setDate(sourceEnd.getDate() + 6);
    sourceEnd.setHours(23, 59, 59, 999);

    const targetStart = new Date(sourceStart);
    targetStart.setDate(targetStart.getDate() + 7);

    let rows = _readShiftRows_();

    const sourceShifts = rows.filter(r => {
      if (String(r[SHIFT_COLS.NAME] || "").trim() !== name) return false;
      const d = _dateObjFromCell_(r[SHIFT_COLS.DATE]);
      return d && d >= sourceStart && d <= sourceEnd;
    });

    if (!sourceShifts.length) return { shiftsCopied: 0 };

    // Remove existing target week shifts for this employee
    rows = rows.filter(r => {
      if (String(r[SHIFT_COLS.NAME] || "").trim() !== name) return true;
      const d = _dateObjFromCell_(r[SHIFT_COLS.DATE]);
      const targetEnd = new Date(targetStart);
      targetEnd.setDate(targetEnd.getDate() + 6);
      targetEnd.setHours(23, 59, 59, 999);
      return !(d && d >= targetStart && d <= targetEnd);
    });

    const toAppend = sourceShifts.map(r => {
      const od = _dateObjFromCell_(r[SHIFT_COLS.DATE]);
      const nd = new Date(od);
      nd.setDate(nd.getDate() + 7);
      const newRow = r.slice();
      newRow[SHIFT_COLS.DATE] = toDateKey(nd);
      return newRow;
    });

    _writeAllShiftRows_(rows.concat(toAppend));
    auditWrite(
      _actorEmail_(),
      "Copy Employee Week Forward",
      `${name} — week of ${sourceKey}`,
      "Success"
    );
    return { shiftsCopied: toAppend.length };
  });
}