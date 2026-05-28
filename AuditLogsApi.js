/* ============================================================
 * AUDIT LOG
 * ============================================================ */

const AUDIT_SHEET_NAME = "AuditLogs";

/* ── RBAC ─────────────────────────────────────────────────── */
function canViewAuditLog(email) {
  return canAccessPage("admin", email);
}

function canExportAuditLog(email) {
  return canAccessPage("admin", email) && canDoAction("admin.access", email);
}

function requireAuditLogView_() {
  const email = getCurrentUserEmail_();
  if (!canViewAuditLog(email)) {
    throw new Error("You do not have permission to access Audit Logs.");
  }
}

/* ── Sheet helper ─────────────────────────────────────────── */
function _auditSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(AUDIT_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(AUDIT_SHEET_NAME);
    sh.appendRow([
      "Timestamp", "User", "Action", "Serial Tag",
      "Entity", "Field", "Old Value", "New Value", "Details (JSON)"
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ── Write ────────────────────────────────────────────────── */
function auditWrite(payload) {
  const p  = payload || {};
  const sh = _auditSheet_();

  const ts        = new Date();
  const user      = getCurrentUserEmail_();
  const action    = String(p.action    || "").trim();
  const serialTag = String(p.serialTag || "").trim();
  const entity    = String(p.entity    || "Masterlist").trim();
  const field     = String(p.field     || "").trim();
  const oldValue  = p.oldValue == null ? "" : String(p.oldValue);
  const newValue  = p.newValue == null ? "" : String(p.newValue);

  const detailsJson = (() => {
    try {
      const d = p.details;
      if (!d || typeof d !== "object") return "";
      return JSON.stringify(d);
    } catch (_) { return ""; }
  })();

  sh.appendRow([ts, user, action, entity, field, oldValue, newValue, serialTag, detailsJson]);
  return true;
}

/* ── List (paginated) ─────────────────────────────────────── */
function audit_list(options) {
  try {
    requireAuditLogView_();

    const opt      = options || {};
    const page     = Math.max(1, parseInt(opt.page,     10) || 1);
    const pageSize = Math.max(1, Math.min(200, parseInt(opt.pageSize, 10) || 50));
    const query    = String(opt.query     || "").trim().toLowerCase();
    const serialTag= String(opt.serialTag || "").trim().toLowerCase();

    const sh     = _auditSheet_();
    const values = sh.getDataRange().getDisplayValues();

    if (values.length < 2) {
      return { ok: true, data: [], total: 0, page, pageSize };
    }

    const rows = values.slice(1).map(r => {
      const col9        = String(r[8] || "").trim();
      const isOldFormat = !r[9] && (col9.startsWith("{") || col9.startsWith("[") || col9 === "");
      return {
        timestamp:    r[0],
        user:         r[1],
        action:       r[2],
        serialTag:    r[3],
        entity:       r[4],
        field:        r[5],
        oldValue:     r[6],
        newValue:     r[7],
        affectedUser: isOldFormat ? ""   : col9,
        details:      isOldFormat ? col9 : String(r[9] || "").trim()
      };
    });

    rows.reverse();

    let filtered = rows;

    if (serialTag) {
      filtered = filtered.filter(x =>
        String(x.serialTag || "").toLowerCase() === serialTag
      );
    }

    if (query) {
      filtered = filtered.filter(x => {
        const blob = [
          x.timestamp, x.user, x.action, x.serialTag,
          x.entity, x.field, x.oldValue, x.newValue, x.details
        ].join(" ").toLowerCase();
        return blob.includes(query);
      });
    }

    const total = filtered.length;
    const start = (page - 1) * pageSize;
    const paged = filtered.slice(start, start + pageSize);

    return { ok: true, data: paged, total, page, pageSize };

  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── Clear ────────────────────────────────────────────────── */
function audit_clear() {
  try {
    requireAuditLogView_();

    const sh      = _auditSheet_();
    const lastRow = sh.getLastRow();
    const lastCol = Math.max(sh.getLastColumn(), 9);

    let deletedRows = 0;
    if (lastRow > 1) {
      deletedRows = lastRow - 1;
      sh.getRange(2, 1, deletedRows, lastCol).clearContent();
    }

    return { ok: true, deletedRows };

  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


/* ============================================================
 * ASSET LOG
 *
 * Sheet columns (1-based):
 *  1  Timestamp
 *  2  User            (actor — who made the change)
 *  3  Action          (ADD | UPDATE | DELETE | CLEAR_ASSIGNMENT)
 *  4  Serial Tag
 *  5  Entity
 *  6  Field
 *  7  Old Value
 *  8  New Value
 *  9  Affected User   (the assignee whose asset was touched)
 * 10  Details (JSON)
 *
 * NOTE: If your existing sheet only has 9 columns, the first time
 * assetLogWrite runs it will just append a 10th value — Sheets
 * handles sparse rows fine. You may want to manually add the
 * "Affected User" header to column 9 and shift "Details (JSON)"
 * to column 10 in the sheet header row for clarity.
 * ============================================================ */

const ASSET_LOG_SHEET_NAME = "Asset Logs";

/* ── Sheet helper ─────────────────────────────────────────── */
function _assetLogSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(ASSET_LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ASSET_LOG_SHEET_NAME);
    sh.appendRow([
      "Timestamp", "User", "Action", "Serial Tag",
      "Entity", "Field", "Old Value", "New Value",
      "Affected User", "Details (JSON)"
    ]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ── Write ────────────────────────────────────────────────── */
function assetLogWrite(payload) {
  const p  = payload || {};
  const sh = _assetLogSheet_();

  const ts           = new Date();
  const user         = getCurrentUserEmail_();
  const action       = String(p.action       || "").trim();
  const serialTag    = String(p.serialTag    || "").trim();
  const entity       = String(p.entity       || "Masterlist").trim();
  const field        = String(p.field        || "").trim();
  const oldValue     = p.oldValue  == null ? "" : String(p.oldValue);
  const newValue     = p.newValue  == null ? "" : String(p.newValue);
  const affectedUser = String(
    (p.details && p.details.affectedUser) || p.affectedUser || ""
  ).trim();

  const detailsJson = (() => {
    try   { return JSON.stringify(p.details || {}); }
    catch (_) { return "{}"; }
  })();

  sh.appendRow([
    ts, user, action, serialTag,
    entity, field, oldValue, newValue,
    affectedUser, detailsJson
  ]);
  return true;
}

/* ── Date bound helper (timezone-aware) ───────────────────── */
function _parseAuditDateBound_(dateStr, endOfDay) {
  if (!dateStr) return null;
  const parts = dateStr.split("-");
  if (parts.length !== 3) return null;
  const y = Number(parts[0]);
  const m = Number(parts[1]) - 1;
  const d = Number(parts[2]);
  // new Date(y, m, d) uses the script's local timezone — no UTC shift
  return endOfDay
    ? new Date(y, m, d, 23, 59, 59, 999)
    : new Date(y, m, d,  0,  0,  0,   0);
}

/* ── List (paginated, filterable) ─────────────────────────── */
function asset_log_list(options) {
  try {
    const opt          = options || {};
    const page         = Math.max(1, parseInt(opt.page,     10) || 1);
    const pageSize     = Math.max(1, Math.min(200, parseInt(opt.pageSize, 10) || 50));
    const query        = String(opt.query        || "").trim().toLowerCase();
    const serialTag    = String(opt.serialTag    || "").trim().toLowerCase();
    const actionFilter = String(opt.action       || "").trim().toUpperCase();
    const dateFrom     = String(opt.dateFrom     || "").trim();  // "YYYY-MM-DD"
    const dateTo       = String(opt.dateTo       || "").trim();  // "YYYY-MM-DD"

    const sh     = _assetLogSheet_();
    const values = sh.getDataRange().getDisplayValues();

    if (values.length < 2) {
      return { ok: true, data: [], total: 0, page, pageSize };
    }

    // Parse date boundaries once, in script timezone
    const tsFrom = _parseAuditDateBound_(dateFrom, false);
    const tsTo   = _parseAuditDateBound_(dateTo,   true);

    const rows = values.slice(1).map(r => {
      const col9        = String(r[8] || "").trim();
      const isOldFormat = !r[9] && (col9.startsWith("{") || col9.startsWith("[") || col9 === "");
      return {
        timestamp:    r[0],
        user:         r[1],
        action:       r[2],
        serialTag:    r[3],
        entity:       r[4],
        field:        r[5],
        oldValue:     r[6],
        newValue:     r[7],
        affectedUser: isOldFormat ? ""   : col9,
        details:      isOldFormat ? col9 : String(r[9] || "").trim()
      };
    });

    rows.reverse(); // newest first

    let filtered = rows;

    /* ── Serial tag exact match ── */
    if (serialTag) {
      filtered = filtered.filter(x =>
        String(x.serialTag || "").toLowerCase() === serialTag
      );
    }

    /* ── Action filter ── */
    if (actionFilter) {
      filtered = filtered.filter(x =>
        String(x.action || "").trim().toUpperCase() === actionFilter
      );
    }

    /* ── Date range (timezone-aware) ── */
    if (tsFrom || tsTo) {
      filtered = filtered.filter(x => {
        const raw = x.timestamp;
        if (!raw) return false;
        const d = (raw instanceof Date) ? raw : new Date(raw);
        if (isNaN(d.getTime())) return false;
        if (tsFrom && d < tsFrom) return false;
        if (tsTo   && d > tsTo)   return false;
        return true;
      });
    }

    /* ── Freetext search ── */
    if (query) {
      filtered = filtered.filter(x => {
        const blob = [
          x.timestamp, x.user, x.action, x.serialTag,
          x.entity, x.field, x.oldValue, x.newValue,
          x.affectedUser, x.details
        ].join(" ").toLowerCase();
        return blob.includes(query);
      });
    }

    const total = filtered.length;
    const start = (page - 1) * pageSize;
    const paged = filtered.slice(start, start + pageSize);

    return { ok: true, data: paged, total, page, pageSize };

  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/**
 * Convenience wrapper for warranty reminder audit entries.
 * Called by sendMonthlyWarrantyReminder() after a successful send.
 */
function logWarrantyReminderSent(cfg, count) {
  assetLogWrite({
    action:    "WARRANTY_REMINDER_SENT",
    entity:    "Masterlist",
    serialTag: "",
    field:     "email",
    oldValue:  "",
    newValue:  cfg.to,
    details:   {
      cc:      cfg.cc    || null,
      days:    cfg.days,
      count:   count,
      subject: cfg.days === 180
        ? "Monthly Warranty Reminder — Assets Expiring Within 6 Months"
        : "Monthly Warranty Reminder — Assets Expiring Within " + cfg.days + " Days"
    }
  });
}

/**
 * Read the last N audit log entries from the Asset Logs sheet.
 * Used by admin pages to surface recent activity.
 *
 * @param  {number} limit  Max rows to return (default 50, max 500)
 * @return {Object}        { ok, data: [ {timestamp, actor, action, ...} ] }
 */
function getRecentAuditLog(limit) {
  try {
    limit = (typeof limit === "number" && limit > 0) ? Math.min(limit, 500) : 50;

    var sh   = _assetLogSheet_();
    var last = sh.getLastRow();

    if (last < 2) return { ok: true, data: [] };

    // Read from the bottom up — most recent first
    var startRow = Math.max(2, last - limit + 1);
    var numRows  = last - startRow + 1;
    var values   = sh.getRange(startRow, 1, numRows, 10).getDisplayValues();

    var entries = values.reverse().map(function(row) {
      var col9        = String(row[8] || "").trim();
      var isOldFormat = !row[9] && (col9.startsWith("{") || col9.startsWith("[") || col9 === "");
      return {
        timestamp:    row[0],
        actor:        row[1],
        action:       row[2],
        serialTag:    row[3],
        entity:       row[4],
        field:        row[5],
        oldValue:     row[6],
        newValue:     row[7],
        affectedUser: isOldFormat ? ""   : col9,
        details:      isOldFormat ? col9 : String(row[9] || "").trim()
      };
    });

    return { ok: true, data: entries };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}