const MASTERLIST_SHEET_NAME = "Masterlist";

/* ============================================================
 * SHARED HELPERS
 * ============================================================ */
function getSS_ml_() {
  return (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getMasterlistSheet_() {
  const sh = getSS_ml_().getSheetByName(MASTERLIST_SHEET_NAME);
  if (!sh) throw new Error("Masterlist sheet not found.");
  return sh;
}

/**
 * Shared date parser — handles MM/DD/YYYY, YYYY-MM-DD, Date objects,
 * and native JS Date strings.
 */
function parseMDY_(s) {
  if (!s) return null;
  if (s instanceof Date) return isNaN(s.getTime()) ? null : s;

  const t = String(s).trim();
  if (!t) return null;

  const mdy = t.split("/");
  if (mdy.length === 3) {
    let mm = Number(mdy[0]), dd = Number(mdy[1]), yy = Number(mdy[2]);
    if (!mm || !dd || !yy) return null;
    if (yy < 100) yy += 2000;
    const d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  const iso = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(t);
  return isNaN(d.getTime()) ? null : d;
}

function isMeaningful_(v) {
  const s = String(v || "").trim().toLowerCase();
  return !!s && !["n/a", "na", "—", "-", "none", "null", "(blank)"].includes(s);
}

function safeText_(row, idx) {
  if (idx < 0) return "";
  return String(row[idx] ?? "").trim();
}

function readDate_(row, idx) {
  if (idx < 0) return null;
  return parseMDY_(row[idx]);
}

function escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g,  "&amp;")
    .replace(/</g,  "&lt;")
    .replace(/>/g,  "&gt;")
    .replace(/"/g,  "&quot;")
    .replace(/'/g,  "&#039;");
}

/** Flexible header index resolver (for code paths that use smart aliases) */
function headerIndex_(headers) {
  const map = {};
  headers.forEach((h, i) => { map[String(h || "").trim().toLowerCase()] = i; });

  const pick = (...names) => {
    for (const n of names) {
      const k = String(n).toLowerCase();
      if (k in map) return map[k];
    }
    return -1;
  };

  return {
    status:      pick("status", "asset status", "state"),
    assignedTo:  pick("assigned to", "assignee", "user", "employee", "assigned"),
    warrantyEnd: pick(
      "end of warranty", "warranty end", "end warranty",
      "warranty expiry", "warranty expiration",
      "warranty expiry date", "warranty",
      "warranty_end", "warranty_expiry"
    )
  };
}

function normalizeStatusBucket_(statusRaw) {
  const s = String(statusRaw || "").trim().toLowerCase();
  if (!s)                                                          return "unassigned";
  if (s.includes("missing")     || s.includes("lost"))            return "missing";
  if (s.includes("faulty")      || s.includes("defective"))       return "faulty";
  if (s.includes("retired")     || s.includes("disposed")  ||
      s.includes("decommission"))                                  return "retired";
  if (s.includes("in use")      || s.includes("assigned"))        return "assigned";
  if (s.includes("available")   || s.includes("stock")    ||
      s.includes("spare"))                                         return "available";
  return "unassigned";
}


/* ============================================================
 * RBAC
 * ============================================================ */
function canViewAssetReport(email) {
  return canAccessPage("assetreport", email);
}

function canExportAssetReport(email) {
  return canAccessPage("assetreport", email) && canDoAction("assetreport.export", email);
}

function requireAssetReportView_(email) {
  if (!canViewAssetReport(email))
    throw new Error("You do not have permission to access Asset Report.");
}

function requireAssetReportExport_(email) {
  if (!canAccessPage("assetreport", email))
    throw new Error("You do not have permission to access Asset Report.");
  if (!canDoAction("assetreport.export", email))
    throw new Error("You do not have permission to export Asset Report.");
}

function requireAssetReportBulk_(email) {
  if (!canAccessPage("assetreport", email))
    throw new Error("You do not have permission to access Asset Report.");
  if (!canDoAction("assetreport.bulk", email))
    throw new Error("You do not have permission to perform bulk actions.");
}


/* ============================================================
 * MASTERLIST REPORT  (dashboard KPIs + charts)
 * ============================================================ */
function masterlist_report() {
  try {
    const email = getCurrentUserEmail_();
    requireAssetReportView_(email);

    const sh     = getMasterlistSheet_();
    const values = sh.getDataRange().getDisplayValues();

    const EMPTY = {
      ok: true,
      data: {
        totals: 0, byStatus: {}, byType: {}, byRam: {},
        warrantyBuckets: {}, byDepartment: {}, byAfStatus: {},
        kpis: {
          total: 0, assigned: 0, unassigned: 0, available: 0,
          inUse: 0, retired: 0, missing: 0, faulty: 0,
          afSigned: 0, afUnsigned: 0,
          warrantyExpired: 0, warrantyExpiringSoon: 0, warrantyNoDate: 0
        },
        typeStatusMatrix: {}, soonToExpire: [], deptLaptopModels: []
      }
    };

    if (values.length < 2) return EMPTY;

    const headers = values[0].map(h => String(h || "").trim().toLowerCase());
    const col     = name => headers.indexOf(String(name).trim().toLowerCase());

    const idx = {
      type:      col("type"),
      status:    col("status"),
      ram:       col("ram"),
      warranty:  col("end of warranty"),
      dept:      col("department"),
      afStatus:  col("af status"),
      signedAf:  col("signed af"),
      model:     col("model"),
      serialTag: col("serial tag")
    };

    const rows = values.slice(1).filter(r => r.some(isMeaningful_));

    const byStatus = {}, byType = {}, byRam = {},
          byDepartment = {}, byAfStatus = {}, typeStatusMatrix = {},
          deptLaptopModelMap = {};

    const warrantyBuckets = {
      "Expired": 0, "0–30 days": 0, "31–90 days": 0,
      "91–180 days": 0, "180+ days": 0
    };

    const kpis = {
      total: rows.length,
      assigned: 0, unassigned: 0, available: 0,
      inUse: 0, retired: 0, missing: 0, faulty: 0,
      afSigned: 0, afUnsigned: 0,
      warrantyExpired: 0, warrantyExpiringSoon: 0, warrantyNoDate: 0
    };

    const soonToExpire = [];
    const today = new Date(); today.setHours(0, 0, 0, 0);

    const inc  = (obj, key) => {
      const k = String(key || "—").trim() || "—";
      obj[k] = (obj[k] || 0) + 1;
    };
    const inc2 = (obj, k1, k2) => {
      const a = String(k1 || "—").trim() || "—";
      const b = String(k2 || "—").trim() || "—";
      obj[a] = obj[a] || {};
      obj[a][b] = (obj[a][b] || 0) + 1;
    };

    rows.forEach(r => {
      const type     = safeText_(r, idx.type)     || "—";
      const status   = safeText_(r, idx.status)   || "—";
      const ram      = safeText_(r, idx.ram)       || "—";
      const warranty = safeText_(r, idx.warranty);
      const dept     = safeText_(r, idx.dept)      || "—";
      const afStatus = safeText_(r, idx.afStatus);
      const signedAf = safeText_(r, idx.signedAf);
      const model    = safeText_(r, idx.model)     || "—";

      inc(byType,       type);
      inc(byStatus,     status);
      if (type.toLowerCase().includes("laptop")) inc(byRam, ram);
      inc(byDepartment, dept);
      inc(byAfStatus,   afStatus || "—");
      inc2(typeStatusMatrix, type, status);

      // Laptop model per dept
      if (type.toLowerCase().includes("laptop") &&
          isMeaningful_(dept) && isMeaningful_(model)) {
        deptLaptopModelMap[dept] = deptLaptopModelMap[dept] || {};
        deptLaptopModelMap[dept][model] = (deptLaptopModelMap[dept][model] || 0) + 1;
      }

      // KPI status buckets
      const bucket = normalizeStatusBucket_(status);
switch (bucket) {
  case "assigned":   kpis.assigned++; kpis.inUse++; break;
  case "available":  kpis.available++;               break;
  case "retired":    kpis.retired++;                 break;
  case "missing":    kpis.missing++;                 break;
  case "faulty":     kpis.faulty++;                  break;
  default:           kpis.unassigned++;
}

      // AF signed / unsigned
      if (idx.signedAf >= 0) {
        const isSignedAf = isMeaningful_(signedAf) &&
          !["no", "false", "n", "0"].includes(signedAf.toLowerCase());
        if (isSignedAf) kpis.afSigned++;
        else            kpis.afUnsigned++;
      }

      // Warranty
      const wd = parseMDY_(warranty);
      if (wd) {
        wd.setHours(0, 0, 0, 0);
        const diffDays = Math.floor((wd - today) / 86400000);

        // Inline bucket classification
        let wBucket;
        if      (diffDays < 0)    wBucket = "Expired";
        else if (diffDays <= 30)  wBucket = "0–30 days";
        else if (diffDays <= 90)  wBucket = "31–90 days";
        else if (diffDays <= 180) wBucket = "91–180 days";
        else                      wBucket = "180+ days";

        warrantyBuckets[wBucket]++;

        if (diffDays < 0) {
          kpis.warrantyExpired++;
        } else if (diffDays <= 180) {
          kpis.warrantyExpiringSoon++;
        }

        if (diffDays >= 0 && diffDays <= 180) {
          soonToExpire.push({
            serialTag:     safeText_(r, idx.serialTag),
            type:          type,
            model:         model,
            endOfWarranty: warranty,
            diffDays
          });
        }
      } else {
        kpis.warrantyNoDate++;
      }
    });

    soonToExpire.sort((a, b) => a.diffDays - b.diffDays);

    const deptLaptopModels = [];
    Object.keys(deptLaptopModelMap).sort().forEach(dept => {
      Object.keys(deptLaptopModelMap[dept]).sort().forEach(model => {
        deptLaptopModels.push({ department: dept, model, count: deptLaptopModelMap[dept][model] });
      });
    });

    return {
      ok: true,
      data: {
        totals: rows.length,
        byStatus, byType, byRam,
        warrantyBuckets, byDepartment, byAfStatus,
        typeStatusMatrix, kpis,
        soonToExpire: soonToExpire.slice(0, 15),
        deptLaptopModels
      }
    };

  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


/* ============================================================
 * INVENTORY ANALYTICS  (export sheet)
 * ============================================================ */
function getInventoryAnalytics() {
  const email = getCurrentUserEmail_();
  requireAssetReportView_(email);

  const sh     = getMasterlistSheet_();
  const values = sh.getDataRange().getValues();

  const EMPTY = {
    totalAssets: 0, assigned: 0, inStock: 0, retired: 0, lost: 0,
    warranty: { expired: 0, d0_30: 0, d31_60: 0, d61_90: 0, d90_plus: 0, no_date: 0 }
  };

  if (values.length < 2) return EMPTY;

  const headers = values[0].map(h => String(h || "").trim());
  const rows    = values.slice(1).filter(r => r.some(v => String(v ?? "").trim() !== ""));
  const idx     = headerIndex_(headers);
  const now     = new Date();
  const msDay   = 86400000;

  const out = { ...EMPTY, totalAssets: rows.length };

  rows.forEach(r => {
    const status     = safeText_(r, idx.status).toLowerCase();
    const assignedTo = safeText_(r, idx.assignedTo);

    if (assignedTo && assignedTo.toLowerCase() !== "unassigned") out.assigned++;

    const bucket = normalizeStatusBucket_(status);
    if      (bucket === "available") out.inStock++;
    else if (bucket === "retired")   out.retired++;
    else if (bucket === "missing")   out.lost++;

    const wd = readDate_(r, idx.warrantyEnd);
    if (!wd) {
      out.warranty.no_date++;
    } else if (wd < now) {
      out.warranty.expired++;
    } else {
      const diff = Math.ceil((wd.getTime() - now.getTime()) / msDay);
      if      (diff <= 30) out.warranty.d0_30++;
      else if (diff <= 60) out.warranty.d31_60++;
      else if (diff <= 90) out.warranty.d61_90++;
      else                 out.warranty.d90_plus++;
    }
  });

  return out;
}

function exportInventoryAnalyticsToSheet() {
  try {
    const email = getCurrentUserEmail_();
    requireAssetReportExport_(email);

    const analytics = getInventoryAnalytics();
    const ss        = getSS_ml_();
    const stamp     = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const name      = `Inventory Analytics - ${stamp}`;

    const existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);

    const sh = ss.insertSheet(name);
    sh.getRange(1, 1).setValue("INVENTORY ANALYTICS").setFontSize(16).setFontWeight("bold");
    sh.getRange(2, 1).setValue(`Generated: ${stamp}`);

    const kpis = [
      ["Metric",                "Value"],
      ["Total Assets",          analytics.totalAssets],
      ["Assigned",              analytics.assigned],
      ["In Stock",              analytics.inStock],
      ["Retired",               analytics.retired],
      ["Lost / Missing",        analytics.lost],
      ["Warranty Expired",      analytics.warranty.expired],
      ["Expiring 0–30 days",    analytics.warranty.d0_30],
      ["Expiring 31–60 days",   analytics.warranty.d31_60],
      ["Expiring 61–90 days",   analytics.warranty.d61_90],
      ["Expiring 90+ days",     analytics.warranty.d90_plus],
      ["No Warranty Date",      analytics.warranty.no_date]
    ];

    sh.getRange(4, 1, kpis.length, 2).setValues(kpis);
    sh.getRange(4, 1, 1, 2).setFontWeight("bold");
    sh.getRange(5, 2, kpis.length - 1, 1).setNumberFormat("0");
    sh.autoResizeColumns(1, 2);

    return { ok: true, url: ss.getUrl() + "#gid=" + sh.getSheetId(), sheetName: name };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


/* ============================================================
 * MASTERLIST TABLE  (full raw table for asset report page)
 * ============================================================ */
function getMasterlistTable() {
  try {
    const email = getCurrentUserEmail_();
    requireAssetReportView_(email);

    const sh     = getMasterlistSheet_();
    const values = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, headers: [], rows: [] };

    const headers = values[0].map(h => String(h || "").trim());
    const rows    = values.slice(1).filter(r => r.some(v => String(v ?? "").trim() !== ""));

    return { ok: true, headers, rows };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

function getMasterlistSheetUrl() {
  try {
    const email = getCurrentUserEmail_();
    requireAssetReportView_(email);

    const ss = getSS_ml_();
    const sh = ss.getSheetByName(MASTERLIST_SHEET_NAME);
    if (!sh) return { ok: false, error: "Masterlist sheet not found" };

    return { ok: true, url: ss.getUrl() + "#gid=" + sh.getSheetId() };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


/* ============================================================
 * EXPORT FILTERED ROWS
 * ============================================================ */
function exportFilteredRows(payload) {
  try {
    const email = getCurrentUserEmail_();
    requireAssetReportExport_(email);

    if (!payload || !Array.isArray(payload.rows) || !Array.isArray(payload.headers))
      return { ok: false, error: "Invalid payload: expected { headers: [], rows: [] }" };

    if (!payload.headers.length)
      return { ok: false, error: "Payload headers are empty." };

    const ss    = getSS_ml_();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    const name  = `Filtered Assets - ${stamp}`;

    const existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);

    const sh     = ss.insertSheet(name);
    const values = [payload.headers].concat(
      payload.rows.map(r => r.map(c => (c == null ? "" : String(c))))
    );

    sh.getRange(1, 1, values.length, values[0].length).setValues(values);
    sh.getRange(1, 1, 1, values[0].length).setFontWeight("bold");
    sh.autoResizeColumns(1, values[0].length);

    return { ok: true, url: ss.getUrl() + "#gid=" + sh.getSheetId(), sheetName: name };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


/* ============================================================
 * WARRANTY REMINDER CONFIG
 *
 * Script Properties consulted at runtime (all optional):
 *   WARRANTY_REMINDER_TO     e.g. "manila-it@fbg.com, ops@fbg.com"
 *   WARRANTY_REMINDER_CC     e.g. "cfo@fbg.com"   (blank is fine)
 *   WARRANTY_REMINDER_DAYS   e.g. "90"            (default 180)
 *
 * Editable via the warrantyconfig.html admin page (RBAC-gated),
 * or directly via Apps Script editor:
 *   setWarrantyReminderConfig({ to: "new@fbg.com", days: 90 });
 * ============================================================ */

const WARRANTY_REMINDER_DEFAULTS_ = {
  to:         "manila-it@fbgphilippines.com",
  cc:         "",
  from:       "manila-it@fbgphilippines.com",
  days:       180,
  senderName: "Manila IT Inventory"
};

/**
 * Read the current warranty reminder config, falling back to defaults
 * for anything unset. Each returned field includes a source marker
 * so callers can tell config from defaults at a glance.
 */
function getWarrantyReminderConfig() {
  const props = PropertiesService.getScriptProperties();
  const to    = String(props.getProperty("WARRANTY_REMINDER_TO")   || "").trim();
  const cc    = String(props.getProperty("WARRANTY_REMINDER_CC")   || "").trim();
  const daysS = String(props.getProperty("WARRANTY_REMINDER_DAYS") || "").trim();

  const daysN = parseInt(daysS, 10);
  const days  = (isFinite(daysN) && daysN > 0 && daysN <= 730)
    ? daysN
    : WARRANTY_REMINDER_DEFAULTS_.days;

  return {
    to:         to || WARRANTY_REMINDER_DEFAULTS_.to,
    cc:         cc,
    from:       WARRANTY_REMINDER_DEFAULTS_.from,
    senderName: WARRANTY_REMINDER_DEFAULTS_.senderName,
    days:       days,
    _toSource:   to     ? "property"  : "default",
    _ccSource:   cc     ? "property"  : "default",
    _daysSource: (isFinite(daysN) && daysN > 0 && daysN <= 730) ? "property" : "default"
  };
}

/**
 * Patch the config. Pass only fields you want to change; omit to leave alone.
 * Pass `null` to explicitly clear a field (reverts to the hardcoded default).
 *
 * Examples:
 *   setWarrantyReminderConfig({ to: "ops@fbg.com", days: 90 });
 *   setWarrantyReminderConfig({ cc: null });   // clears CC
 */
function setWarrantyReminderConfig(patch) {
  const props = PropertiesService.getScriptProperties();
  patch = patch || {};

  if ("to" in patch) {
    if (patch.to === null || patch.to === "") props.deleteProperty("WARRANTY_REMINDER_TO");
    else props.setProperty("WARRANTY_REMINDER_TO", String(patch.to).trim());
  }
  if ("cc" in patch) {
    if (patch.cc === null || patch.cc === "") props.deleteProperty("WARRANTY_REMINDER_CC");
    else props.setProperty("WARRANTY_REMINDER_CC", String(patch.cc).trim());
  }
  if ("days" in patch) {
    if (patch.days === null) props.deleteProperty("WARRANTY_REMINDER_DAYS");
    else props.setProperty("WARRANTY_REMINDER_DAYS", String(patch.days));
  }

  return getWarrantyReminderConfig();
}


/* ============================================================
 * WARRANTY REMINDER CONFIG — CLIENT WRAPPERS
 * These are what warrantyconfig.html calls. All gated by
 * assetreport.export. Unsuffixed functions above stay usable
 * from the Apps Script editor and from the scheduled trigger.
 * ============================================================ */

function getWarrantyReminderConfigClient() {
  try {
    const email = getCurrentUserEmail_();
    requireWarrantyConfigView_(email);
    return { ok: true, data: getWarrantyReminderConfig() };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


function setWarrantyReminderConfigClient(patch) {
  try {
    const email = getCurrentUserEmail_();
    requireWarrantyConfigManage_(email);

    patch = patch || {};

    // Input validation before we hit Script Properties
    if ("to" in patch && patch.to !== null && patch.to !== "") {
      const toStr = String(patch.to).trim();
      if (!toStr) return { ok: false, error: "Recipient cannot be blank." };
      if (!/[^\s@,]+@[^\s@,]+/.test(toStr))
        return { ok: false, error: "Recipient must contain a valid email address." };
    }
    if ("cc" in patch && patch.cc !== null && patch.cc !== "") {
      const ccStr = String(patch.cc).trim();
      if (ccStr && !/[^\s@,]+@[^\s@,]+/.test(ccStr))
        return { ok: false, error: "CC must contain a valid email address (or be blank)." };
    }
    if ("days" in patch && patch.days !== null) {
      const daysN = parseInt(String(patch.days), 10);
      if (!isFinite(daysN) || daysN <= 0 || daysN > 730)
        return { ok: false, error: "Days must be a number between 1 and 730." };
    }

    const cfg = setWarrantyReminderConfig(patch);

    // Audit log — non-blocking
    try {
      if (typeof assetLogWrite === "function") {
        assetLogWrite({
          action:   "WARRANTY_CONFIG_UPDATED",
          entity:   "ScriptProperties",
          field:    "warranty_reminder_config",
          oldValue: "",
          newValue: JSON.stringify({ to: cfg.to, cc: cfg.cc, days: cfg.days }),
          details:  { actor: email || "" }
        });
      }
    } catch (logErr) {
      console.error("Audit log failed (non-blocking):", logErr);
    }

    return { ok: true, data: cfg };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

function previewWarrantyReminderClient() {
  try {
    const email = getCurrentUserEmail_();
    requireWarrantyConfigView_(email);  // preview is read-only
    const preview = sendMonthlyWarrantyReminder({ preview: true });
    return { ok: true, data: preview };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}


/* ============================================================
 * MONTHLY WARRANTY REMINDER EMAIL  (time-based trigger)
 *
 * Options:
 *   { preview: true }  — return the planned email without sending.
 *                        Used by the warrantyconfig admin page preview.
 *
 * Returns a summary:
 *   { ok, sent, previewed, count, days, to, cc, subject, htmlBody?, plainBody?, reason? }
 * ============================================================ */
function sendMonthlyWarrantyReminder(options) {
  options = options || {};
  const preview = !!options.preview;

  try {
    // No RBAC check here — runs as a scheduled trigger (not a user request).
    // Client callers go through previewWarrantyReminderClient() which IS gated.

    const cfg = getWarrantyReminderConfig();

    const sh     = getMasterlistSheet_();
    const values = sh.getDataRange().getDisplayValues();
    if (values.length < 2) {
      Logger.log("[WARRANTY REMINDER] Masterlist is empty — nothing to send.");
      return { ok: true, sent: false, previewed: preview, count: 0, reason: "empty_sheet" };
    }

    const headers = values[0].map(h => String(h || "").trim().toLowerCase());
    const col     = name => headers.indexOf(String(name).trim().toLowerCase());

    const idx = {
      type:      col("type"),
      status:    col("status"),
      warranty:  col("end of warranty"),
      dept:      col("department"),
      assignee:  col("assignee"),
      model:     col("model"),
      serialTag: col("serial tag")
    };

    const today = new Date(); today.setHours(0, 0, 0, 0);

    const soon = values.slice(1).reduce((acc, r) => {
      const warranty = idx.warranty >= 0 ? r[idx.warranty] : "";
      const wd       = parseMDY_(warranty);
      if (!wd) return acc;

      wd.setHours(0, 0, 0, 0);
      const diff = Math.floor((wd.getTime() - today.getTime()) / 86400000);

      if (diff >= 0 && diff <= cfg.days) {
        acc.push({
          department:    safeText_(r, idx.dept),
          assignee:      safeText_(r, idx.assignee),
          type:          safeText_(r, idx.type),
          model:         safeText_(r, idx.model),
          serialTag:     safeText_(r, idx.serialTag),
          status:        safeText_(r, idx.status),
          endOfWarranty: String(warranty || "").trim(),
          diffDays:      diff
        });
      }
      return acc;
    }, []);

    if (!soon.length) {
      Logger.log(`[WARRANTY REMINDER] No assets expiring within ${cfg.days} days. Email skipped.`);
      return { ok: true, sent: false, previewed: preview, count: 0, days: cfg.days, reason: "no_matches" };
    }

    soon.sort((a, b) => a.diffDays - b.diffDays);

    // Urgency bucket labels scale with the configured window so a 90-day
    // window doesn't show "Within 6 months" buckets that don't apply.
    const urgencyLabel = (d) => {
      if (d <= 30)  return "⚠️ Within 30 days";
      if (d <= 90)  return "Within 3 months";
      if (d <= 180) return "Within 6 months";
      return "Within " + Math.ceil(d / 30) + " months";
    };

    const htmlRows = soon.map(item => `
      <tr>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.department)}</td>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.assignee)}</td>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.type)}</td>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.model)}</td>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.serialTag)}</td>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.status)}</td>
        <td style="padding:8px;border:1px solid #ddd;">${escapeHtml_(item.endOfWarranty)}</td>
        <td style="padding:8px;border:1px solid #ddd;text-align:center;">${item.diffDays}</td>
        <td style="padding:8px;border:1px solid #ddd;">${urgencyLabel(item.diffDays)}</td>
      </tr>`).join("");

    const windowLabel = cfg.days === 180
      ? "within the next <strong>6 months</strong>"
      : `within the next <strong>${cfg.days}</strong> day(s)`;

    const today_label = today.toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" });
    const htmlBody = `
<div style="font-family:Arial,sans-serif;color:#222;">
  <h2 style="margin-bottom:8px;">Warranty Reminder</h2>
  <p style="margin-top:0;">
    The following <strong>${soon.length}</strong> asset(s) have warranties expiring
    ${windowLabel}.
  </p>
        <table style="border-collapse:collapse;width:100%;font-size:13px;">
          <thead>
            <tr style="background:#f3f4f6;">
              <th style="padding:8px;border:1px solid #ddd;">Department</th>
              <th style="padding:8px;border:1px solid #ddd;">Assignee</th>
              <th style="padding:8px;border:1px solid #ddd;">Type</th>
              <th style="padding:8px;border:1px solid #ddd;">Model</th>
              <th style="padding:8px;border:1px solid #ddd;">Serial Tag</th>
              <th style="padding:8px;border:1px solid #ddd;">Status</th>
              <th style="padding:8px;border:1px solid #ddd;">Warranty End</th>
              <th style="padding:8px;border:1px solid #ddd;">Days Left</th>
              <th style="padding:8px;border:1px solid #ddd;">Urgency</th>
            </tr>
          </thead>
          <tbody>${htmlRows}</tbody>
        </table>
      <hr style="margin-top:24px;border:none;border-top:1px solid #eee;">
  <p style="font-size:11px;color:#888;margin-top:8px;">
    This is an automated monthly reminder generated on ${today_label} by the
    Manila IT Inventory System. Please do not reply to this email.
  </p>
</div>`;

    const plainBody = `Assets with warranty expiring within ${cfg.days} day(s):\n\n` +
      soon.map(item =>
        [item.department, item.assignee, item.type, item.model,
         item.serialTag, item.status, item.endOfWarranty,
         item.diffDays + " days left"].join(" | ")
      ).join("\n");

    const subject = cfg.days === 180
      ? "Monthly Warranty Reminder — Assets Expiring Within 6 Months"
      : `Monthly Warranty Reminder — Assets Expiring Within ${cfg.days} Days`;

    // Preview mode — return without sending
    if (preview) {
      Logger.log(`[WARRANTY REMINDER] Preview — would send to ${cfg.to} (cc: ${cfg.cc || "—"}), ${soon.length} asset(s), window=${cfg.days}d`);
      return {
        ok:        true,
        sent:      false,
        previewed: true,
        count:     soon.length,
        days:      cfg.days,
        to:        cfg.to,
        cc:        cfg.cc,
        subject,
        htmlBody,
        plainBody
      };
    }

    // Send
    const mailOpts = {
      htmlBody,
      from: cfg.from,
      name: cfg.senderName
    };
    if (cfg.cc) mailOpts.cc = cfg.cc;

    GmailApp.sendEmail(cfg.to, subject, plainBody, mailOpts);

    Logger.log(`[WARRANTY REMINDER] Sent to ${cfg.to} (cc: ${cfg.cc || "—"}) — ${soon.length} asset(s), window=${cfg.days}d`);

    // Optional audit-log entry — non-blocking
    try {
      logWarrantyReminderSent(cfg, soon.length);
    } catch (logErr) {
      Logger.log("[WARRANTY REMINDER] Audit log failed (non-blocking): " +
        (logErr && logErr.message ? logErr.message : String(logErr)));
    }

    return {
      ok:        true,
      sent:      true,
      previewed: false,
      count:     soon.length,
      days:      cfg.days,
      to:        cfg.to,
      cc:        cfg.cc,
      subject
    };

  } catch (e) {
    Logger.log("[WARRANTY REMINDER] Error: " + (e?.message || String(e)));
    throw e;
  }
}


/* ============================================================
 * TRIGGER MANAGEMENT
 * ============================================================ */
function createMonthlyWarrantyReminderTrigger() {
  _deleteTriggers_("sendMonthlyWarrantyReminder");
  ScriptApp.newTrigger("sendMonthlyWarrantyReminder")
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();
  Logger.log("[TRIGGER] Monthly warranty reminder trigger created.");
}

function deleteMonthlyWarrantyReminderTrigger() {
  _deleteTriggers_("sendMonthlyWarrantyReminder");
  Logger.log("[TRIGGER] Monthly warranty reminder trigger deleted.");
}

function _deleteTriggers_(fnName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fnName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}


/* ============================================================
 * BULK UPDATE STATUS
 * req = { serialTags: [String], newStatus: String }
 * Returns { ok, data: { updated, failed: [{serialTag, reason}] }, error? }
 * ============================================================ */

function masterlist_bulkUpdateStatus(req) {
  try {
    var email = getCurrentUserEmail_();       // resolve ONCE
    requireAssetReportBulk_(email);           // pass down — no internal Session call

    req = req || {};
    var tags      = Array.isArray(req.serialTags) ? req.serialTags : [];
    var newStatus = String(req.newStatus || "").trim();

    if (!tags.length)  return { ok: false, error: "No serial tags provided." };
    if (!newStatus)    return { ok: false, error: "No newStatus provided." };

    // Restrict to the statuses the UI exposes so callers can't sneak in
    // workflow-specific values like "Assigned" or "Return" via bulk
    var ALLOWED = ["In-stock", "Retired", "Defective"];
    var allowedOk = false;
    for (var a = 0; a < ALLOWED.length; a++) {
      if (ALLOWED[a].toLowerCase() === newStatus.toLowerCase()) {
        newStatus = ALLOWED[a]; // normalise case
        allowedOk = true;
        break;
      }
    }
    if (!allowedOk) {
      return { ok: false, error: "Status \"" + newStatus + "\" is not allowed for bulk change. Use the single-row edit flow." };
    }

    var sh        = _mlSheet_();
    var hm        = _mlHeaderMap_(sh);
    var statusCol = hm.map.status;
    if (!statusCol) return { ok: false, error: "Masterlist is missing the Status column." };

    var updated    = 0;
    var failed     = [];
    var beforeByTag = {};
    var actor      = email;                  // reuse — no extra Session call

    for (var i = 0; i < tags.length; i++) {
      var tag = String(tags[i] || "").trim();
      if (!tag) { failed.push({ serialTag: "", reason: "empty tag" }); continue; }

      // Snapshot before state for audit
      var beforeResp = masterlist_get(tag);
      var before     = (beforeResp && beforeResp.ok) ? (beforeResp.data || {}) : null;
      beforeByTag[tag] = before;

      var row = _mlFindRowBySerialTag_(sh, tag);
      if (row < 0) { failed.push({ serialTag: tag, reason: "not found" }); continue; }

      var oldStatus = before ? String(before.status || "") : "";

      // No-op if already at target status
      if (oldStatus.toLowerCase() === newStatus.toLowerCase()) {
        // still counts as "no change" — don't increment updated, don't log
        continue;
      }

      try {
        sh.getRange(row, statusCol).setValue(newStatus);
        updated++;

        // Audit log — one entry per asset, field=status
        try {
          assetLogWrite({
            action:    "BULK_UPDATE_STATUS",
            serialTag: tag,
            entity:    "Masterlist",
            field:     "status",
            oldValue:  oldStatus,
            newValue:  newStatus,
            details:   { actor: actor, bulk: true }
          });
        } catch (logErr) {
          console.error("assetLogWrite failed for " + tag + ":", logErr);
        }
      } catch (writeErr) {
        failed.push({ serialTag: tag, reason: (writeErr && writeErr.message) || String(writeErr) });
      }
    }

    SpreadsheetApp.flush();

    // Single Slack notification summarizing the bulk op
    try {
      if (updated > 0) {
        _sendBulkStatusSlack_(updated, failed.length, newStatus, actor, tags.slice(0, 10));
      }
    } catch (slackErr) {
      console.error("Slack notify failed (non-blocking):", slackErr);
    }

    return {
      ok: true,
      data: {
        updated:   updated,
        failed:    failed,
        requested: tags.length,
        newStatus: newStatus
      }
    };
  } catch (e) {
    console.error("masterlist_bulkUpdateStatus error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}

/* ============================================================
 * BULK REASSIGN
 * req = { serialTags: [String], newAssignee: String, newEmail?: String, newDepartment?: String }
 * Returns { ok, data: { updated, failed: [...] }, error? }
 * ============================================================ */

function masterlist_bulkReassign(req) {
  try {
    var email = getCurrentUserEmail_();
    requireAssetReportBulk_(email);

    req = req || {};
    var tags         = Array.isArray(req.serialTags) ? req.serialTags : [];
    var newAssignee  = String(req.newAssignee   || "").trim();
    var newEmail     = String(req.newEmail      || "").trim();
    var newDept      = String(req.newDepartment || "").trim();

    if (!tags.length)    return { ok: false, error: "No serial tags provided." };
    if (!newAssignee)    return { ok: false, error: "No newAssignee provided." };

    // If email/dept weren't supplied, resolve from the Employees sheet
    if (!newEmail || !newDept) {
      try {
        var list = employees_dropdown_list();
        if (list && list.ok && Array.isArray(list.data)) {
          for (var e = 0; e < list.data.length; e++) {
            if (String(list.data[e].displayName || "").trim() === newAssignee) {
              if (!newEmail) newEmail = String(list.data[e].email      || "").trim();
              if (!newDept)  newDept  = String(list.data[e].department || "").trim();
              break;
            }
          }
        }
      } catch (resolveErr) {
        console.warn("employees_dropdown_list resolve failed:", resolveErr);
      }
    }

    var sh = _mlSheet_();
    var hm = _mlHeaderMap_(sh);
    var colAssignee   = hm.map.assignee;
    var colEmail      = hm.map.email;
    var colDepartment = hm.map.department;
    var colStatus     = hm.map.status;
    if (!colAssignee) return { ok: false, error: "Masterlist is missing the Assignee column." };

    var updated = 0;
    var failed  = [];
    var actor = email;

    for (var i = 0; i < tags.length; i++) {
      var tag = String(tags[i] || "").trim();
      if (!tag) { failed.push({ serialTag: "", reason: "empty tag" }); continue; }

      var beforeResp = masterlist_get(tag);
      var before     = (beforeResp && beforeResp.ok) ? (beforeResp.data || {}) : null;

      var row = _mlFindRowBySerialTag_(sh, tag);
      if (row < 0) { failed.push({ serialTag: tag, reason: "not found" }); continue; }

      var oldAssignee = before ? String(before.assignee   || "") : "";
      var oldEmail    = before ? String(before.email      || "") : "";
      var oldDept     = before ? String(before.department || "") : "";
      var oldStatus   = before ? String(before.status     || "") : "";

      try {
        sh.getRange(row, colAssignee).setValue(newAssignee);
        if (colEmail)      sh.getRange(row, colEmail).setValue(newEmail);
        if (colDepartment) sh.getRange(row, colDepartment).setValue(newDept);
        // Reassignment implies "Assigned" status — set it if not already in use
        if (colStatus) {
          var oldLC = oldStatus.toLowerCase();
          if (oldLC !== "assigned" && oldLC !== "in use") {
            sh.getRange(row, colStatus).setValue("Assigned");
          }
        }

        updated++;

        // Audit log — one entry per changed field
        try {
          if (oldAssignee !== newAssignee) {
            assetLogWrite({
              action: "BULK_REASSIGN", serialTag: tag, entity: "Masterlist",
              field: "assignee", oldValue: oldAssignee, newValue: newAssignee,
              details: { actor: actor, bulk: true }
            });
          }
          if (colEmail && oldEmail !== newEmail) {
            assetLogWrite({
              action: "BULK_REASSIGN", serialTag: tag, entity: "Masterlist",
              field: "email", oldValue: oldEmail, newValue: newEmail,
              details: { actor: actor, bulk: true }
            });
          }
          if (colDepartment && oldDept !== newDept) {
            assetLogWrite({
              action: "BULK_REASSIGN", serialTag: tag, entity: "Masterlist",
              field: "department", oldValue: oldDept, newValue: newDept,
              details: { actor: actor, bulk: true }
            });
          }
          if (colStatus) {
            var oldLC2 = oldStatus.toLowerCase();
            if (oldLC2 !== "assigned" && oldLC2 !== "in use") {
              assetLogWrite({
                action: "BULK_REASSIGN", serialTag: tag, entity: "Masterlist",
                field: "status", oldValue: oldStatus, newValue: "Assigned",
                details: { actor: actor, bulk: true }
              });
            }
          }
        } catch (logErr) {
          console.error("assetLogWrite failed for " + tag + ":", logErr);
        }
      } catch (writeErr) {
        failed.push({ serialTag: tag, reason: (writeErr && writeErr.message) || String(writeErr) });
      }
    }

    SpreadsheetApp.flush();

    // Single Slack notification
    try {
      if (updated > 0) {
        _sendBulkReassignSlack_(updated, failed.length, newAssignee, newDept, actor, tags.slice(0, 10));
      }
    } catch (slackErr) {
      console.error("Slack notify failed (non-blocking):", slackErr);
    }

    return {
      ok: true,
      data: {
        updated:     updated,
        failed:      failed,
        requested:   tags.length,
        newAssignee: newAssignee,
        newEmail:    newEmail,
        newDepartment: newDept
      }
    };
  } catch (e) {
    console.error("masterlist_bulkReassign error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


/* ============================================================
 * SLACK NOTIFICATIONS  (used by bulk ops above)
 * ============================================================ */

function _bulkTagSample_(tags) {
  if (!tags || !tags.length) return "\u2014";
  var shown  = tags.slice(0, 10).join(", ");
  var more   = tags.length > 10 ? " …and " + (tags.length - 10) + " more" : "";
  return shown + more;
}

function _bulkNowPHT_() {
  return "\uD83D\uDD50 " +
    new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) +
    " PHT";
}

function _bulkAppUrl_() {
  try {
    return ScriptApp.getService().getUrl() + "?page=assetreport";
  } catch (e) {
    return "";
  }
}

function _sendBulkStatusSlack_(updated, failedCount, newStatus, actor, sampleTags) {
  var url = _bulkAppUrl_();

  var fields = [
    { type: "mrkdwn", text: "*Assets Updated:*\n" + updated },
    { type: "mrkdwn", text: "*New Status:*\n"     + newStatus },
    { type: "mrkdwn", text: "*Updated by:*\n"     + (actor || "\u2014") },
    { type: "mrkdwn", text: "*Failed:*\n"         + (failedCount || 0) }
  ];

  var blocks = [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: "*\uD83D\uDCE6 Bulk Status Change — Manila IT Inventory*"
      }
    },
    { type: "section", fields: fields },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: "*Affected Serial Tags:*\n`" + _bulkTagSample_(sampleTags) + "`"
      }
    }
  ];

  if (url) {
    blocks.push({
      type: "actions",
      elements: [
        {
          type: "button",
          text: { type: "plain_text", text: "Open Asset Report", emoji: true },
          url:  url,
          style: "primary"
        }
      ]
    });
  }

  blocks.push({
    type: "context",
    elements: [{ type: "mrkdwn", text: _bulkNowPHT_() }]
  });

  // Color: newStatus drives sidebar — retired=red, defective=gray, in-stock=blue
  var color = "#3b82f6";
  var s = newStatus.toLowerCase();
  if (s === "retired")        color = "#ef4444";
  else if (s === "defective") color = "#6b7280";

  sendSlackNotification_(null, blocks, color);
}

function _sendBulkReassignSlack_(updated, failedCount, newAssignee, newDept, actor, sampleTags) {
  var url = _bulkAppUrl_();

  var fields = [
    { type: "mrkdwn", text: "*Assets Reassigned:*\n" + updated },
    { type: "mrkdwn", text: "*New Assignee:*\n"     + (newAssignee || "\u2014") },
    { type: "mrkdwn", text: "*Department:*\n"       + (newDept || "\u2014") },
    { type: "mrkdwn", text: "*Reassigned by:*\n"    + (actor || "\u2014") },
    { type: "mrkdwn", text: "*Failed:*\n"           + (failedCount || 0) }
  ];

  var blocks = [
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: "*\uD83D\uDC65 Bulk Reassignment — Manila IT Inventory*"
      }
    },
    { type: "section", fields: fields },
    {
      type: "section",
      text: {
        type: "mrkdwn",
        text: "*Affected Serial Tags:*\n`" + _bulkTagSample_(sampleTags) + "`"
      }
    }
  ];

  if (url) {
    blocks.push({
      type: "actions",
      elements: [
        {
          type: "button",
          text: { type: "plain_text", text: "Open Asset Report", emoji: true },
          url:  url,
          style: "primary"
        }
      ]
    });
  }

  blocks.push({
    type: "context",
    elements: [{ type: "mrkdwn", text: _bulkNowPHT_() }]
  });

  sendSlackNotification_(null, blocks, "#22c55e");
}

function requireWarrantyConfigView_(email) {
  if (!canAccessPage("warrantyconfig", email))
    throw new Error("You do not have permission to access Warranty Config.");
}

function requireWarrantyConfigManage_(email) {
  if (!canAccessPage("warrantyconfig", email))
    throw new Error("You do not have permission to access Warranty Config.");
  if (!canDoAction("warrantyconfig.manage", email))
    throw new Error("You do not have permission to change Warranty Config.");
}


/* ============================================================
 * ASSET REPORT BUNDLE
 * Single sheet read that replaces three separate calls:
 *   masterlist_report() + getMasterlistTable() + getDashboardStats()
 * Called by assetreport.html on every load.
 * ============================================================ */

function getAssetReportBundle() {
  try {
    var email = getCurrentUserEmail_();
    requireAssetReportView_(email);

    var sh      = getMasterlistSheet_();
    var display = sh.getDataRange().getDisplayValues();

    if (display.length < 2) {
      return {
        ok:   true,
        report: {
          totals: 0, byStatus: {}, byType: {}, byRam: {},
          warrantyBuckets: {}, byDepartment: {}, byAfStatus: {},
          kpis: {
            total: 0, assigned: 0, unassigned: 0, available: 0,
            inUse: 0, retired: 0, missing: 0, faulty: 0,
            afSigned: 0, afUnsigned: 0,
            warrantyExpired: 0, warrantyExpiringSoon: 0, warrantyNoDate: 0
          },
          typeStatusMatrix: {}, soonToExpire: [], deptLaptopModels: []
        },
        table:          { headers: [], rows: [] },
        dashboardStats: { assignedByTeam: {}, inStockByModel: {} },
        masterlistUrl:  ""
      };
    }

    // ── Build table (raw display values) ─────────────────────────
    var headers = display[0].map(function(h) { return String(h || "").trim(); });
    var rows    = display.slice(1).filter(function(r) {
      return r.some(function(v) { return String(v == null ? "" : v).trim() !== ""; });
    });

    // ── Build report data (mirrors masterlist_report logic) ───────
    var col = {};
    var lc  = headers.map(function(h) { return String(h || "").trim().toLowerCase(); });
    col.type      = lc.indexOf("type");
    col.status    = lc.indexOf("status");
    col.ram       = lc.indexOf("ram");
    col.warranty  = lc.indexOf("end of warranty");
    col.dept      = lc.indexOf("department");
    col.afStatus  = lc.indexOf("af status");
    col.signedAf  = lc.indexOf("signed af");
    col.model     = lc.indexOf("model");
    col.serialTag = lc.indexOf("serial tag");
    col.assignee  = lc.indexOf("assignee");
    col.status    = lc.indexOf("status");

    var byStatus = {}, byType = {}, byRam = {},
        byDepartment = {}, byAfStatus = {}, typeStatusMatrix = {},
        deptLaptopModelMap = {};

    var warrantyBuckets = {
      "Expired": 0, "0\u201330 days": 0, "31\u201390 days": 0,
      "91\u2013180 days": 0, "180+ days": 0
    };

    var kpis = {
      total: rows.length,
      assigned: 0, unassigned: 0, available: 0,
      inUse: 0, retired: 0, missing: 0, faulty: 0,
      afSigned: 0, afUnsigned: 0,
      warrantyExpired: 0, warrantyExpiringSoon: 0, warrantyNoDate: 0
    };

    var soonToExpire = [];
    var today = new Date(); today.setHours(0, 0, 0, 0);

    // Dashboard stats accumulators
    var assignedByTeam  = {};
    var inStockByModel  = {};

    function safe(r, i) { return i >= 0 ? String(r[i] == null ? "" : r[i]).trim() : ""; }
    function inc(obj, key) {
      var k = String(key || "\u2014").trim() || "\u2014";
      obj[k] = (obj[k] || 0) + 1;
    }
    function inc2(obj, k1, k2) {
      var a = String(k1 || "\u2014").trim() || "\u2014";
      var b = String(k2 || "\u2014").trim() || "\u2014";
      if (!obj[a]) obj[a] = {};
      obj[a][b] = (obj[a][b] || 0) + 1;
    }

    for (var i = 0; i < rows.length; i++) {
      var r        = rows[i];
      var type     = safe(r, col.type)     || "\u2014";
      var status   = safe(r, col.status)   || "\u2014";
      var ram      = safe(r, col.ram)      || "\u2014";
      var warranty = safe(r, col.warranty);
      var dept     = safe(r, col.dept)     || "\u2014";
      var afStatus = safe(r, col.afStatus);
      var signedAf = safe(r, col.signedAf);
      var model    = safe(r, col.model)    || "\u2014";
      var assignee = safe(r, col.assignee);

      var typeLC = type.toLowerCase();

      inc(byType,       type);
      inc(byStatus,     status);
      if (typeLC.indexOf("laptop") >= 0) inc(byRam, ram);
      inc(byDepartment, dept);
      inc(byAfStatus,   afStatus || "\u2014");
      inc2(typeStatusMatrix, type, status);

      // Laptop model per dept
      if (typeLC.indexOf("laptop") >= 0 && dept !== "\u2014" && model !== "\u2014") {
        if (!deptLaptopModelMap[dept]) deptLaptopModelMap[dept] = {};
        deptLaptopModelMap[dept][model] = (deptLaptopModelMap[dept][model] || 0) + 1;
      }

      // KPI status bucket
      var statusLC = status.toLowerCase();
      if (statusLC.indexOf("missing") >= 0 || statusLC.indexOf("lost") >= 0) {
        kpis.missing++;
      } else if (statusLC.indexOf("faulty") >= 0 || statusLC.indexOf("defective") >= 0) {
        kpis.faulty++;
      } else if (statusLC.indexOf("retired") >= 0 || statusLC.indexOf("disposed") >= 0) {
        kpis.retired++;
      } else if (statusLC.indexOf("in use") >= 0 || statusLC.indexOf("assigned") >= 0) {
        kpis.assigned++; kpis.inUse++;
      } else if (statusLC.indexOf("available") >= 0 || statusLC.indexOf("stock") >= 0 || statusLC.indexOf("spare") >= 0) {
        kpis.available++;
      } else {
        kpis.unassigned++;
      }

      var signedAfNorm = String(signedAf || "").trim().toLowerCase();
      var isAfSigned   = ["yes", "true", "1", "signed", "✓", "x"].indexOf(signedAfNorm) >= 0;
      if (isAfSigned) kpis.afSigned++; else kpis.afUnsigned++;

      // Warranty buckets
      var wd = _parseWarrantyDate_(warranty);
      if (wd) {
        wd.setHours(0, 0, 0, 0);
        var diff = Math.floor((wd.getTime() - today.getTime()) / 86400000);
        if (diff < 0) {
          warrantyBuckets["Expired"]++;
          kpis.warrantyExpired++;
        } else if (diff <= 30) {
          warrantyBuckets["0\u201330 days"]++;
          kpis.warrantyExpiringSoon++;
        } else if (diff <= 90) {
          warrantyBuckets["31\u201390 days"]++;
          kpis.warrantyExpiringSoon++;
        } else if (diff <= 180) {
          warrantyBuckets["91\u2013180 days"]++;
          kpis.warrantyExpiringSoon++;
        } else {
          warrantyBuckets["180+ days"]++;
        }

        if (diff >= 0 && diff <= 180) {
          soonToExpire.push({
            serialTag:     safe(r, col.serialTag),
            type:          type,
            model:         model,
            endOfWarranty: warranty,
            diffDays:      diff
          });
        }
      } else {
        kpis.warrantyNoDate++;
      }

      // Dashboard stats — assigned laptops by dept, in-stock laptops by model
      if (typeLC.indexOf("laptop") >= 0) {
        var isAssigned = statusLC.indexOf("assigned") >= 0 || statusLC.indexOf("in use") >= 0;
        var isStock    = statusLC.indexOf("available") >= 0 || statusLC.indexOf("stock") >= 0 || statusLC.indexOf("spare") >= 0;

        if (isAssigned && dept !== "\u2014") {
          assignedByTeam[dept] = (assignedByTeam[dept] || 0) + 1;
        }
        if (isStock && model !== "\u2014") {
          inStockByModel[model] = (inStockByModel[model] || 0) + 1;
        }
      }
    }

    soonToExpire.sort(function(a, b) { return a.diffDays - b.diffDays; });

    var deptLaptopModels = [];
    var deptKeys = Object.keys(deptLaptopModelMap).sort();
    for (var di = 0; di < deptKeys.length; di++) {
      var deptKey    = deptKeys[di];
      var modelKeys  = Object.keys(deptLaptopModelMap[deptKey]).sort();
      for (var mi = 0; mi < modelKeys.length; mi++) {
        var modelKey = modelKeys[mi];
        deptLaptopModels.push({
          department: deptKey,
          model:      modelKey,
          count:      deptLaptopModelMap[deptKey][modelKey]
        });
      }
    }

    // ── Masterlist URL ────────────────────────────────────────────
    var masterlistUrl = "";
    try {
      var ss = getSS_ml_();
      masterlistUrl = ss.getUrl() + "#gid=" + sh.getSheetId();
    } catch (urlErr) {
      console.warn("getAssetReportBundle: URL lookup failed:", urlErr);
    }

    return {
      ok: true,
      report: {
        totals: rows.length,
        byStatus:         byStatus,
        byType:           byType,
        byRam:            byRam,
        warrantyBuckets:  warrantyBuckets,
        byDepartment:     byDepartment,
        byAfStatus:       byAfStatus,
        typeStatusMatrix: typeStatusMatrix,
        kpis:             kpis,
        soonToExpire:     soonToExpire.slice(0, 15),
        deptLaptopModels: deptLaptopModels
      },
      table: {
        headers: headers,
        rows:    rows
      },
      dashboardStats: {
        assignedByTeam: assignedByTeam,
        inStockByModel: inStockByModel
      },
      masterlistUrl: masterlistUrl
    };

  } catch (e) {
    console.error("getAssetReportBundle error:", e);
    return { ok: false, error: (e && e.message) || String(e) };
  }
}


/* ── Internal date parser used only by getAssetReportBundle ── */
function _parseWarrantyDate_(s) {
  if (!s) return null;
  var t = String(s).trim();
  if (!t) return null;

  var mdy = t.split("/");
  if (mdy.length === 3) {
    var mm = Number(mdy[0]), dd = Number(mdy[1]), yy = Number(mdy[2]);
    if (!mm || !dd || !yy) return null;
    if (yy < 100) yy += 2000;
    var d = new Date(yy, mm - 1, dd);
    return isNaN(d.getTime()) ? null : d;
  }

  var iso = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    var di = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return isNaN(di.getTime()) ? null : di;
  }

  var d2 = new Date(t);
  return isNaN(d2.getTime()) ? null : d2;
}