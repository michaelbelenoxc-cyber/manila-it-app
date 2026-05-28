// ============================================================
// KPI METRICS — GAS backend
// ============================================================

const KPI_SHEET_NAME = "KPIMetrics";

/* ── helpers (reuse your shared ones if already defined) ── */
function _kpiOpenSS_() {
  return (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function _kpiSheet_() {
  const sh = _kpiOpenSS_().getSheetByName(KPI_SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${KPI_SHEET_NAME}" not found.`);
  return sh;
}

/* ── READ ── */
function kpi_getAll() {
  try {
    requireKpiView_();

    const sh     = _kpiSheet_();
    const values = sh.getDataRange().getDisplayValues();
    if (values.length < 2) return { ok: true, data: [] };

    // skip header row
    const rows = values.slice(1).filter(r => r.some(c => String(c).trim()));

    const data = rows.map((r, i) => ({
      id:         String(i + 1),          // row-based id (1-indexed after header)
      shift:      String(r[0] || "").trim(),
      kra:        String(r[1] || "").trim(),
      kpi:        String(r[2] || "").trim(),
      definition: String(r[3] || "").trim(),
      weight:     String(r[4] || "").trim(),
      rate5:      String(r[5] || "").trim(),
      rate4:      String(r[6] || "").trim(),
      rate3:      String(r[7] || "").trim(),
      rate2:      String(r[8] || "").trim(),
      rate1:      String(r[9] || "").trim(),
      highlight:  String(r[10] || "").trim().toLowerCase() === "yes",  // optional col K
    }));

    return { ok: true, data };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── SAVE (full sheet rewrite) ── */
function kpi_save(rows) {
  try {
    requireKpiManage_();

    if (!Array.isArray(rows)) return { ok: false, error: "Invalid payload." };

    const sh      = _kpiSheet_();
    const header  = ["Shift","KRA","KPI","Definition","Weight",
                     "Rate5","Rate4","Rate3","Rate2","Rate1","Highlight"];

    const values  = [header].concat(rows.map(r => [
      String(r.shift      || "").trim(),
      String(r.kra        || "").trim(),
      String(r.kpi        || "").trim(),
      String(r.definition || "").trim(),
      String(r.weight     || "").trim(),
      String(r.rate5      || "").trim(),
      String(r.rate4      || "").trim(),
      String(r.rate3      || "").trim(),
      String(r.rate2      || "").trim(),
      String(r.rate1      || "").trim(),
      r.highlight ? "yes" : "no",
    ]));

    sh.clearContents();
    sh.getRange(1, 1, values.length, values[0].length).setValues(values);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── DELETE ONE ROW ── */
function kpi_deleteRow(rowId) {
  try {
    requireKpiManage_();

    const sh     = _kpiSheet_();
    const values = sh.getDataRange().getDisplayValues();
    // rowId is 1-based index into data rows (after header)
    const sheetRow = parseInt(rowId) + 1; // +1 for header
    if (isNaN(sheetRow) || sheetRow < 2 || sheetRow > values.length) {
      return { ok: false, error: "Row not found." };
    }
    sh.deleteRow(sheetRow);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

function canViewKpi(email) {
  return canAccessPage("kpi", email);
}

function canManageKpi(email) {
  return canAccessPage("kpi", email) && canDoAction("kpi.manage", email);
}

function requireKpiView_() {
  const actor = Session.getActiveUser().getEmail();
  if (!canViewKpi(actor)) {
    throw new Error("You do not have permission to access KPI Metrics.");
  }
}

function requireKpiManage_() {
  const actor = Session.getActiveUser().getEmail();
  if (!canAccessPage("kpi", actor)) {
    throw new Error("You do not have permission to access KPI Metrics.");
  }
  if (!canDoAction("kpi.manage", actor)) {
    throw new Error("You do not have permission to manage KPI Metrics.");
  }
}