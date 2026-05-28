// ============================================================
//  TOP PERFORMERS — Sheet-based, admin-managed
//
//  SETUP: Create a sheet tab named "TopPerformers" in your
//  spreadsheet with these exact column headers in row 1:
//
//    A: Name
//    B: Role / Note       (e.g. "Most tickets resolved", "Best response time")
//    C: Badge             (emoji, e.g. 🥇 or leave blank to auto-assign)
//    D: Period            (e.g. "April 2026", "Q1 2026" — display only)
//    E: Active            (TRUE or FALSE — only TRUE rows are shown)
//    F: Order             (number — 1 = top, 2 = second, 3 = third)
//
//  Rows are sorted by column F (Order) ascending.
//  Only rows where Active = TRUE are returned.
//  Max 3 rows are shown on the home page.
// ============================================================


// ── GUARDS ───────────────────────────────────────────────────

/**
 * View guard — requires topperformers page access.
 * Used by admin read functions (still requires an authenticated user
 * with permission to view the admin editor).
 */
function requireTopPerformerView_() {
  const email  = getCurrentUserEmail_();
  const role   = (typeof getUserRoleByEmail_ === 'function')
                   ? getUserRoleByEmail_(email)
                   : '(unknown)';
  const canView = canAccessPage("topperformer", email);

  console.log("[TP-VIEW]",
    "email:",        email,
    "| role:",       role,
    "| canAccess:",  canView);

  if (!canView) {
    throw new Error("You do not have permission to access Top Performers.");
  }
}

/**
 * Manage guard — requires both page access and manage action.
 * Used by save/write functions.
 */
function requireTopPerformerManage_() {
  const email   = getCurrentUserEmail_();
  const role    = (typeof getUserRoleByEmail_ === 'function')
                    ? getUserRoleByEmail_(email)
                    : '(unknown)';
  const canView   = canAccessPage("topperformer",      email);
  const canManage = canDoAction  ("topperformer.manage", email);

  console.log("[TP-MANAGE]",
    "email:",       email,
    "| role:",      role,
    "| canAccess:", canView,
    "| canManage:", canManage);

  if (!canView) {
    throw new Error("You do not have permission to access Top Performers.");
  }
  if (!canManage) {
    throw new Error("You do not have permission to manage Top Performers.");
  }
}


// ── READ (called by home page — no auth required) ─────────────

/**
 * Public read — returns up to 3 active performers, sorted by Order.
 * Called by the home page for all users. No permission check.
 * Results cached for 2 minutes.
 */
function getTopPerformers() {
  try {
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'top_performers_v1';
    const cached   = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (_) {}
    }

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('TopPerformers');

    if (!sheet || sheet.getLastRow() < 2) {
      return { ok: true, data: [] };
    }

    const values  = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim().toLowerCase());
    const ci      = _tpColumnIndex_(headers);

    const AUTO_BADGES = ['🥇', '🥈', '🥉'];
    const AVATAR_STYLES = [
      { avatarBg: 'rgba(79,142,247,.18)',  avatarColor: '#4f8ef7' },
      { avatarBg: 'rgba(52,211,153,.18)',  avatarColor: '#34d399' },
      { avatarBg: 'rgba(167,139,250,.18)', avatarColor: '#a78bfa' },
    ];

    const rows = values.slice(1)
      .filter(row => _tpIsActive_(row[ci.active]))
      .sort((a, b) => (Number(a[ci.order]) || 99) - (Number(b[ci.order]) || 99))
      .slice(0, 3);

    const data = rows.map((row, i) => {
      const name = String(row[ci.name] || '').trim();
      const initials = name
        .split(/\s+/)
        .map(w => w[0] || '')
        .join('')
        .slice(0, 2)
        .toUpperCase();

      return {
        name,
        role     : String(row[ci.role]   || '').trim(),
        badge    : String(row[ci.badge]  || '').trim() || AUTO_BADGES[i] || '🏅',
        period   : String(row[ci.period] || '').trim(),
        initials,
        ...AVATAR_STYLES[i % AVATAR_STYLES.length],
      };
    });

    try {
      cache.put(cacheKey, JSON.stringify({ ok: true, data }), 120);
    } catch (_) {}

    return { ok: true, data };

  } catch (e) {
    console.error('getTopPerformers error:', e);
    return { ok: false, error: e.message, data: [] };
  }
}


// ── READ FOR ADMIN (all rows, including inactive) ─────────────

/**
 * Loads all top performer rows (including inactive) for the admin editor.
 * Requires topperformers page access.
 */
function getTopPerformersForAdmin() {
  try {
    requireTopPerformerView_();

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('TopPerformers');

    if (!sheet || sheet.getLastRow() < 2) {
      return { ok: true, data: [] };
    }

    const values  = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim().toLowerCase());
    const ci      = _tpColumnIndex_(headers);

    const data = values.slice(1)
      .filter(row => String(row[ci.name] || '').trim())
      .sort((a, b) => (Number(a[ci.order]) || 99) - (Number(b[ci.order]) || 99))
      .map((row, i) => ({
        name  : String(row[ci.name]   || '').trim(),
        role  : String(row[ci.role]   || '').trim(),
        badge : String(row[ci.badge]  || '').trim(),
        period: String(row[ci.period] || '').trim(),
        active: _tpIsActive_(row[ci.active]),
        order : Number(row[ci.order]) || (i + 1),
      }));

    return { ok: true, data };

  } catch (e) {
    console.error('getTopPerformersForAdmin error:', e);
    return { ok: false, error: e.message, data: [] };
  }
}


// ── WRITE (called by admin UI) ────────────────────────────────

/**
 * Saves the full top performers list.
 * Requires topperformers.manage action permission.
 *
 * @param {Array} performers  Array of { name, role, badge, period, active, order }
 * @returns {{ ok: boolean, error?: string }}
 */
function saveTopPerformers(performers) {
  try {
    requireTopPerformerManage_();

    if (!Array.isArray(performers)) {
      return { ok: false, error: 'Invalid data: expected an array.' };
    }

    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let   sheet = ss.getSheetByName('TopPerformers');

    // Create sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet('TopPerformers');
    }

    // Always write headers in row 1
    sheet.getRange(1, 1, 1, 6)
      .setValues([['Name', 'Role / Note', 'Badge', 'Period', 'Active', 'Order']])
      .setFontWeight('bold');

    // Clear existing data rows
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    }

    // Invalidate cache regardless — even a save of zero rows should clear stale data
    CacheService.getScriptCache().remove('top_performers_v1');

    if (!performers.length) {
      return { ok: true };
    }

    // Write new rows — cap at 3
    const rows = performers.slice(0, 3).map((p, i) => [
      String(p.name   || '').trim(),
      String(p.role   || '').trim(),
      String(p.badge  || '').trim(),
      String(p.period || '').trim(),
      p.active !== false,          // default true
      Number(p.order) || (i + 1),
    ]);

    sheet.getRange(2, 1, rows.length, 6).setValues(rows);

    // Audit log
    try {
      const email = getCurrentUserEmail_();
      auditWrite({
        action:   'topperformer.save',
        entity:   'TopPerformer',
        field:    'performers',
        oldValue: '',
        newValue: String(rows.length),
        details:  { savedCount: rows.length, savedBy: email }
      });
    } catch (_) {}

    return { ok: true };

  } catch (e) {
    console.error('saveTopPerformers error:', e);
    return { ok: false, error: e.message };
  }
}


// ── PRIVATE HELPERS ───────────────────────────────────────────

/** Build a column-index map from the sheet headers. */
function _tpColumnIndex_(headers) {
  const ci = {
    name   : headers.indexOf('name'),
    role   : headers.indexOf('role / note'),
    badge  : headers.indexOf('badge'),
    period : headers.indexOf('period'),
    active : headers.indexOf('active'),
    order  : headers.indexOf('order'),
  };

  // Positional fallbacks if headers differ slightly
  if (ci.name   < 0) ci.name   = 0;
  if (ci.role   < 0) ci.role   = 1;
  if (ci.badge  < 0) ci.badge  = 2;
  if (ci.period < 0) ci.period = 3;
  if (ci.active < 0) ci.active = 4;
  if (ci.order  < 0) ci.order  = 5;

  return ci;
}

/** Normalise an "Active" cell value to a boolean. */
function _tpIsActive_(v) {
  if (v === true)  return true;
  if (v === false) return false;
  return ['true', 'yes', '1'].includes(String(v || '').trim().toLowerCase());
}