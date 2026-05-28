const RBAC_CONFIG = Object.freeze({
  RBAC_SHEET: "RBAC",
  USERS_SHEET: "Users",
  CACHE_SECONDS: 300,
  CACHE_KEYS: Object.freeze({
    RULES: "rbac_rules_v2",
    RULE_MAP: "rbac_rule_map_v2",
    ROLES: "rbac_roles_v2",
    USER_ROLE_MAP: "rbac_user_role_map_v2"
  }),
  SUPER_ROLES: Object.freeze(["admin"]),
  PAGES: Object.freeze([
    "home", "masterlist", "employees", "events", "companyengagement",
    "reports", "assetreport", "signedaf", "tasks", "vendor", "schedule",
    "aboutus", "admin", "users", "profile", "rbac", "request",
    "myrequest", "viewrequest", "publish", "maintenance", "cheatsheet",
    "topperformer", "warrantyconfig", "kpi"
  ]),
  ACTIONS: Object.freeze([
  "schedule.manage", "schedule.clear", "schedule.copy", "schedule.weekend",
  "masterlist.edit", "masterlist.delete",
  "employees.manage",
  "employees.delete",
  "events.manage",
  "companyengagement.manage",
  "reports.export",
  "assetreport.export",
  "assetreport.add",          // ← NEW
  "assetreport.bulk",
  "tasks.manage", "tasks.bulk", "tasks.reminders",
  "vendor.add", "vendor.edit", "vendor.delete",
  "users.manage",
  "admin.access",
  "rbac.manage",
  "request.manage", "request.review", "request.submit",
  "publish.write", "publish.delete",
  "topperformer.manage",
  "warrantyconfig.manage",
  "kpi.manage"
])
});

const RBAC_PAGES = RBAC_CONFIG.PAGES.slice();
const RBAC_ACTIONS = RBAC_CONFIG.ACTIONS.slice();
const RBAC_SHEET = RBAC_CONFIG.RBAC_SHEET;
const USERS_SHEET_NAME = RBAC_CONFIG.USERS_SHEET;
const RBAC_CACHE_SECONDS = RBAC_CONFIG.CACHE_SECONDS;


/* ============================================================
 * NORMALIZE / PARSE
 * ============================================================ */

function _clean_(value) {
  return String(value == null ? "" : value)
    .replace(/\u00A0/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

const normalizeRbacType_ = _clean_;
const normalizeRbacKey_ = _clean_;

function toBool_(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  return ["true", "yes", "1"].includes(String(value == null ? "" : value).trim().toLowerCase());
}

function isValidRoleName_(role) {
  return /^[a-z0-9 _-]+$/.test(_clean_(role));
}


/* ============================================================
 * SPREADSHEET ACCESS
 * ============================================================ */

if (typeof _ssInstance_ === "undefined") var _ssInstance_ = null;

function getSS_() {
  if (_ssInstance_) return _ssInstance_;
  _ssInstance_ = (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  return _ssInstance_;
}

function getUsersSheet_() {
  const sh = getSS_().getSheetByName(RBAC_CONFIG.USERS_SHEET);
  if (!sh) throw new Error("Users sheet not found.");
  return sh;
}

function ensureRbacSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(RBAC_CONFIG.RBAC_SHEET);

  if (!sh) sh = ss.insertSheet(RBAC_CONFIG.RBAC_SHEET);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([["Role", "Type", "Key", "Allowed"]]);
  }

  return sh;
}

function getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (_) {
    return "";
  }
}

function getHeaderMap_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};

  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};

  headers.forEach((header, i) => {
    const key = _clean_(header);
    if (key) map[key] = i;
  });

  return map;
}

function getUsersColumnMap_() {
  const sh = getUsersSheet_();
  const map = getHeaderMap_(sh);

  const required = {
    username: map["username"] ?? 0,
    email: map["email"] ?? 1,
    status: map["status"] ?? 3,
    role: map["role"] ?? 4,
    group: map["group"] ?? 5,
    avatarurl: map["avatarurl"] ?? 6
  };

  return required;
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================
 * CACHE HELPERS
 * ============================================================ */

function _scriptCache_() {
  try {
    return CacheService.getScriptCache();
  } catch (_) {
    return null;
  }
}

function cacheGetJson_(key) {
  const cache = _scriptCache_();
  if (!cache) return null;

  try {
    const raw = cache.get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (_) {
    return null;
  }
}

function cachePutJson_(key, value, seconds) {
  const cache = _scriptCache_();
  if (!cache) return;

  try {
    cache.put(key, JSON.stringify(value), seconds || RBAC_CONFIG.CACHE_SECONDS);
  } catch (_) {}
}

function clearShellCaches_() {
  const cache = _scriptCache_();
  if (!cache) return;

  const shellKeys = RBAC_CONFIG.PAGES.map(page => `shell_v4_${_clean_(page)}`);

  try {
    cache.removeAll(shellKeys);
  } catch (_) {}
}

function clearRbacCaches_() {
  const cache = _scriptCache_();

  if (cache) {
    try {
      cache.removeAll([
        RBAC_CONFIG.CACHE_KEYS.RULES,
        RBAC_CONFIG.CACHE_KEYS.RULE_MAP,
        RBAC_CONFIG.CACHE_KEYS.ROLES,
        RBAC_CONFIG.CACHE_KEYS.USER_ROLE_MAP
      ]);
    } catch (_) {}
  }

  try {
    if (typeof _clearDoGetUserCache_ === "function") _clearDoGetUserCache_();
  } catch (_) {}

  try {
    clearShellCaches_();
  } catch (_) {}
}


/* ============================================================
 * USERS / ROLES
 * ============================================================ */

function getUserRoleMap_() {
  const cached = cacheGetJson_(RBAC_CONFIG.CACHE_KEYS.USER_ROLE_MAP);
  if (cached && typeof cached === "object") return cached;

  const sh = getUsersSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow >= 2) {
    const values = sh.getDataRange().getValues();
    const col = getUsersColumnMap_();

    for (let i = 1; i < values.length; i++) {
      const email = _clean_(values[i][col.email]);
      const role = _clean_(values[i][col.role] || "viewer") || "viewer";
      if (email) map[email] = role;
    }
  }

  cachePutJson_(RBAC_CONFIG.CACHE_KEYS.USER_ROLE_MAP, map, RBAC_CONFIG.CACHE_SECONDS);
  return map;
}

function getUserRoleByEmail_(email) {
  const target = _clean_(email);
  if (!target) return "viewer";
  return getUserRoleMap_()[target] || "viewer";
}

function getCurrentUserRole_() {
  return getUserRoleByEmail_(getCurrentUserEmail_());
}


/* ============================================================
 * RBAC RULES
 * ============================================================ */

function getRbacRules_() {
  const cached = cacheGetJson_(RBAC_CONFIG.CACHE_KEYS.RULES);
  if (Array.isArray(cached)) return cached;

  const sh = ensureRbacSheet_();
  const values = sh.getDataRange().getValues();

  const deduped = new Map();

  for (let i = 1; i < values.length; i++) {
    const role = _clean_(values[i][0]);
    const type = _clean_(values[i][1]);
    const key = _clean_(values[i][2]);
    const allowed = toBool_(values[i][3]);

    if (!role || !type || !key) continue;
    if (type !== "page" && type !== "action") continue;

    deduped.set(`${role}|${type}|${key}`, { role, type, key, allowed });
  }

  const rules = Array.from(deduped.values());

  cachePutJson_(RBAC_CONFIG.CACHE_KEYS.RULES, rules, RBAC_CONFIG.CACHE_SECONDS);
  return rules;
}

function getRbacRuleMap_() {
  const cached = cacheGetJson_(RBAC_CONFIG.CACHE_KEYS.RULE_MAP);
  if (cached && typeof cached === "object") return cached;

  const map = {};
  const rules = getRbacRules_();

  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    map[`${r.role}|${r.type}|${r.key}`] = !!r.allowed;
  }

  cachePutJson_(RBAC_CONFIG.CACHE_KEYS.RULE_MAP, map, RBAC_CONFIG.CACHE_SECONDS);
  return map;
}

function isAllowedFromMap_(role, type, key, ruleMap) {
  const r = _clean_(role);
  const t = _clean_(type);
  const k = _clean_(key);

  if (!r || !t || !k) return false;
  if (RBAC_CONFIG.SUPER_ROLES.includes(r)) return true;

  return !!ruleMap[`${r}|${t}|${k}`];
}

function isAllowed_(role, type, key) {
  return isAllowedFromMap_(role, type, key, getRbacRuleMap_());
}

function canAccessPage(page, email) {
  const role = email ? getUserRoleByEmail_(email) : getCurrentUserRole_();
  return isAllowedFromMap_(role, "page", page, getRbacRuleMap_());
}

function canDoAction(action, email) {
  const role = email ? getUserRoleByEmail_(email) : getCurrentUserRole_();
  return isAllowedFromMap_(role, "action", action, getRbacRuleMap_());
}


/* ============================================================
 * PERMISSIONS PAYLOAD
 * ============================================================ */

function getMyPermissions() {
  const email = getCurrentUserEmail_();
  const role = getCurrentUserRole_();
  const ruleMap = getRbacRuleMap_();

  const pages = {};
  const actions = {};

  RBAC_CONFIG.PAGES.forEach(page => {
    pages[_clean_(page)] = isAllowedFromMap_(role, "page", page, ruleMap);
  });

  RBAC_CONFIG.ACTIONS.forEach(action => {
    actions[_clean_(action)] = isAllowedFromMap_(role, "action", action, ruleMap);
  });

  return { email, role, pages, actions };
}


/* ============================================================
 * COMPANY ENGAGEMENT HELPERS
 * ============================================================ */

function canViewCompanyEngagement(email) {
  return canAccessPage("companyengagement", email);
}

function canManageCompanyEngagement(email) {
  return canAccessPage("companyengagement", email) &&
         canDoAction("companyengagement.manage", email);
}

function requireCompanyEngagementView_(email) {
  if (!canViewCompanyEngagement(email)) {
    throw new Error("You do not have permission to access Company Engagement.");
  }
}

function requireCompanyEngagementManage_(email) {
  if (!canAccessPage("companyengagement", email)) {
    throw new Error("You do not have permission to access Company Engagement.");
  }
  if (!canDoAction("companyengagement.manage", email)) {
    throw new Error("You do not have permission to manage Company Engagement.");
  }
}


/* ============================================================
 * GUARDS
 * ============================================================ */

function requireUsersManage_() {
  if (!canDoAction("users.manage")) {
    throw new Error("You do not have permission to manage users.");
  }
}

function requireRbacManage_() {
  if (!canDoAction("rbac.manage")) {
    throw new Error("You do not have permission to manage RBAC.");
  }
}

function requireRoleAssignmentManage_() {
  if (canDoAction("users.manage") || canDoAction("rbac.manage")) return;
  throw new Error("You do not have permission to assign roles.");
}


/* ============================================================
 * DISTINCT ROLES
 * ============================================================ */

function getDistinctRoles_() {
  const cached = cacheGetJson_(RBAC_CONFIG.CACHE_KEYS.ROLES);
  if (Array.isArray(cached)) return cached;

  const roles = new Set(["admin", "viewer"]);

  Object.values(getUserRoleMap_()).forEach(role => {
    const cleaned = _clean_(role);
    if (cleaned) roles.add(cleaned);
  });

  getRbacRules_().forEach(rule => {
    const cleaned = _clean_(rule.role);
    if (cleaned) roles.add(cleaned);
  });

  const out = Array.from(roles).sort();
  cachePutJson_(RBAC_CONFIG.CACHE_KEYS.ROLES, out, RBAC_CONFIG.CACHE_SECONDS);
  return out;
}

function getDistinctRolesClient() {
  if (!(canDoAction("users.manage") || canDoAction("rbac.manage"))) {
    throw new Error("You do not have permission to load roles.");
  }
  return getDistinctRoles_();
}


/* ============================================================
 * MATRIX
 * ============================================================ */

function getRbacMatrix() {
  requireRbacManage_();
  return {
    roles: getDistinctRoles_(),
    pages: RBAC_CONFIG.PAGES.slice(),
    actions: RBAC_CONFIG.ACTIONS.slice(),
    rules: getRbacRules_()
  };
}

function saveRbacMatrix(payload) {
  return withScriptLock_(() => {
    requireRbacManage_();

    const validPages = new Set(RBAC_CONFIG.PAGES.map(_clean_));
    const validActions = new Set(RBAC_CONFIG.ACTIONS.map(_clean_));
    const entries = Array.isArray(payload && payload.entries) ? payload.entries : [];
    const deduped = new Map();

    for (let i = 0; i < entries.length; i++) {
      const item = entries[i];
      const role = _clean_(item && item.role);
      const type = _clean_(item && item.type);
      const key = _clean_(item && item.key);
      const allowed = !!(item && item.allowed);

      if (!role || !type || !key) continue;
      if (!isValidRoleName_(role)) continue;
      if (type !== "page" && type !== "action") continue;
      if (type === "page" && !validPages.has(key)) continue;
      if (type === "action" && !validActions.has(key)) continue;

      deduped.set(`${role}|${type}|${key}`, [role, type, key, allowed]);
    }

    const rows = [["Role", "Type", "Key", "Allowed"], ...Array.from(deduped.values())];
    const sh = ensureRbacSheet_();

    sh.clearContents();
    sh.getRange(1, 1, rows.length, 4).setValues(rows);

    clearRbacCaches_();
    return { ok: true, count: rows.length - 1 };
  });
}

function seedDefaultRbac_() {
  return withScriptLock_(() => {
    const viewerPages = new Set([
      "home", "masterlist", "employees", "events", "companyengagement",
      "reports", "assetreport", "signedaf", "schedule", "aboutus", "profile"
    ].map(_clean_));

    const roles = ["admin", "viewer"];
    const rows = [["Role", "Type", "Key", "Allowed"]];

    roles.forEach(role => {
      RBAC_CONFIG.PAGES.forEach(page => {
        const allowed = role === "admin" || viewerPages.has(_clean_(page));
        rows.push([role, "page", _clean_(page), allowed]);
      });

      RBAC_CONFIG.ACTIONS.forEach(action => {
        rows.push([role, "action", _clean_(action), role === "admin"]);
      });
    });

    const sh = ensureRbacSheet_();
    sh.clearContents();
    sh.getRange(1, 1, rows.length, 4).setValues(rows);

    clearRbacCaches_();
    return { ok: true, count: rows.length - 1 };
  });
}


/* ============================================================
 * USER ROLE ASSIGNMENT
 * ============================================================ */

function getUsersForRoleAssignment() {
  if (!(canDoAction("users.manage") || canDoAction("rbac.manage"))) {
    throw new Error("You do not have permission to view assignable users.");
  }

  const sh = getUsersSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = getUsersColumnMap_();

  return values.slice(1)
    .map((row, idx) => ({
      rowIndex: idx + 2,
      username: String(row[col.username] || "").trim(),
      email: String(row[col.email] || "").trim(),
      status: String(row[col.status] || "").trim(),
      role: _clean_(row[col.role] || "viewer"),
      group: String(row[col.group] || "").trim(),
      avatarUrl: String(row[col.avatarurl] || "").trim()
    }))
    .filter(user => !!user.email);
}

function assignRoleToUser(payload) {
  return withScriptLock_(() => {
    requireRoleAssignmentManage_();

    const email = _clean_(payload && payload.email);
    const newRole = _clean_(payload && payload.role);

    if (!email) throw new Error("Email is required.");
    if (!newRole) throw new Error("Role is required.");
    if (!isValidRoleName_(newRole)) throw new Error("Invalid role name.");
    if (!new Set(getDistinctRoles_()).has(newRole)) throw new Error("Role does not exist in RBAC.");

    const sh = getUsersSheet_();
    const values = sh.getDataRange().getValues();
    const col = getUsersColumnMap_();

    for (let i = 1; i < values.length; i++) {
      if (_clean_(values[i][col.email]) === email) {
        sh.getRange(i + 1, col.role + 1).setValue(newRole);
        clearRbacCaches_();
        return { ok: true, email, role: newRole };
      }
    }

    throw new Error("User not found.");
  });
}

function bulkAssignRoleToUsers(payload) {
  return withScriptLock_(() => {
    requireRoleAssignmentManage_();

    const emails = Array.isArray(payload && payload.emails)
      ? payload.emails.map(_clean_).filter(Boolean)
      : [];
    const emailSet = new Set(emails);
    const newRole = _clean_(payload && payload.role);

    if (!emails.length) throw new Error("No users selected.");
    if (!newRole) throw new Error("Role is required.");
    if (!isValidRoleName_(newRole)) throw new Error("Invalid role name.");
    if (!new Set(getDistinctRoles_()).has(newRole)) throw new Error("Role does not exist in RBAC.");

    const sh = getUsersSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, updated: 0, role: newRole };

    const col = getUsersColumnMap_();
    const lastCol = Math.max(col.email, col.role) + 1;
    const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const roleUpdates = [];
    let updated = 0;

    for (let i = 0; i < values.length; i++) {
      const email = _clean_(values[i][col.email]);
      const currentRole = _clean_(values[i][col.role] || "viewer");
      const nextRole = emailSet.has(email) ? newRole : currentRole;

      roleUpdates.push([nextRole]);
      if (emailSet.has(email) && currentRole !== nextRole) updated++;
    }

    if (roleUpdates.length) {
      sh.getRange(2, col.role + 1, roleUpdates.length, 1).setValues(roleUpdates);
    }

    clearRbacCaches_();
    return { ok: true, updated, role: newRole };
  });
}

function getUsersByRole(role) {
  requireRbacManage_();
  const target = _clean_(role);
  return getUsersForRoleAssignment().filter(user => _clean_(user.role) === target);
}


/* ============================================================
 * DEBUG HELPERS
 * ============================================================ */

function debugMyRbac() {
  const email = getCurrentUserEmail_();
  const role = getCurrentUserRole_();

  return {
    email,
    role,
    permissions: getMyPermissions(),
    canViewCompanyEngagement: canViewCompanyEngagement(email),
    canManageCompanyEngagement: canManageCompanyEngagement(email),
    rules: getRbacRules_().filter(rule => rule.role === role)
  };
}

function testPermission_(email, type, key) {
  const role = getUserRoleByEmail_(email);
  return {
    email: _clean_(email),
    role,
    type: _clean_(type),
    key: _clean_(key),
    allowed: isAllowed_(role, type, key)
  };
}