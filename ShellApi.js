/* =========================================================
 * SHELL API v4 — Google / Okta SSO
 * ---------------------------------------------------------
 * - No custom session tokens
 * - No localStorage auth
 * - No login redirects from API layer
 * - Auth comes from Session.getActiveUser().getEmail()
 * ======================================================= */

const SHELL_ALLOWED_PAGES_ = new Set([
  "home",
  "masterlist",
  "employees",
  "events",
  "reports",
  "schedule",
  "aboutus",
  "admin",
  "users",
  "profile",
  "settings",
  "inventory",
  "request",
  "viewrequest",
  "myrequest",
  "companyengagement",
  "tasks",
  "assetreport",
  "vendor",
  "signedaf",
  "rbac",
  "kpi",
  "publish",
  "maintenance",
  "cheatsheet",
  "topperformer",
  "warrantyconfig"
]);

const SHELL_NO_SERVER_CACHE_PAGES_ = new Set([
  "masterlist",
  "employees",
  "users",
  "reports",
  "assetreport",
  "signedaf",
  "tasks",
  "vendor",
  "rbac",
  "profile",
  "maintenance"
]);

const SHELL_CACHE_PREFIX_ = "shell_v4_";
const SHELL_CACHE_TTL_SECONDS_ = 10 * 60; // 10 minutes
const SHELL_MAX_CACHEABLE_HTML_LENGTH_ = 90000;

/* =========================================================
 * MAIN ENTRY
 * Called by shell.html navigation engine
 * ======================================================= */
function getPageHtml(page, query, _sessionID) {
  try {
    const normalizedPage = normalizeShellPage_(page);
    const normalizedQuery = normalizeShellQuery_(query);

    console.log("[Shell] getPageHtml called", {
      rawPage: page,
      normalizedPage,
      hasQuery: Object.keys(normalizedQuery).length > 0
    });

    if (!SHELL_ALLOWED_PAGES_.has(normalizedPage)) {
      console.warn("[Shell] Page not allowed:", normalizedPage);
      return { ok: false, error: "access_denied" };
    }

    const email = getShellUserEmail_();
    if (!email) {
      console.warn("[Shell] No active user email");
      return { ok: false, error: "not_authenticated" };
    }

    const html = getPageFragment_(normalizedPage, normalizedQuery, email);
    const bootstrap = buildBootstrapSafe_(email, normalizedPage);

    return {
      ok: true,
      page: normalizedPage,
      title: "Manila IT Inventory",
      html: html || "",
      bootstrap
    };

  } catch (err) {
    console.error("[Shell] getPageHtml error:", err && err.message, err && err.stack);
    return {
      ok: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

/* =========================================================
 * PAGE FRAGMENT RENDERER
 * ======================================================= */
function getPageFragment_(page, query, email) {
  const safePage = normalizeShellPage_(page);
  const safeQuery = normalizeShellQuery_(query);
  const hasQuery = Object.keys(safeQuery).length > 0;
  const canUseServerCache = !hasQuery && !SHELL_NO_SERVER_CACHE_PAGES_.has(safePage);
  const cacheKey = buildShellCacheKey_(safePage);

  // ── PERMISSION CHECKS per page ──────────────────────────
  if (safePage === "publish" && !canAccessPage("publish", email)) {
    throw new Error("access_denied");
  }

  if (safePage === "rbac" && !canDoAction("rbac.manage", email)) {
    throw new Error("access_denied");
  }

  if (safePage === "users" && !canDoAction("users.manage", email)) {
    throw new Error("access_denied");
  }

  if (safePage === "admin" && !canDoAction("admin.access", email)) {
    throw new Error("access_denied");
  }

  if (safePage === "warrantyconfig") {
  if (!canAccessPage("warrantyconfig", email)) {
    throw new Error("access_denied");
  }
}

  if (safePage === "cheatsheet" && !canAccessPage("cheatsheet", email)) {
  throw new Error("access_denied");
}

if (safePage === "topperformer") {
    if (!canAccessPage("topperformer", email)) {
      throw new Error("access_denied");
    }
  }

  if (canUseServerCache) {
    const cached = readShellFragmentCache_(cacheKey);
    if (cached) return cached;
  }

  const template = HtmlService.createTemplateFromFile(safePage);
  template.QUERY = safeQuery;
  template.SESSION_EMAIL = email;
  template.IS_SHELL_MODE = true;

  if (safeQuery.tag != null && safeQuery.tag !== "") {
    template.SERIAL_TAG = String(safeQuery.tag).trim();
  }

  const fullHtml = template.evaluate().getContent();
  const fragmentHtml = extractBody_(fullHtml);

  if (
    canUseServerCache &&
    fragmentHtml &&
    fragmentHtml.length <= SHELL_MAX_CACHEABLE_HTML_LENGTH_
  ) {
    writeShellFragmentCache_(cacheKey, fragmentHtml);
  }

  return fragmentHtml;
}

/* =========================================================
 * HTML BODY / FRAGMENT EXTRACTOR
 * - Preserves inline <style> blocks
 * - Removes CDN assets already loaded by shell
 * ======================================================= */
function extractBody_(html) {
  if (!html) return "";

  const source = String(html);
  const styleBlocks = [];
  const styleRegex = /<style\b[^>]*>[\s\S]*?<\/style>/gi;

  let styleMatch;
  while ((styleMatch = styleRegex.exec(source)) !== null) {
    styleBlocks.push(styleMatch[0]);
  }

  const bodyMatch = source.match(/<body\b[^>]*>([\s\S]*?)<\/body>/i);
  let content = bodyMatch ? bodyMatch[1].trim() : source.trim();

  const assetPatterns = [
    /<script[^>]+cdn\.jsdelivr\.net[^>]*><\/script>/gi,
    /<script[^>]+cdnjs\.cloudflare\.com[^>]*><\/script>/gi,
    /<link[^>]+cdn\.jsdelivr\.net[^>]*>/gi,
    /<link[^>]+cdnjs\.cloudflare\.com[^>]*>/gi,
    /<link[^>]+fonts\.googleapis\.com[^>]*>/gi,
    /<link[^>]+fonts\.gstatic\.com[^>]*>/gi
  ];

  assetPatterns.forEach(function(re) {
    content = content.replace(re, "");
  });

  return (styleBlocks.join("\n") + "\n" + content).trim();
}

/* =========================================================
 * BOOTSTRAP
 * ======================================================= */
function buildBootstrapSafe_(email, page) {
  const safeEmail = String(email || "").trim().toLowerCase();
  const safePage = normalizeShellPage_(page);

  try {
    return buildBootstrap_(safeEmail, safePage);
  } catch (err) {
    console.warn("[Shell] buildBootstrap_ failed:", err && err.message);
    return {
      email: safeEmail,
      username: "",
      perms: null,
      page: safePage
    };
  }
}

function buildBootstrap_(email, page) {
  const safeEmail = String(email || "").trim().toLowerCase();
  let username = "";

  try {
    // Check cache first
    const cacheKey = "bootstrap_user_" + safeEmail;
    const cached = CacheService.getScriptCache().get(cacheKey);

    if (cached) {
      username = cached;
    } else {
      const users = getUsersData();
      const me = Array.isArray(users)
        ? users.find(function(user) {
            return normalizeEmailForShell_(user && user.email) === safeEmail;
          })
        : null;

      if (me) {
        username = String(me.username || "").trim();
        try {
          CacheService.getScriptCache().put(cacheKey, username, 300); // 5 min cache
        } catch (_) {}
      }
    }
  } catch (err) {
    console.warn("[Shell] User lookup failed:", err && err.message);
  }

  return {
    email: safeEmail,
    username: username,
    perms: getMyPermissions(),
    page: normalizeShellPage_(page)
  };
}


/* =========================================================
 * CACHE HELPERS
 * ======================================================= */
function buildShellCacheKey_(page) {
  return SHELL_CACHE_PREFIX_ + normalizeShellPage_(page);
}

function readShellFragmentCache_(cacheKey) {
  try {
    return CacheService.getScriptCache().get(cacheKey) || "";
  } catch (err) {
    console.warn("[Shell] Cache read failed:", err && err.message);
    return "";
  }
}

function writeShellFragmentCache_(cacheKey, html) {
  try {
    CacheService.getScriptCache().put(
      cacheKey,
      String(html || ""),
      SHELL_CACHE_TTL_SECONDS_
    );
  } catch (err) {
    console.warn("[Shell] Cache write failed:", err && err.message);
  }
}

/* =========================================================
 * CACHE INVALIDATION
 * ======================================================= */
function invalidateShellCache(page) {
  try {
    const cache = CacheService.getScriptCache();

    if (page) {
      cache.remove(buildShellCacheKey_(page));
    } else {
      const keys = Array.from(SHELL_ALLOWED_PAGES_).map(buildShellCacheKey_);
      cache.removeAll(keys);
    }

    return { ok: true };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

function clearMasterlistCache() {
  try {
    CacheService.getScriptCache().remove(buildShellCacheKey_("masterlist"));
    console.log("[Shell] Masterlist cache cleared");
    return { ok: true };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

/* =========================================================
 * WARM CACHE
 * ======================================================= */
function warmShellCache() {
  try {
    const email = getShellUserEmail_();
    if (!email) return { ok: false, error: "not_authenticated" };

    try { getUserRoleMap_(); } catch (_) {}
    try { getRbacRuleMap_(); } catch (_) {}
    try { getMyPermissions(); } catch (_) {}

    // Warm cacheable pages that users visit most often
    // NO_CLIENT_CACHE pages (masterlist, employees, etc.) are excluded
    // because they are not stored in server cache anyway
    const pagesToWarm = ["home", "aboutus", "schedule", "events", "cheatsheet"];
    pagesToWarm.forEach(function(page) {
      try {
        getPageFragment_(page, {}, email);
      } catch (err) {
        console.warn("[Shell] Failed warming page:", page, err && err.message);
      }
    });

    return { ok: true, email: email };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

/* =========================================================
 * NORMALIZERS / HELPERS
 * ======================================================= */
function normalizeShellPage_(page) {
  return String(page || "home")
    .trim()
    .toLowerCase()
    .replace(/^["']+|["']+$/g, "");
}

function normalizeShellQuery_(query) {
  return query && typeof query === "object" ? query : {};
}

function normalizeEmailForShell_(email) {
  return String(email || "").trim().toLowerCase();
}

function getShellUserEmail_() {
  return normalizeEmailForShell_(Session.getActiveUser().getEmail());
}