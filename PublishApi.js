/* ============================================================
 * publishapi.gs
 * CRUD for the Posts sheet — used by publish.html
 * ============================================================ */

const POSTS_SHEET_NAME = "Posts";

/* ── SHEET BOOTSTRAP ────────────────────────────────────────
 * Call once manually to create the Posts sheet with headers.
 * Run from Apps Script editor: bootstrapPostsSheet()
 * ─────────────────────────────────────────────────────────── */

 if (typeof _openSS_ === 'undefined') {
  var _openSS_ = function() {
    return (typeof SHEET_ID !== 'undefined' && SHEET_ID)
      ? SpreadsheetApp.openById(SHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
  };
}

function bootstrapPostsSheet() {
  const ss = _openSS_();
  if (ss.getSheetByName(POSTS_SHEET_NAME)) {
    Logger.log("Posts sheet already exists.");
    return;
  }

  const sh = ss.insertSheet(POSTS_SHEET_NAME);
  const headers = ["id", "type", "pinned", "title", "body", "author", "date", "status"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  Logger.log("Posts sheet created.");
}

/* ── RBAC GUARDS ──────────────────────────────────────────── */

function renderPublish_() {
  const email = getShellUserEmail_();   // same helper your other render functions use
  if (!canAccessPage("publish", email)) {
    return { ok: false, error: "access_denied" };
  }
  const html = _getCachedShellFragment_("publish");
  return { ok: true, html, bootstrap: null };
}

function requirePublishView_() {
  const email = Session.getActiveUser().getEmail();
  if (!canAccessPage("publish", email)) {
    throw new Error("You do not have permission to access the Publish page.");
  }
}

function requirePublishWrite_() {
  requirePublishView_();
  const email = Session.getActiveUser().getEmail();
  if (!canDoAction("publish.write", email)) {
    throw new Error("You do not have permission to publish or edit posts.");
  }
}

function requirePublishDelete_() {
  requirePublishView_();
  const email = Session.getActiveUser().getEmail();
  if (!canDoAction("publish.delete", email)) {
    throw new Error("You do not have permission to delete posts.");
  }
}

/* ── SHEET HELPERS ──────────────────────────────────────────  */
function _getPostsSheet_() {
  const sh = _openSS_().getSheetByName(POSTS_SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${POSTS_SHEET_NAME}" not found. Run bootstrapPostsSheet() first.`);
  return sh;
}

function _getPostsData_() {
  const sh     = _getPostsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { sh, headers: values[0] || [], rows: [], startRow: 2 };

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const rows    = values.slice(1);
  return { sh, headers, rows, startRow: 2 };
}

function _rowToPost_(headers, row) {
  const g = name => {
    const i = headers.indexOf(name);
    return i >= 0 ? row[i] : "";
  };
  return {
    id:     Number(g("id"))                 || 0,
    type:   String(g("type")  || "").trim() || "announcement",
    pinned: String(g("pinned")).toLowerCase() === "true",
    title:  String(g("title")  || "").trim(),
    body:   String(g("body")   || "").trim(),
    author: String(g("author") || "").trim(),
    date:   String(g("date")   || "").trim(),
    status: String(g("status") || "").trim() || "draft",
  };
}

function _nextId_(rows, headers) {
  const idIdx = headers.indexOf("id");
  if (idIdx < 0 || !rows.length) return 1;
  const ids = rows.map(r => Number(r[idIdx]) || 0);
  return Math.max(...ids) + 1;
}

/* ── READ ───────────────────────────────────────────────────
 * Returns all posts (all statuses) — publish page filters client-side.
 * Home page calls getHomePosts() which filters for published only.
 * ─────────────────────────────────────────────────────────── */
function getPosts() {
  try {
    requirePublishView_();
    const { headers, rows } = _getPostsData_();
    const posts = rows
      .map(r => _rowToPost_(headers, r))
      .filter(p => p.id > 0)
      .sort((a, b) => b.id - a.id);   // newest first
    return { ok: true, data: posts };
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── READ (HOME) ────────────────────────────────────────────
 * Returns only published posts + the pinned post separately.
 * Called by home.html via getDashboardStats or directly.
 * ─────────────────────────────────────────────────────────── */
function getHomePosts() {
  try {
    requireAssetReportView_();   // home is viewable by anyone with dashboard access

    const cache     = CacheService.getScriptCache();
    const CACHE_KEY = "home_posts_v1";
    const cached    = cache.get(CACHE_KEY);
    if (cached) return JSON.parse(cached);

    const { headers, rows } = _getPostsData_();
    const all = rows
      .map(r => _rowToPost_(headers, r))
      .filter(p => p.id > 0 && p.status === "published")
      .sort((a, b) => b.id - a.id);

    const pinned  = all.find(p => p.pinned) || null;
    const feed    = all.filter(p => !p.pinned).slice(0, 20);

    const result = { ok: true, data: { pinned, feed } };
    cache.put(CACHE_KEY, JSON.stringify(result), 600);   // 10 min TTL
    return result;
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── CREATE / UPDATE ────────────────────────────────────────
 * payload: { id?, type, title, body, author, status, pinned }
 * If id is present and > 0 → UPDATE, else → INSERT.
 * ─────────────────────────────────────────────────────────── */
function savePost(payload) {
  try {
    requirePublishWrite_();

    const p = payload || {};
    if (!String(p.title || "").trim()) throw new Error("Title is required.");
    if (!String(p.body  || "").trim()) throw new Error("Body is required.");

    const VALID_TYPES   = ["announcement", "achievement", "news", "tip"];
    const VALID_STATUSES = ["published", "draft", "archived"];

    const type   = VALID_TYPES.includes(p.type)     ? p.type   : "announcement";
    const status = VALID_STATUSES.includes(p.status) ? p.status : "draft";
    const pinned = !!p.pinned;
    const author = String(p.author || "").trim() || Session.getActiveUser().getEmail();

    const { sh, headers, rows, startRow } = _getPostsData_();

    /* If pinning, clear all existing pinned flags first */
    if (pinned) {
      const pinnedIdx = headers.indexOf("pinned");
      if (pinnedIdx >= 0) {
        rows.forEach((row, i) => {
          if (String(row[pinnedIdx]).toLowerCase() === "true") {
            sh.getRange(startRow + i, pinnedIdx + 1).setValue("false");
          }
        });
      }
    }

    if (p.id && Number(p.id) > 0) {
      /* ── UPDATE ── */
      const idIdx = headers.indexOf("id");
      const rowIdx = rows.findIndex(r => Number(r[idIdx]) === Number(p.id));
      if (rowIdx < 0) throw new Error(`Post id ${p.id} not found.`);

      const shRow = startRow + rowIdx;
      const write = (col, val) => {
        const i = headers.indexOf(col);
        if (i >= 0) sh.getRange(shRow, i + 1).setValue(val);
      };
      write("type",   type);
      write("pinned", String(pinned));
      write("title",  p.title.trim());
      write("body",   p.body.trim());
      write("author", author);
      write("status", status);

    } else {
      /* ── INSERT ── */
      const newId  = _nextId_(rows, headers);
      const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy");
      const newRow = headers.map(h => {
        switch(h) {
          case "id":     return newId;
          case "type":   return type;
          case "pinned": return String(pinned);
          case "title":  return p.title.trim();
          case "body":   return p.body.trim();
          case "author": return author;
          case "date":   return today;
          case "status": return status;
          default:       return "";
        }
      });
      sh.appendRow(newRow);
    }

    _clearPostsCache_();
    return { ok: true };

  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── ARCHIVE ────────────────────────────────────────────────  */
function archivePost(id) {
  try {
    requirePublishWrite_();
    return _setPostStatus_(id, "archived");
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── RESTORE (archived → draft) ─────────────────────────────  */
function restorePost(id) {
  try {
    requirePublishWrite_();
    return _setPostStatus_(id, "draft");
  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── DELETE ─────────────────────────────────────────────────  */
function deletePost(id) {
  try {
    requirePublishDelete_();

    const { sh, headers, rows, startRow } = _getPostsData_();
    const idIdx  = headers.indexOf("id");
    const rowIdx = rows.findIndex(r => Number(r[idIdx]) === Number(id));
    if (rowIdx < 0) throw new Error(`Post id ${id} not found.`);

    sh.deleteRow(startRow + rowIdx);
    _clearPostsCache_();
    return { ok: true };

  } catch (e) {
    return { ok: false, error: e?.message || String(e) };
  }
}

/* ── INTERNAL HELPERS ───────────────────────────────────────  */
function _setPostStatus_(id, status) {
  const { sh, headers, rows, startRow } = _getPostsData_();
  const idIdx     = headers.indexOf("id");
  const statusIdx = headers.indexOf("status");
  const pinnedIdx = headers.indexOf("pinned");

  if (idIdx < 0 || statusIdx < 0) throw new Error("Posts sheet is missing required columns.");

  const rowIdx = rows.findIndex(r => Number(r[idIdx]) === Number(id));
  if (rowIdx < 0) throw new Error(`Post id ${id} not found.`);

  const shRow = startRow + rowIdx;
  sh.getRange(shRow, statusIdx + 1).setValue(status);

  /* If archiving a pinned post, also un-pin it */
  if (status === "archived" && pinnedIdx >= 0) {
    sh.getRange(shRow, pinnedIdx + 1).setValue("false");
  }

  _clearPostsCache_();
  return { ok: true };
}

function _clearPostsCache_() {
  CacheService.getScriptCache().remove("home_posts_v1");
}