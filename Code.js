// Single place to update your spreadsheet ID
const SHEET_ID = "1M2zf8c5J39rGkH-yNAy8Xn2LNoVuj9GbnK6gl8x3xEU";
const SS_ID = SHEET_ID;
const AVATAR_FOLDER_ID = "1Jvf0HkuPG2VwmfvalsmaWhR0CT9RJp1f";


const SHELL_PAGES_ = new Set([
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
  "warrantyconfig",
  "topperformer"
]);

function doGet(e) {
  try {
    // ── Image proxy ───────────────────────────────────────────────
    // Serves Drive image files as base64 text, bypassing the GAS
    // message channel (avoids base64 timeout issues in iframes).
    // URL format: ?action=img&id=DRIVE_FILE_ID
    if (e && e.parameter && e.parameter.action === 'img' && e.parameter.id) {
      try {
        const file     = DriveApp.getFileById(e.parameter.id);
        const blob     = file.getBlob();
        const mime     = blob.getContentType() || '';

        // Security: only serve actual image files — reject anything else
        // to prevent this endpoint being used to read arbitrary Drive files
        if (!mime.startsWith('image/')) {
          console.warn('[doGet] Image proxy rejected non-image file:', e.parameter.id, mime);
          return ContentService
            .createTextOutput('')
            .setMimeType(ContentService.MimeType.TEXT);
        }

        return ContentService
          .createTextOutput(Utilities.base64Encode(blob.getBytes()))
          .setMimeType(ContentService.MimeType.TEXT);

      } catch (imgErr) {
        console.warn('[doGet] Image proxy failed:', imgErr && imgErr.message);
        return ContentService
          .createTextOutput('')
          .setMimeType(ContentService.MimeType.TEXT);
      }
    }
    // ─────────────────────────────────────────────────────────────

    const params = getRequestParams_(e);

    if (isApprovalRequest_(params)) {
      return handleApproval(e);
    }

    const route = resolveRoute_(params);

    if (route.type === 'uploadsigned') {
      return renderUploadSignedPage_(params);
    }

    const email = getSignedInEmail_();
    if (!email) {
      return renderSignInRequired_();
    }

    ensureUserExists_(email);

    return renderShell_(route.page, params, email);

  } catch (err) {
    console.error('doGet error:', err && err.stack ? err.stack : err);
    return renderDoGetError_();
  }
}

function getRequestParams_(e) {
  const raw = (e && e.parameter && typeof e.parameter === "object")
    ? e.parameter
    : {};

  return raw || {};
}

function isApprovalRequest_(params) {
  return hasTruthyParam_(params.approve) || hasTruthyParam_(params.reject);
}

function resolveRoute_(params) {
  const requestedPage = normalizePage_(params.page);

  if (requestedPage === "uploadsigned" || hasTruthyParam_(params.uploadSigned)) {
    return { type: "uploadsigned", page: "uploadsigned" };
  }

  return {
    type: "shell",
    page: SHELL_PAGES_.has(requestedPage) ? requestedPage : "home"
  };
}

function normalizePage_(page) {
  return String(page || "home")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, "");
}

function hasTruthyParam_(value) {
  const v = String(value || "").trim().toLowerCase();
  return v === "1" || v === "true" || v === "yes" || v === "y";
}

function getSignedInEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || "")
      .trim()
      .toLowerCase();
  } catch (err) {
    console.error("getSignedInEmail_ error:", err && err.stack ? err.stack : err);
    return "";
  }
}

function renderUploadSignedPage_(params) {
  const tpl = HtmlService.createTemplateFromFile("uploadsigned");
  tpl.QUERY = params || {};
  tpl.SERIAL_TAG = String((params && params.tag) || "").trim();

  return tpl.evaluate()
    .setTitle("Upload Signed Accountability Form")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderShell_(initialPage, params, email) {
  const tpl = HtmlService.createTemplateFromFile("shell");
  tpl.QUERY = params || {};
  tpl.INITIAL_PAGE = initialPage || "home";
  tpl.INITIAL_QUERY = params || {};
  tpl.INITIAL_EMAIL = String(email || "").trim().toLowerCase();

  return tpl.evaluate()
    .setTitle("Manila IT Inventory")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderSignInRequired_() {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body {
            font-family: system-ui, sans-serif;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            margin: 0;
            background: #060608;
            color: #fff;
          }
          .box {
            text-align: center;
            padding: 40px;
          }
          p {
            color: rgba(255,255,255,.55);
            margin-top: 8px;
          }
        </style>
      </head>
      <body>
        <div class="box">
          <h2>Sign in required</h2>
          <p>Please sign in with your company Google account to continue.</p>
        </div>
      </body>
    </html>
  `)
    .setTitle("Sign in required")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function renderDoGetError_() {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <style>
          body {
            font-family: system-ui, sans-serif;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100vh;
            margin: 0;
            background: #060608;
            color: #fff;
          }
          .box {
            text-align: center;
            padding: 40px;
            max-width: 520px;
          }
          p {
            color: rgba(255,255,255,.6);
            margin-top: 8px;
          }
        </style>
      </head>
      <body>
        <div class="box">
          <h2>Something went wrong</h2>
          <p>The app could not be loaded. Please refresh and try again.</p>
        </div>
      </body>
    </html>
  `)
    .setTitle("Error")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function _clearDoGetUserCache_() {
  // Stub — no per-user doGet cache in SSO mode
  // Called by clearRbacCaches_() after role changes
}

/**
 * Auto-provision a new user with viewer role if they don't exist yet.
 * Called on every doGet so new org members get access automatically.
 */
function ensureUserExists_(email) {
  var normalizedEmail = String(email || "").trim().toLowerCase();
  if (!normalizedEmail) return;

  try {
    var cache    = CacheService.getScriptCache();
    var cacheKey = "user_exists_" + normalizedEmail;
    if (cache.get(cacheKey) === "1") return;

    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("Users");
    if (!sh) return;

    var lastRow = sh.getLastRow();
    if (lastRow >= 2) {
      var values = sh.getRange(2, 2, lastRow - 1, 1).getDisplayValues();
      for (var i = 0; i < values.length; i++) {
        if (String(values[i][0] || "").trim().toLowerCase() === normalizedEmail) {
          try { cache.put(cacheKey, "1", 300); } catch (e) {}
          return;
        }
      }
    }

    var parts    = normalizedEmail.split("@")[0].split(/[._-]+/);
    var username = parts.map(function(p) {
      return p.charAt(0).toUpperCase() + p.slice(1);
    }).join(" ").trim();

    sh.appendRow([
      username,        // A Username
      normalizedEmail, // B Email
      new Date(),      // C Created
      "Active",        // D Status
      "viewer",        // E Role
      "Morning",       // F Group
      ""               // G AvatarUrl
    ]);

    try { cache.put(cacheKey, "1", 300); } catch (e) {}
    invalidateUsersCache_();

  } catch (err) {
    console.error("ensureUserExists_ failed:", err);
  }
}


// Keep below OUTSIDE doGet()
function include(file){ return HtmlService.createHtmlOutputFromFile(file).getContent(); }


function upsertPendingUser_(email, code, now) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("PendingUsers") || ss.insertSheet("PendingUsers");

    // Ensure headers once
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Email", "Code", "Status", "Date"]);
    }

    const last = sh.getLastRow();
    if (last < 2) {
      sh.appendRow([email, code, "Pending Approval", now]);
      return;
    }

    // Search A2:A for email (TextFinder)
    const emailRange = sh.getRange(2, 1, last - 1, 1);
    const cell = emailRange.createTextFinder(email).matchEntireCell(true).findNext();

    if (cell) {
      const r = cell.getRow();
      sh.getRange(r, 2, 1, 3).setValues([[code, "Pending Approval", now]]); // B:C:D
    } else {
      sh.appendRow([email, code, "Pending Approval", now]);
    }
  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
             APPROVAL HANDLER – ADMIN EMAIL CLICK
=========================================================*/
function handleApproval(e) {
  const params     = (e && e.parameter) ? e.parameter : {};
const approveRaw = params.approve ? decodeURIComponent(params.approve) : "";
const rejectRaw  = params.reject  ? decodeURIComponent(params.reject)  : "";

  const action = approveRaw ? "approve" : (rejectRaw ? "reject" : "");
  const email  = (approveRaw || rejectRaw || "").trim().toLowerCase();

  if (!action) {
    return renderDecisionPage_({
      kind: "error",
      title: "No action provided",
      message: "Missing approve/reject parameter."
    });
  }

  if (!isValidEmail_(email)) {
    return renderDecisionPage_({
      kind: "error",
      title: "Invalid email",
      message: `Invalid email: ${escapeHtml_(email)}`
    });
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("PendingUsers");
  if (!sh) {
    return renderDecisionPage_({
      kind: "error",
      title: "Missing sheet",
      message: "PendingUsers sheet not found."
    });
  }

  const desiredStatus = action === "approve" ? "Approved" : "Rejected";

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const result = updatePendingStatus_(sh, email, desiredStatus);

    if (result.kind === "not_found") {
      return renderDecisionPage_({
        kind: "error",
        title: "Not found",
        message: `No pending registration found for ${escapeHtml_(email)}.`
      });
    }

    if (result.kind === "already") {
      return renderDecisionPage_({
        kind: "info",
        title: "Already processed",
        message: `${escapeHtml_(email)} is already <b>${escapeHtml_(result.currentStatus)}</b>.`,
        meta: result
      });
    }

    const subject = desiredStatus === "Approved"
      ? "✔ Manila IT – Account Approved"
      : "❌ Manila IT – Registration Rejected";

    const htmlBody = desiredStatus === "Approved"
      ? approvedEmailHtml_(email)
      : rejectedEmailHtml_(email);

    try {
      GmailApp.sendEmail(email, subject, "", {
        from: "manila-it@fbgphilippines.com",
        name: "FBG Manila IT",
        replyTo: "manila-it@fbgphilippines.com",
        htmlBody
      });
    } catch (mailErr) {
      console.error("Approval email failed:", mailErr);
    }

    return renderDecisionPage_({
      kind: desiredStatus === "Approved" ? "success" : "danger",
      title: desiredStatus,
      message: `${escapeHtml_(email)} has been <b>${escapeHtml_(desiredStatus)}</b>.`,
      meta: result
    });

  } finally {
    lock.releaseLock();
  }
}


/**
 * Updates PendingUsers status and returns what happened.
 * Sheet columns assumed: A Email, B Code, C Status, D Date
 */
function updatePendingStatus_(sh, email, newStatus) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { kind: "not_found" };

  const data = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // A:D

  let best = null;
  for (let i = 0; i < data.length; i++) {
    const rowEmail = String(data[i][0] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;

    const status = String(data[i][2] || "").trim();
    const dateVal = data[i][3];
    const issuedMs = dateVal instanceof Date ? dateVal.getTime() : 0;

    if (!best || issuedMs >= best.issuedMs) {
      best = { idx: i, row: i + 2, currentStatus: status, issuedMs };
    }
  }

  if (!best) return { kind: "not_found" };

  if (best.currentStatus === newStatus) {
    return { kind: "already", email, currentStatus: best.currentStatus, row: best.row, issuedMs: best.issuedMs };
  }

  if (best.currentStatus === "Approved" || best.currentStatus === "Rejected") {
    return { kind: "already", email, currentStatus: best.currentStatus, row: best.row, issuedMs: best.issuedMs };
  }

  sh.getRange(best.row, 3).setValue(newStatus);

  return { kind: "updated", email, previousStatus: best.currentStatus, currentStatus: newStatus, row: best.row, issuedMs: best.issuedMs, updatedAt: Date.now() };
}

function approvedEmailHtml_(email) {
  const loginUrl = getLoginUrl();
  const safeEmail = escapeHtml_(email);

  return `
<div style="margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
  <div style="max-width:560px;margin:28px auto;padding:0 16px;">
    <div style="background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);">

      <div style="padding:18px 20px;background:linear-gradient(135deg,#30D158,#2ACB70);color:#fff;">
        <div style="font-size:13px;opacity:.92;">Manila IT Inventory</div>
        <div style="font-size:18px;font-weight:800;margin-top:4px;">Account Approved</div>
      </div>

      <div style="padding:20px;color:#f5f5f7;">
        <p style="margin:0 0 12px;font-size:14px;line-height:1.55;color:#d1d1d6;">
          Your registration for <b style="color:#fff;">${safeEmail}</b> has been approved.
        </p>

        <div style="margin:16px 0;padding:14px 16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;text-align:center;">
          <a href="${loginUrl}"
             style="display:inline-block;background:#ffffff;color:#000;padding:12px 18px;border-radius:12px;text-decoration:none;font-weight:800;font-size:13px;">
            Go to Login
          </a>
        </div>

        <p style="margin:0;font-size:12px;color:#9a9aa0;line-height:1.45;">
          If you didn't request this registration, please contact IT.
        </p>
      </div>
    </div>

    <div style="text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;">
      © ${new Date().getFullYear()} Manila IT Inventory
    </div>
  </div>
</div>`;
}

function rejectedEmailHtml_(email) {
  const safeEmail = escapeHtml_(email);

  return `
<div style="margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
  <div style="max-width:560px;margin:28px auto;padding:0 16px;">
    <div style="background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);">

      <div style="padding:18px 20px;background:linear-gradient(135deg,#FF453A,#FF6A3D);color:#fff;">
        <div style="font-size:13px;opacity:.92;">Manila IT Inventory</div>
        <div style="font-size:18px;font-weight:800;margin-top:4px;">Registration Rejected</div>
      </div>

      <div style="padding:20px;color:#f5f5f7;">
        <p style="margin:0 0 12px;font-size:14px;line-height:1.55;color:#d1d1d6;">
          Your registration for <b style="color:#fff;">${safeEmail}</b> was rejected.
        </p>

        <div style="margin:16px 0;padding:14px 16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;">
          <div style="font-size:12px;color:#b0b0b5;line-height:1.5;">
            If you think this is a mistake, please contact IT.
          </div>
        </div>

        <p style="margin:0;font-size:12px;color:#9a9aa0;line-height:1.45;">
          You can re-register using a different email if needed.
        </p>
      </div>
    </div>

    <div style="text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;">
      © ${new Date().getFullYear()} Manila IT Inventory
    </div>
  </div>
</div>`;
}

function renderDecisionPage_(opts) {
  var kind    = opts.kind;
  var title   = opts.title;
  var message = opts.message;
  var meta    = opts.meta;

  var colorMap = {
    success: { bar: "#30d158", glow: "rgba(48,209,88,.22)",   bg: "rgba(48,209,88,0.10)",  chip: "rgba(48,209,88,0.16)"  },
    danger:  { bar: "#ff453a", glow: "rgba(255,69,58,.22)",   bg: "rgba(255,69,58,0.10)",  chip: "rgba(255,69,58,0.16)"  },
    info:    { bar: "#0a84ff", glow: "rgba(10,132,255,.22)",  bg: "rgba(10,132,255,0.10)", chip: "rgba(10,132,255,0.16)" },
    error:   { bar: "#ff9f0a", glow: "rgba(255,159,10,.22)",  bg: "rgba(255,159,10,0.10)", chip: "rgba(255,159,10,0.16)" }
  };
  var colors = colorMap[kind] || { bar: "#86868b", glow: "rgba(255,255,255,.10)", bg: "rgba(255,255,255,0.06)", chip: "rgba(255,255,255,0.10)" };

  var iconMap = { success: "✅", danger: "⛔", info: "ℹ️", error: "⚠️" };
  var icon = iconMap[kind] || "ℹ️";

  var metaRow    = meta ? (meta.row != null ? meta.row : "-") : "-";
  var metaStatus = meta ? (meta.currentStatus != null ? meta.currentStatus : (meta.status != null ? meta.status : "-")) : "-";
  var metaTime   = meta ? (meta.updatedAt ? new Date(meta.updatedAt).toLocaleString() : "-") : "-";

  var metaHtml = meta
    ? '<div class="meta">' +
        '<div class="meta-row"><span class="k">Row</span><span class="v">' + metaRow + '</span></div>' +
        '<div class="meta-row"><span class="k">Status</span><span class="v">' + escapeHtml_(metaStatus) + '</span></div>' +
        '<div class="meta-row"><span class="k">Time</span><span class="v">' + metaTime + '</span></div>' +
      '</div>'
    : "";

  const html = `
<!doctype html>
<html>
<head>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <base target="_top" />
  <title>Manila IT Approval</title>
  <style>
    :root{
      --bg0:#0b0b0c;
      --bg1:#111114;
      --card: rgba(17,17,20,0.78);
      --stroke: rgba(255,255,255,0.12);
      --stroke2: rgba(255,255,255,0.08);
      --text:#f5f5f7;
      --muted:#a1a1aa;
      --muted2:#8e8e93;
      --shadow: 0 22px 60px rgba(0,0,0,.60);
      --radius: 22px;
      --radius2: 16px;
      --bar:${colors.bar};
      --glow:${colors.glow};
      --msgbg:${colors.bg};
      --chip:${colors.chip};
    }

    *{ box-sizing:border-box; }
    body{
      margin:0;
      min-height:100vh;
      display:flex;
      align-items:center;
      justify-content:center;
      padding:24px;
      font-family:-apple-system,BlinkMacSystemFont,"SF Pro Text","Segoe UI",Roboto,Arial,sans-serif;
      background:
        radial-gradient(1200px 700px at 70% 20%, rgba(10,132,255,0.10), transparent 60%),
        radial-gradient(900px 600px at 20% 30%, var(--glow), transparent 62%),
        linear-gradient(135deg, rgba(0,0,0,0.94), rgba(45,45,45,0.55), rgba(0,0,0,0.90));
      color:var(--text);
    }

    .wrap{ width:560px; max-width:100%; }

    .card{
      position:relative;
      overflow:hidden;
      border-radius:var(--radius);
      background: var(--card);
      border:1px solid var(--stroke);
      box-shadow: var(--shadow);
      backdrop-filter:saturate(180%) blur(18px);
      -webkit-backdrop-filter:saturate(180%) blur(18px);
    }

    .topbar{
      height:6px;
      background: var(--bar);
      box-shadow: 0 0 0 1px rgba(255,255,255,.06) inset, 0 0 28px var(--glow);
    }

    .head{
      padding:18px 20px 12px;
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
      border-bottom:1px solid var(--stroke2);
      background: linear-gradient(135deg, rgba(10,132,255,0.18), rgba(41,151,255,0.06), rgba(255,255,255,0.02));
    }

    .brand{ display:flex; flex-direction:column; gap:2px; min-width:0; }
    .brand .app{
      font-size:13px; letter-spacing:.02em;
      color:rgba(255,255,255,0.86); font-weight:700;
      white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
    }
    .brand .sub{ font-size:12px; color:var(--muted2); font-weight:600; }

    .pill{
      display:inline-flex; align-items:center; gap:8px;
      padding:8px 12px; border-radius:999px;
      background: var(--chip); border:1px solid rgba(255,255,255,0.10);
      font-size:12px; font-weight:800; color: rgba(255,255,255,0.92); white-space:nowrap;
    }

    .body{ padding:18px 20px 20px; }

    .title{
      margin:4px 0 10px; font-size:22px; line-height:1.2;
      font-weight:900; letter-spacing:-0.02em;
      display:flex; align-items:center; gap:10px;
    }
    .title .ic{
      width:34px; height:34px; border-radius:12px;
      display:grid; place-items:center;
      background: rgba(255,255,255,0.06);
      border:1px solid rgba(255,255,255,0.10);
      box-shadow: 0 0 0 4px rgba(255,255,255,0.02) inset;
      flex:0 0 auto;
    }

    .msg{
      margin-top:10px; padding:14px 14px; border-radius:14px;
      background: var(--msgbg); border:1px solid rgba(255,255,255,0.10);
      color: rgba(245,245,247,0.92); font-size:14px; line-height:1.5; font-weight:650;
    }

    .meta{
      margin-top:14px; border-radius:14px;
      border:1px solid rgba(255,255,255,0.10);
      background: rgba(255,255,255,0.04); overflow:hidden;
    }
    .meta-row{
      display:flex; justify-content:space-between; gap:12px;
      padding:10px 12px; border-top:1px solid rgba(255,255,255,0.06);
    }
    .meta-row:first-child{ border-top:none; }
    .k{ color:var(--muted2); font-size:12px; font-weight:800; letter-spacing:.06em; text-transform:uppercase; }
    .v{ color:rgba(245,245,247,0.92); font-size:12px; font-weight:800; text-align:right; word-break:break-word; }

    .foot{ margin-top:14px; text-align:center; font-size:11px; color:rgba(142,142,147,0.9); }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <div class="topbar"></div>

      <div class="head">
        <div class="brand">
          <div class="app">Manila IT Inventory</div>
          <div class="sub">Approval Result</div>
        </div>
        <div class="pill">${icon} ${escapeHtml_(title)}</div>
      </div>

      <div class="body">
        <div class="title">
          <span class="ic">${icon}</span>
          <span>${escapeHtml_(title)}</span>
        </div>

        <div class="msg">${message}</div>

        ${metaHtml}

        <div class="foot">
          You can close this tab after reviewing the result.
        </div>
      </div>
    </div>
  </div>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Manila IT Approval")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function escapeApproved_(s){ return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function escapeHtml_(s){ return escapeApproved_(s); }

function safeGetAppUrl_() {
  try {
    return typeof getAppUrl === "function" ? getAppUrl() : ScriptApp.getService().getUrl();
  } catch (_) {
    return "";
  }
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}


/* =========================================================
           VERIFY CODE BEFORE FINISHING REGISTRATION
=========================================================*/
function verifyCode(email, code) {
  email = String(email || "").trim().toLowerCase();
  code  = String(code  || "").trim();
  if (!email || !code) return "invalid";

  const TTL_MS = 10 * 60 * 1000;
  const now = Date.now();
  const key = `otp_${email}`;

  const normalizeIssued = (v) => {
    if (!v) return 0;
    if (v instanceof Date) return v.getTime();
    let n = (typeof v === "number") ? v : Number(String(v).replace(/[^\d.]/g, ""));
    if (!isFinite(n) || n <= 0) return 0;
    if (n < 60000) {
      const ms = Math.round((n - 25569) * 86400 * 1000);
      return ms > 0 ? ms : 0;
    }
    if (n < 1e12) return Math.round(n * 1000);
    return Math.round(n);
  };

  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);

  if (cached) {
    try {
      const obj = JSON.parse(cached);
      const issued = normalizeIssued(obj.issued);
      const cachedCode = String(obj.code || "").trim();

      if (!issued || (now - issued) > TTL_MS) {
        cache.remove(key);
        return "expired";
      }
      if (cachedCode === code) return checkApprovalStatus_(email);
      return "invalid";
    } catch (e) {
      // fall through
    }
  }

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("PendingUsers");
  if (!sh || sh.getLastRow() < 2) return "invalid";

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(); // A:D
  let best = null;

  for (let i = 0; i < data.length; i++) {
    const rowEmail = String(data[i][0] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;

    const issued = normalizeIssued(data[i][3]);

    if (!best || issued > best.issued) {
      best = {
        code: String(data[i][1] || "").trim(),
        status: String(data[i][2] || "").trim(),
        issued
      };
    }
  }

  if (!best) return "invalid";
  if (!best.issued || (now - best.issued) > TTL_MS) return "expired";
  if (best.code !== code) return "invalid";

  return best.status === "Approved" ? "verified" : "waiting_approval";
}


function checkApprovalStatus_(email) {
  email = String(email || "").trim().toLowerCase();

  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("PendingUsers");
  if (!sh || sh.getLastRow() < 2) return "invalid";

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(); // A:D
  let best = null;

  for (let i = 0; i < data.length; i++) {
    const rowEmail = String(data[i][0] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;

    const v = data[i][3];
    let issued = 0;
    if (v instanceof Date) issued = v.getTime();
    else {
      const t = Date.parse(String(v));
      if (!isNaN(t)) issued = t;
    }

    if (!best || issued > best.issued) {
      best = { status: String(data[i][2] || "").trim(), issued };
    }
  }

  if (!best) return "invalid";
  return best.status === "Approved" ? "verified" : "waiting_approval";
}


/* =========================================================
     UPDATE APPROVAL STATUS IN SHEET (REQUIRED FUNC)
=========================================================*/
function updateStatus(sheet, email, newStatus) {
  if (!sheet) return { ok: false, error: "no_sheet" };

  email = String(email || "").trim().toLowerCase();
  newStatus = String(newStatus || "").trim();
  if (!email || !newStatus) return { ok: false, error: "bad_args" };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, error: "empty" };

  const emailRange = sheet.getRange(2, 1, lastRow - 1, 1);
  const finder = emailRange.createTextFinder(email).matchEntireCell(true);

  let bestRow = null;
  let bestIssued = -1;
  let currentStatus = "";

  let cell = finder.findNext();
  while (cell) {
    const r = cell.getRow();
    const issuedVal = sheet.getRange(r, 4).getValue(); // D
    const issued = issuedVal instanceof Date ? issuedVal.getTime() : 0;

    if (issued >= bestIssued) {
      bestIssued = issued;
      bestRow = r;
      currentStatus = String(sheet.getRange(r, 3).getValue() || "").trim(); // C
    }
    cell = finder.findNext();
  }

  if (!bestRow) return { ok: false, error: "not_found" };

  if (currentStatus === newStatus) {
    return { ok: true, changed: false, row: bestRow, status: currentStatus };
  }

  if (currentStatus === "Approved" || currentStatus === "Rejected") {
    return { ok: false, error: "already_final", row: bestRow, status: currentStatus };
  }

  sheet.getRange(bestRow, 3).setValue(newStatus);
  return { ok: true, changed: true, row: bestRow, previous: currentStatus, status: newStatus };
}


/* =========================================================
         REGISTER USER AFTER VERIFIED + APPROVED
=========================================================*/
function registerUser(email, username, password) {
  email    = String(email    || "").trim().toLowerCase();
  username = String(username || "").trim();

  if (!email || !username || !password) {
    return { ok: false, error: "missing_fields" };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(8000);

  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const users = ss.getSheetByName("Users") || ss.insertSheet("Users");

    // FIX #2: Write only the headers row when the sheet is brand new.
    // The original code used passHash/salt before they were declared,
    // and also wrote a real data row inside the header-setup block.
    if (users.getLastRow() === 0) {
      users.appendRow([
        "Username", "Email", "PasswordHash", "Salt", "Created",
        "Status", "Role", "FailCount", "Position", "Group",
        "MustChange", "AvatarUrl"
      ]);
    }

    // Prevent duplicates (email in column B)
    const last = users.getLastRow();
    if (last > 1) {
      const emails = users.getRange(2, 2, last - 1, 1).getValues().flat()
        .map(e => String(e || "").trim().toLowerCase());
      if (emails.includes(email)) {
        return { ok: false, error: "already_registered" };
      }
    }

    const salt     = generateSalt();
    const passHash = hashPassword(password, salt);

    users.appendRow([
      username,
      email,
      passHash,
      salt,
      new Date(),
      "Active",
      "User",
      0,
      "",
      "Morning",
      false,
      ""
    ]);

    return { ok: true, loginUrl: getLoginUrl() };

  } finally {
    lock.releaseLock();
  }
}


/* ================= Helpers ================= */

function ensureUsersHeaders_(sh) {
  const headers = ["Username", "Email", "Created", "Status", "Role", "Group", "AvatarUrl"];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  const firstRow   = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const normalized = firstRow.map(v => String(v || "").trim());
  const ok         = headers.every((h, i) => normalized[i] === h);

  if (!ok) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}



function isValidEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i.test(s);
}
function validUsername_(u) {
  return /^[a-zA-Z0-9._-]{3,20}$/.test(u);
}
function validPassword_(p) {
  return (
    p.length >= 8 &&
    /[a-z]/.test(p) &&
    /[A-Z]/.test(p) &&
    /[0-9]/.test(p) &&
    /[^A-Za-z0-9]/.test(p)
  );
}


/* =========================================================
      REMOVE ENTRY AFTER ACCOUNT VERIFIED SUCCESSFULLY
=========================================================*/
function removePending(email) {
  const sheet = SpreadsheetApp.openById(SHEET_ID)
    .getSheetByName("PendingUsers");
  if (!sheet) return "no_pending_sheet";

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == email) {
      sheet.deleteRow(i + 1);
      return "removed";
    }
  }
  return "not_found";
}

/* =========================================================
            LIVE EMAIL CHECK — Prevent duplicate use
=========================================================*/
function checkEmailExists(email) {
  email = String(email || "").trim().toLowerCase();
  if (!email) return "available";

  const cache = CacheService.getScriptCache();
  const cacheKey = `emailcheck_${email}`;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const usersSh = ss.getSheetByName("Users");
  if (usersSh && usersSh.getLastRow() >= 2) {
    const rng = usersSh.getRange(2, 2, usersSh.getLastRow() - 1, 1); // B2:B
    const hit = rng.createTextFinder(email).matchEntireCell(true).findNext();
    if (hit) {
      cache.put(cacheKey, "exists", 30);
      return "exists";
    }
  }

  const pendingSh = ss.getSheetByName("PendingUsers");
  if (!pendingSh || pendingSh.getLastRow() < 2) {
    cache.put(cacheKey, "available", 30);
    return "available";
  }

  const TTL_MS = 10 * 60 * 1000;
  const now = Date.now();

  const emailRng = pendingSh.getRange(2, 1, pendingSh.getLastRow() - 1, 1); // A2:A
  const cell = emailRng.createTextFinder(email).matchEntireCell(true).findNext();

  if (!cell) {
    cache.put(cacheKey, "available", 30);
    return "available";
  }

  const row = cell.getRow();
  const status  = String(pendingSh.getRange(row, 3).getValue() || "").trim();
  const dateVal = pendingSh.getRange(row, 4).getValue();
  const issued  = (dateVal instanceof Date) ? dateVal.getTime() : 0;

  const expired = !issued || (now - issued) > TTL_MS;

  if (expired || status === "Rejected") {
    cache.put(cacheKey, "available", 15);
    return "available";
  }

  if (status === "Pending Approval" || status === "Approved") {
    cache.put(cacheKey, "awaiting", 15);
    return "awaiting";
  }

  cache.put(cacheKey, "available", 15);
  return "available";
}


const RESEND_COOLDOWN_SECONDS = 60;
const APPROVAL_TOKEN_TTL_SECONDS = 86400; // 24 hours

function normalizeEmail_(s){ return String(s || "").trim().toLowerCase(); }

function ensureSheet_(ss, name, headers){
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 && headers && headers.length) sh.appendRow(headers);
  return sh;
}

function findRowByEmail_(sheet, email, emailColIndex){
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeEmail_(data[i][emailColIndex]) === email) return i + 1; // 1-based
  }
  return -1;
}

function signApprovalToken_(email){
  let secret = PropertiesService.getScriptProperties().getProperty("APPROVAL_SECRET");
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty("APPROVAL_SECRET", secret);
  }
  const payload = `${email}|${Math.floor(Date.now() / 1000)}`;
  const sig = Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret));
  return { payload, sig };
}

function verifyApprovalToken_(payload, sig){
  const secret = PropertiesService.getScriptProperties().getProperty("APPROVAL_SECRET");
  if (!secret) return false;

  const expected = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(payload, secret)
  );
  if (expected !== sig) return false;

  const parts = String(payload).split("|");
  if (parts.length !== 2) return false;

  const ts = Number(parts[1]);
  if (!ts) return false;

  const age = Math.floor(Date.now() / 1000) - ts;
  return age >= 0 && age <= APPROVAL_TOKEN_TTL_SECONDS;
}


function getUsersData() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = "users_data_v2";

  try {
    const cached = cache.get(CACHE_KEY);
    if (cached) return JSON.parse(cached);

    const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Users");
    if (!sh) return [];

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];

    const data = sh.getRange(2, 1, lastRow - 1, 7).getValues(); // 7 columns now
    const tz = Session.getScriptTimeZone();

    const out = data.map((row, i) => ({
      index:     i,
      username:  row[0] || "",
      email:     row[1] || "",
      created:   row[2] instanceof Date
                   ? Utilities.formatDate(row[2], tz, "yyyy-MM-dd")
                   : (row[2] || ""),
      status:    row[3] || "Inactive",
      role:      row[4] || "Staff",
      group:     row[5] || "Morning",
      avatarUrl: row[6] || ""
    }));

    out.sort((a, b) => String(a.username || "").localeCompare(String(b.username || "")));

    try { cache.put(CACHE_KEY, JSON.stringify(out), 300); } catch (e) {}

    return out;

  } catch (err) {
    console.error("getUsersData error:", err && err.message ? err.message : err);
    return [];
  }
}

function invalidateUsersCache_() {
  try {
    CacheService.getScriptCache().remove("users_data_v2");
  } catch (err) {
    console.error("invalidateUsersCache_ error:", err && err.message ? err.message : err);
  }
}


function getUsersPageBootstrap() {
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

  return {
    me:    { email },
    perms: getMyPermissions(),
    roles: getDistinctRoles_(),  
    users: getUsersData()
  };
}


/* 🔹 Convert date to readable text */
function formatDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "MM/dd/yyyy");
}



/* ========= UPDATE USER IN SHEET ========= */
function updateUserInSheet(index, name, email, status, role, group) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  const row = index + 2;

  sh.getRange(row, 1).setValue(name);    // A Username
  sh.getRange(row, 2).setValue(email);   // B Email
  sh.getRange(row, 4).setValue(status || "Active");  // D Status
  sh.getRange(row, 5).setValue(role   || "Staff");   // E Role
  sh.getRange(row, 6).setValue(group  || "Morning"); // F Group

  invalidateUsersCache_();
  return true;
}


/* ========= DELETE USER FROM SHEET ========= */
function deleteUserFromSheet(index, actorEmail){
  assertCanManageUsers_(actorEmail);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  if (!sh) throw new Error('Sheet "Users" not found.');

  const sheetRow = Number(index) + 2;
  if (!Number.isFinite(sheetRow) || sheetRow < 2 || sheetRow > sh.getLastRow()){
    throw new Error("Invalid row.");
  }

  sh.deleteRow(sheetRow);
  return true;
}


function resetUserAccess(email) {
  email = String(email || "").trim().toLowerCase();

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowEmail = String(data[i][1] || "").trim().toLowerCase(); // B
    if (rowEmail === email) {
      sheet.getRange(i + 1, 6).setValue("Active"); // F Status
      sheet.getRange(i + 1, 8).setValue(0);        // H FailCount
      return true;
    }
  }
  return false;
}

function getLoginUrl() {
  const base = ScriptApp.getService().getUrl();
  return base + "?page=login";
}


/**
 * Server-side function to get names from the Users sheet
 */
function getUserNames() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];

  const data  = sheet.getDataRange().getValues();
  const names = data.slice(1).map(row => row[0]).filter(String);

  return names;
}


// Helper to map labels back to CSS classes
function getCategoryTag(label) {
  const tags = {
    'Meeting': 'tag-blue',
    'Holiday': 'tag-purple',
    'Inventory': 'tag-orange',
    'Other': 'tag-green'
  };
  return tags[label] || 'tag-green';
}


/**
 * Fetches the last 20 audit logs from the 'AuditLogs' sheet.
 */
function getAuditLogs() {
  // ── Cache check ──────────────────────────────────
  const cache    = CacheService.getScriptCache();
  const cacheKey = "audit_logs_v1";
  const cached   = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }
  // ────────────────────────────────────────────────

  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("AuditLogs");
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    values.shift(); // remove header row

    const tz = Session.getScriptTimeZone();

    const toDate = (v) => {
      if (v instanceof Date) return v;
      if (typeof v === "string") {
        const d = new Date(v);
        return isNaN(d.getTime()) ? null : d;
      }
      return null;
    };

    values.sort((a, b) => {
      const da = toDate(a[0]);
      const db = toDate(b[0]);
      return ((db && db.getTime()) || 0) - ((da && da.getTime()) || 0);
    });

    const result = values.slice(0, 20).map(row => {
      const d = toDate(row[0]);
      return {
        timestamp: d ? Utilities.formatDate(d, tz, "M/d/yyyy H:mm:ss") : "",
        user:      String(row[1] || ""),
        action:    String(row[2] || ""),
        resource:  String(row[3] || ""),
        status:    String(row[4] || "")
      };
    });

    // Cache for 60 seconds
    try { cache.put(cacheKey, JSON.stringify(result), 60); } catch (_) {}

    return result;

  } catch (e) {
    console.error("getAuditLogs error:", e);
    return [];
  }
}



function flushAuditLogs_() {
  const cache = CacheService.getScriptCache();
  const key   = "audit_buffer";

  const existing = cache.get(key);
  if (!existing) return;

  let rows;
  try {
    rows = JSON.parse(existing);
    if (!Array.isArray(rows) || !rows.length) return;
  } catch (e) {
    cache.remove(key);
    return;
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh   = ss.getSheetByName("AuditLogs");
    if (!sh) {
      sh = ss.insertSheet("AuditLogs");
      sh.getRange(1, 1, 1, 5)
        .setValues([["Timestamp", "User", "Action", "Resource", "Status"]])
        .setFontWeight("bold")
        .setBackground("#f3f3f3");
    }

    const normalized = rows.map(r => {
      const v = r[0];
      const d =
        v instanceof Date    ? v          :
        typeof v === "number" ? new Date(v) :
        typeof v === "string" ? new Date(v) :
        new Date();
      return [d, r[1], r[2], r[3], r[4]];
    });

    const startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, normalized.length, 5).setValues(normalized);
    sh.getRange(startRow, 1, normalized.length, 1).setNumberFormat("m/d/yyyy h:mm:ss");

    cache.remove(key);

  } catch (e) {
    console.error("flushAuditLogs_ error:", e);
  } finally {
    lock.releaseLock();
  }
}


/**
 * Fetches real-time counts for the Dashboard widgets.
 */
function getDashboardStats() {
  // ── Cache check ──────────────────────────────────
  const cache    = CacheService.getScriptCache();
  const cacheKey = "dashboard_stats_v1";
  const cached   = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }
  // ────────────────────────────────────────────────

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    const master   = ss.getSheetByName("Masterlist");
    const usersS   = ss.getSheetByName("Users");
    const reportsS = ss.getSheetByName("AuditLogs");

    let assignedLaptopCount  = 0;
    let inStockLaptopCount   = 0;
    let warrantyExpiringSoon = 0;
    let totalAssets          = 0;
    let assignedByTeam       = {};
    let inStockByModel       = {};
    let afPending            = 0;
    let afSigned             = 0;

    const WARRANTY_WARN_MS = 30 * 24 * 60 * 60 * 1000; // 30 days in ms
    const now              = new Date();

    if (master && master.getLastRow() > 1) {
      // Use getValues() (not getDisplayValues()) for the warranty date column
      // so we get real Date objects instead of formatted strings
      const displayValues = master.getDataRange().getDisplayValues();
      const rawValues     = master.getDataRange().getValues();

      const headers = displayValues[0].map(h => String(h || "").trim().toLowerCase());

      const col = (name) => headers.indexOf(name);

      const cType         = col("type");
      const cModel        = col("model");
      const cStatus       = col("status");
      const cDepartment   = col("department");
      const cAfStatus     = col("af status");
      const cSignedAf     = col("signed af");
      const cWarrantyEnd  = col("warranty end date") >= 0
                              ? col("warranty end date")
                              : col("warranty end");   // fallback

      displayValues.slice(1).forEach((row, i) => {
        const type   = String(row[cType]   || "").trim().toLowerCase();
        const status = String(row[cStatus] || "").trim().toLowerCase();

        // Skip completely empty rows
        if (!type && !status) return;

        totalAssets++;

        if (type === "laptop") {
          if (status === "assigned") {
            assignedLaptopCount++;
            const dept = String(row[cDepartment] || "").trim() || "No Department";
            assignedByTeam[dept] = (assignedByTeam[dept] || 0) + 1;
          }

          if (["in-stock", "in stock", "instock", "available"].includes(status)) {
            inStockLaptopCount++;
            const model = String(row[cModel] || "").trim() || "Unknown Model";
            inStockByModel[model] = (inStockByModel[model] || 0) + 1;
          }

          // Warranty expiry check — use raw values for reliable Date parsing
          if (cWarrantyEnd >= 0) {
            const rawDate = rawValues[i + 1][cWarrantyEnd]; // +1 to skip header
            if (rawDate) {
              const expiry = rawDate instanceof Date ? rawDate : new Date(rawDate);
              if (!isNaN(expiry.getTime())) {
                const diff = expiry.getTime() - now.getTime();
                if (diff > 0 && diff <= WARRANTY_WARN_MS) warrantyExpiringSoon++;
              }
            }
          }
        }

        if (cAfStatus >= 0) {
          const af = String(row[cAfStatus] || "").trim().toLowerCase();
          if      (af === "pending") afPending++;
          else if (af === "signed")  afSigned++;
        } else if (cSignedAf >= 0) {
          const url = String(row[cSignedAf] || "").trim();
          if (url) afSigned++;
          else     afPending++;
        }
      });
    }

    const result = {
      assignedLaptopCount,
      inStockLaptopCount,
      warrantyExpiringSoon,
      totalAssets,
      reports: reportsS ? Math.max(0, reportsS.getLastRow() - 1) : 0,
      users:   usersS   ? Math.max(0, usersS.getLastRow()   - 1) : 0,
      afPending,
      afSigned,
      assignedByTeam,
      inStockByModel,
    };

    // Cache for 3 minutes
    try { cache.put(cacheKey, JSON.stringify(result), 180); } catch (_) {}

    return result;

  } catch (e) {
    console.error("getDashboardStats error:", e);
    return {
      assignedLaptopCount:  0,
      inStockLaptopCount:   0,
      warrantyExpiringSoon: 0,
      totalAssets:          0,
      reports:              0,
      users:                0,
      afPending:            0,
      afSigned:             0,
      assignedByTeam:       {},
      inStockByModel:       {},
    };
  }
}


/**
 * Generates a PDF from a Sheet and returns the Drive URL.
 */
function createReportPDF(type) {
 const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetName = (type === 'inventory') ? "Inventory" : "AuditLogs";
  const sheet     = ss.getSheetByName(sheetName);

  if (!sheet) throw new Error("Sheet not found");

  const data = sheet.getDataRange().getValues();
  let rowsHtml = "";

  data.forEach((row, index) => {
    const tag = (index === 0) ? "th" : "td";
    rowsHtml += `<tr>${row.map(cell => `<${tag} style="border:1px solid #ddd;padding:8px;">${cell}</${tag}>`).join("")}</tr>`;
  });

  const html = `
    <html>
      <body style="font-family:sans-serif;">
        <h1 style="text-align:center;">Manila IT - ${type.toUpperCase()} Report</h1>
        <p style="text-align:center;">Generated on: ${new Date().toLocaleString()}</p>
        <table style="width:100%;border-collapse:collapse;">${rowsHtml}</table>
      </body>
    </html>`;

  const blob = HtmlService.createHtmlOutput(html).getAs("application/pdf");
  blob.setName(`ManilaIT_${type}_Report.pdf`);

  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}


function toggleUserStatusInSheet(index, newStatus, actorEmail) {
  assertCanManageUsers_(actorEmail);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  if (!sh) throw new Error('Sheet "Users" not found.');

  const sheetRow = Number(index) + 2;
  if (!Number.isFinite(sheetRow) || sheetRow < 2 || sheetRow > sh.getLastRow()) {
    throw new Error("Invalid row.");
  }

  const headers   = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const statusCol = headers.indexOf("Status");
  if (statusCol === -1) throw new Error('Users sheet missing "Status" column.');

  sh.getRange(sheetRow, statusCol + 1).setValue(newStatus);
  invalidateUsersCache_();
  return true;
}


function resetUserPassword(index, actorEmail) {
  assertCanManageUsers_(actorEmail);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  if (!sh) throw new Error("Users sheet not found");

  const sheetRow = Number(index) + 2;
  if (!Number.isFinite(sheetRow) || sheetRow < 2 || sheetRow > sh.getLastRow()){
    throw new Error("Invalid row.");
  }

  const EMAIL_COL       = 2;  // B
  const HASH_COL        = 3;  // C
  const SALT_COL        = 4;  // D
  const STATUS_COL      = 6;  // F
  const FAIL_COL        = 8;  // H
  const MUST_CHANGE_COL = 11; // K

  const email = String(sh.getRange(sheetRow, EMAIL_COL).getValue() || "").trim();
  if (!email) throw new Error("User email is empty");

  const salt = String(sh.getRange(sheetRow, SALT_COL).getValue() || "").trim();
  if (!salt) throw new Error("Missing salt for this user");

  const tempPassword = generateTempPassword_(10);
  const newHash      = hashPassword(tempPassword, salt);

  sh.getRange(sheetRow, HASH_COL).setValue(newHash);
  sh.getRange(sheetRow, FAIL_COL).setValue(0);
  sh.getRange(sheetRow, MUST_CHANGE_COL).setValue(true);

  const currentStatus = String(sh.getRange(sheetRow, STATUS_COL).getValue() || "").trim();
  if (currentStatus === "Disabled") sh.getRange(sheetRow, STATUS_COL).setValue("Active");

  let emailed = false;
  try {
    const subject   = "Your temporary password";
    const plainBody = `Your password has been reset.\n\nTemporary password: ${tempPassword}\n\nPlease change it after logging in.`;

    const htmlBody = `
<div style="margin:0;padding:0;background:#f5f5f7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
  <div style="max-width:560px;margin:32px auto;padding:0 16px;">
    <div style="background:#ffffff;border:1px solid #e5e5ea;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.06);">
      <div style="padding:18px 20px;background:linear-gradient(135deg,#0A84FF,#2997FF);color:#fff;">
        <div style="font-size:14px;opacity:.9;">Manila IT Inventory</div>
        <div style="font-size:18px;font-weight:700;margin-top:4px;">Password Reset</div>
      </div>
      <div style="padding:20px;color:#1c1c1e;">
        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;">Your password has been reset. Use the temporary password below to sign in.</p>
        <div style="margin:16px 0;padding:14px 16px;background:#f2f2f7;border:1px solid #e5e5ea;border-radius:12px;">
          <div style="font-size:12px;color:#6e6e73;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;">Temporary Password</div>
          <div style="font-size:20px;font-weight:800;letter-spacing:.06em;font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Courier New',monospace;">${tempPassword}</div>
        </div>
        <p style="margin:0;font-size:13px;line-height:1.5;color:#3a3a3c;">After logging in, please change your password immediately.</p>
        <div style="margin-top:16px;padding-top:14px;border-top:1px solid #e5e5ea;">
          <p style="margin:0;font-size:12px;line-height:1.4;color:#6e6e73;">If you did not request this reset, please contact your administrator.</p>
        </div>
      </div>
    </div>
    <div style="text-align:center;margin-top:14px;font-size:11px;color:#8e8e93;">© ${new Date().getFullYear()} Manila IT Inventory</div>
  </div>
</div>`;

    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, name: "Manila IT Inventory" });
    emailed = true;
  } catch (e) {
    emailed = false;
  }

  return { tempPassword, emailed };
}


function generateTempPassword_(len) {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789";
  let out = "";
  for (let i = 0; i < len; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}


function getEmailFromSession(sessionID){
  sessionID = String(sessionID || "").trim();
  if (!sessionID) return "";

  const raw = CacheService.getScriptCache().get(`sess_${sessionID}`);
  if (!raw) return "";

  try {
    const obj = JSON.parse(raw);
    return String(obj.email || "").trim().toLowerCase();
  } catch (e) {
    return "";
  }
}


function getCurrentUserProfile() {
  const email  = Session.getActiveUser().getEmail() || "";
  const domain = email.includes("@") ? email.split("@")[1] : "";

  const nameGuess = email
    ? email.split("@")[0]
        .replace(/[._-]+/g, " ")
        .replace(/\b\w/g, (c) => c.toUpperCase())
    : "Unknown User";

  return { email, name: nameGuess, domain };
}

function getCurrentUser() {
  const email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!email) return { email: "", username: "" };

  const users = getUsersData();
  const me = users.find(u => String(u.email || "").trim().toLowerCase() === email);

  return me || { email, username: "" };
}


function getUserByEmail(email) {
  email = String(email || "").trim().toLowerCase();
  if (!email) return { email: "", role: "viewer", username: "", avatarUrl: "" };

  const ss = (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Users");
  if (!sh) return { email, role: "viewer", username: "", avatarUrl: "" };

  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return { email, role: "viewer", username: "", avatarUrl: "" };

  const header = (vals[0] || []).map(h => String(h || "").trim().toLowerCase());

  const idxUsername = header.findIndex(h => /^(username|name)$/i.test(h));
  const idxEmail    = header.findIndex(h => /^email$/i.test(h));
  const idxRole     = header.findIndex(h => /^(role|roles?)$/i.test(h));
  const idxAvatar   = header.findIndex(h => /^(avatarurl|avatar|avatar_url)$/i.test(h));

  const defaultIdx = {
    username: idxUsername >= 0 ? idxUsername : 0,
    email:    idxEmail    >= 0 ? idxEmail    : 1,
    role:     idxRole     >= 0 ? idxRole     : 6,
    avatar:   idxAvatar   >= 0 ? idxAvatar   : 11
  };

  for (let r = 1; r < vals.length; r++) {
    const row      = vals[r];
    const rowEmail = String(row[defaultIdx.email] || "").trim().toLowerCase();
    if (rowEmail !== email) continue;

    const username = String(row[defaultIdx.username] || "").trim();
    const role     = String(row[defaultIdx.role]     || "viewer").trim();
    let avatarUrl  = String(row[defaultIdx.avatar]   || "").trim();

    if (avatarUrl) {
      avatarUrl = avatarUrl.replace(/\r?\n|\r/g, "").trim();
      if (!/^data:|^https?:\/\//i.test(avatarUrl)) {
        const maybeB64 = avatarUrl.replace(/\s+/g, "");
        if (/^[A-Za-z0-9+/=]+$/.test(maybeB64) && maybeB64.length > 40) {
          avatarUrl = "data:image/jpeg;base64," + maybeB64;
        } else {
          avatarUrl = "";
        }
      }
    }

    return {
      username:  username || email.split("@")[0],
      email:     rowEmail,
      role:      role || "viewer",
      avatarUrl: avatarUrl || ""
    };
  }

  return { email, role: "viewer", username: "", avatarUrl: "" };
}


/**
 * FIX #3: resetMyPassword previously passed the target email as the
 * actorEmail argument to resetUserPassword(), which caused the admin
 * RBAC check inside that function to reject the call for non-admin users.
 * Self-service resets bypass the admin check via a dedicated helper.
 */
function resetMyPassword(email) {
  email = String(email || "").trim().toLowerCase();
  if (!email) throw new Error("Missing email.");

  const users = getUsersData();
  const me    = users.find(u => String(u.email || "").trim().toLowerCase() === email);
  if (!me) throw new Error("User not found in Users sheet.");

  // Perform the reset directly without going through the admin RBAC check.
  return _resetPasswordByIndex_(me.index, email);
}

/**
 * Internal helper: resets password for a given row index and sends email.
 * Used by both resetUserPassword (admin path) and resetMyPassword (self-service).
 */
function _resetPasswordByIndex_(index, targetEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  if (!sh) throw new Error("Users sheet not found");

  const sheetRow = Number(index) + 2;
  if (!Number.isFinite(sheetRow) || sheetRow < 2 || sheetRow > sh.getLastRow()){
    throw new Error("Invalid row.");
  }

  const EMAIL_COL       = 2;
  const HASH_COL        = 3;
  const SALT_COL        = 4;
  const STATUS_COL      = 6;
  const FAIL_COL        = 8;
  const MUST_CHANGE_COL = 11;

  const email = targetEmail || String(sh.getRange(sheetRow, EMAIL_COL).getValue() || "").trim();
  if (!email) throw new Error("User email is empty");

  const salt = String(sh.getRange(sheetRow, SALT_COL).getValue() || "").trim();
  if (!salt)  throw new Error("Missing salt for this user");

  const tempPassword = generateTempPassword_(10);
  const newHash      = hashPassword(tempPassword, salt);

  sh.getRange(sheetRow, HASH_COL).setValue(newHash);
  sh.getRange(sheetRow, FAIL_COL).setValue(0);
  sh.getRange(sheetRow, MUST_CHANGE_COL).setValue(true);

  const currentStatus = String(sh.getRange(sheetRow, STATUS_COL).getValue() || "").trim();
  if (currentStatus === "Disabled") sh.getRange(sheetRow, STATUS_COL).setValue("Active");

  let emailed = false;
  try {
    const subject   = "Your temporary password";
    const plainBody = `Your password has been reset.\n\nTemporary password: ${tempPassword}\n\nPlease change it after logging in.`;

    const htmlBody = `
<div style="margin:0;padding:0;background:#f5f5f7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
  <div style="max-width:560px;margin:32px auto;padding:0 16px;">
    <div style="background:#ffffff;border:1px solid #e5e5ea;border-radius:16px;overflow:hidden;box-shadow:0 8px 24px rgba(0,0,0,.06);">
      <div style="padding:18px 20px;background:linear-gradient(135deg,#0A84FF,#2997FF);color:#fff;">
        <div style="font-size:14px;opacity:.9;">Manila IT Inventory</div>
        <div style="font-size:18px;font-weight:700;margin-top:4px;">Password Reset</div>
      </div>
      <div style="padding:20px;color:#1c1c1e;">
        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;">Your password has been reset. Use the temporary password below to sign in.</p>
        <div style="margin:16px 0;padding:14px 16px;background:#f2f2f7;border:1px solid #e5e5ea;border-radius:12px;">
          <div style="font-size:12px;color:#6e6e73;text-transform:uppercase;letter-spacing:.08em;margin-bottom:6px;">Temporary Password</div>
          <div style="font-size:20px;font-weight:800;letter-spacing:.06em;font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Courier New',monospace;">${tempPassword}</div>
        </div>
        <p style="margin:0;font-size:13px;line-height:1.5;color:#3a3a3c;">After logging in, please change your password immediately.</p>
        <div style="margin-top:16px;padding-top:14px;border-top:1px solid #e5e5ea;">
          <p style="margin:0;font-size:12px;line-height:1.4;color:#6e6e73;">If you did not request this reset, please contact your administrator.</p>
        </div>
      </div>
    </div>
    <div style="text-align:center;margin-top:14px;font-size:11px;color:#8e8e93;">© ${new Date().getFullYear()} Manila IT Inventory</div>
  </div>
</div>`;

    GmailApp.sendEmail(email, subject, plainBody, { htmlBody, name: "Manila IT Inventory" });
    emailed = true;
  } catch (e) {
    emailed = false;
  }

  return { tempPassword, emailed };
}


function changeMyPassword(email, oldPass, newPass){
  email   = String(email   || "").trim().toLowerCase();
  oldPass = String(oldPass || "");
  newPass = String(newPass || "");

  if (!email)          throw new Error("Missing email");
  if (!oldPass)        throw new Error("Missing current password");
  if (newPass.length < 8) throw new Error("New password too short (min 8)");
  if (oldPass === newPass) throw new Error("New password must be different");

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  if (!sh) throw new Error("Users sheet not found");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error("No users");

  const cache    = CacheService.getScriptCache();
  const cacheKey = `urow_${email}`;
  let row = Number(cache.get(cacheKey) || 0);

  if (row >= 2 && row <= lastRow) {
    const cachedEmail = String(sh.getRange(row, 2).getValue() || "").trim().toLowerCase();
    if (cachedEmail !== email) row = 0;
  }

  if (!row) {
    const emailCol = sh.getRange(2, 2, lastRow - 1, 1);
    const cell     = emailCol.createTextFinder(email).matchEntireCell(true).findNext();
    if (!cell) throw new Error("Account not found");
    row = cell.getRow();
    cache.put(cacheKey, String(row), 30 * 60);
  }

  const HASH_COL        = 3;
  const SALT_COL        = 4;
  const STATUS_COL      = 6;
  const FAIL_COL        = 8;
  const MUST_CHANGE_COL = 11;

  const lock = LockService.getDocumentLock();
  lock.waitLock(5000);

  try {
    const emailCell = String(sh.getRange(row, 2).getValue() || "").trim().toLowerCase();
    if (emailCell !== email) throw new Error("Account not found");

    const rng  = sh.getRange(row, HASH_COL, 1, MUST_CHANGE_COL - HASH_COL + 1);
    const vals = rng.getValues()[0];

    const passHash = String(vals[0] || "").trim();
    const salt     = String(vals[1] || "").trim();
    const status   = String(vals[STATUS_COL - HASH_COL] || "").trim();

    if (status !== "Active")    throw new Error("Account is not active");
    if (!salt || !passHash)     throw new Error("Account not configured");

    const oldHash = hashPassword(oldPass, salt);
    if (oldHash !== passHash)   throw new Error("Current password is incorrect");

    const newSalt = Utilities.getUuid().replace(/-/g, "");
    const newHash = hashPassword(newPass, newSalt);

    vals[0] = newHash;
    vals[1] = newSalt;
    vals[FAIL_COL        - HASH_COL] = 0;
    vals[MUST_CHANGE_COL - HASH_COL] = false;

    rng.setValues([vals]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try { logActivity(email, "Change Password", "Users", "Success"); } catch(e){}

  return "ok";
}


function setMyAvatarDataUrl(email, dataUrl) {
  email = String(email || "").trim().toLowerCase();
  if (!email)    return "missing_email";
  if (!dataUrl)  return "missing_image";

  const s = String(dataUrl).trim();

  const m = s.match(/^data:image\/(png|jpe?g|webp|gif)(;[^,]+)?,(base64,)?(.+)$/i);
  if (!m) return "invalid_format";

  const commaIndex = s.indexOf(",");
  if (commaIndex < 0) return "invalid_format";
  const b64 = s.slice(commaIndex + 1);
  if (!b64) return "invalid_format";

  const paddedLen = b64.length;
  const padding   = (b64.endsWith("==") ? 2 : b64.endsWith("=") ? 1 : 0);
  const byteSize  = Math.floor((paddedLen * 3) / 4) - padding;

  const MAX_BYTES = 50 * 1024;
  if (byteSize > MAX_BYTES) return "too_large";

  let sessionEmail = "";
  try {
    const sess = Session.getActiveUser();
    if (sess && sess.getEmail) sessionEmail = String(sess.getEmail() || "").trim().toLowerCase();
  } catch (e) {
    sessionEmail = "";
  }

  if (sessionEmail && sessionEmail !== email) return "session_mismatch";

  const ss = (typeof SHEET_ID !== "undefined" && SHEET_ID)
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  const sh = ss.getSheetByName("Users");
  if (!sh) return "Users sheet not found";

  const values = sh.getDataRange().getValues();
  if (!values || values.length === 0) return "Users sheet empty";

  const header    = values[0].map(h => String(h || "").trim());
  const emailCol  = header.findIndex(h => /^email$/i.test(h));
  const avatarCol = header.findIndex(h => /^avatarurl$/i.test(h));

  if (emailCol  === -1) return "Email column not found";
  if (avatarCol === -1) return "AvatarUrl column not found";

  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][emailCol] || "").trim().toLowerCase();
    if (rowEmail === email) { rowIndex = i; break; }
  }
  if (rowIndex < 0) return "User not found";

  sh.getRange(rowIndex + 1, avatarCol + 1).setValue(s);
  return "ok";
}

function toDateOrEmpty_(v){
  if (!v) return "";
  const d = new Date(v);
  return isNaN(d.getTime()) ? "" : d;
}


function clearAuditLogs() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sh = ss.getSheetByName("AuditLogs");

    if (!sh) return { ok: true, deletedRows: 0 };

    try { flushAuditLogs_(); } catch (e) {}

    var lastRow = sh.getLastRow();
    var lastCol = Math.max(sh.getLastColumn(), 5);

    var deletedRows = 0;
    if (lastRow > 1) {
      deletedRows = lastRow - 1;
      sh.getRange(2, 1, deletedRows, lastCol).clearContent();
    }

    try { CacheService.getScriptCache().remove("audit_buffer"); } catch (e) {}
    try { CacheService.getScriptCache().remove("audit_logs_v1"); } catch (e) {}  // ← invalidate audit log cache

    sh.getRange(1, 1, 1, 5)
      .setValues([["Timestamp", "User", "Action", "Resource", "Status"]])
      .setFontWeight("bold")
      .setBackground("#f3f3f3");

    return { ok: true, deletedRows: deletedRows };

  } catch (e) {
    console.error("clearAuditLogs error:", e);
    throw new Error("clearAuditLogs failed: " + e);
  }
}


function formatDate_(val) {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    const mm = String(val.getMonth() + 1).padStart(2, "0");
    const dd = String(val.getDate()).padStart(2, "0");
    const yyyy = val.getFullYear();
    return `${yyyy}-${mm}-${dd}`;
  }
  return String(val).trim();
}


function normalizeRole_(role){
  return String(role || "").trim().toLowerCase();
}


function assertBulkAdmin_(actorEmail) {
  var email = String(actorEmail || "").trim().toLowerCase();
  if (!email) throw new Error("Missing actor email.");

  var admins   = new Set(["admin", "super admin"]);
  var allUsers = getUsersData();
  var me       = (allUsers || []).find(function(u) {
    return String(u.email || "").toLowerCase() === email;
  });
  var role = String((me && me.role) || "").trim().toLowerCase();

  if (!admins.has(role)) {
    throw new Error("Bulk delete is restricted to Admins only.");
  }
}


function getRecentActivity(limit) {
  limit = limit || 6;
 
  const ACTION_COLOR = {
    'assign'   : '#4f8ef7',
    'unassign' : '#f87171',
    'clear'    : '#f87171',
    'edit'     : '#fbbf24',
    'update'   : '#fbbf24',
    'add'      : '#34d399',
    'create'   : '#34d399',
    'delete'   : '#f87171',
    'retire'   : '#a78bfa',
    'sign'     : '#22d3ee',
    'generate' : '#22d3ee',
    'login'    : '#94a3b8',
    'logout'   : '#94a3b8',
    'change password' : '#a78bfa',
  };
 
  try {
    // Reuse your existing getAuditLogs() — no need to duplicate sheet logic
    const logs = getAuditLogs();  // returns last 20, newest first
 
    const result = logs.slice(0, limit).map(entry => {
      const action   = String(entry.action   || '').trim();
      const user     = String(entry.user     || 'IT').trim();
      const resource = String(entry.resource || '').trim();
 
      // Pick colour by matching action keywords
      let color = '#94a3b8';
      for (const [key, val] of Object.entries(ACTION_COLOR)) {
        if (action.toLowerCase().includes(key)) { color = val; break; }
      }
 
      // Build readable summary
      let text = '';
      if (action && resource) {
        text = `${_capFirst_(action)} — ${resource}`;
      } else if (action && user) {
        text = `${user}: ${_capFirst_(action)}`;
      } else {
        text = action || 'Activity recorded';
      }
 
      return {
        color,
        text,
        time: entry.timestamp ? _relativeTime_(entry.timestamp) : '',
        user,
      };
    });
 
    return { ok: true, data: result };
 
  } catch (e) {
    console.error('getRecentActivity error:', e);
    return { ok: false, error: e.message, data: [] };
  }
}
 
 
 
// ============================================================
//  PRIVATE HELPERS (add alongside the new functions above)
// ============================================================
 
function _capFirst_(str) {
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : '';
}
 
function _relativeTime_(raw) {
  if (!raw) return '';
  const ts = new Date(raw);
  if (isNaN(ts.getTime())) return String(raw);
 
  const diffMs  = Date.now() - ts.getTime();
  const diffMin = Math.floor(diffMs / 60000);
  const diffHr  = Math.floor(diffMin / 60);
  const diffDay = Math.floor(diffHr  / 24);
 
  if (diffMin < 2)   return 'Just now';
  if (diffMin < 60)  return `${diffMin}m ago`;
  if (diffHr  < 24)  return `${diffHr}h ago`;
  if (diffDay === 1) return 'Yesterday';
  if (diffDay < 7)   return `${diffDay}d ago`;
 
  return ts.toLocaleDateString('en-PH', { month: 'short', day: 'numeric' });
}
 

 function getMyAssignedTasks(username) {
  username = String(username || "").trim();
  if (!username) return [];

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("Tasks");
    if (!sh || sh.getLastRow() < 2) return [];

    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || "").trim().toLowerCase());

    const idxAssignee      = headers.indexOf("assignee");
    const idxTask          = headers.indexOf("task");
    const idxStatus        = headers.indexOf("status");
    const idxDateRequired  = headers.indexOf("date required");
    const idxDateSubmitted = headers.indexOf("date submitted");
    const idxNote          = headers.indexOf("note");

    if (idxAssignee === -1 || idxTask === -1) return [];

    const tz = Session.getScriptTimeZone();

    function fmtDate(v) {
      if (!v) return "";
      const d = v instanceof Date ? v : new Date(v);
      if (isNaN(d.getTime())) return String(v);
      return Utilities.formatDate(d, tz, "MM/dd/yyyy");
    }

    return values.slice(1)
      .filter(row => {
        const assignee = String(row[idxAssignee] || "").trim().toLowerCase();
        return assignee === username.toLowerCase();
      })
      .map(row => ({
        taskDescription: idxTask          >= 0 ? String(row[idxTask]          || "") : "",
        status:          idxStatus        >= 0 ? String(row[idxStatus]        || "") : "",
        dateRequired:    idxDateRequired  >= 0 ? fmtDate(row[idxDateRequired])       : "",
        dateSubmitted:   idxDateSubmitted >= 0 ? fmtDate(row[idxDateSubmitted])      : "",
        note:            idxNote          >= 0 ? String(row[idxNote]          || "") : "",
      }));

  } catch (e) {
    console.error("getMyAssignedTasks error:", e);
    return [];
  }
}
 