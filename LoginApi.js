const CFG = {
  SHEETS: {
    USERS: "Users",
    PENDING: "PendingUsers",
  },

  USERS_COLS: 11, // A:K

  OTP: {
    TTL_SECONDS:     5  * 60,
    RATE_SECONDS:    30,
    REG_TTL_SECONDS: 10 * 60,
  },

  LOGIN: {
    MAX_FAILS: 5,
  },

  EMAIL: {
    FROM_ALIAS:    "manila-it@fbgphilippines.com",
    FROM_NAME:     "FBG Manila IT",
    ADMIN_EMAIL:   "michael.beleno.xc@betfanatics.com",
    QUEUE_PROP:    "EMAIL_QUEUE",
    SENT_PROP:     "EMAIL_SENT_MAP",
    TRIGGER_PROP:  "EMAIL_QUEUE_TRIGGER",
    MAX_QUEUE_SIZE: 500,
    MAX_PER_RUN:    10,
    SENT_TTL_MS:   7 * 24 * 60 * 60 * 1000,
  },

  CACHE_KEYS: {
    AUDIT_BUFFER: "audit_buffer",
  },

  CACHE_TTL: {
    SESSION_SECONDS:      60 * 60,
    USER_ROW_SECONDS:     30 * 60,
    AUDIT_BUFFER_SECONDS: 6  * 60 * 60,
  },

  USER_COL: {
    USERNAME:    1,  // A
    EMAIL:       2,  // B
    HASH:        3,  // C
    SALT:        4,  // D
    CREATED:     5,  // E
    STATUS:      6,  // F
    ROLE:        7,  // G
    FAILS:       8,  // H
    POSITION:    9,  // I
    GROUP:       10, // J
    MUST_CHANGE: 11, // K
  },

  PENDING_COL: {
    EMAIL:  1, // A
    CODE:   2, // B
    STATUS: 3, // C
    DATE:   4, // D
  },
};

/* =========================================================
 * APP URL / SPREADSHEET
 * ======================================================= */
function getAppUrl() {
  // Hardcode the production exec URL to avoid returning wrong deployment
  return 'https://script.google.com/a/macros/betfanatics.com/s/AKfycbwzoyd7lWmwp6o1B7pB84INF733pWLWiV7DsUMkK_HmhCj1Ydx4U2GcsZzmUR28Kgk/exec';
}

function getSS_() {
  if (typeof SHEET_ID === "undefined" || !SHEET_ID) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
  return SpreadsheetApp.openById(SHEET_ID);
}


function getPendingSheet_() {
  const ss = getSS_();
  let sh = ss.getSheetByName(CFG.SHEETS.PENDING);
  if (!sh) sh = ss.insertSheet(CFG.SHEETS.PENDING);
  ensurePendingHeaders_(sh);
  return sh;
}

function ensurePendingHeaders_(sh) {
  if (sh.getLastRow() > 0) return;
  sh.getRange(1, 1, 1, 4).setValues([["Email", "Code", "Status", "Date"]]);
}

/* =========================================================
 * SMALL HELPERS
 * ======================================================= */
function normEmail_(email) {
  return String(email || "").trim().toLowerCase();
}

function trim_(value) {
  return String(value ?? "").trim();
}

function toInt_(value, fallback) {
  const n = parseInt(String(value ?? ""), 10);
  return Number.isFinite(n) ? n : (fallback ?? 0);
}

function safeJsonParse_(raw, fallback) {
  try {
    return JSON.parse(raw);
  } catch (_) {
    return fallback;
  }
}

function boolish_(value) {
  if (value === true) return true;
  const s = String(value ?? "").trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes";
}

function withScriptLock_(timeoutMs, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(timeoutMs || 5000);
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function withDocumentLock_(timeoutMs, fn) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(timeoutMs || 5000);
  try {
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function getScriptCache_() {
  return CacheService.getScriptCache();
}

function getScriptProps_() {
  return PropertiesService.getScriptProperties();
}

/* =========================================================
 * SESSION HELPERS
 * ======================================================= */
function getEmailFromSession(sessionID) {
  const sid = trim_(sessionID);
  if (!sid) return "";

  const raw = getScriptCache_().get(`sess_${sid}`);
  if (!raw) return "";

  const obj = safeJsonParse_(raw, null);
  return normEmail_(obj && obj.email);
}

function verifySession(sessionID) {
  return !!getEmailFromSession(sessionID);
}

function isLogged(email, sessionID) {
  const normalized = normEmail_(email);
  if (!normalized) return false;
  return getEmailFromSession(sessionID) === normalized;
}

function setAuthSession(email, sessionID) {
  const em  = normEmail_(email);
  const sid = trim_(sessionID);
  if (!em || !sid) throw new Error("Missing email/sessionID");

  const cache = getScriptCache_();
  const ttl   = CFG.CACHE_TTL.SESSION_SECONDS;

  cache.put(`sess_${sid}`, JSON.stringify({
    email:     em,
    createdAt: Date.now(),
  }), ttl);

  cache.put(`sessByEmail_${em}`, sid, ttl);
}

function findSessionID(email) {
  const em = normEmail_(email);
  if (!em) return "";
  return getScriptCache_().get(`sessByEmail_${em}`) || "";
}

/* =========================================================
 * AUDIT LOG
 * ======================================================= */
function logActivity(email, action, resource, status) {
  try {
    const normalizedEmail = normEmail_(email || "anonymous");
    const entry = [
      Date.now(),
      normalizedEmail,
      trim_(action),
      trim_(resource),
      trim_(status),
    ];

    withScriptLock_(2500, () => {
      const cache = getScriptCache_();
      const key   = CFG.CACHE_KEYS.AUDIT_BUFFER;

      let rows = safeJsonParse_(cache.get(key), []);
      if (!Array.isArray(rows)) rows = [];

      rows.push(entry);
      if (rows.length > 200) rows = rows.slice(-200);

      cache.put(key, JSON.stringify(rows), CFG.CACHE_TTL.AUDIT_BUFFER_SECONDS);
    });

    try {
      if (typeof flushAuditLogs_ === "function") flushAuditLogs_();
    } catch (_) {}

    return true;
  } catch (err) {
    console.error("logActivity error:", err);
    return false;
  }
}

/* =========================================================
 * USER LOOKUP
 * ======================================================= */
function getUserRowByEmail_(sh, email) {
  const em = normEmail_(email);
  if (!em) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const cache    = getScriptCache_();
  const cacheKey = `urow_${em}`;

  let cachedRow = toInt_(cache.get(cacheKey), 0);
  if (cachedRow >= 2 && cachedRow <= lastRow) {
    const rowEmail = normEmail_(
      sh.getRange(cachedRow, CFG.USER_COL.EMAIL).getDisplayValue()
    );
    if (rowEmail === em) return cachedRow;
  }

  const range = sh.getRange(2, CFG.USER_COL.EMAIL, lastRow - 1, 1);
  const found = range.createTextFinder(em).matchEntireCell(true).findNext();
  if (!found) return 0;

  const row = found.getRow();
  cache.put(cacheKey, String(row), CFG.CACHE_TTL.USER_ROW_SECONDS);
  return row;
}

/* =========================================================
 * PASSWORD HASHING  (v3 — HMAC-SHA256, single round)
 *
 * Version history:
 *   v1 — bare base64, 10 000 × SHA-256 iterations (original)
 *   v2 — "v2:<hash>",  1 000 × SHA-256 iterations (intermediate)
 *   v3 — "v3:<hash>",  HMAC-SHA256, single GAS call  ← current
 *
 * The HMAC secret lives in Script Properties under "HMAC_SECRET".
 * Run setHmacSecret() once to generate and store it.
 * ======================================================= */

/**
 * Read the HMAC secret from Script Properties.
 * Throws clearly if it has not been set yet.
 */
function getHmacSecret_() {
  const secret = getScriptProps_().getProperty("HMAC_SECRET");
  if (!secret) throw new Error(
    "HMAC_SECRET is not set. Run setHmacSecret() once in the Apps Script editor."
  );
  return secret;
}

/**
 * Hash a password with HMAC-SHA256 (v3).
 * Single GAS call — effectively instant.
 *
 * @param  {string} password
 * @param  {string} userSalt  Per-user value (email or UUID from the sheet)
 * @returns {string}          "v3:<base64url>"
 */
function hashPassword(password, userSalt) {
  if (!password) throw new Error("hashPassword: password is required");
  if (!userSalt) throw new Error("hashPassword: userSalt is required");

  const raw = Utilities.computeHmacSha256Signature(
    `${password}|${userSalt}`,
    getHmacSecret_()
  );
  return "v3:" + Utilities.base64EncodeWebSafe(raw);
}

/**
 * Constant-time string comparison — prevents timing side-channel attacks.
 */
function safeEquals_(a, b) {
  if (typeof a !== "string" || typeof b !== "string") return false;
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

/**
 * Verify a password against whatever version of hash is stored in the sheet.
 * Returns { ok: boolean, needsUpgrade: boolean }.
 *
 * needsUpgrade = true means the hash is v1 or v2 and should be silently
 * rewritten to v3 on the next successful login.
 */
function verifyPassword_(password, userSalt, storedHash) {
  if (!password || !userSalt || !storedHash) {
    return { ok: false, needsUpgrade: false };
  }

  // v3 — current HMAC hash
  if (storedHash.startsWith("v3:")) {
    const attempt = hashPassword(password, userSalt);
    return { ok: safeEquals_(attempt, storedHash), needsUpgrade: false };
  }

  // v2 — 1 000-round SHA-256
  if (storedHash.startsWith("v2:")) {
    const attempt = hashPasswordLegacy_(password, userSalt, 1000, "v2");
    return { ok: safeEquals_(attempt, storedHash), needsUpgrade: true };
  }

  // v1 — bare base64, 10 000-round SHA-256 (no prefix)
  const attempt = hashPasswordLegacy_(password, userSalt, 10000, null);
  return { ok: safeEquals_(attempt, storedHash), needsUpgrade: true };
}

/**
 * Silently upgrade a legacy hash to v3 after a successful login.
 * Non-fatal — if the write fails the user is still logged in and
 * the upgrade is retried on the next login.
 *
 * @param {string}   password
 * @param {string}   userSalt
 * @param {string}   storedHash
 * @param {Function} saveFn     Called with the new hash string to persist it
 */
function upgradeHashIfNeeded_(password, userSalt, storedHash, saveFn) {
  if (storedHash.startsWith("v3:")) return; // already current
  try {
    saveFn(hashPassword(password, userSalt));
  } catch (err) {
    console.warn("Hash upgrade failed, will retry next login:", err.message);
  }
}

/**
 * Legacy iterated-SHA-256 hasher — only used during the transition window
 * to verify v1/v2 hashes that still exist in the sheet.
 * Delete this function (and the v1/v2 branches in verifyPassword_) once
 * all users have logged in at least once after the v3 migration.
 */
function hashPasswordLegacy_(password, userSalt, iterations, prefix) {
  let hash = `${password}|${userSalt}`;
  for (let i = 0; i < iterations; i++) {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      hash
    );
    hash = Utilities.base64EncodeWebSafe(digest);
  }
  return prefix ? `${prefix}:${hash}` : hash;
}

/**
 * Generate a random per-user salt.
 * Still used when creating new accounts or changing passwords.
 */
function generateSalt() {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${Math.random()}|${Date.now()}|${Utilities.getUuid()}`
  );
  return Utilities.base64EncodeWebSafe(digest).slice(0, 16);
}

/* =========================================================
 * LOGIN
 * ======================================================= */
function checkLogin(email, pass) {
  const normalizedEmail = normEmail_(email);
  const password        = String(pass || "");

  if (!normalizedEmail || !password) return "missing_fields";

  const sh  = getUsersSheet_();
  const row = getUserRowByEmail_(sh, normalizedEmail);
  if (!row) return "no_account";

  const values = sh.getRange(row, 1, 1, CFG.USERS_COLS).getValues()[0];

  const rowEmail   = normEmail_(values[CFG.USER_COL.EMAIL       - 1]);
  const storedHash = trim_(values[CFG.USER_COL.HASH             - 1]);
  const salt       = trim_(values[CFG.USER_COL.SALT             - 1]);
  const status     = trim_(values[CFG.USER_COL.STATUS           - 1] || "Inactive");
  const failCount  = toInt_(values[CFG.USER_COL.FAILS           - 1], 0);
  const mustChange = boolish_(values[CFG.USER_COL.MUST_CHANGE   - 1]);

  if (rowEmail !== normalizedEmail) return "no_account";
  if (status   === "Disabled")      return "account_locked";
  if (!salt)                        return "server_error_missing_salt";

  // ── Verify password (handles v1, v2, and v3 hashes) ──
  const result = verifyPassword_(password, salt, storedHash);

  if (!result.ok) {
    try {
      return withDocumentLock_(3000, () => {
        const currentFails = toInt_(
          sh.getRange(row, CFG.USER_COL.FAILS).getValue(), 0
        );
        const newFails = currentFails + 1;

        sh.getRange(row, CFG.USER_COL.FAILS).setValue(newFails);

        if (newFails >= CFG.LOGIN.MAX_FAILS) {
          sh.getRange(row, CFG.USER_COL.STATUS).setValue("Disabled");
          return "account_locked";
        }
        return "wrong_password";
      });
    } catch (err) {
      console.error("checkLogin fail update error:", err);
      return "server_busy";
    }
  }

  if (status !== "Active") return "not_active";

  // ── Silently upgrade legacy hash to v3 ──
  if (result.needsUpgrade) {
    upgradeHashIfNeeded_(password, salt, storedHash,
      (newHash) => sh.getRange(row, CFG.USER_COL.HASH).setValue(newHash)
    );
  }

  // ── Reset fail counter only when it is non-zero ──
  if (failCount !== 0) {
    try {
      withDocumentLock_(3000, () => {
        sh.getRange(row, CFG.USER_COL.FAILS).setValue(0);
      });
    } catch (err) {
      console.error("checkLogin reset fail count error:", err);
      return "server_busy";
    }
  }

  const sessionID = Utilities.getUuid();
  setAuthSession(normalizedEmail, sessionID);

  return mustChange ? `pw_change_required:${sessionID}` : sessionID;
}

/* =========================================================
 * LOGIN OTP
 * ======================================================= */
function canSendOtp_(email, sessionID) {
  const em    = normEmail_(email);
  const sid   = trim_(sessionID);
  const cache = getScriptCache_();

  const sessionKey = sid ? `otp_rate_sid_${sid}` : "";
  const emailKey   = em  ? `otp_rate_em_${em}`   : "";

  if (sessionKey && cache.get(sessionKey)) return false;
  if (emailKey   && cache.get(emailKey))   return false;

  if (sessionKey) cache.put(sessionKey, "1", CFG.OTP.RATE_SECONDS);
  if (emailKey)   cache.put(emailKey,   "1", CFG.OTP.RATE_SECONDS);

  return true;
}

function sendAuthCode(email, sessionID) {
  const em  = normEmail_(email);
  let   sid = trim_(sessionID);

  if (!sid) sid = findSessionID(em);
  if (!em || !sid) return false;

  const sessionEmail = getEmailFromSession(sid);
  if (!sessionEmail || sessionEmail !== em) return false;

  if (!canSendOtp_(em, sid)) return false;

  const code = String(Math.floor(100000 + Math.random() * 900000));
  getScriptCache_().put(`otp_${sid}`, code, CFG.OTP.TTL_SECONDS);

  const subject   = `${code} is your Manila IT login code`;
  const plainBody =
`Your Manila IT login verification code is: ${code}

This code expires in 5 minutes.
If you didn't request this, you can ignore this email.`;

  const htmlBody = `
<div style="margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
  <div style="max-width:560px;margin:28px auto;padding:0 16px;">
    <div style="background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);">
      <div style="padding:18px 20px;background:linear-gradient(135deg,#0A84FF,#2997FF);color:#fff;">
        <div style="font-size:13px;opacity:.92;">Manila IT Inventory</div>
        <div style="font-size:18px;font-weight:800;margin-top:4px;">Login Verification</div>
      </div>
      <div style="padding:20px;color:#f5f5f7;">
        <p style="margin:0 0 12px;font-size:14px;line-height:1.5;color:#d1d1d6;">
          Use the verification code below to finish signing in.
        </p>
        <div style="margin:16px 0;padding:14px 16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;text-align:center;">
          <div style="font-size:11px;letter-spacing:.12em;color:#b0b0b5;text-transform:uppercase;margin-bottom:6px;">
            Your Code
          </div>
          <div style="font-size:28px;font-weight:900;letter-spacing:.22em;font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Courier New',monospace;">
            ${code}
          </div>
        </div>
        <p style="margin:0;font-size:12px;color:#9a9aa0;line-height:1.4;">
          This code expires in <b>5 minutes</b>. If you didn't request this, you can ignore this email.
        </p>
      </div>
    </div>
    <div style="text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;">
      © ${new Date().getFullYear()} Manila IT Inventory
    </div>
  </div>
</div>`;

  return sendEmailWithAliasFallback_(em, subject, plainBody, htmlBody);
}

function verifyAuthCode(sessionID, code, email) {
  const sid = trim_(sessionID);
  const otp = trim_(code);
  const em  = normEmail_(email);

  if (!sid || !otp || !em) return false;

  const sessionEmail = getEmailFromSession(sid);
  if (!sessionEmail || sessionEmail !== em) return false;

  const cache  = getScriptCache_();
  const stored = cache.get(`otp_${sid}`);
  const ok     = stored === otp;

  if (ok) cache.remove(`otp_${sid}`);

  return ok;
}

/* =========================================================
 * REGISTRATION OTP + QUEUE
 * ======================================================= */
function sendVerificationCode(email) {
  const em = normEmail_(email);
  if (!em) return "invalid_email";

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const now  = new Date();

  // Store OTP in cache first (no lock needed — cache is atomic)
  getScriptCache_().put(
    `otp_${em}`,
    JSON.stringify({ code, issued: now.getTime() }),
    CFG.OTP.REG_TTL_SECONDS
  );

  // Get the sheet reference outside the lock so getPendingSheet_
  // (which may itself use a lock for sheet creation) doesn't nest locks.
  const sh = getPendingSheet_();

  withDocumentLock_(5000, () => {
    const lastRow = sh.getLastRow();
    let targetRow = 0;

    if (lastRow >= 2) {
      const emails = sh.getRange(2, CFG.PENDING_COL.EMAIL, lastRow - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (normEmail_(emails[i][0]) === em) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow) {
      sh.getRange(targetRow, CFG.PENDING_COL.CODE, 1, 3)
        .setValues([[code, "Pending Approval", now]]);
    } else {
      sh.getRange(lastRow + 1, 1, 1, 4)
        .setValues([[em, code, "Pending Approval", now]]);
    }
  });

  enqueueEmailJob_(em, code, "registration");
  ensureQueueTrigger_();

  return "sent";
}

/* =========================================================
 * CHANGE PASSWORD
 * ======================================================= */
function changePassword(sessionID, newPass) {
  const sid  = trim_(sessionID);
  const pass = String(newPass || "");

  if (!sid)             throw new Error("Missing sessionID");
  if (pass.length < 8)  throw new Error("Password too short (min 8)");

  const email = getEmailFromSession(sid);
  if (!email) throw new Error("Invalid session");

  const sh  = getUsersSheet_();
  const row = getUserRowByEmail_(sh, email);
  if (!row) throw new Error("Account not found");

  const newSalt = generateSalt();
  const newHash = hashPassword(pass, newSalt); // always writes v3

  return withDocumentLock_(5000, () => {
    const startCol = CFG.USER_COL.HASH;
    const width    = CFG.USER_COL.MUST_CHANGE - CFG.USER_COL.HASH + 1;
    const range    = sh.getRange(row, startCol, 1, width);
    const values   = range.getValues()[0];

    values[CFG.USER_COL.HASH        - startCol] = newHash;
    values[CFG.USER_COL.SALT        - startCol] = newSalt;
    values[CFG.USER_COL.FAILS       - startCol] = 0;
    values[CFG.USER_COL.MUST_CHANGE - startCol] = false;

    range.setValues([values]);
    return "ok";
  });
}

/* =========================================================
 * TOUCH ID
 * ======================================================= */
function getCredentialBySession(sessionID) {
  const sid = trim_(sessionID);
  if (!sid) return null;

  const email = getEmailFromSession(sid);
  if (!email) return null;

  const credIdBase64 = getScriptProps_().getProperty(`webauthn_cred_${email}`);
  if (!credIdBase64) return null;

  try {
    Utilities.base64Decode(credIdBase64);
    return credIdBase64;
  } catch (err) {
    console.error("Invalid stored credential ID for", email, err);
    return null;
  }
}

/* =========================================================
 * EMAIL QUEUE
 * ======================================================= */
function enqueueEmailJob_(email, code, type) {
  const em      = normEmail_(email);
  const otp     = trim_(code);
  const jobType = trim_(type || "registration");

  if (!em || !otp) return false;

  return withScriptLock_(5000, () => {
    const props = getScriptProps_();
    let queue   = safeJsonParse_(props.getProperty(CFG.EMAIL.QUEUE_PROP), []);
    if (!Array.isArray(queue)) queue = [];

    queue.push({
      id:       Utilities.getUuid(),
      type:     jobType,
      email:    em,
      code:     otp,
      created:  Date.now(),
      attempts: 0,
    });

    if (queue.length > CFG.EMAIL.MAX_QUEUE_SIZE) {
      queue = queue.slice(-CFG.EMAIL.MAX_QUEUE_SIZE);
    }

    props.setProperty(CFG.EMAIL.QUEUE_PROP, JSON.stringify(queue));
    return true;
  });
}

function ensureQueueTrigger_() {
  const props = getScriptProps_();

  if (props.getProperty(CFG.EMAIL.TRIGGER_PROP) === "1") return;

  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === "processEmailQueue_");

  if (!exists) {
    ScriptApp.newTrigger("processEmailQueue_")
      .timeBased()
      .everyMinutes(1)
      .create();
  }

  props.setProperty(CFG.EMAIL.TRIGGER_PROP, "1");
}

function processEmailQueue_() {
  try {
    withScriptLock_(20000, () => {
      const props = getScriptProps_();

      let queue = safeJsonParse_(props.getProperty(CFG.EMAIL.QUEUE_PROP), []);
      if (!Array.isArray(queue)) queue = [];

      if (queue.length === 0) {
        cleanupQueueTrigger_();
        return;
      }

      let sentMap = safeJsonParse_(props.getProperty(CFG.EMAIL.SENT_PROP), {});
      if (!sentMap || typeof sentMap !== "object" || Array.isArray(sentMap)) {
        sentMap = {};
      }

      const now = Date.now();
      for (const [id, ts] of Object.entries(sentMap)) {
        if (!ts || (now - Number(ts)) > CFG.EMAIL.SENT_TTL_MS) {
          delete sentMap[id];
        }
      }

      const batch     = queue.slice(0, CFG.EMAIL.MAX_PER_RUN);
      const rest      = queue.slice(CFG.EMAIL.MAX_PER_RUN);
      const retryJobs = [];

      for (const job of batch) {
        if (!job || !job.id || !job.email || !job.code) continue;
        if (sentMap[job.id]) continue;

        let ok = false;
        try {
          switch (job.type) {
            case "registration":
            default:
              ok = sendRegistrationEmails_(job.email, job.code);
              break;
          }
        } catch (err) {
          console.error("Email job send error:", job.id, err);
          ok = false;
        }

        if (ok) {
          sentMap[job.id] = now;
        } else {
          job.attempts = toInt_(job.attempts, 0) + 1;
          retryJobs.push(job);
        }
      }

      const newQueue = retryJobs.concat(rest);

      props.setProperty(CFG.EMAIL.QUEUE_PROP, JSON.stringify(newQueue));
      props.setProperty(CFG.EMAIL.SENT_PROP,  JSON.stringify(sentMap));

      if (newQueue.length === 0) cleanupQueueTrigger_();
    });
  } catch (err) {
    console.error("processEmailQueue_ error:", err);
  }
}

function cleanupQueueTrigger_() {
  const props    = getScriptProps_();
  const triggers = ScriptApp.getProjectTriggers();

  for (const t of triggers) {
    if (t.getHandlerFunction() === "processEmailQueue_") {
      ScriptApp.deleteTrigger(t);
    }
  }

  props.deleteProperty(CFG.EMAIL.TRIGGER_PROP);
}

/* =========================================================
 * REGISTRATION EMAILS
 * ======================================================= */
function sendRegistrationEmails_(email, code) {
  const appUrl = getAppUrl();

  const userHtml = `
<div style="background:#111;color:#fff;padding:30px;border-radius:10px;font-family:Arial,sans-serif;">
  <h2 style="text-align:center;margin-top:0;">MANILA IT PORTAL</h2>
  <p style="text-align:center;color:#ccc;">Your verification code:</p>
  <div style="background:#fff;color:#000;padding:15px;font-size:32px;text-align:center;border-radius:6px;">
    <b>${code}</b>
  </div>
  <p style="text-align:center;color:#bbb;margin-top:20px;font-size:13px;">
    Enter this in the registration page. This code expires in 10 minutes.
  </p>
</div>`;

  const adminHtml = `
<div style="font-family:Arial,sans-serif;padding:25px;background:white;border-radius:10px;border:1px solid #ddd;">
  <h2 style="text-align:center;margin:0;color:#222;">New Pending Registration</h2>
  <p style="font-size:14px;color:#444;margin-top:10px;">Approve or reject new user request:</p>
  <div style="background:#f5f5f5;padding:10px;border-radius:6px;margin:15px 0;border:1px solid #e3e3e3;">
    <b>Email:</b> ${email}
  </div>
  <div style="text-align:center;margin-top:25px;">
    <a href="${appUrl}?approve=${encodeURIComponent(email)}"
       style="background:#0d8a2f;padding:13px 26px;color:white;border-radius:8px;text-decoration:none;font-weight:bold;">
      APPROVE
    </a>
    <a href="${appUrl}?reject=${encodeURIComponent(email)}"
       style="background:#b50000;padding:13px 26px;color:white;border-radius:8px;margin-left:10px;text-decoration:none;font-weight:bold;">
      REJECT
    </a>
  </div>
</div>`;

  const userOk  = sendEmailWithAliasFallback_(
    email, "Manila IT Inventory – Verification Code", "", userHtml
  );
  const adminOk = sendEmailWithAliasFallback_(
    CFG.EMAIL.ADMIN_EMAIL, "User Approval – Manila IT Registration", "", adminHtml
  );

  return userOk && adminOk;
}

function sendEmailWithAliasFallback_(to, subject, plainBody, htmlBody) {
  const recipient = trim_(to);
  if (!recipient) return false;

  const primaryOptions = {
    htmlBody: htmlBody || "",
    name:     CFG.EMAIL.FROM_NAME,
    from:     CFG.EMAIL.FROM_ALIAS,
    replyTo:  CFG.EMAIL.FROM_ALIAS,
  };

  try {
    GmailApp.sendEmail(recipient, subject, plainBody || "", primaryOptions);
    return true;
  } catch (err) {
    console.error("sendEmail alias error:", err);
    try {
      GmailApp.sendEmail(recipient, subject, plainBody || "", {
        htmlBody: htmlBody || "",
        name:     CFG.EMAIL.FROM_NAME,
      });
      return true;
    } catch (fallbackErr) {
      console.error("sendEmail fallback error:", fallbackErr);
      return false;
    }
  }
}

/* =========================================================
 * DEV SESSION
 * ======================================================= */
function createDevSession() {
  const sid = `DEV_${Utilities.getUuid()}`;
  setAuthSession("developer", sid);
  return sid;
}