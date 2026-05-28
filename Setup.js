function setHmacSecret() {
  PropertiesService.getScriptProperties()
    .setProperty('HMAC_SECRET', Utilities.base64EncodeWebSafe(
      Utilities.computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        Math.random().toString() + Date.now()
      )
    ));
  console.log('HMAC_SECRET set successfully.');
}


function fixMyRole() {
  const email = "michael.beleno.xc@betfanatics.com";

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === email.toLowerCase()) {
      sh.getRange(i + 1, 5).setValue("admin"); // E = Role
      invalidateUsersCache_();
      console.log("Role set to admin for: " + email);
      return;
    }
  }
  console.log("User not found.");
}


function clearRbac() {
  clearRbacCaches_();
  console.log("RBAC cache cleared.");
}

function seedAndClear() {
  seedDefaultRbac_();      // from RbacApi.gs — seeds admin + viewer rules
  clearRbacCaches_();      // clears stale cache
  console.log("Done");
}

function fixEverything() {
  clearRbacCaches_();
  seedDefaultRbac_();
  clearRbacCaches_();
  console.log("Done");
}

function debugMyRole() {
  const email = "michael.beleno.xc@betfanatics.com";
  clearRbacCaches_();
  const map = getUserRoleMap_();
  console.log("Full role map:", JSON.stringify(map));
  console.log("My role:", map[email]);
  console.log("canAccessPage home:", canAccessPage("home", email));
}

function warmCache() {
  clearRbacCaches_();
  const email = "michael.beleno.xc@betfanatics.com";
  const role  = getUserRoleByEmail_(email);  // builds and caches role map
  getRbacRuleMap_();                          // builds and caches rule map
  console.log("Cache warmed. Role:", role);
  console.log("Can access home:", canAccessPage("home", email));
}

function testGetPageHtml() {
  const result = getPageHtml('home', {}, '');
  console.log(JSON.stringify(result));
}

function checkMissingFunctions() {
  const fns = [
    '_clearDoGetUserCache_',
    'getUserRoleByEmail_',
    'canAccessPage',
    'clearRbacCaches_',
    'getPageFragment_',
    'buildBootstrap_'
  ];
  fns.forEach(fn => {
    try {
      const exists = typeof eval(fn) === 'function';
      console.log(fn + ':', exists ? 'EXISTS' : 'MISSING');
    } catch(e) {
      console.log(fn + ': ERROR -', e.message);
    }
  });
}

function debugGetPageHtml() {
  try {
    const result = getPageHtml('home', {}, '');
    console.log('Result:', JSON.stringify(result).substring(0, 500));
  } catch(e) {
    console.log('Caught error:', e.message);
    console.log('Stack:', e.stack);
  }
}

function clearAllCaches() {
  // Clear RBAC caches (also flushes the per-user doGet cache)
  clearRbacCaches_();

  const cache = CacheService.getScriptCache();

  // Clear shell fragment caches — derived from the authoritative
  // allowlist so new pages are covered automatically
  const shellKeys = Array.from(SHELL_ALLOWED_PAGES_).map(buildShellCacheKey_);

  // Clear other caches the app maintains
  const otherKeys = [
    'users_data_v2'
  ];

  cache.removeAll(shellKeys.concat(otherKeys));

  console.log('All caches cleared (' +
    shellKeys.length + ' shell pages + ' +
    otherKeys.length + ' other keys)');
}

function simulateColdStart() {
  // Clear everything to simulate cold start
  CacheService.getScriptCache().removeAll([
    'rbac_rules_v1', 'rbac_rule_map_v1', 
    'rbac_roles_v1', 'rbac_user_role_map_v1'
  ]);
  
  // Now immediately call getPageHtml like the shell does
  const result = getPageHtml('home', {}, '');
  console.log('ok:', result.ok);
  console.log('error:', result.error);
  console.log('html length:', result.html?.length);
}


function clearMasterlistCache() {
  CacheService.getScriptCache().remove('shell_v3_masterlist');
  console.log('Done');
}

function testAdminPerms() {
  const email = 'michael.beleno.xc@betfanatics.com';
  const perms = getMyPermissions();
  console.log(JSON.stringify(perms));
}

function invalidateHomeAndUserCaches_() {
  const cache = CacheService.getScriptCache();
  cache.removeAll([
    "users_data_v2",
    "shell_v3_home",
    "shell_v3_profile"
  ]);
}


function testRbacMatrix() {
  const result = getRbacMatrix();
  console.log(JSON.stringify(result));
}

function testSeedRbac() {
  seedDefaultRbac_();
  const result = getRbacMatrix();
  console.log(JSON.stringify(result));
}


function debugApprovedRequestFlow() {
  const shifts = getEmployeesFromShifts();
  const users = getTeamList();

  Logger.log("=== SHIFTS ===");
  shifts.slice(-20).forEach(s => Logger.log(JSON.stringify(s)));

  Logger.log("=== USERS ===");
  users.slice(0, 50).forEach(u => Logger.log(JSON.stringify(u)));
}

function debugScheduleCreation() {
  const sh = getRequestsSheet_();
  const values = sh.getDataRange().getValues();
  
  // Find the approved VL row
  const headers = values[0];
  const hm = getRequestHeaderMap_(headers);
  
  const row = values.find((r, i) => 
    i > 0 && String(r[hm.status] || "").trim().toLowerCase() === "approved"
  );
  
  if (!row) { Logger.log("No approved row found"); return; }
  
  const rowIndex = values.indexOf(row);
  Logger.log("Found row at index: " + rowIndex + " (sheet row " + (rowIndex + 1) + ")");
  Logger.log("Type: " + row[hm.type]);
  Logger.log("Email: " + row[hm.email]);
  Logger.log("StartDate raw: " + row[hm.startDate]);
  Logger.log("EndDate raw: " + row[hm.endDate]);
  Logger.log("ScheduleRowIds col index: " + hm.scheduleRowIds);
  Logger.log("ScheduleRowIds value: " + (hm.scheduleRowIds > -1 ? row[hm.scheduleRowIds] : "COLUMN NOT FOUND"));

  // Test date parsing
  const startRaw = row[hm.startDate];
  const startDate = (startRaw instanceof Date) ? startRaw : parseMMDDYYYY_(String(startRaw));
  Logger.log("startDate parsed: " + startDate);

  // Test user lookup
  try {
    const user = lookupUserInfoByEmail_(String(row[hm.email]));
    Logger.log("User lookup result: " + JSON.stringify(user));
  } catch(e) {
    Logger.log("lookupUserInfoByEmail_ ERROR: " + e.message);
  }

  // Test getShiftsSheet_
  try {
    const shifts = getShiftsSheet_();
    Logger.log("Shifts sheet found: " + shifts.getName());
  } catch(e) {
    Logger.log("getShiftsSheet_ ERROR: " + e.message);
  }
}

function reprocessApprovedRequests() {
  const sh = getRequestsSheet_();
  const values = sh.getDataRange().getValues();
  const hm = getRequestHeaderMap_(values[0]);

  values.forEach((row, i) => {
    if (i === 0) return;
    const status = String(row[hm.status] || "").trim().toLowerCase();
    const scheduleVal = hm.scheduleRowIds > -1 ? String(row[hm.scheduleRowIds] || "").trim() : "";
    
    if (status === "approved" && !scheduleVal) {
      Logger.log("Reprocessing row " + (i + 1));
      createScheduleIfNeeded_(sh, { map: buildMapFromHm_(hm) }, i + 1);
    }
  });
}

function debugUsersSheet() {
  const sh = getUsersSheet_();
  Logger.log("Sheet found: " + (sh ? sh.getName() : "NULL"));
  if (!sh) return;
  
  Logger.log("Last row: " + sh.getLastRow());
  Logger.log("Last col: " + sh.getLastColumn());
  
  if (sh.getLastRow() > 0 && sh.getLastColumn() > 0) {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    Logger.log("Headers: " + JSON.stringify(headers));
  }
  
  const hm = getHeaderMap_(sh);
  Logger.log("Header map: " + JSON.stringify(hm.map));
  Logger.log("hm.headers.length: " + hm.headers.length);
}

// Helper to convert hm back to the raw {ColumnName: index} map that createScheduleIfNeeded_ expects
function buildMapFromHm_(hm) {
  const map = {};
  const keyToHeader = {
    type: "Type", email: "Email", startDate: "StartDate",
    endDate: "EndDate", timeOrHours: "TimeorHours",
    scheduleRowIds: "ScheduleRowIds"
  };
  Object.entries(keyToHeader).forEach(([key, header]) => {
    if (hm[key] > -1) map[header] = hm[key];
  });
  return map;
}

function nukeCacheNow() {
  const cache = CacheService.getScriptCache();
  
  cache.removeAll([
    'shell_v3_publish',
    'shell_publish',
    'publish',
    'shell_v4_cheatsheet',
    'home_posts_v1',
    'rbac_rules_v1',
    'rbac_rule_map_v1',
    'rbac_roles_v1',
    'rbac_user_role_map_v1',
    'users_data_v2'
  ]);
  
  Logger.log('All caches nuked.');
}


function clearHomeCache() {
  const cache = CacheService.getScriptCache();
  const key = buildShellCacheKey_('home');
  cache.remove(key);
  Logger.log('Home cache key removed: ' + key);

  // Also nuke common variants
  cache.removeAll([
    'shell_v3_home',
    'shell_home',
    'home',
    'home_posts_v1'
  ]);
  Logger.log('Done.');
}


function nukeWarrantyCache() {
  CacheService.getScriptCache().remove("shell_v4_warrantyconfig");
  Logger.log("Fragment cache cleared. Hard-refresh your browser now.");
}


function diagnoseWarrantyConfig() {
  var email = getCurrentUserEmail_();
  Logger.log("=== Diagnostic ===");
  Logger.log("email: " + email);
  Logger.log("role: " + getCurrentUserRole_());
  Logger.log("canAccessPage(assetreport): " + canAccessPage("assetreport", email));
  Logger.log("canAccessPage(warrantyconfig): " + canAccessPage("warrantyconfig", email));
  Logger.log("canDoAction(assetreport.export): " + canDoAction("assetreport.export", email));
  
  Logger.log("--- Testing getPageHtml ---");
  try {
    var result = getPageHtml("warrantyconfig", {}, "");
    Logger.log("ok: " + result.ok);
    Logger.log("error: " + (result.error || "(none)"));
    Logger.log("html length: " + (result.html ? result.html.length : 0));
    Logger.log("html first 300 chars: " + (result.html ? result.html.substring(0, 300) : "(empty)"));
  } catch (e) {
    Logger.log("getPageHtml threw: " + e.message);
  }
  
  Logger.log("--- Testing getWarrantyReminderConfigClient ---");
  try {
    var cfg = getWarrantyReminderConfigClient();
    Logger.log("config result: " + JSON.stringify(cfg));
  } catch (e) {
    Logger.log("getWarrantyReminderConfigClient threw: " + e.message);
  }
  
  Logger.log("--- Cache state ---");
  try {
    var cached = CacheService.getScriptCache().get("shell_v4_warrantyconfig");
    Logger.log("shell_v4_warrantyconfig cache: " + (cached ? (cached.length + " bytes cached") : "(not cached)"));
  } catch (e) {
    Logger.log("cache check threw: " + e.message);
  }
}

function debugShellTopPerformers() {
  // Force-fresh cache — bypasses any stale data
  CacheService.getScriptCache().removeAll([
    "rbac_rules_v1", "rbac_rule_map_v1", "rbac_roles_v1", 
    "rbac_user_role_map_v1", "shell_v4_topperformers"
  ]);
  
  const email = Session.getActiveUser().getEmail();
  
  // Simulate the EXACT check the shell does
  const shellCheckResult = canAccessPage("topperformers", email);
  
  // Read the actual running source of getPageFragment_
  const fnSource = getPageFragment_.toString();
  const hasOldAdminCheck = fnSource.includes('admin.access') && fnSource.includes('topperformers');
  const hasNewPageCheck  = fnSource.includes('canAccessPage("topperformers"');
  
  // Try to actually call getPageHtml and see what it returns
  let result;
  try {
    result = getPageHtml('topperformers', {}, '');
  } catch(e) {
    result = { error: e.message };
  }
  
  const output = {
    email,
    shellCheck_canAccessPage_topperformers: shellCheckResult,
    getPageFragment_stillHasOldAdminCheck: hasOldAdminCheck,
    getPageFragment_hasNewPageCheck: hasNewPageCheck,
    getPageHtml_result_ok: result && result.ok,
    getPageHtml_result_error: result && result.error,
    getPageHtml_result_htmlLength: result && result.html ? result.html.length : 0
  };
  
  Logger.log(JSON.stringify(output, null, 2));
  return output;
}


function testSlackWebhook() {
  try {
    sendSlackNotification_(null, [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: "*\uD83E\uDDEA Slack webhook test from Manila IT Inventory*"
        }
      },
      {
        type: "section",
        fields: [
          { type: "mrkdwn", text: "*Test:*\nWebhook reachability" },
          { type: "mrkdwn", text: "*Run by:*\n" + (Session.getActiveUser().getEmail() || "unknown") },
          { type: "mrkdwn", text: "*Time:*\n" + new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" }) + " PHT" }
        ]
      },
      {
        type: "context",
        elements: [
          { type: "mrkdwn", text: "If you see this in Slack, the webhook is working. \uD83D\uDC4D" }
        ]
      }
    ], "#3b82f6");

    Logger.log("Test posted (no exception thrown).");
  } catch (e) {
    Logger.log("FAILED: " + (e && e.message ? e.message : String(e)));
    throw e;
  }
}

function pingSlackDirect() {
  var url = PropertiesService.getScriptProperties()
    .getProperty("SLACK_WEBHOOK_URL");

  if (!url) {
    Logger.log("Property SLACK_WEBHOOK_URL is empty. Run setSlackWebhookUrl() first.");
    return;
  }

  Logger.log("URL prefix: " + url.substring(0, 45) + "...");

  var resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ text: "Direct ping test " + new Date().toISOString() }),
    muteHttpExceptions: true
  });

  Logger.log("HTTP " + resp.getResponseCode());
  Logger.log("Body: " + resp.getContentText());
}

function cleanRamColumn() {
  var sh = _mlSheet_();
  var hm = _mlHeaderMap_(sh);
  var ramCol = hm.map.ram;
  if (!ramCol) {
    Logger.log("RAM column not found.");
    return;
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  var range = sh.getRange(2, ramCol, lastRow - 1, 1);
  var values = range.getValues();
  var changed = 0;

  for (var i = 0; i < values.length; i++) {
    var raw = String(values[i][0] || '');
    var cleaned = raw
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
    if (cleaned !== raw) {
      values[i][0] = cleaned;
      changed++;
    }
  }

  if (changed > 0) {
    range.setValues(values);
    Logger.log("Cleaned " + changed + " RAM cells.");
  } else {
    Logger.log("No dirty RAM values found.");
  }
}