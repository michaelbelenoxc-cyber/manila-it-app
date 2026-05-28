function assertCanViewEngagement_(actorEmail) {
  const email = normalizeEmail_(actorEmail || getCurrentUserEmail_());

  if (!canAccessPage("companyengagement", email)) {
    throw new Error("You do not have permission to access Company Engagement.");
  }
}

function assertCanManageEngagement_(actorEmail) {
  const email = normalizeEmail_(actorEmail || getCurrentUserEmail_());

  if (!canAccessPage("companyengagement", email)) {
    throw new Error("You do not have permission to access Company Engagement.");
  }

  if (!canDoAction("companyengagement.manage", email)) {
    throw new Error("You do not have permission to manage Company Engagement.");
  }
}

function assertCanManageUsers_(actorEmail) {
  const email = String(actorEmail || "").trim().toLowerCase();
  if (!email) throw new Error("Missing actor email.");

  if (!canDoAction("users.manage", email)) {
    throw new Error("You do not have permission to manage users.");
  }
}

function assertValidUserRow_(sh, index) {
  const sheetRow = Number(index) + 2;
  const lastRow = sh.getLastRow();

  if (!Number.isFinite(sheetRow) || sheetRow < 2 || sheetRow > lastRow) {
    throw new Error("Invalid row index.");
  }

  return sheetRow;
}

function parseDateForSheet_(dateStr) {
  const s = String(dateStr || "").trim();
  if (!s) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const parts = s.split("-").map(Number);
    return new Date(parts[0], parts[1] - 1, parts[2]);
  }

  return s;
}

function addCompanyEngagementBulk(rows, actorEmail) {
  assertCanManageEngagement_(actorEmail);

  if (!Array.isArray(rows) || !rows.length) return true;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Company Engagement");
  if (!sh) throw new Error('Sheet "Company Engagement" not found.');

  const values = rows.map(r => [
    String(r?.employee || "").trim(),
    parseDateForSheet_(r?.date),
    String(r?.event || "").trim(),
    String(r?.status || "").trim(),
    String(r?.comment || "").trim()
  ]);

  sh.getRange(sh.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
  return true;
}

function deleteCompanyEngagementByEvent(eventName, actorEmail) {
  assertCanManageEngagement_(actorEmail);

  const targetEvent = String(eventName || "").trim();
  if (!targetEvent) throw new Error("Missing event name.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Company Engagement");
  if (!sh) throw new Error('Sheet "Company Engagement" not found.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return true;

  const events = sh.getRange(2, 3, lastRow - 1, 1).getValues().flat();
  const rowsToDelete = [];

  for (let i = 0; i < events.length; i++) {
    if (String(events[i] || "").trim() === targetEvent) {
      rowsToDelete.push(i + 2);
    }
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sh.deleteRow(rowsToDelete[i]);
  }

  return true;
}

function getCompanyEngagementData(actorEmail) {
  assertCanViewEngagement_(actorEmail);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Company Engagement");
  if (!sh) throw new Error('Sheet "Company Engagement" not found.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h || "").trim());
  const idx = {
    employee: headers.indexOf("Employee"),
    date: headers.indexOf("Date"),
    event: headers.indexOf("Event"),
    status: headers.indexOf("Status"),
    comment: headers.indexOf("Comment")
  };

  const missing = Object.entries(idx)
    .filter(([, v]) => v === -1)
    .map(([k]) => k);

  if (missing.length) {
    throw new Error("Missing columns: " + missing.join(", "));
  }

  return values.slice(1).map((r, i) => ({
    id: i + 2,
    employee: String(r[idx.employee] ?? "").trim(),
    date: formatDate_(r[idx.date]),
    event: String(r[idx.event] ?? "").trim(),
    status: String(r[idx.status] ?? "").trim(),
    comment: String(r[idx.comment] ?? "").trim()
  })).filter(x => x.employee || x.date || x.event || x.status || x.comment);
}

function addCompanyEngagement(payload) {
  assertCanManageEngagement_(payload?.actorEmail);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company Engagement");
  if (!sh) throw new Error('Sheet "Company Engagement" not found.');

  sh.appendRow([
    String(payload?.employee || "").trim(),
    parseDateForSheet_(payload?.date),
    String(payload?.event || "").trim(),
    String(payload?.status || "").trim(),
    String(payload?.comment || "").trim()
  ]);

  return true;
}

function updateCompanyEngagement(payload) {
  assertCanManageEngagement_(payload?.actorEmail);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company Engagement");
  if (!sh) throw new Error('Sheet "Company Engagement" not found.');
  if (!payload || !payload.id) throw new Error("Missing row id.");

  const rowNumber = Number(payload.id);
  if (!Number.isFinite(rowNumber) || rowNumber < 2 || rowNumber > sh.getLastRow()) {
    throw new Error("Invalid row id.");
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
  const col = {
    employee: headers.indexOf("Employee") + 1,
    date: headers.indexOf("Date") + 1,
    event: headers.indexOf("Event") + 1,
    status: headers.indexOf("Status") + 1,
    comment: headers.indexOf("Comment") + 1
  };

  Object.entries(col).forEach(([k, c]) => {
    if (c === 0) throw new Error(`Missing column: ${k}`);
  });

  sh.getRange(rowNumber, col.employee).setValue(String(payload.employee || "").trim());
  sh.getRange(rowNumber, col.date).setValue(parseDateForSheet_(payload.date));
  sh.getRange(rowNumber, col.event).setValue(String(payload.event || "").trim());
  sh.getRange(rowNumber, col.status).setValue(String(payload.status || "").trim());
  sh.getRange(rowNumber, col.comment).setValue(String(payload.comment || "").trim());

  return true;
}

function deleteCompanyEngagement(rowId, actorEmail) {
  assertCanManageEngagement_(actorEmail);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company Engagement");
  if (!sh) throw new Error('Sheet "Company Engagement" not found.');

  const rowNumber = Number(rowId);
  if (!Number.isFinite(rowNumber) || rowNumber < 2 || rowNumber > sh.getLastRow()) {
    throw new Error("Invalid row id.");
  }

  sh.deleteRow(rowNumber);
  return true;
}