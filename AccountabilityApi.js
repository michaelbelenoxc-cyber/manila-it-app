// ===== CONFIG =====
var AC_TEMPLATE_DOC_ID    = "1GbeQZukrGCq_YDdbnQa1Ejbm3X_hAM0YJS2sETHJvjQ";
var AC_OUTPUT_FOLDER_ID   = "1e3xT6FYK9TMeRbL3DNccefLVstKqJojH";
var SIGNED_AF_FOLDER_ID   = "1uJvQ5uZbrb8iACcu4RBakemKF_01Kk6z";
var SIGNED_AF_ROOT_FOLDER_ID = "1uJvQ5uZbrb8iACcu4RBakemKF_01Kk6z";
var WEBAPP_BASE_URL       = "https://script.google.com/a/macros/betfanatics.com/s/AKfycbwAa0_5WuloiEPjwqELYOzB08-igHNfPfTcf3nQeZbJZ9Ea5K2lxD61XgzMVZsvyB-r/exec";
var EMPLOYEES_SHEET_NAME  = "Employees";
var SEND_EMAIL_AFTER_GENERATE = true;
var TRASH_AC_DOC_COPY     = true;


/* ============================================================
 * PUBLIC — called by masterlist_generateAccountabilityForm
 * ============================================================ */
function generateAccountabilityForm(serialTag) {
  var tag = _trim_(serialTag);
  if (!tag)                throw new Error("Missing Serial Tag");
  if (!AC_TEMPLATE_DOC_ID) throw new Error("Missing AC_TEMPLATE_DOC_ID");

  var asset = _getMasterlistAssetBySerialTag_(tag);
  if (!asset) throw new Error("Asset not found: " + tag);

  var emp = _getEmployeeByNameOrEmail_(asset.assignee, asset.email);
  if (!emp) {
    throw new Error(
      "Employee not found. Assignee=\"" + asset.assignee + "\" Email=\"" + asset.email + "\""
    );
  }

  var folder   = _getOutputFolder_();
  var tz       = Session.getScriptTimeZone();
  var today    = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");
  var safeName = _safeName_(emp.displayName || emp.name || "Employee");
  var filename = "Accountability Form - " + safeName + " - " + tag;

  // 1) Copy template and fill placeholders
  var docCopy = DriveApp.getFileById(AC_TEMPLATE_DOC_ID).makeCopy(filename, folder);
  var docId   = docCopy.getId();

  var accessoriesText = _formatAccessoriesForAf_(asset.accessories);

  _fillDocPlaceholders_(docId, {
    NAME:        safeName,
    DATE:        today,
    DEPARTMENT:  emp.department,
    EMPLOYEE_ID: emp.employeeId,
    ASSET_TYPE:  asset.type,
    BRAND_MODEL: _joinDefined_([asset.type, asset.model], " / "),
    SERIAL_TAG:  asset.serialTag || tag,
    CONDITION:   asset.condition,
    ACCESSORIES: accessoriesText
  });

  // 2) Export PDF
  var pdfFile   = _exportDocToPdfFile_(docId, filename + ".pdf", folder);
  var pdfUrl    = pdfFile.getUrl();
  var pdfFileId = pdfFile.getId();

  // 3) Optionally trash the doc copy
  if (TRASH_AC_DOC_COPY) docCopy.setTrashed(true);

  // 4) Write link back to Masterlist
  _setAfAfterGenerate_(tag, pdfUrl);

  // 5) Optionally email the PDF
  if (SEND_EMAIL_AFTER_GENERATE) {
    _emailAccountabilityPdf_(
      {
        email:        asset.email        || emp.email,
        assignee:     asset.assignee     || safeName,
        type:         asset.type,
        model:        asset.model,
        serialNumber: asset.serialNumber,
        accessories:  asset.accessories
      },
      tag,
      pdfUrl,
      pdfFileId,
      Session.getActiveUser().getEmail()
    );
  }

  return pdfUrl;
}


/* ============================================================
 * ACCESSORIES FORMATTER
 * ============================================================ */
function _formatAccessoriesForAf_(raw) {
  var parts  = String(raw || "").split(",");
  var items  = [];
  for (var i = 0; i < parts.length; i++) {
    var s = String(parts[i] || "").trim();
    if (s) items.push(s);
  }
  return items.length ? items.join(", ") : "None";
}


/* ============================================================
 * DOC TEMPLATE HELPERS
 * ============================================================ */
function _fillDocPlaceholders_(docId, data) {
  var doc  = DocumentApp.openById(docId);
  var body = doc.getBody();
  var keys = Object.keys(data || {});
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    body.replaceText("{{" + k + "}}", String(data[k] == null ? "" : data[k]));
  }
  doc.saveAndClose();
}

function _exportDocToPdfFile_(docId, pdfName, folder) {
  var pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(pdfName);
  return folder.createFile(pdfBlob);
}

function _getOutputFolder_() {
  if (AC_OUTPUT_FOLDER_ID) return DriveApp.getFolderById(AC_OUTPUT_FOLDER_ID);
  return DriveApp.getRootFolder();
}

function _safeName_(name) {
  return String(name || "Employee")
    .trim()
    .replace(/[\/\\:*?"<>|]+/g, "")
    .replace(/\s+/g, " ")
    .slice(0, 120) || "Employee";
}

function _trim_(v) {
  return String(v == null ? "" : v).trim();
}

function _joinDefined_(arr, sep) {
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var s = String(arr[i] == null ? "" : arr[i]).trim();
    if (s) out.push(s);
  }
  return out.join(sep);
}


/* ============================================================
 * MASTERLIST LOOKUP
 * ============================================================ */
function _getMasterlistAssetBySerialTag_(serialTag) {
  var tag = String(serialTag == null ? "" : serialTag)
    .trim()
    .replace(/^"+|"+$/g, "");
  if (!tag) return null;

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MASTERLIST_SHEET);
  if (!sh) throw new Error("Sheet \"" + MASTERLIST_SHEET + "\" not found.");

  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return null;

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(String(values[0][h] || "").trim().toLowerCase());
  }

  function col(name) {
    return headers.indexOf(String(name).trim().toLowerCase());
  }

  var cSerial = col("Serial Tag");
  if (cSerial < 0) throw new Error("Header \"Serial Tag\" not found.");

  var row = null;
  for (var i = 1; i < values.length; i++) {
    if (_trim_(values[i][cSerial]) === tag) { row = values[i]; break; }
  }
  if (!row) return null;

  function pick(headerName, fallback) {
    var idx = col(headerName);
    return idx >= 0 ? _trim_(row[idx]) : (fallback == null ? "" : fallback);
  }

  return {
    type:         pick("Type"),
    model:        pick("Model"),
    serialTag:    pick("Serial Tag"),
    serialNumber: pick("Serial Number"),
    assignee:     pick("Assignee"),
    email:        pick("Email"),
    condition:    pick("Notes", "") || pick("Condition", ""),
    accessories:  pick("Accessories", "")
  };
}


/* ============================================================
 * EMPLOYEES LOOKUP
 * ============================================================ */
function _getEmployeeByNameOrEmail_(name, email) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMPLOYEES_SHEET_NAME);
  if (!sh) throw new Error("Sheet \"" + EMPLOYEES_SHEET_NAME + "\" not found.");

  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return null;

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(String(values[0][h] || "").trim().toLowerCase());
  }

  function col(n) { return headers.indexOf(String(n).trim().toLowerCase()); }

  var cEmail = col("Email");
  var cName  = col("Display Name");
  if (cEmail < 0) throw new Error("Employees header \"Email\" not found.");
  if (cName  < 0) throw new Error("Employees header \"Display Name\" not found.");

  var targetEmail = _trim_(email).toLowerCase();
  var targetName  = _trim_(name).toLowerCase();

  var row = null;
  for (var i = 1; i < values.length; i++) {
    var rowEmail = _trim_(values[i][cEmail]).toLowerCase();
    var rowName  = _trim_(values[i][cName]).toLowerCase();
    if ((targetEmail && rowEmail === targetEmail) ||
        (targetName  && rowName  === targetName)) {
      row = values[i];
      break;
    }
  }
  if (!row) return null;

  function pick(headerName) {
    var idx = col(headerName);
    return idx >= 0 ? _trim_(row[idx]) : "";
  }

  return {
    displayName: _trim_(row[cName]),
    email:       _trim_(row[cEmail]),
    department:  pick("Department"),
    employeeId:  pick("Employee ID")
  };
}


/* ============================================================
 * MASTERLIST WRITE-BACK — ACCOUNTABILITY FORM LINK
 * ============================================================ */
function _setAccountabilityLink_(serialTag, url) {
  var tag = _trim_(serialTag);
  if (!tag) throw new Error("Missing serialTag");

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MASTERLIST_SHEET);
  if (!sh) throw new Error("Sheet \"" + MASTERLIST_SHEET + "\" not found.");

  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error("Masterlist has no data.");

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(_normHeader_(values[0][h]));
  }

  function col(name) { return headers.indexOf(_normHeader_(name)); }

  var cSerial = col("Serial Tag");
  var cForm   = col("Accountability Form");
  if (cSerial < 0) throw new Error("Header \"Serial Tag\" not found.");
  if (cForm   < 0) throw new Error("Header \"Accountability Form\" not found.");

  var rIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (_trim_(values[i][cSerial]) === tag) { rIndex = i; break; }
  }
  if (rIndex < 0) throw new Error("Serial Tag not found in Masterlist: " + tag);

  sh.getRange(rIndex + 1, cForm + 1).setValue(String(url || ""));
  return true;
}


/* ============================================================
 * MASTERLIST WRITE-BACK — AF AFTER GENERATE
 * ============================================================ */
function _setAfAfterGenerate_(serialTag, pdfUrl) {
  var tag = _trim_(serialTag);
  if (!tag) throw new Error("Missing Serial Tag");

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MASTERLIST_SHEET);
  if (!sh) throw new Error("Sheet \"" + MASTERLIST_SHEET + "\" not found.");

  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error("Masterlist has no data.");

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(String(values[0][h] || "").trim().toLowerCase());
  }

  function col(name) { return headers.indexOf(String(name).trim().toLowerCase()); }

  var cSerial = col("serial tag");
  var cForm   = col("accountability form");
  var cStatus = col("af status");
  var cSigned = col("signed af");
  var cDate   = col("af signed date");

  if (cSerial < 0) throw new Error("Header \"Serial Tag\" not found.");
  if (cForm   < 0) throw new Error("Header \"Accountability Form\" not found.");
  if (cStatus < 0) throw new Error("Header \"AF Status\" not found.");

  var rIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][cSerial] || "").trim() === tag) { rIndex = i; break; }
  }
  if (rIndex < 0) throw new Error("Serial Tag not found in Masterlist: " + tag);

  var row = rIndex + 1;

  sh.getRange(row, cForm   + 1).setValue(String(pdfUrl || ""));
  sh.getRange(row, cStatus + 1).setValue("Pending");
  if (cSigned >= 0) sh.getRange(row, cSigned + 1).setValue("");
  if (cDate   >= 0) sh.getRange(row, cDate   + 1).setValue("");

  return true;
}


/* ============================================================
 * MASTERLIST WRITE-BACK — SIGNED AF
 * ============================================================ */
function _setMasterlistSignedAf_(serialTag, signedUrl) {
  var tag = _trim_(serialTag);
  if (!tag) throw new Error("Missing serialTag");

  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(MASTERLIST_SHEET);
  if (!sh) throw new Error("Sheet \"" + MASTERLIST_SHEET + "\" not found.");

  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error("Masterlist has no data.");

  var headers = [];
  for (var h = 0; h < values[0].length; h++) {
    headers.push(String(values[0][h] || "").trim().toLowerCase());
  }

  function col(name) { return headers.indexOf(String(name).trim().toLowerCase()); }

  var cSerial = col("serial tag");
  var cSigned = col("signed af");
  var cStatus = col("af status");
  var cDate   = col("af signed date");

  if (cSerial < 0) throw new Error("Header \"Serial Tag\" not found.");
  if (cSigned < 0) throw new Error("Header \"Signed AF\" not found.");
  if (cStatus < 0) throw new Error("Header \"AF Status\" not found.");
  if (cDate   < 0) throw new Error("Header \"AF Signed Date\" not found.");

  var rIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][cSerial] || "").trim() === tag) { rIndex = i; break; }
  }
  if (rIndex < 0) throw new Error("Serial Tag not found in Masterlist: " + tag);

  var row = rIndex + 1;
  sh.getRange(row, cSigned + 1).setValue(String(signedUrl || ""));
  sh.getRange(row, cStatus + 1).setValue("Signed");
  sh.getRange(row, cDate   + 1).setValue(new Date());

  SpreadsheetApp.flush(); 
}


/* ============================================================
 * EMAIL — ACCOUNTABILITY FORM
 * ============================================================ */
function _emailAccountabilityPdf_(asset, serialTag, pdfUrl, pdfFileId, generatedByEmail) {
  var tag = _trim_(serialTag);
  if (!tag)   throw new Error("Missing Serial Tag");
  if (!asset) throw new Error("Missing asset data");

  var to = _trim_(asset.email);
  if (!to) throw new Error("No assignee email found for Serial Tag \"" + tag + "\".");

  var fileId = _trim_(pdfFileId);
  if (!fileId && pdfUrl) fileId = _extractDriveFileId_(pdfUrl);
  if (!fileId) throw new Error("Could not determine PDF fileId.");

  var file = DriveApp.getFileById(fileId);

  var safeAssignee = String(asset.assignee || "Employee")
    .trim()
    .replace(/[\/\\:*?"<>|]+/g, "")
    .replace(/\s+/g, " ")
    .slice(0, 80) || "Employee";

  var pdfBlob = file
    .getAs(MimeType.PDF)
    .setName("Accountability Form - " + safeAssignee + " - " + tag + ".pdf");

  var assignee      = String(asset.assignee || "Employee").trim() || "Employee";
  var generatedBy   = String(generatedByEmail || "Manila IT").trim();
  var generatedDate = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "MMMM d, yyyy 'at' h:mm a"
  );

  var FROM_ALIAS  = "manila-it@fbgphilippines.com";
  var SENDER_NAME = "FBG Manila IT";
  var CC          = "manila-it@fbgphilippines.com";

  var subject =
    "Accountability Form - " + tag + " \u2022 " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  var uploadUrl = WEBAPP_BASE_URL + "?page=uploadsigned&tag=" + encodeURIComponent(tag);

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g,  "&amp;")
      .replace(/</g,  "&lt;")
      .replace(/>/g,  "&gt;")
      .replace(/"/g,  "&quot;")
      .replace(/'/g,  "&#39;");
  }

  var htmlBody =
    "<div style=\"margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;\">" +
      "<div style=\"max-width:560px;margin:28px auto;padding:0 16px;\">" +
        "<div style=\"background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);\">" +

          "<div style=\"padding:18px 20px;background:linear-gradient(135deg,#0A84FF,#2997FF);color:#fff;\">" +
            "<div style=\"font-size:13px;opacity:.92;\">Manila IT Inventory</div>" +
            "<div style=\"font-size:18px;font-weight:800;margin-top:4px;\">Accountability Form</div>" +
          "</div>" +

          "<div style=\"padding:20px;color:#f5f5f7;\">" +
            "<p style=\"margin:0 0 12px;font-size:14px;line-height:1.6;color:#d1d1d6;\">" +
              "Hi <b style=\"color:#fff;\">" + esc(assignee) + "</b>, please see your attached" +
              "<b>Accountability Form</b> for the assigned asset below." +
            "</p>" +

            "<div style=\"margin:14px 0;padding:16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;\">" +
              "<table role=\"presentation\" cellspacing=\"0\" cellpadding=\"0\" border=\"0\" width=\"100%\" style=\"border-collapse:collapse;\">" +
                "<tr>" +
                  "<td style=\"width:50%;padding:0 12px 14px 0;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Serial Tag</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(tag) + "</div>" +
                  "</td>" +
                  "<td style=\"width:50%;padding:0 0 14px 12px;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Type</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(asset.type || "-") + "</div>" +
                  "</td>" +
                "</tr>" +
                "<tr>" +
                  "<td style=\"width:50%;padding:0 12px 14px 0;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Model</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(asset.model || "-") + "</div>" +
                  "</td>" +
                  "<td style=\"width:50%;padding:0 0 14px 12px;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Serial Number</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.5;margin-top:4px;\">" + esc(asset.serialNumber || "-") + "</div>" +
                  "</td>" +
                "</tr>" +
                "<tr>" +
                  "<td colspan=\"2\" style=\"padding:0 0 14px 0;vertical-align:top;\">" +
                    "<div style=\"font-size:12px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;\">Accessories</div>" +
                    "<div style=\"font-size:15px;color:#ffffff;font-weight:700;line-height:1.6;margin-top:4px;word-break:break-word;\">" + esc(asset.accessories || "None") + "</div>" +
                  "</td>" +
                "</tr>" +
              "</table>" +
            "</div>" +

            "<div style=\"margin:0 0 14px;padding:13px 16px;background:rgba(10,132,255,0.10);border:1px solid rgba(10,132,255,0.25);border-radius:10px;\">" +
              "<div style=\"font-size:11px;color:#9a9aa0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:4px;\">Generated by</div>" +
              "<div style=\"font-size:13px;color:#ffffff;font-weight:700;\">" + esc(generatedBy) + "</div>" +
              "<div style=\"font-size:11px;color:#9a9aa0;margin-top:3px;\">" + esc(generatedDate) + "</div>" +
            "</div>" +

            "<p style=\"margin:0 0 10px;font-size:13px;color:#d1d1d6;line-height:1.5;\">" +
              "After signing, upload the signed PDF using the button below:" +
            "</p>" +
            "<div style=\"text-align:center;margin:14px 0 6px;\">" +
              "<a href=\"" + esc(uploadUrl) + "\"" +
                " style=\"display:inline-block;padding:12px 18px;border-radius:12px;" +
                "background:#ffffff;color:#000;text-decoration:none;font-weight:800;font-size:13px;\">" +
                "Upload Signed PDF" +
              "</a>" +
            "</div>" +

            "<p style=\"margin:12px 0 0;font-size:12px;color:#9a9aa0;line-height:1.4;\">" +
              "Note: The unsigned accountability form is attached to this email." +
            "</p>" +
          "</div>" +
        "</div>" +

        "<div style=\"text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;\">" +
          "&#169; " + new Date().getFullYear() + " Manila IT Inventory" +
        "</div>" +
      "</div>" +
    "</div>";

  var plainBody =
    "Hi " + assignee + ",\n\n" +
    "Please see attached your Accountability Form for the assigned asset below.\n\n" +
    "Serial Tag:    " + tag                          + "\n" +
    "Type:          " + (asset.type         || "-")  + "\n" +
    "Model:         " + (asset.model        || "-")  + "\n" +
    "Serial Number: " + (asset.serialNumber || "-")  + "\n" +
    "Accessories:   " + (asset.accessories  || "None") + "\n\n" +
    "Generated by:  " + generatedBy   + "\n" +
    "Generated on:  " + generatedDate + "\n\n" +
    "After signing, upload the signed PDF here:\n" +
    uploadUrl + "\n\n" +
    "Thanks,\n" +
    "Manila IT Team";

  try {
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody:    htmlBody,
      name:        SENDER_NAME,
      from:        FROM_ALIAS,
      replyTo:     FROM_ALIAS,
      attachments: [pdfBlob],
      cc:          CC
    });
  } catch (e) {
    console.error("_emailAccountabilityPdf_ alias send failed:", e);
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody:    htmlBody,
      name:        SENDER_NAME,
      attachments: [pdfBlob],
      cc:          CC
    });
  }

  return { ok: true, to: to, fileId: fileId };
}


/* ============================================================
 * HTML ESCAPE HELPER
 * ============================================================ */
function _escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g,  "&amp;")
    .replace(/</g,  "&lt;")
    .replace(/>/g,  "&gt;")
    .replace(/"/g,  "&quot;")
    .replace(/'/g,  "&#39;");
}


/* ============================================================
 * URL → FILE ID HELPER
 * ============================================================ */
function _extractDriveFileId_(url) {
  var s = String(url || "");
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m && m[1]) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m && m[1]) return m[1];
  return null;
}


/* ============================================================
 * UPLOAD SIGNED AF
 * ============================================================ */
function uploadSignedAccountabilityPdf(serialTag, base64Data, originalFilename) {
  try {
    var tag = String(serialTag || "").trim().replace(/^"+|"+$/g, "");
    if (!tag)          throw new Error("Missing Serial Tag");
    if (!base64Data)   throw new Error("Missing file data");
    if (!SIGNED_AF_ROOT_FOLDER_ID) throw new Error("SIGNED_AF_ROOT_FOLDER_ID is not set");

    var asset = _getMasterlistAssetBySerialTag_(tag);
    if (!asset) throw new Error("Asset not found for Serial Tag: " + tag);

    var empDisplayName = String(asset.assignee || "Employee").trim();
    var empEmail       = String(asset.email    || "").trim();

    try {
      var emp = _getEmployeeByNameOrEmail_(asset.assignee, asset.email);
      if (emp && emp.displayName) empDisplayName = emp.displayName;
      if (emp && emp.email)       empEmail       = emp.email;
    } catch (empErr) {
      console.error("uploadSignedAccountabilityPdf: employee lookup failed:", empErr);
    }

    var empFolder = _getOrCreateEmployeeSignedFolder_(empDisplayName, empEmail);

    var bytes = Utilities.base64Decode(base64Data);
    var blob  = Utilities.newBlob(
      bytes,
      MimeType.PDF,
      originalFilename || ("Signed AF - " + tag + ".pdf")
    );

    var tz      = Session.getScriptTimeZone();
    var stamp   = Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");
    var safeEmp = _safeFolderName_(empDisplayName);
    var filename = "Signed AF - " + safeEmp + " - " + tag + " - " + stamp + ".pdf";

    var file = empFolder.createFile(blob.setName(filename));
    var url  = file.getUrl();

    var writeResult = false;
var writeError  = null;
try {
  _setMasterlistSignedAf_(tag, url);
  _mlBustCache_();
  writeResult = true;
} catch (writeErr) {
  writeError = writeErr.message || String(writeErr);
  console.error('_setMasterlistSignedAf_ failed for ' + tag + ':', writeError);
}

if (!writeResult) {
  // File is saved to Drive but sheet wasn't updated — return partial success with error detail
  return {
    ok:      false,
    error:   'File uploaded to Drive but sheet update failed: ' + writeError,
    url:     url,
    fileId:  file.getId()
  };
}

var to = String(empEmail || "").trim();
    if (to) {
      sendSignedAfConfirmationEmail_({
        to:       to,
        assignee: empDisplayName || safeEmp,
        tag:      tag,
        url:      url
      });
    } else {
      console.log("[AF CONFIRM] No recipient email. Skipping confirmation email.");
    }

    return {
      ok:                 true,
      url:                url,
      fileId:             file.getId(),
      filename:           filename,
      employeeFolderId:   empFolder.getId(),
      employeeFolderName: empFolder.getName()
    };

  } catch (err) {
    return { ok: false, error: (err && err.message) || String(err) };
  }
}


/* ============================================================
 * EMAIL — SIGNED AF CONFIRMATION
 * ============================================================ */
function sendSignedAfConfirmationEmail_(opts) {
  var to       = String((opts && opts.to)       || "").trim();
  var assignee = String((opts && opts.assignee) || "Employee").trim() || "Employee";
  var tag      = String((opts && opts.tag)      || "").trim();
  var url      = String((opts && opts.url)      || "").trim();

  if (!to)  throw new Error("Missing recipient email for confirmation.");
  if (!tag) throw new Error("Missing serial tag for confirmation.");
  if (!url) throw new Error("Missing url for confirmation.");

  var FROM_ALIAS  = "manila-it@fbgphilippines.com";
  var SENDER_NAME = "FBG Manila IT";
  var CC          = "manila-it@fbgphilippines.com";
  var subject     = "Signed Accountability Form Received - " + tag;

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g,  "&amp;")
      .replace(/</g,  "&lt;")
      .replace(/>/g,  "&gt;")
      .replace(/"/g,  "&quot;")
      .replace(/'/g,  "&#39;");
  }

  var htmlBody =
    "<div style=\"margin:0;padding:0;background:#0b0b0c;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;\">" +
      "<div style=\"max-width:560px;margin:28px auto;padding:0 16px;\">" +
        "<div style=\"background:#111114;border:1px solid rgba(255,255,255,0.10);border-radius:18px;overflow:hidden;box-shadow:0 18px 40px rgba(0,0,0,.55);\">" +

          "<div style=\"padding:18px 20px;background:linear-gradient(135deg,#0A84FF,#2997FF);color:#fff;\">" +
            "<div style=\"font-size:13px;opacity:.92;\">Manila IT Inventory</div>" +
            "<div style=\"font-size:18px;font-weight:800;margin-top:4px;\">Signed Form Received</div>" +
          "</div>" +

          "<div style=\"padding:20px;color:#f5f5f7;\">" +
            "<p style=\"margin:0 0 12px;font-size:14px;line-height:1.5;color:#d1d1d6;\">" +
              "Hi <b style=\"color:#fff;\">" + esc(assignee) + "</b>, we have successfully received your " +
              "<b>signed Accountability Form</b>." +
            "</p>" +

            "<div style=\"margin:14px 0;padding:14px 16px;background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.10);border-radius:14px;\">" +
              "<div style=\"font-size:13px;color:#d1d1d6;line-height:1.6;\">" +
                "<b style=\"color:#fff;\">Serial Tag:</b> " + esc(tag) +
              "</div>" +
            "</div>" +

            "<p style=\"margin:0 0 10px;font-size:13px;color:#d1d1d6;line-height:1.5;\">" +
              "You can view the uploaded file here:" +
            "</p>" +
            "<div style=\"text-align:center;margin:14px 0 6px;\">" +
              "<a href=\"" + esc(url) + "\"" +
                " style=\"display:inline-block;padding:12px 18px;border-radius:12px;" +
                "background:#ffffff;color:#000;text-decoration:none;font-weight:800;font-size:13px;\">" +
                "View Signed PDF" +
              "</a>" +
            "</div>" +

            "<p style=\"margin:12px 0 0;font-size:12px;color:#9a9aa0;line-height:1.4;\">" +
              "No further action is required." +
            "</p>" +
          "</div>" +
        "</div>" +

        "<div style=\"text-align:center;margin-top:12px;font-size:11px;color:#8e8e93;\">" +
          "&#169; " + new Date().getFullYear() + " Manila IT Inventory" +
        "</div>" +
      "</div>" +
    "</div>";

  var plainBody =
    "Hi " + assignee + ",\n\n" +
    "We have successfully received your signed Accountability Form.\n\n" +
    "Serial Tag: " + tag + "\n\n" +
    "View Signed PDF:\n" +
    url + "\n\n" +
    "No further action is required.\n\n" +
    "Thanks,\n" +
    "Manila IT Team";

  try {
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody: htmlBody,
      name:     SENDER_NAME,
      from:     FROM_ALIAS,
      replyTo:  FROM_ALIAS,
      cc:       CC
    });
  } catch (e) {
    console.error("sendSignedAfConfirmationEmail_ alias send failed:", e);
    GmailApp.sendEmail(to, subject, plainBody, {
      htmlBody: htmlBody,
      name:     SENDER_NAME,
      cc:       CC
    });
  }

  console.log("[AF CONFIRM] Email sent OK.");
}


/* ============================================================
 * FOLDER HELPERS
 * ============================================================ */
function _normKey_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function _safeFolderName_(name) {
  return String(name || "Employee")
    .trim()
    .replace(/[\/\\:*?"<>|]+/g, "")
    .replace(/\s+/g, " ")
    .slice(0, 120) || "Employee";
}

function _getOrCreateEmployeeSignedFolder_(displayName, email) {
  var root      = DriveApp.getFolderById(SIGNED_AF_ROOT_FOLDER_ID);
  var safeName  = _safeFolderName_(displayName || "Employee");
  var safeEmail = String(email || "").trim().toLowerCase();

  var it = root.getFolders();
  while (it.hasNext()) {
    var f     = it.next();
    var fname = f.getName();

    if (safeEmail && _normKey_(fname).includes(_normKey_(safeEmail))) return f;
    if (!safeEmail && _normKey_(fname).indexOf(_normKey_(safeName)) === 0) return f;
    if (_normKey_(fname).includes(_normKey_(safeName))) return f;
  }

  var folderName = safeEmail
    ? (safeName + " (" + safeEmail + ")")
    : safeName;

  return root.createFolder(folderName);
}