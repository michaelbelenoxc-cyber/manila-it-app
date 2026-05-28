/* =========================================================
   SLACK NOTIFICATIONS
=========================================================*/
function sendSlackNotification_(message, blocks, color) {
  try {
    const url = PropertiesService.getScriptProperties()
      .getProperty("SLACK_WEBHOOK_URL");
    if (!url) return false;

    let payload;

    if (color && blocks) {
      payload = {
        attachments: [
          {
            color: color,
            blocks: blocks
          }
        ]
      };
    } else if (blocks) {
      payload = { blocks };
    } else {
      payload = { text: String(message || "") };
    }

    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      console.error("Slack returned non-200:", res.getResponseCode(), res.getContentText());
      return false;
    }

    return true;
  } catch (e) {
    console.error("sendSlackNotification_ failed:", e);
    return false;
  }
}


function setSlackWebhookUrl() {
  PropertiesService.getScriptProperties()
    .setProperty("SLACK_WEBHOOK_URL", "https://hooks.slack.com/services/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");
}


/* ========= ADD USER TO SHEET ========= */
function addUserToSheet(name, email, status, role, group) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName("Users") || ss.insertSheet("Users");

  ensureUsersHeaders_(sh);

  sh.appendRow([
    name,
    email,
    new Date(),
    status || "Active",
    role   || "Staff",
    group  || "Morning",
    ""
  ]);

  invalidateUsersCache_();

  // ── Slack notification ──────────────────────────────
  try {
    const actor = String(Session.getActiveUser().getEmail() || "unknown").trim();

    sendSlackNotification_(null, [
      {
        type: "section",
        text: {
          type: "mrkdwn",
          text: `*👤 New User Added to Manila IT Inventory*`
        }
      },
      {
        type: "section",
        fields: [
          { type: "mrkdwn", text: `*Name:*\n${name}`                  },
          { type: "mrkdwn", text: `*Email:*\n${email}`                },
          { type: "mrkdwn", text: `*Role:*\n${role   || "Staff"}`     },
          { type: "mrkdwn", text: `*Group:*\n${group  || "Morning"}`  },
          { type: "mrkdwn", text: `*Status:*\n${status || "Active"}`  },
          { type: "mrkdwn", text: `*Added by:*\n${actor}`             }
        ]
      },
      {
        type: "actions",
        elements: [
          {
            type: "button",
            text: { type: "plain_text", text: "Open Users Page", emoji: true },
            url: ScriptApp.getService().getUrl() + "?page=users",
            style: "primary"
          }
        ]
      },
      {
        type: "context",
        elements: [
          {
            type: "mrkdwn",
            text: `🕐 ${new Date().toLocaleString("en-PH", { timeZone: "Asia/Manila" })} PHT`
          }
        ]
      }
    ], "#22c55e");

  } catch (slackErr) {
    console.error("Slack notify failed (non-blocking):", slackErr);
  }
  // ────────────────────────────────────────────────────

  return true;
}