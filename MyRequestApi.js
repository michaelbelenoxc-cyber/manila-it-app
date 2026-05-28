function canViewMyRequests(email) {
  return canAccessPage("myrequest", email);
}

function canEditMyRequests(email) {
  return canViewMyRequests(email) && canDoAction("request.submit", email);
}

function requireMyRequestsView_(email) {
  if (!canViewMyRequests(email)) {
    throw new Error("You do not have permission to access My Requests.");
  }
}

function requireMyRequestsEdit_(email) {
  requireMyRequestsView_(email);

  if (!canDoAction("request.submit", email)) {
    throw new Error("You do not have permission to edit or delete your requests.");
  }
}


function deleteRequestById(requestId, callerEmail) {
  requestId = normRequest_(requestId);
  if (!requestId) return { ok: false, error: "missing_id" };

  const actorEmail = normalizeEmail_(getRequestActorEmail_(callerEmail));
  if (!actorEmail) return { ok: false, error: "missing_email_param" };

  try {
    requireMyRequestsEdit_(actorEmail);
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }

  return lockRequests_(() => {
    const sh = getRequestsSheet_();
    const values = sh.getDataRange().getValues();
    if (values.length < 2) return { ok: false, error: "empty" };

    const hm = getRequestHeaderMap_(values[0]);
    if (hm.id === -1) return { ok: false, error: "missing_id_header" };
    if (hm.email === -1) return { ok: false, error: "missing_email_header" };

    const rowIndex = values.findIndex((row, i) =>
      i > 0 && normRequest_(row[hm.id]) === requestId
    );
    if (rowIndex < 0) return { ok: false, error: "not_found" };

    const rowEmail = normalizeEmail_(values[rowIndex][hm.email]);
    if (rowEmail !== actorEmail) {
      return {
        ok: false,
        error: "not_owner",
        debug: {
          actorEmail: actorEmail,
          rowEmail: rowEmail,
          requestId: requestId
        }
      };
    }

    const currentStatus = hm.status > -1 ? normRequest_(values[rowIndex][hm.status]) : "";
    const status = lowerRequest_(currentStatus);

    if (status === "approved" || status === "rejected") {
      return { ok: false, error: "locked_status" };
    }

    sh.deleteRow(rowIndex + 1);
    return { ok: true };
  });
}
