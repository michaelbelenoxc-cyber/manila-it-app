
function canViewSignedAf(email) {
  return canAccessPage("signedaf", email);
}

function requireSignedAfView_(email) {
  if (!canViewSignedAf(email)) {
    throw new Error("You do not have permission to access Signed Accountability Forms.");
  }
}

/* =========================================================
 * DRIVE HELPERS
 * =======================================================*/
function getSignedAfRootFolder_() {
  if (!SIGNED_AF_FOLDER_ID) {
    throw new Error("SIGNED_AF_FOLDER_ID is not configured.");
  }

  try {
    return DriveApp.getFolderById(SIGNED_AF_FOLDER_ID);
  } catch (e) {
    throw new Error("Signed AF root folder not found or inaccessible.");
  }
}

function getSignedAfChildFolderMap_() {
  const root = getSignedAfRootFolder_();
  const foldersIt = root.getFolders();
  const map = {};

  while (foldersIt.hasNext()) {
    const f = foldersIt.next();
    map[String(f.getId())] = f;
  }

  return map;
}

function getSafeSignedAfFolder_(folderId) {
  const id = String(folderId || "").trim();

  if (!id || id === "ROOT") {
    return getSignedAfRootFolder_();
  }

  const childMap = getSignedAfChildFolderMap_();
  const folder = childMap[id];

  if (!folder) {
    throw new Error("Invalid Signed AF folder.");
  }

  return folder;
}

function toIsoSafe_(value) {
  try {
    const d = value instanceof Date ? value : new Date(value);
    if (isNaN(d.getTime())) return "";
    return d.toISOString();
  } catch (e) {
    return "";
  }
}

/* =========================================================
 * INDEX
 * Return folders + root files count
 * =======================================================*/
function getSignedAfIndex() {
  requireSignedAfView_();

  const root = getSignedAfRootFolder_();

  const folders = [];
  const foldersIt = root.getFolders();

  while (foldersIt.hasNext()) {
    const f = foldersIt.next();
    folders.push({
      id: f.getId(),
      name: f.getName()
    });
  }

  folders.sort((a, b) => {
    const aName = String(a.name || "").toLowerCase();
    const bName = String(b.name || "").toLowerCase();
    return aName.localeCompare(bName);
  });

  let rootCount = 0;
  const filesIt = root.getFiles();
  while (filesIt.hasNext()) {
    filesIt.next();
    rootCount++;
  }

  return {
    folders,
    rootCount
  };
}

/* =========================================================
 * FILE LIST
 * List files in a given folder
 * folderId === "ROOT" -> files directly in main folder
 * =======================================================*/
function getSignedAfFilesIn(folderId) {
  requireSignedAfView_();

  const folder = getSafeSignedAfFolder_(folderId);
  const filesIt = folder.getFiles();
  const files = [];

  while (filesIt.hasNext()) {
    const f = filesIt.next();
    const updated = toIsoSafe_(f.getLastUpdated ? f.getLastUpdated() : "");

    files.push({
      id: f.getId(),
      name: f.getName(),
      mimeType: f.getMimeType(),
      updated: updated,
      previewUrl: "https://drive.google.com/file/d/" + f.getId() + "/preview",
      openUrl: "https://drive.google.com/file/d/" + f.getId() + "/view"
    });
  }

  files.sort((a, b) => {
    const aUpdated = String(a.updated || "");
    const bUpdated = String(b.updated || "");
    return bUpdated.localeCompare(aUpdated);
  });

  return files;
}