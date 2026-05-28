function serveImage(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const mime = blob.getContentType() || '';
    if (!mime.startsWith('image/')) return null;
    return { mime: mime, b64: Utilities.base64Encode(blob.getBytes()) };
  } catch(e) {
    console.warn('[serveImage] Failed:', fileId, e.message);
    return null;
  }
}