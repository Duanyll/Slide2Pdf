export function getPresentationBaseName(
  presentationTitle: string,
  documentUrl: string,
): string {
  const documentName = getLastPathSegment(documentUrl);
  return sanitizeFileName(documentName || presentationTitle || "Presentation");
}

function getLastPathSegment(documentUrl: string): string {
  if (!documentUrl) {
    return "";
  }

  let path: string;
  try {
    path = new URL(documentUrl).pathname;
  } catch {
    path = documentUrl.split(/[?#]/, 1)[0];
  }

  const segment = path.replaceAll("\\", "/").split("/").filter(Boolean).at(-1);
  if (!segment) {
    return "";
  }

  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function sanitizeFileName(fileName: string): string {
  return fileName
    .replace(/[\u0000-\u001f\\/:*?"<>|]/g, "_")
    .replace(/\.(?:ppt|pptx|pptm|pps|ppsx|ppsm|pot|potx|potm)$/i, "");
}
