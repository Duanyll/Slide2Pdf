const DOWNLOAD_SEQUENCE_PREFIX = "slide2pdf.download-sequence.";
const fallbackSequences = new Map<string, number>();

export function downloadPdf(data: Uint8Array, fileName: string): string {
  const sequencedFileName = nextDownloadFileName(fileName);
  const url = URL.createObjectURL(
    new Blob([toArrayBuffer(data)], { type: "application/pdf" }),
  );
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = sequencedFileName;
  anchor.hidden = true;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  return sequencedFileName;
}

function nextDownloadFileName(fileName: string): string {
  const sequence = incrementDownloadSequence(fileName);
  const extensionIndex = fileName.lastIndexOf(".");

  if (extensionIndex <= 0) {
    return `${fileName}_${sequence}`;
  }

  return `${fileName.slice(0, extensionIndex)}_${sequence}${fileName.slice(extensionIndex)}`;
}

function incrementDownloadSequence(fileName: string): number {
  try {
    const key = `${DOWNLOAD_SEQUENCE_PREFIX}${fileName}`;
    const storedSequence = Number.parseInt(
      window.localStorage.getItem(key) ?? "0",
      10,
    );
    const previousSequence =
      Number.isSafeInteger(storedSequence) && storedSequence >= 0
        ? storedSequence
        : 0;
    const nextSequence = previousSequence + 1;
    window.localStorage.setItem(key, String(nextSequence));
    return nextSequence;
  } catch {
    const nextSequence = (fallbackSequences.get(fileName) ?? 0) + 1;
    fallbackSequences.set(fileName, nextSequence);
    return nextSequence;
  }
}

function toArrayBuffer(data: Uint8Array): ArrayBuffer {
  return data.buffer.slice(
    data.byteOffset,
    data.byteOffset + data.byteLength,
  ) as ArrayBuffer;
}
