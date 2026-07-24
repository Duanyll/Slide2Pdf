interface SaveFilePickerOptions {
  suggestedName?: string;
  types?: Array<{
    description: string;
    accept: Record<string, string[]>;
  }>;
}

interface FilePickerWindow extends Window {
  showSaveFilePicker?: (
    options?: SaveFilePickerOptions,
  ) => Promise<FileSystemFileHandle>;
}

export interface SavePdfResult {
  method: "file" | "download";
  fileName: string;
}

const fileHandles = new Map<string, FileSystemFileHandle>();

export async function savePdf(
  data: Uint8Array,
  slideKey: string,
  suggestedName: string,
  forceNewPath: boolean,
): Promise<SavePdfResult> {
  const hasLocalCompanion =
    window.location.hostname === "localhost" && window.location.port === "3000";
  if (hasLocalCompanion) {
    return saveWithLocalCompanion(
      data,
      slideKey,
      suggestedName,
      forceNewPath,
    );
  }

  const pickerWindow = window as FilePickerWindow;
  if (pickerWindow.showSaveFilePicker) {
    let handle = forceNewPath ? undefined : fileHandles.get(slideKey);
    if (!handle) {
      handle = await pickerWindow.showSaveFilePicker({
        suggestedName,
        types: [
          {
            description: "PDF document",
            accept: { "application/pdf": [".pdf"] },
          },
        ],
      });
      fileHandles.set(slideKey, handle);
    }

    try {
      const writable = await handle.createWritable();
      await writable.write(toArrayBuffer(data));
      await writable.close();
      return { method: "file", fileName: handle.name };
    } catch (error) {
      fileHandles.delete(slideKey);
      throw error;
    }
  }

  downloadPdf(data, suggestedName);
  return { method: "download", fileName: suggestedName };
}

async function saveWithLocalCompanion(
  data: Uint8Array,
  slideKey: string,
  suggestedName: string,
  forceNewPath: boolean,
): Promise<SavePdfResult> {
  const response = await fetch("/slide2pdf/save", {
    method: "POST",
    headers: {
      "Content-Type": "application/pdf",
      "X-Slide2Pdf-Key": encodeURIComponent(slideKey),
      "X-Slide2Pdf-Name": encodeURIComponent(suggestedName),
      "X-Slide2Pdf-New-Path": forceNewPath ? "1" : "0",
    },
    body: toArrayBuffer(data),
  });
  const result = (await response.json()) as {
    cancelled?: boolean;
    error?: string;
    fileName?: string;
  };

  if (response.status === 409 && result.cancelled) {
    throw new DOMException("Save cancelled.", "AbortError");
  }
  if (!response.ok || !result.fileName) {
    throw new Error(
      result.error || "The local save helper could not write the PDF.",
    );
  }

  return { method: "file", fileName: result.fileName };
}

function downloadPdf(data: Uint8Array, fileName: string): void {
  const url = URL.createObjectURL(
    new Blob([toArrayBuffer(data)], { type: "application/pdf" }),
  );
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.hidden = true;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function toArrayBuffer(data: Uint8Array): ArrayBuffer {
  return data.buffer.slice(
    data.byteOffset,
    data.byteOffset + data.byteLength,
  ) as ArrayBuffer;
}
