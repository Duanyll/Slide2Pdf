const SLICE_SIZE = 4 * 1024 * 1024;

export async function getPresentationPdf(): Promise<Uint8Array> {
  const file = await openPdfFile();

  try {
    const slices: Uint8Array[] = [];
    for (let index = 0; index < file.sliceCount; index += 1) {
      slices.push(await getSlice(file, index));
    }

    const result = new Uint8Array(
      slices.reduce((total, slice) => total + slice.byteLength, 0),
    );
    let offset = 0;
    for (const slice of slices) {
      result.set(slice, offset);
      offset += slice.byteLength;
    }

    return result;
  } finally {
    await closeFile(file);
  }
}

function openPdfFile(): Promise<Office.File> {
  return runOfficeAsync((callback) => {
    Office.context.document.getFileAsync(
      Office.FileType.Pdf,
      { sliceSize: SLICE_SIZE },
      callback,
    );
  });
}

async function getSlice(file: Office.File, index: number): Promise<Uint8Array> {
  const slice = await runOfficeAsync<Office.Slice>((callback) => {
    file.getSliceAsync(index, callback);
  });
  const data = slice.data as ArrayBuffer | ArrayLike<number>;

  return data instanceof ArrayBuffer
    ? new Uint8Array(data)
    : Uint8Array.from(data);
}

function closeFile(file: Office.File): Promise<void> {
  return runOfficeAsync((callback) => file.closeAsync(callback));
}

function runOfficeAsync<T>(
  operation: (callback: (result: Office.AsyncResult<T>) => void) => void,
): Promise<T> {
  return new Promise((resolve, reject) => {
    operation((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}
