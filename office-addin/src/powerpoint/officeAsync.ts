export function runOfficeAsync<T>(
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
