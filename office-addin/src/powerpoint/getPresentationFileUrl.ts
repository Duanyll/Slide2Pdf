import { runOfficeAsync } from "./officeAsync";

export async function getPresentationFileUrl(): Promise<string> {
  try {
    const properties = await runOfficeAsync<Office.FileProperties>((callback) => {
      Office.context.document.getFilePropertiesAsync(callback);
    });
    return properties.url || "";
  } catch {
    return Office.context.document.url || "";
  }
}
