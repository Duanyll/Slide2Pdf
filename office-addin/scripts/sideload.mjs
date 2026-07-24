import { copyFile, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = dirname(fileURLToPath(import.meta.url));
const manifestName = process.argv[2] || "manifest.xml";
const source = join(scriptDirectory, "..", manifestName);
const catalog = join(
  homedir(),
  "Library",
  "Containers",
  "com.microsoft.Powerpoint",
  "Data",
  "Documents",
  "wef",
);
const destination = join(catalog, "Slide2Pdf.xml");

await mkdir(catalog, { recursive: true });
await copyFile(source, destination);

console.log(`Sideloaded manifest to ${destination}`);
console.log("Restart PowerPoint, then open Home > Slide2Pdf > Export PDF.");
