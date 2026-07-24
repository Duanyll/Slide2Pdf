const { execFile } = require("node:child_process");
const { writeFile } = require("node:fs/promises");
const path = require("node:path");
const { promisify } = require("node:util");

const execFileAsync = promisify(execFile);
const MAX_PDF_SIZE = 512 * 1024 * 1024;
const SAVE_SCRIPT = `
set defaultName to system attribute "SLIDE2PDF_DEFAULT_NAME"
set chosenFile to choose file name with prompt "Save current slide as PDF" default name defaultName
return POSIX path of chosenFile
`;

class LocalPdfSaver {
  constructor({ choosePath = chooseSavePath, writeFile: write = writeFile } = {}) {
    this.choosePath = choosePath;
    this.writeFile = write;
    this.pathsBySlide = new Map();
  }

  async save({ slideKey, suggestedName, forceNewPath, data }) {
    let outputPath = forceNewPath ? undefined : this.pathsBySlide.get(slideKey);
    if (!outputPath) {
      outputPath = await this.choosePath(suggestedName);
      if (!outputPath.toLowerCase().endsWith(".pdf")) {
        outputPath += ".pdf";
      }
    }

    try {
      await this.writeFile(outputPath, data);
      this.pathsBySlide.set(slideKey, outputPath);
      return { fileName: path.basename(outputPath) };
    } catch (error) {
      this.pathsBySlide.delete(slideKey);
      throw error;
    }
  }
}

function createLocalSaveMiddleware() {
  const saver = new LocalPdfSaver();

  return async function localSaveMiddleware(request, response, next) {
    if (request.method !== "POST" || request.url !== "/slide2pdf/save") {
      next();
      return;
    }

    const origin = request.headers.origin;
    if (origin && origin !== "https://localhost:3000") {
      sendJson(response, 403, {
        error: "The save request did not come from Slide2Pdf.",
      });
      return;
    }

    try {
      const slideKey = decodeHeader(request.headers["x-slide2pdf-key"]);
      const suggestedName = decodeHeader(
        request.headers["x-slide2pdf-name"],
      );
      if (!slideKey || !suggestedName) {
        sendJson(response, 400, {
          error: "The slide key or file name is missing.",
        });
        return;
      }

      const chunks = [];
      let size = 0;
      for await (const chunk of request) {
        size += chunk.length;
        if (size > MAX_PDF_SIZE) {
          sendJson(response, 413, {
            error: "The generated PDF is larger than 512 MB.",
          });
          return;
        }
        chunks.push(chunk);
      }

      const result = await saver.save({
        slideKey,
        suggestedName,
        forceNewPath: request.headers["x-slide2pdf-new-path"] === "1",
        data: Buffer.concat(chunks),
      });
      sendJson(response, 200, result);
    } catch (error) {
      if (error?.name === "AbortError") {
        sendJson(response, 409, { cancelled: true });
      } else {
        sendJson(response, 500, {
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }
  };
}

async function chooseSavePath(suggestedName) {
  try {
    const { stdout } = await execFileAsync(
      "/usr/bin/osascript",
      ["-e", SAVE_SCRIPT],
      { env: { ...process.env, SLIDE2PDF_DEFAULT_NAME: suggestedName } },
    );
    return stdout.trim();
  } catch (error) {
    if (
      error?.stderr?.includes("-128") ||
      error?.stderr?.includes("User canceled")
    ) {
      const cancellation = new Error("Save cancelled.");
      cancellation.name = "AbortError";
      throw cancellation;
    }
    throw error;
  }
}

function decodeHeader(value) {
  const raw = Array.isArray(value) ? value[0] : value;
  return raw ? decodeURIComponent(raw) : "";
}

function sendJson(response, status, value) {
  response.statusCode = status;
  response.setHeader("Content-Type", "application/json; charset=utf-8");
  response.end(JSON.stringify(value));
}

module.exports = { LocalPdfSaver, createLocalSaveMiddleware };
