import "fake-indexeddb/auto";

import { Buffer as NodeBuffer } from "node:buffer";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";

import {
  seedBareGitProject,
  startGitHttpServer,
} from "./helpers/gitHttpServer";

const temporaryDirectories: string[] = [];
const serverClosers: Array<() => Promise<void>> = [];

afterEach(async () => {
  await Promise.all(serverClosers.splice(0).map((close) => close()));
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { force: true, recursive: true });
  }
});

describe("pushPdfFromBrowser", () => {
  it("pushes when the host does not provide Node Buffer", async () => {
    const root = makeTemporaryDirectory("slide2pdf-browser-remote-");
    const seed = makeTemporaryDirectory("slide2pdf-browser-seed-");
    const remote = seedBareGitProject(root, seed);
    const server = await startGitHttpServer(root);
    serverClosers.push(server.close);

    const nativeBuffer = globalThis.Buffer;
    Reflect.deleteProperty(globalThis, "Buffer");
    try {
      const { pushPdfFromBrowser } = await import(
        "../src/overleaf/browserGitClient"
      );
      const result = await pushPdfFromBrowser({
        data: new Uint8Array([0x25, 0x50, 0x44, 0x46]),
        filePath: "figures/browser.pdf",
        remoteUrl: `${server.origin}/project.git`,
        token: "test-token",
      });

      expect(result.changed).toBe(true);
      const checkout = makeTemporaryDirectory("slide2pdf-browser-checkout-");
      execFileSync("git", ["clone", remote, checkout]);
      expect(fs.existsSync(path.join(checkout, "main.tex"))).toBe(true);
      expect(
        fs.readFileSync(path.join(checkout, "figures", "browser.pdf")),
      ).toEqual(NodeBuffer.from([0x25, 0x50, 0x44, 0x46]));
    } finally {
      globalThis.Buffer = nativeBuffer;
    }
  });
});

function makeTemporaryDirectory(prefix: string): string {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  temporaryDirectories.push(directory);
  return directory;
}
