import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterAll, describe, expect, it } from "vitest";
import * as git from "isomorphic-git";
import http from "isomorphic-git/http/node";

import { OverleafGitClient } from "../src/overleaf/overleafGitClient";
import { parseOverleafTarget } from "../src/overleaf/overleafTarget";

const enabled =
  process.env.SLIDE2PDF_LIVE_OVERLEAF === "1" &&
  Boolean(process.env.OVERLEAF_GIT_REPO) &&
  Boolean(process.env.OVERLEAF_GIT_KEY);
const temporaryDirectories: string[] = [];

afterAll(() => {
  for (const directory of temporaryDirectories) {
    fs.rmSync(directory, { force: true, recursive: true });
  }
});

describe.runIf(enabled)("live Overleaf Git Bridge", () => {
  it(
    "pushes a PDF and restores the original project tree",
    async () => {
      const remoteUrl = parseOverleafTarget(
        process.env.OVERLEAF_GIT_REPO ?? "",
        "slide2pdf-test.pdf",
      ).remoteUrl;
      const token = process.env.OVERLEAF_GIT_KEY ?? "";
      const dir = makeTemporaryDirectory();
      const onAuth = () => ({ username: "git", password: token });

      await git.clone({
        fs,
        http,
        dir,
        url: remoteUrl,
        singleBranch: true,
        depth: 1,
        noTags: true,
        onAuth,
      });
      const initialOid = await git.resolveRef({ fs, dir, ref: "HEAD" });
      const initialTree = (await git.readCommit({ fs, dir, oid: initialOid }))
        .commit.tree;
      const filePath = `slide2pdf-feasibility/implementation-${Date.now()}.pdf`;

      try {
        const client = new OverleafGitClient({ dir, fs, http });
        const result = await client.pushPdf({
          data: new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37]),
          filePath,
          remoteUrl,
          token,
        });
        expect(result.changed).toBe(true);
      } finally {
        const cleanupDir = makeTemporaryDirectory();
        await git.clone({
          fs,
          http,
          dir: cleanupDir,
          url: remoteUrl,
          singleBranch: true,
          depth: 1,
          noTags: true,
          onAuth,
        });
        if (fs.existsSync(path.join(cleanupDir, filePath))) {
          fs.unlinkSync(path.join(cleanupDir, filePath));
          await git.remove({ fs, dir: cleanupDir, filepath: filePath });
          await git.commit({
            fs,
            dir: cleanupDir,
            message: `Remove ${filePath} after Slide2Pdf test`,
            author: { name: "Slide2Pdf", email: "slide2pdf@localhost" },
          });
          const branch = await git.currentBranch({ fs, dir: cleanupDir });
          if (!branch) throw new Error("Live test clone has no current branch.");
          const result = await git.push({
            fs,
            http,
            dir: cleanupDir,
            ref: branch,
            force: false,
            onAuth,
          });
          if (!result.ok) {
            throw new Error(result.error ?? "Live cleanup push failed.");
          }
        }
      }

      const verificationDir = makeTemporaryDirectory();
      await git.clone({
        fs,
        http,
        dir: verificationDir,
        url: remoteUrl,
        singleBranch: true,
        depth: 1,
        noTags: true,
        onAuth,
      });
      const finalOid = await git.resolveRef({
        fs,
        dir: verificationDir,
        ref: "HEAD",
      });
      const finalTree = (
        await git.readCommit({ fs, dir: verificationDir, oid: finalOid })
      ).commit.tree;
      expect(finalTree).toBe(initialTree);
      expect(fs.existsSync(path.join(verificationDir, filePath))).toBe(false);
    },
    120_000,
  );
});

function makeTemporaryDirectory(): string {
  const directory = fs.mkdtempSync(
    path.join(os.tmpdir(), "slide2pdf-live-overleaf-"),
  );
  temporaryDirectories.push(directory);
  return directory;
}
