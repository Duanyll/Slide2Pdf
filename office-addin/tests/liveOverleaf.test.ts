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
      let added = false;

      try {
        const client = new OverleafGitClient({ dir, fs, http });
        const result = await client.pushPdf({
          data: new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x37]),
          filePath,
          remoteUrl,
          token,
        });
        added = result.changed;
        expect(result.changed).toBe(true);
      } finally {
        if (added) {
          fs.unlinkSync(path.join(dir, filePath));
          await git.remove({ fs, dir, filepath: filePath });
          await git.commit({
            fs,
            dir,
            message: `Remove ${filePath} after Slide2Pdf test`,
            author: { name: "Slide2Pdf", email: "slide2pdf@localhost" },
          });
          const branch = await git.currentBranch({ fs, dir });
          if (!branch) throw new Error("Live test clone has no current branch.");
          const result = await git.push({
            fs,
            http,
            dir,
            ref: branch,
            force: false,
            onAuth,
          });
          if (!result.ok) {
            throw new Error(result.error ?? "Live cleanup push failed.");
          }
        }
      }

      const finalOid = await git.resolveRef({ fs, dir, ref: "HEAD" });
      const finalTree = (await git.readCommit({ fs, dir, oid: finalOid })).commit
        .tree;
      expect(finalTree).toBe(initialTree);
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
