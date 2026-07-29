import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { afterEach, describe, expect, it } from "vitest";
import http from "isomorphic-git/http/node";

import { OverleafGitClient } from "../src/overleaf/overleafGitClient";
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

describe("OverleafGitClient", () => {
  it("pushes a binary PDF without deleting the existing project tree", async () => {
    const fixture = createRemoteFixture();
    const server = await startGitHttpServer(fixture.root);
    serverClosers.push(server.close);
    const workingDirectory = makeTemporaryDirectory("slide2pdf-git-work-");
    const client = new OverleafGitClient({
      dir: workingDirectory,
      fs,
      http,
    });
    const progress: string[] = [];

    const result = await client.pushPdf({
      data: new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d]),
      filePath: "figures/result.pdf",
      remoteUrl: `${server.origin}/project.git`,
      token: "test-token",
      onProgress: (step) => progress.push(step),
    });

    expect(result).toMatchObject({ changed: true, branch: "master" });
    expect(progress.at(-1)).toBe("verifying");

    const checkout = makeTemporaryDirectory("slide2pdf-git-checkout-");
    execFileSync("git", ["clone", fixture.remote, checkout]);
    expect(fs.readFileSync(path.join(checkout, "main.tex"), "utf8")).toBe(
      "original project\n",
    );
    expect(
      fs.readFileSync(path.join(checkout, "figures", "result.pdf")),
    ).toEqual(Buffer.from([0x25, 0x50, 0x44, 0x46, 0x2d]));
  });

  it("pulls remote edits before updating a cached project", async () => {
    const fixture = createRemoteFixture();
    const server = await startGitHttpServer(fixture.root);
    serverClosers.push(server.close);
    const workingDirectory = makeTemporaryDirectory("slide2pdf-git-work-");
    const client = new OverleafGitClient({ dir: workingDirectory, fs, http });
    const remoteUrl = `${server.origin}/project.git`;

    await client.pushPdf({
      data: new Uint8Array([1]),
      filePath: "figures/result.pdf",
      remoteUrl,
      token: "test-token",
    });

    const editor = makeTemporaryDirectory("slide2pdf-git-editor-");
    execFileSync("git", ["clone", fixture.remote, editor]);
    execFileSync("git", ["-C", editor, "config", "user.name", "Web Editor"]);
    execFileSync("git", ["-C", editor, "config", "user.email", "web@example.com"]);
    fs.writeFileSync(path.join(editor, "references.bib"), "remote edit\n");
    execFileSync("git", ["-C", editor, "add", "references.bib"]);
    execFileSync("git", ["-C", editor, "commit", "-m", "Edit from Overleaf"]);
    execFileSync("git", ["-C", editor, "push", "origin", "master"]);

    await client.pushPdf({
      data: new Uint8Array([2]),
      filePath: "figures/result.pdf",
      remoteUrl,
      token: "test-token",
    });

    const checkout = makeTemporaryDirectory("slide2pdf-git-checkout-");
    execFileSync("git", ["clone", fixture.remote, checkout]);
    expect(fs.readFileSync(path.join(checkout, "references.bib"), "utf8")).toBe(
      "remote edit\n",
    );
    expect(
      fs.readFileSync(path.join(checkout, "figures", "result.pdf")),
    ).toEqual(Buffer.from([2]));
  });

  it("restores a clean base after a non-fast-forward push is rejected", async () => {
    const fixture = createRemoteFixture();
    const editor = makeTemporaryDirectory("slide2pdf-git-editor-");
    execFileSync("git", ["clone", fixture.remote, editor]);
    execFileSync("git", ["-C", editor, "config", "user.name", "Web Editor"]);
    execFileSync("git", ["-C", editor, "config", "user.email", "web@example.com"]);
    fs.writeFileSync(path.join(editor, "remote-change.tex"), "remote wins\n");
    execFileSync("git", ["-C", editor, "add", "remote-change.tex"]);
    execFileSync("git", ["-C", editor, "commit", "-m", "Concurrent edit"]);

    let concurrentPushPending = true;
    const server = await startGitHttpServer(fixture.root, {
      beforeRequest: (_method, url) => {
        if (
          concurrentPushPending &&
          url.includes("info/refs?service=git-receive-pack")
        ) {
          concurrentPushPending = false;
          execFileSync("git", ["-C", editor, "push", "origin", "master"]);
        }
      },
    });
    serverClosers.push(server.close);
    const workingDirectory = makeTemporaryDirectory("slide2pdf-git-work-");
    const client = new OverleafGitClient({ dir: workingDirectory, fs, http });
    const options = {
      data: new Uint8Array([9]),
      filePath: "figures/result.pdf",
      remoteUrl: `${server.origin}/project.git`,
      token: "test-token",
    };

    await expect(client.pushPdf(options)).rejects.toThrow();
    await expect(client.pushPdf(options)).resolves.toMatchObject({
      changed: true,
    });

    const checkout = makeTemporaryDirectory("slide2pdf-git-checkout-");
    execFileSync("git", ["clone", fixture.remote, checkout]);
    expect(fs.readFileSync(path.join(checkout, "remote-change.tex"), "utf8")).toBe(
      "remote wins\n",
    );
    expect(
      fs.readFileSync(path.join(checkout, "figures", "result.pdf")),
    ).toEqual(Buffer.from([9]));
  });

  it("rejects PDFs larger than the default Overleaf file limit before connecting", async () => {
    const workingDirectory = makeTemporaryDirectory("slide2pdf-git-work-");
    const client = new OverleafGitClient({ dir: workingDirectory, fs, http });

    await expect(
      client.pushPdf({
        data: new Uint8Array(50 * 1024 * 1024 + 1),
        filePath: "figures/too-large.pdf",
        remoteUrl: "https://127.0.0.1:1/git/project",
        token: "test-token",
      }),
    ).rejects.toThrow("50 MiB");
  });

  it("refuses to reuse a cached client for a different remote", async () => {
    const fixture = createRemoteFixture();
    const server = await startGitHttpServer(fixture.root);
    serverClosers.push(server.close);
    const workingDirectory = makeTemporaryDirectory("slide2pdf-git-work-");
    const client = new OverleafGitClient({ dir: workingDirectory, fs, http });

    await client.pushPdf({
      data: new Uint8Array([1]),
      filePath: "figures/result.pdf",
      remoteUrl: `${server.origin}/project.git`,
      token: "test-token",
    });

    await expect(
      client.pushPdf({
        data: new Uint8Array([2]),
        filePath: "figures/result.pdf",
        remoteUrl: "https://overleaf.example/git/another-project",
        token: "another-token",
      }),
    ).rejects.toThrow("不同的 Git 仓库");
  });
});

function createRemoteFixture(): { root: string; remote: string } {
  const root = makeTemporaryDirectory("slide2pdf-git-remote-");
  const seed = makeTemporaryDirectory("slide2pdf-git-seed-");
  const remote = seedBareGitProject(root, seed);

  return { root, remote };
}

function makeTemporaryDirectory(prefix: string): string {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), prefix));
  temporaryDirectories.push(directory);
  return directory;
}
